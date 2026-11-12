"""Sync records from a REST API into SQLite, CSV and a formatted Excel report.

Usage:
    python api_data_sync.py [--sample] [--url URL] [-o OUTPUT_DIR] [--strict]
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sqlite3
import sys
import time
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE_DIR = Path(__file__).resolve().parent
API_URL = "https://jsonplaceholder.typicode.com/users"
SAMPLE_FILE = BASE_DIR / "sample_api_response.json"
OUTPUT_DIR = BASE_DIR / "output"
DATABASE_FILE = OUTPUT_DIR / "records.db"
CSV_FILE = OUTPUT_DIR / "api_records_clean.csv"
EXCEL_FILE = OUTPUT_DIR / "api_records_report.xlsx"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

FIELDS = [
    "source_id", "name", "username", "email", "phone",
    "website", "city", "company", "synced_at",
]


# --------------------------------------------------------------------------
# Fetching
# --------------------------------------------------------------------------

def load_sample_data(path: Path = SAMPLE_FILE) -> list[dict[str, Any]]:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Sample file not found: {path}")

    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)

    if not isinstance(data, list):
        raise ValueError("Sample JSON must contain a list of objects")

    return data


def _request_json(api_url: str, timeout_seconds: float, retries: int,
                  backoff_seconds: float) -> list[Any]:
    """GET a JSON list. Retries connection errors, timeouts and 5xx responses."""
    last_error: Exception | None = None

    for attempt in range(retries + 1):
        try:
            response = requests.get(api_url, timeout=timeout_seconds)
            response.raise_for_status()
            data = response.json()
        except requests.HTTPError as error:
            last_error = error
            status = error.response.status_code if error.response is not None else 500
            if status < 500:  # a client error will not fix itself
                break
        except requests.RequestException as error:
            last_error = error
        else:
            if not isinstance(data, list):
                raise ValueError("API response must contain a list of objects")
            return data

        if attempt < retries:
            time.sleep(backoff_seconds * (2 ** attempt))

    raise last_error if last_error else RuntimeError("API request failed")


def fetch_records(
    use_sample: bool = False,
    api_url: str = API_URL,
    timeout_seconds: float = 15,
    retries: int = 2,
    backoff_seconds: float = 1.0,
    strict: bool = False,
    sample_file: Path = SAMPLE_FILE,
) -> tuple[list[dict[str, Any]], str]:
    """Return (records, source) where source is 'api' or 'sample'.

    Falls back to the local demo data when the API fails, unless ``strict``.
    """
    if use_sample:
        return load_sample_data(sample_file), "sample"

    try:
        return _request_json(api_url, timeout_seconds, retries, backoff_seconds), "api"
    except (requests.RequestException, ValueError) as error:
        if strict:
            raise
        print(f"API request failed; using sample data instead: {error}",
              file=sys.stderr)
        return load_sample_data(sample_file), "sample"


def fetch_api_data(use_sample: bool = False, api_url: str = API_URL,
                   timeout_seconds: float = 15) -> list[dict[str, Any]]:
    """Backward-compatible wrapper that returns only the records."""
    return fetch_records(use_sample, api_url, timeout_seconds)[0]


# --------------------------------------------------------------------------
# Cleaning
# --------------------------------------------------------------------------

def normalize_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_email(value: Any) -> str:
    return normalize_text(value).lower()


def extract_nested_text(record: dict[str, Any], field: str, nested_key: str) -> str:
    value = record.get(field, "")
    if isinstance(value, dict):
        return normalize_text(value.get(nested_key, ""))
    return normalize_text(value)


def clean_records(
    raw_records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    clean: list[dict[str, Any]] = []
    invalid: list[dict[str, Any]] = []
    seen_emails: set[str] = set()
    seen_ids: set[str] = set()
    synced_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

    for record in raw_records:
        if not isinstance(record, dict):
            invalid.append({
                **{field: "" for field in FIELDS},
                "synced_at": synced_at,
                "error": "Record is not an object",
            })
            continue

        cleaned = {
            "source_id": normalize_text(record.get("id")),
            "name": normalize_text(record.get("name")).title(),
            "username": normalize_text(record.get("username")),
            "email": normalize_email(record.get("email")),
            "phone": normalize_text(record.get("phone")),
            "website": normalize_text(record.get("website")),
            "city": extract_nested_text(record, "address", "city"),
            "company": extract_nested_text(record, "company", "name"),
            "synced_at": synced_at,
            "error": "",
        }

        errors: list[str] = []
        if not cleaned["source_id"]:
            errors.append("Missing source id")
        elif cleaned["source_id"] in seen_ids:
            errors.append("Duplicate source id")
        if not cleaned["name"]:
            errors.append("Missing name")
        if not EMAIL_RE.match(cleaned["email"]):
            errors.append("Invalid email")
        elif cleaned["email"] in seen_emails:
            errors.append("Duplicate email")

        if errors:
            cleaned["error"] = "; ".join(errors)
            invalid.append(cleaned)
            continue

        seen_ids.add(cleaned["source_id"])
        seen_emails.add(cleaned["email"])
        clean.append(cleaned)

    return clean, invalid


# --------------------------------------------------------------------------
# Storage and exports
# --------------------------------------------------------------------------

def save_to_sqlite(
    records: list[dict[str, Any]],
    database_file: Path = DATABASE_FILE,
) -> dict[str, int]:
    """Upsert records by source_id and return {'inserted': n, 'updated': n}."""
    database_file = Path(database_file)
    database_file.parent.mkdir(parents=True, exist_ok=True)

    with closing(sqlite3.connect(database_file)) as connection:
        with connection:  # one transaction: commit on success, roll back on error
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS api_records (
                    source_id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    username TEXT,
                    email TEXT UNIQUE NOT NULL,
                    phone TEXT,
                    website TEXT,
                    city TEXT,
                    company TEXT,
                    synced_at TEXT NOT NULL
                )
                """
            )

            existing = {row[0] for row in
                        connection.execute("SELECT source_id FROM api_records")}
            inserted = sum(1 for r in records if r["source_id"] not in existing)

            # If the API moved an email to a different id, free it first so the
            # UNIQUE constraint does not abort the whole sync.
            connection.executemany(
                "DELETE FROM api_records WHERE email = ? AND source_id != ?",
                [(r["email"], r["source_id"]) for r in records],
            )

            connection.executemany(
                """
                INSERT INTO api_records (
                    source_id, name, username, email, phone,
                    website, city, company, synced_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_id) DO UPDATE SET
                    name = excluded.name,
                    username = excluded.username,
                    email = excluded.email,
                    phone = excluded.phone,
                    website = excluded.website,
                    city = excluded.city,
                    company = excluded.company,
                    synced_at = excluded.synced_at
                """,
                [tuple(record[field] for field in FIELDS) for record in records],
            )

    return {"inserted": inserted, "updated": len(records) - inserted}


def export_to_csv(
    records: list[dict[str, Any]],
    output_file: Path = CSV_FILE,
) -> None:
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with output_file.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(
            {field: record.get(field, "") for field in FIELDS}
            for record in records
        )


def style_header(worksheet, row: int = 1) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in worksheet[row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border


def write_records_sheet(worksheet, records: list[dict[str, Any]], include_error: bool) -> None:
    headers = list(FIELDS) + (["error"] if include_error else [])

    worksheet.append(headers)
    for record in records:
        worksheet.append([record.get(header, "") for header in headers])

    style_header(worksheet)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    if include_error:
        error_font = Font(color="C00000")
        for row in worksheet.iter_rows(min_row=2, min_col=len(headers),
                                       max_col=len(headers)):
            row[0].font = error_font

    for column in worksheet.columns:
        column_letter = column[0].column_letter
        max_length = max(len(str(cell.value or "")) for cell in column)
        worksheet.column_dimensions[column_letter].width = min(max_length + 3, 40)


def create_dashboard_sheet(
    worksheet,
    clean_records: list[dict[str, Any]],
    invalid_records: list[dict[str, Any]],
    sync_stats: dict[str, int] | None = None,
) -> None:
    duplicates = sum(1 for r in invalid_records if "Duplicate" in r.get("error", ""))
    metrics = [
        ("Raw records", len(clean_records) + len(invalid_records)),
        ("Clean records", len(clean_records)),
        ("Invalid records", len(invalid_records)),
        ("Duplicates rejected", duplicates),
        ("Unique companies", len({r["company"] for r in clean_records if r["company"]})),
        ("Unique cities", len({r["city"] for r in clean_records if r["city"]})),
    ]
    if sync_stats:
        metrics += [
            ("New in database", sync_stats.get("inserted", 0)),
            ("Updated in database", sync_stats.get("updated", 0)),
        ]

    worksheet["A1"] = "API Data Sync Report"
    worksheet["A1"].font = Font(size=18, bold=True, color="1F4E78")
    worksheet.merge_cells("A1:B1")
    worksheet.append(["Metric", "Value"])
    for metric, value in metrics:
        worksheet.append([metric, value])

    style_header(worksheet, row=2)
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 18

    # The chart shows the first three metrics (raw / clean / invalid) only,
    # so the scale stays readable.
    chart = BarChart()
    chart.title = "Sync Summary"
    chart.legend = None
    chart.y_axis.title = "Count"
    chart.add_data(Reference(worksheet, min_col=2, min_row=3, max_row=5),
                   titles_from_data=False)
    chart.set_categories(Reference(worksheet, min_col=1, min_row=3, max_row=5))
    chart.height = 8
    chart.width = 14
    worksheet.add_chart(chart, "D2")


def export_to_excel(
    clean_records: list[dict[str, Any]],
    invalid_records: list[dict[str, Any]],
    output_file: Path = EXCEL_FILE,
    sync_stats: dict[str, int] | None = None,
) -> None:
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    dashboard = workbook.active
    dashboard.title = "Dashboard"
    clean_sheet = workbook.create_sheet("Clean Records")
    invalid_sheet = workbook.create_sheet("Invalid Records")

    create_dashboard_sheet(dashboard, clean_records, invalid_records, sync_stats)
    write_records_sheet(clean_sheet, clean_records, include_error=False)
    write_records_sheet(invalid_sheet, invalid_records, include_error=True)
    workbook.save(output_file)


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def run(
    use_sample: bool = False,
    api_url: str = API_URL,
    output_dir: Path = OUTPUT_DIR,
    timeout_seconds: float = 15,
    retries: int = 2,
    strict: bool = False,
) -> dict[str, Any]:
    output_dir = Path(output_dir)
    raw_records, source = fetch_records(
        use_sample, api_url, timeout_seconds, retries, strict=strict)
    clean, invalid = clean_records(raw_records)

    database = output_dir / DATABASE_FILE.name
    csv_file = output_dir / CSV_FILE.name
    excel_file = output_dir / EXCEL_FILE.name

    stats = save_to_sqlite(clean, database)
    export_to_csv(clean, csv_file)
    export_to_excel(clean, invalid, excel_file, stats)

    return {
        "source": source,
        "raw": len(raw_records),
        "clean": len(clean),
        "invalid": len(invalid),
        "inserted": stats["inserted"],
        "updated": stats["updated"],
        "database": database,
        "csv": csv_file,
        "excel": excel_file,
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync API records into SQLite, CSV and an Excel report.")
    parser.add_argument("--sample", action="store_true",
                        help="use sample_api_response.json instead of calling the API")
    parser.add_argument("--url", default=API_URL,
                        help=f"API endpoint returning a JSON list (default: {API_URL})")
    parser.add_argument("-o", "--output-dir", type=Path, default=OUTPUT_DIR,
                        help="directory for the database, CSV and Excel files (default: output/)")
    parser.add_argument("--timeout", type=float, default=15,
                        help="request timeout in seconds (default: 15)")
    parser.add_argument("--retries", type=int, default=2,
                        help="retries after a failed request (default: 2)")
    parser.add_argument("--strict", action="store_true",
                        help="exit with an error instead of falling back to sample data")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        result = run(args.sample, args.url, args.output_dir,
                     args.timeout, args.retries, args.strict)
    except (requests.RequestException, FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print("API data sync complete")
    print(f"Source:             {result['source']}")
    print(f"Raw records:        {result['raw']}")
    print(f"Clean records:      {result['clean']}")
    print(f"Invalid records:    {result['invalid']}")
    print(f"New / updated:      {result['inserted']} / {result['updated']}")
    print(f"SQLite database:    {result['database']}")
    print(f"CSV export:         {result['csv']}")
    print(f"Excel report:       {result['excel']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
