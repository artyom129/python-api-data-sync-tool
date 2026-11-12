import csv
import sqlite3
from contextlib import closing
from pathlib import Path

import pytest
import requests
from openpyxl import load_workbook

import api_data_sync
from api_data_sync import (
    clean_records,
    export_to_csv,
    export_to_excel,
    fetch_records,
    load_sample_data,
    main,
    run,
    save_to_sqlite,
)


def sample_records():
    return [
        {
            "id": 1,
            "name": " john smith ",
            "username": "john",
            "email": "JOHN@example.com",
            "phone": "123",
            "website": "example.com",
            "address": {"city": "Austin"},
            "company": {"name": "Example LLC"},
        },
        {
            "id": 2,
            "name": "Duplicate Contact",
            "username": "duplicate",
            "email": "john@example.com",
            "phone": "456",
            "website": "example.net",
            "address": {"city": "Dallas"},
            "company": {"name": "Another LLC"},
        },
        {
            "id": 3,
            "name": "Broken Email",
            "username": "broken",
            "email": "not-an-email",
            "phone": "789",
            "website": "example.org",
            "address": {"city": "Houston"},
            "company": {"name": "Broken LLC"},
        },
    ]


def make_record(source_id, email, name="Test User"):
    return {"id": source_id, "name": name, "email": email}


class FakeResponse:
    def __init__(self, payload=None, status=200):
        self._payload = payload
        self.status_code = status

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f"{self.status_code} error", response=self)

    def json(self):
        return self._payload


@pytest.fixture(autouse=True)
def no_sleep(monkeypatch):
    """Retries must never slow the test suite down."""
    monkeypatch.setattr(api_data_sync.time, "sleep", lambda _seconds: None)


# --------------------------------------------------------------------------
# Cleaning
# --------------------------------------------------------------------------

def test_clean_records_normalizes_and_validates():
    clean, invalid = clean_records(sample_records())

    assert len(clean) == 1
    assert len(invalid) == 2
    assert clean[0]["name"] == "John Smith"
    assert clean[0]["email"] == "john@example.com"
    assert clean[0]["city"] == "Austin"
    assert any("Duplicate email" in record["error"] for record in invalid)
    assert any("Invalid email" in record["error"] for record in invalid)


def test_duplicate_source_id_is_rejected():
    clean, invalid = clean_records([
        make_record(7, "a@example.com"),
        make_record(7, "b@example.com"),
    ])

    assert len(clean) == 1
    assert invalid[0]["error"] == "Duplicate source id"


def test_all_problems_of_a_record_are_listed():
    _, invalid = clean_records([{"id": "", "name": "", "email": "bad"}])

    assert invalid[0]["error"] == "Missing source id; Missing name; Invalid email"


def test_non_object_items_are_reported_not_crashing():
    clean, invalid = clean_records(["oops", None, make_record(1, "ok@example.com")])

    assert len(clean) == 1
    assert [r["error"] for r in invalid] == ["Record is not an object"] * 2


def test_missing_nested_fields_become_empty_strings():
    clean, _ = clean_records([make_record(1, "ok@example.com")])

    assert clean[0]["city"] == ""
    assert clean[0]["company"] == ""


# --------------------------------------------------------------------------
# Fetching
# --------------------------------------------------------------------------

def test_sample_file_is_loadable_and_has_expected_split():
    clean, invalid = clean_records(load_sample_data())

    assert (len(clean), len(invalid)) == (3, 2)


def test_use_sample_never_calls_the_api(monkeypatch):
    def boom(*_args, **_kwargs):
        raise AssertionError("network must not be used")

    monkeypatch.setattr(api_data_sync.requests, "get", boom)
    records, source = fetch_records(use_sample=True)

    assert source == "sample"
    assert records


def test_api_success_returns_api_records(monkeypatch):
    payload = [make_record(1, "a@example.com")]
    monkeypatch.setattr(api_data_sync.requests, "get",
                        lambda *_a, **_k: FakeResponse(payload))

    assert fetch_records() == (payload, "api")


def test_transient_failure_is_retried(monkeypatch):
    calls = []

    def flaky(*_args, **_kwargs):
        calls.append(1)
        if len(calls) < 3:
            return FakeResponse(status=503)
        return FakeResponse([make_record(1, "a@example.com")])

    monkeypatch.setattr(api_data_sync.requests, "get", flaky)
    _, source = fetch_records(retries=2)

    assert source == "api"
    assert len(calls) == 3


def test_client_error_is_not_retried(monkeypatch):
    calls = []

    def not_found(*_args, **_kwargs):
        calls.append(1)
        return FakeResponse(status=404)

    monkeypatch.setattr(api_data_sync.requests, "get", not_found)
    _, source = fetch_records(retries=3)

    assert source == "sample"
    assert len(calls) == 1


def test_failure_falls_back_to_sample_data(monkeypatch, capsys):
    def down(*_args, **_kwargs):
        raise requests.ConnectionError("offline")

    monkeypatch.setattr(api_data_sync.requests, "get", down)
    records, source = fetch_records(retries=1)

    assert source == "sample"
    assert records
    assert "using sample data" in capsys.readouterr().err


def test_strict_mode_raises_instead_of_falling_back(monkeypatch):
    def down(*_args, **_kwargs):
        raise requests.ConnectionError("offline")

    monkeypatch.setattr(api_data_sync.requests, "get", down)

    with pytest.raises(requests.ConnectionError):
        fetch_records(retries=0, strict=True)


def test_non_list_api_payload_is_rejected(monkeypatch):
    monkeypatch.setattr(api_data_sync.requests, "get",
                        lambda *_a, **_k: FakeResponse({"error": "nope"}))

    with pytest.raises(ValueError):
        fetch_records(strict=True)


def test_missing_sample_file_raises(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        load_sample_data(tmp_path / "missing.json")


# --------------------------------------------------------------------------
# Storage and exports
# --------------------------------------------------------------------------

def test_sqlite_and_csv_exports(tmp_path: Path):
    clean, _ = clean_records(sample_records())
    database_file = tmp_path / "records.db"
    csv_file = tmp_path / "records.csv"

    save_to_sqlite(clean, database_file)
    export_to_csv(clean, csv_file)

    with closing(sqlite3.connect(database_file)) as connection:
        count = connection.execute("SELECT COUNT(*) FROM api_records").fetchone()[0]
    assert count == 1

    with csv_file.open("r", encoding="utf-8-sig", newline="") as file:
        rows = list(csv.DictReader(file))
    assert rows[0]["email"] == "john@example.com"


def test_sqlite_sync_reports_inserted_and_updated(tmp_path: Path):
    database_file = tmp_path / "records.db"
    first, _ = clean_records([make_record(1, "a@example.com"),
                              make_record(2, "b@example.com")])
    assert save_to_sqlite(first, database_file) == {"inserted": 2, "updated": 0}

    second, _ = clean_records([make_record(2, "b@example.com", "Renamed User"),
                               make_record(3, "c@example.com")])
    assert save_to_sqlite(second, database_file) == {"inserted": 1, "updated": 1}

    with closing(sqlite3.connect(database_file)) as connection:
        rows = dict(connection.execute("SELECT source_id, name FROM api_records"))
    assert rows == {"1": "Test User", "2": "Renamed User", "3": "Test User"}


def test_email_moving_to_another_id_does_not_abort_the_sync(tmp_path: Path):
    database_file = tmp_path / "records.db"
    save_to_sqlite(clean_records([make_record(1, "a@example.com")])[0], database_file)

    moved, _ = clean_records([make_record(2, "a@example.com")])
    save_to_sqlite(moved, database_file)

    with closing(sqlite3.connect(database_file)) as connection:
        rows = connection.execute("SELECT source_id, email FROM api_records").fetchall()
    assert rows == [("2", "a@example.com")]


def test_excel_export_contains_expected_sheets(tmp_path: Path):
    clean, invalid = clean_records(sample_records())
    excel_file = tmp_path / "report.xlsx"

    export_to_excel(clean, invalid, excel_file)
    workbook = load_workbook(excel_file)

    assert workbook.sheetnames == ["Dashboard", "Clean Records", "Invalid Records"]
    assert workbook["Clean Records"]["D2"].value == "john@example.com"


def test_dashboard_metrics_are_correct(tmp_path: Path):
    clean, invalid = clean_records(sample_records())
    excel_file = tmp_path / "report.xlsx"

    export_to_excel(clean, invalid, excel_file, {"inserted": 1, "updated": 0})
    sheet = load_workbook(excel_file)["Dashboard"]
    metrics = {row[0].value: row[1].value
               for row in sheet.iter_rows(min_row=3, max_col=2)}

    assert metrics["Raw records"] == 3
    assert metrics["Clean records"] == 1
    assert metrics["Invalid records"] == 2
    assert metrics["Duplicates rejected"] == 1
    assert metrics["New in database"] == 1


# --------------------------------------------------------------------------
# End to end and CLI
# --------------------------------------------------------------------------

def test_run_produces_all_outputs(tmp_path: Path):
    result = run(use_sample=True, output_dir=tmp_path)

    assert (result["clean"], result["invalid"]) == (3, 2)
    assert result["inserted"] == 3
    for key in ("database", "csv", "excel"):
        assert result[key].exists()

    again = run(use_sample=True, output_dir=tmp_path)
    assert (again["inserted"], again["updated"]) == (0, 3)


def test_cli_success(tmp_path: Path, capsys):
    code = main(["--sample", "-o", str(tmp_path)])

    assert code == 0
    assert "API data sync complete" in capsys.readouterr().out
    assert (tmp_path / "api_records_report.xlsx").exists()


def test_cli_strict_failure_returns_exit_code_1(tmp_path: Path, monkeypatch, capsys):
    def down(*_args, **_kwargs):
        raise requests.ConnectionError("offline")

    monkeypatch.setattr(api_data_sync.requests, "get", down)
    code = main(["--strict", "--retries", "0", "-o", str(tmp_path)])

    assert code == 1
    assert "Error:" in capsys.readouterr().err
