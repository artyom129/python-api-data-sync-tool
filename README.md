<div align="center">

# 🔄 Python API Data Sync Tool

**Pull records from a REST API, clean and validate them, and ship them to SQLite, CSV, and a formatted Excel report.**

[![Tests](https://github.com/artyom129/python-api-data-sync-tool/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/python-api-data-sync-tool/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-3776AB?logo=python&logoColor=white)
![requests](https://img.shields.io/badge/requests-2.x-2A6DB0)
![SQLite](https://img.shields.io/badge/SQLite-3-003B57?logo=sqlite&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)

**🌐 Language:** **English** · [Русский](README.ru.md) · [Қазақша](README.kk.md)

</div>

---

## 📌 The problem

Businesses constantly move API data into spreadsheets, CRMs, dashboards, and internal tools. Copy-pasting it by hand creates inconsistent formatting, duplicate contacts, and validation errors, and a flaky API can break the whole routine.

## ✨ What it does

| | Feature |
|---|---|
| 🌐 | **Fetches** records from any REST endpoint that returns a JSON list, with a timeout |
| 🔁 | **Retries** connection errors, timeouts, and 5xx responses with exponential backoff (client errors such as 404 are not retried) |
| 🛟 | **Falls back to bundled sample data** when the API is down, or fails loudly with `--strict` for scheduled jobs |
| 🔤 | **Normalizes** names, emails, and nested fields (city, company) |
| ✅ | **Validates** source id, name, and email format; **rejects duplicate emails and duplicate ids** |
| 🔍 | **Explains every rejection**: all problems of a record are listed in the `error` column |
| 💾 | **Syncs into SQLite** with upserts and reports how many records are new and how many were updated |
| 📊 | **Exports CSV and a formatted Excel workbook** with a dashboard and chart |
| 💻 | **Command-line interface** that works from any directory and returns a proper exit code |
| 🧪 | **22 automated tests** (network is mocked) run on every push by GitHub Actions |

## 📑 The Excel report

| Sheet | Contents |
|---|---|
| **Dashboard** | Raw / clean / invalid records, duplicates rejected, unique companies and cities, new and updated rows in the database, and a *Sync Summary* chart |
| **Clean Records** | Validated records with filters and a frozen header |
| **Invalid Records** | Rejected records with the reason highlighted in red |

## 🚀 Quick start

```bash
python -m venv .venv
```

Activate the environment:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Install the dependencies and run:

```bash
pip install -r requirements.txt
python api_data_sync.py
```

Generated files (in `output/`):

```text
output/records.db
output/api_records_clean.csv
output/api_records_report.xlsx
```

## 🛠️ Usage

```bash
python api_data_sync.py                          # call the default demo API
python api_data_sync.py --sample                 # offline: use sample_api_response.json
python api_data_sync.py --url https://example.com/api/users
python api_data_sync.py -o reports               # write the files to ./reports
python api_data_sync.py --strict --retries 3     # fail instead of using sample data
python api_data_sync.py --help
```

| Option | Description | Default |
|---|---|---|
| `--sample` | Use the bundled sample data, no network | off |
| `--url URL` | API endpoint that returns a JSON list | JSONPlaceholder `/users` |
| `-o`, `--output-dir` | Folder for the database, CSV, and Excel files | `output/` |
| `--timeout SECONDS` | Request timeout | `15` |
| `--retries N` | Retries after a failed request | `2` |
| `--strict` | Exit with an error instead of falling back to sample data | off |

Example output for the bundled sample:

```text
API data sync complete
Source:             sample
Raw records:        5
Clean records:      3
Invalid records:    2
New / updated:      3 / 0
```

> **Exit codes:** `0` on success, `1` if the API fails in `--strict` mode or the sample file is missing. Run it twice and the second run reports `0 / 3`: records are updated in place, never duplicated.

## 🧪 Tests

```bash
python -m pytest -q
```

The suite covers cleaning and validation rules, duplicate ids and emails, malformed API items, retries and the sample-data fallback (all with a mocked network), SQLite upserts, the Excel dashboard values, an end-to-end run, and the CLI.

## 📂 Project structure

```text
.
├── api_data_sync.py              # fetch, clean, SQLite, CSV, Excel, CLI
├── sample_api_response.json      # offline demo data
├── tests/test_api_data_sync.py   # pytest suite
├── requirements.txt
└── .github/workflows/tests.yml   # CI: Python 3.11 and 3.12
```

## 💼 Example adaptations

- API to Google Sheets automation
- CRM contact synchronization
- Order and product data imports
- Scheduled reporting pipelines
- Lead-data normalization
- Internal API-to-database integrations

## 🧰 Tech stack

Python · requests · SQLite · openpyxl · pytest · GitHub Actions

---

<div align="center">

*A personal demonstration project showcasing API integration, data cleaning, validation, SQLite persistence, and CSV/Excel reporting.*

</div>
