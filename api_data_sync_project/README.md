# Python API Data Sync Tool

A Python automation project that fetches data from an API, cleans and validates records, removes duplicates, stores data in SQLite, and exports the results to CSV and Excel.

## What this project does

- Fetches records from a REST API
- Uses a local sample JSON file if the API is unavailable
- Cleans names, emails, company names, cities, and timestamps
- Validates required fields and email format
- Detects duplicate emails
- Saves clean records to a local SQLite database
- Exports clean records to CSV
- Generates an Excel report with:
  - Dashboard
  - Clean Records
  - Invalid Records

## Tech Stack

- Python
- requests
- SQLite
- CSV
- openpyxl
- REST API automation
- Data cleaning

## How to run

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Run the script:

```bash
python api_data_sync.py
```

3. The generated files will appear in the `output` folder:

```text
output/api_records_clean.csv
output/api_records_report.xlsx
output/records.db
```

## Why this project is useful

Many businesses need to move data from APIs into spreadsheets, CRMs, dashboards, or internal tools. This project shows how a Python script can automate API data import, validation, duplicate prevention, local storage, and reporting.

## Example use cases

This project can be adapted for:

- API to Google Sheets automation
- API to Excel reports
- CRM data import
- lead data sync
- order data sync
- webhook/API data processing
- daily scheduled data exports

## Portfolio description

I built a Python API automation tool that fetches data from a REST API, cleans and validates the records, removes duplicates, stores clean data in SQLite, and exports the results to CSV and Excel. This type of automation helps businesses save time on repetitive data import, cleanup, and reporting tasks.