# Python API Data Sync Tool

A Python automation pipeline for importing records from a REST API, cleaning and validating the data, preventing duplicates, storing clean records in SQLite, and exporting CSV and Excel reports.

## Business problem

Businesses often need to move API data into spreadsheets, CRMs, dashboards, or internal tools. Manual copy-paste creates inconsistent formatting, duplicate contacts, and validation errors. This project demonstrates an end-to-end sync workflow with an offline fallback.

## Features

- Fetches records from a REST API with timeout and error handling
- Falls back to a local JSON dataset when the API is unavailable
- Normalizes names, emails, companies, cities, and timestamps
- Validates required fields and email format
- Detects duplicate email addresses
- Stores clean records in SQLite
- Exports structured CSV data
- Generates a formatted Excel workbook with:
  - Dashboard
  - Clean Records
  - Invalid Records
  - Summary chart
- Includes unit tests

## Tech stack

Python, Requests, SQLite, CSV, OpenPyXL, REST API integration, data validation.

## Quick start

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

Install and run:

```bash
pip install -r requirements.txt
python api_data_sync.py
```

Generated files:

```text
output/api_records_clean.csv
output/api_records_report.xlsx
output/records.db
```

## Offline demo

The project automatically uses `sample_api_response.json` if the external API request fails. You can also call `fetch_api_data(use_sample=True)` during development or testing.

## Example adaptations

- API to Google Sheets automation
- CRM contact synchronization
- Order and product data imports
- Scheduled reporting pipelines
- Lead-data normalization
- Internal API-to-database integrations

## Portfolio positioning

This is a personal demonstration project built to showcase API integration, data cleaning, validation, SQLite persistence, and CSV/Excel reporting.
