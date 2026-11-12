# Upwork Portfolio Entry

## Title

Python API Data Sync and Excel Reporting Tool

## Description

Built a Python automation pipeline that imports records from a REST API, normalizes and validates the data, rejects duplicate emails and ids, syncs clean records into SQLite with upserts, and exports CSV and formatted Excel reports.

The tool retries transient API failures with exponential backoff, falls back to an offline JSON dataset (or fails loudly in strict mode for scheduled jobs), runs from the command line with a proper exit code, reports how many records are new or updated on every run, and is covered by 22 automated tests with a mocked network.

## Technologies

Python, Requests, SQLite, CSV, OpenPyXL, REST APIs, data validation, pytest, GitHub Actions.

## Client value

The tool replaces repetitive API-to-spreadsheet work, improves data consistency, prevents duplicate imports, survives flaky APIs, and creates reusable reports for operations or CRM workflows. Nothing is dropped silently: every rejected record is kept with the reason it failed.
