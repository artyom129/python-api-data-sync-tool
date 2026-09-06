# Python API Data Sync Tool

[English](README.md) | **Русский**

Python-пайплайн для загрузки данных из REST API, очистки и проверки записей, удаления дублей, сохранения результата в SQLite и формирования CSV/Excel-отчётов.

## Что решает проект

Компании часто переносят данные из API в таблицы, CRM, внутренние панели и базы. Ручной copy-paste приводит к дублям, разным форматам и ошибкам. Этот проект показывает полный цикл синхронизации данных с резервным офлайн-режимом.

## Возможности

- запрос данных из REST API с timeout и обработкой ошибок;
- автоматический fallback на локальный JSON, если API недоступен;
- нормализация имён, email, компаний, городов и временных меток;
- проверка обязательных полей и email;
- поиск дублей по email;
- сохранение очищенных записей в SQLite;
- экспорт в CSV;
- Excel-отчёт с Dashboard, Clean Records, Invalid Records и диаграммой;
- unit-тесты.

## Стек

Python, Requests, SQLite, CSV, OpenPyXL, REST API integration, data validation.

## Запуск

```bash
python -m venv .venv
```

Windows:

```bat
.venv\Scripts\activate
pip install -r requirements.txt
python api_data_sync.py
```

macOS / Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
python api_data_sync.py
```

Результаты:

```text
output/api_records_clean.csv
output/api_records_report.xlsx
output/records.db
```

## Примеры применения

- API → Google Sheets;
- синхронизация контактов CRM;
- импорт заказов и товаров;
- автоматическая отчётность;
- нормализация лидов;
- интеграции API → база данных.

Проект демонстрирует API-интеграции, очистку данных, валидацию, SQLite и отчётность.
