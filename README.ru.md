<div align="center">

# 🔄 Python API Data Sync Tool

**Забирает записи из REST API, очищает и проверяет их и выгружает в SQLite, CSV и отформатированный Excel-отчёт.**

[![Tests](https://github.com/artyom129/python-api-data-sync-tool/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/python-api-data-sync-tool/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-3776AB?logo=python&logoColor=white)
![requests](https://img.shields.io/badge/requests-2.x-2A6DB0)
![SQLite](https://img.shields.io/badge/SQLite-3-003B57?logo=sqlite&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)

**🌐 Язык:** [English](README.md) · **Русский** · [Қазақша](README.kk.md)

</div>

---

## 📌 Проблема

Бизнесу постоянно нужно переносить данные из API в таблицы, CRM, дашборды и внутренние сервисы. Ручное копирование приводит к неоднородному формату, дублям контактов и ошибкам, а нестабильный API может сломать всю рутину.

## ✨ Что делает инструмент

| | Возможность |
|---|---|
| 🌐 | **Получает** записи с любого REST-адреса, который возвращает JSON-список, с таймаутом |
| 🔁 | **Повторяет запрос** при ошибках соединения, таймаутах и ответах 5xx с экспоненциальной задержкой (ошибки клиента вроде 404 не повторяются) |
| 🛟 | **Переходит на встроенные демо-данные**, если API недоступен, или падает с ошибкой при `--strict` (для заданий по расписанию) |
| 🔤 | **Нормализует** имена, email и вложенные поля (город, компания) |
| ✅ | **Проверяет** идентификатор, имя и формат email; **отклоняет дубликаты email и идентификаторов** |
| 🔍 | **Объясняет каждый отказ**: все проблемы записи перечислены в колонке `error` |
| 💾 | **Синхронизирует в SQLite** через upsert и сообщает, сколько записей новых, а сколько обновлено |
| 📊 | **Экспортирует CSV и отформатированный Excel-файл** с дашбордом и графиком |
| 💻 | **Работает из командной строки** из любой папки и возвращает корректный код выхода |
| 🧪 | **22 автотеста** (сеть подменена) запускаются в GitHub Actions при каждом пуше |

## 📑 Excel-отчёт

| Лист | Содержимое |
|---|---|
| **Dashboard** | Исходные / чистые / невалидные записи, отклонённые дубликаты, уникальные компании и города, новые и обновлённые строки в базе и график *Sync Summary* |
| **Clean Records** | Проверенные записи с фильтрами и закреплённой шапкой |
| **Invalid Records** | Отклонённые записи, причина выделена красным |

## 🚀 Быстрый старт

```bash
python -m venv .venv
```

Активируйте окружение:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Установите зависимости и запустите:

```bash
pip install -r requirements.txt
python api_data_sync.py
```

Созданные файлы (в папке `output/`):

```text
output/records.db
output/api_records_clean.csv
output/api_records_report.xlsx
```

## 🛠️ Использование

```bash
python api_data_sync.py                          # демо-API по умолчанию
python api_data_sync.py --sample                 # офлайн: sample_api_response.json
python api_data_sync.py --url https://example.com/api/users
python api_data_sync.py -o reports               # сохранить файлы в ./reports
python api_data_sync.py --strict --retries 3     # ошибка вместо демо-данных
python api_data_sync.py --help
```

| Параметр | Описание | По умолчанию |
|---|---|---|
| `--sample` | Использовать встроенные демо-данные, без сети | выкл. |
| `--url URL` | Адрес API, возвращающий JSON-список | JSONPlaceholder `/users` |
| `-o`, `--output-dir` | Папка для базы, CSV и Excel | `output/` |
| `--timeout SECONDS` | Таймаут запроса | `15` |
| `--retries N` | Число повторов после неудачного запроса | `2` |
| `--strict` | Завершиться с ошибкой вместо перехода на демо-данные | выкл. |

Пример вывода для встроенных данных:

```text
API data sync complete
Source:             sample
Raw records:        5
Clean records:      3
Invalid records:    2
New / updated:      3 / 0
```

> **Коды выхода:** `0` при успехе, `1`, если API недоступен в режиме `--strict` или нет файла с демо-данными. При втором запуске будет `0 / 3`: записи обновляются на месте и не дублируются.

## 🧪 Тесты

```bash
python -m pytest -q
```

Тесты покрывают правила очистки и проверки, дубликаты идентификаторов и email, некорректные элементы ответа API, повторы запросов и переход на демо-данные (с подменённой сетью), upsert в SQLite, значения на дашборде Excel, полный прогон и CLI.

## 📂 Структура проекта

```text
.
├── api_data_sync.py              # получение, очистка, SQLite, CSV, Excel, CLI
├── sample_api_response.json      # офлайн демо-данные
├── tests/test_api_data_sync.py   # тесты pytest
├── requirements.txt
└── .github/workflows/tests.yml   # CI: Python 3.11 и 3.12
```

## 💼 Примеры применения

- Автоматизация API → Google Sheets
- Синхронизация контактов с CRM
- Импорт заказов и товаров
- Отчёты по расписанию
- Нормализация данных о лидах
- Интеграции между внутренними API и базами данных

## 🧰 Технологии

Python · requests · SQLite · openpyxl · pytest · GitHub Actions

---

<div align="center">

*Личный демонстрационный проект: интеграция с API, очистка данных, валидация, хранение в SQLite и отчёты в CSV/Excel.*

</div>
