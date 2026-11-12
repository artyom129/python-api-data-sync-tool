<div align="center">

# 🔄 Python API Data Sync Tool

**REST API-дан жазбаларды алады, тазалап тексереді және SQLite, CSV мен пішімделген Excel есебіне шығарады.**

[![Tests](https://github.com/artyom129/python-api-data-sync-tool/actions/workflows/tests.yml/badge.svg)](https://github.com/artyom129/python-api-data-sync-tool/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/python-3.11%20%7C%203.12-3776AB?logo=python&logoColor=white)
![requests](https://img.shields.io/badge/requests-2.x-2A6DB0)
![SQLite](https://img.shields.io/badge/SQLite-3-003B57?logo=sqlite&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-217346?logo=microsoftexcel&logoColor=white)

**🌐 Тіл:** [English](README.md) · [Русский](README.ru.md) · **Қазақша**

</div>

---

## 📌 Мәселе

Бизнес API деректерін кестелерге, CRM-ге, дашбордтарға және ішкі жүйелерге үнемі көшіреді. Қолмен көшіру біркелкі емес пішімге, қайталанған контактілерге және қателерге әкеледі, ал тұрақсыз API бүкіл үдерісті бұзуы мүмкін.

## ✨ Құрал не істейді

| | Мүмкіндік |
|---|---|
| 🌐 | JSON тізімін қайтаратын кез келген REST мекенжайдан жазбаларды күту уақытымен **алады** |
| 🔁 | Байланыс қателерінде, күту уақыты асқанда және 5xx жауаптарында экспоненциалды үзіліспен **қайталайды** (404 сияқты клиент қателері қайталанбайды) |
| 🛟 | API қолжетімсіз болса, **кірістірілген демо-деректерге ауысады**, ал `--strict` режимінде (кесте бойынша жұмыстар үшін) қатемен тоқтайды |
| 🔤 | Аттарды, email-дерді және ішкі өрістерді (қала, компания) **қалыпқа келтіреді** |
| ✅ | Идентификаторды, атты және email пішімін **тексереді**; **қайталанған email мен идентификаторларды қабылдамайды** |
| 🔍 | **Әр бас тартуды түсіндіреді**: жазбаның барлық мәселесі `error` бағанында көрсетіледі |
| 💾 | upsert арқылы **SQLite-қа синхрондайды** және қанша жазба жаңа, қанша жазба жаңартылғанын хабарлайды |
| 📊 | Дашборды мен диаграммасы бар **CSV және пішімделген Excel файлын экспорттайды** |
| 💻 | Кез келген қалтадан **командалық жолмен жұмыс істейді** және дұрыс шығу кодын қайтарады |
| 🧪 | **22 автотест** (желі алмастырылған) әр push кезінде GitHub Actions арқылы іске қосылады |

## 📑 Excel есебі

| Парақ | Мазмұны |
|---|---|
| **Dashboard** | Бастапқы / таза / жарамсыз жазбалар, қабылданбаған қайталанулар, бірегей компаниялар мен қалалар, базадағы жаңа және жаңартылған жолдар және *Sync Summary* диаграммасы |
| **Clean Records** | Сүзгілері мен бекітілген тақырыбы бар тексерілген жазбалар |
| **Invalid Records** | Қабылданбаған жазбалар, себебі қызыл түспен белгіленеді |

## 🚀 Жылдам бастау

```bash
python -m venv .venv
```

Ортаны іске қосыңыз:

```bash
# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

Тәуелділіктерді орнатып, іске қосыңыз:

```bash
pip install -r requirements.txt
python api_data_sync.py
```

Жасалатын файлдар (`output/` қалтасында):

```text
output/records.db
output/api_records_clean.csv
output/api_records_report.xlsx
```

## 🛠️ Қолдану

```bash
python api_data_sync.py                          # әдепкі демо-API
python api_data_sync.py --sample                 # офлайн: sample_api_response.json
python api_data_sync.py --url https://example.com/api/users
python api_data_sync.py -o reports               # файлдарды ./reports қалтасына сақтау
python api_data_sync.py --strict --retries 3     # демо-деректердің орнына қате
python api_data_sync.py --help
```

| Параметр | Сипаттама | Әдепкі мәні |
|---|---|---|
| `--sample` | Желісіз, кірістірілген демо-деректерді қолдану | өшірулі |
| `--url URL` | JSON тізімін қайтаратын API мекенжайы | JSONPlaceholder `/users` |
| `-o`, `--output-dir` | База, CSV және Excel үшін қалта | `output/` |
| `--timeout SECONDS` | Сұраудың күту уақыты | `15` |
| `--retries N` | Сәтсіз сұраудан кейінгі қайталау саны | `2` |
| `--strict` | Демо-деректерге ауыспай, қатемен аяқтау | өшірулі |

Кірістірілген деректер үшін шығыс мысалы:

```text
API data sync complete
Source:             sample
Raw records:        5
Clean records:      3
Invalid records:    2
New / updated:      3 / 0
```

> **Шығу кодтары:** сәтті болса `0`, `--strict` режимінде API қолжетімсіз болса немесе демо-деректер файлы табылмаса `1`. Екінші іске қосуда `0 / 3` болады: жазбалар орнында жаңартылады, қайталанбайды.

## 🧪 Тесттер

```bash
python -m pytest -q
```

Тесттер тазалау мен тексеру ережелерін, қайталанған идентификаторлар мен email-дерді, API жауабындағы дұрыс емес элементтерді, сұрауларды қайталауды және демо-деректерге ауысуды (желі алмастырылған), SQLite-тағы upsert-ті, Excel дашбордындағы мәндерді, толық жұмысты және CLI-ды қамтиды.

## 📂 Жоба құрылымы

```text
.
├── api_data_sync.py              # алу, тазалау, SQLite, CSV, Excel, CLI
├── sample_api_response.json      # офлайн демо-деректер
├── tests/test_api_data_sync.py   # pytest тесттері
├── requirements.txt
└── .github/workflows/tests.yml   # CI: Python 3.11 және 3.12
```

## 💼 Қолдану мысалдары

- API → Google Sheets автоматтандыруы
- CRM контактілерін синхрондау
- Тапсырыстар мен тауарларды импорттау
- Кесте бойынша есептер
- Лид деректерін қалыпқа келтіру
- Ішкі API мен дерекқорлар арасындағы интеграциялар

## 🧰 Технологиялар

Python · requests · SQLite · openpyxl · pytest · GitHub Actions

---

<div align="center">

*Жеке демонстрациялық жоба: API интеграциясы, деректерді тазалау, тексеру, SQLite-та сақтау және CSV/Excel есептері.*

</div>
