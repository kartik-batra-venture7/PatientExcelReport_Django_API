# PatientsExcelReport – Django API

Python 3.13.1 / Django 5.1 port of the original ASP.NET Core 8 API.

Accepts patient care Excel timesheets (`.xlsx`) and returns a structured
audit report containing:

- Caregiver / patient names and contract
- Day and date (Sun – Sat)
- Personal care code validation
- Mobile text note validation
- Patient and caregiver signature presence (`signed` / `unsigned`)
- GPS coordinates from start/end times (when present)
- Service type (`cls`, `respite`, or `cls, respite`)
- Overall validation result (`TRUE` / `FALSE`)

---

## Endpoints

| Method | URL | Description |
|--------|-----|-------------|
| POST | `/api/Patient/CheckSignatureFile` | Simple check – returns `{sheet1: {date: "signed"\|"unsigned"}}` |
| POST | `/api/PatientReaderReport/CheckSignatureFilePath` | Full audit from a server-side file path (query param `filePath`) |
| POST | `/api/PatientReaderReport/CheckSignatureFileUpload` | Full audit from uploaded `.xlsx` or base64 `.txt` |
| POST | `/api/PatientReaderReport/CheckSignatureFromBase64String` | Full audit from base64 JSON body (query param `filename`) |
| GET  | `/swagger/` | Swagger UI |

---

## Setup

### Requirements

- Python >= 3.13
- pip

### Install

```bash
git clone https://github.com/kartik-batra-venture7/PatientExcelReport_Django_API.git
cd PatientExcelReport_Django_API

python -m venv .venv

# Windows:
.venv\Scripts\activate

# Linux / macOS:
source .venv/bin/activate

pip install -r requirements.txt
```

### Configure

```bash
cp .env.example .env
# Edit .env as needed
```

### Run (development)

```bash
python manage.py migrate
python manage.py runserver
```

Swagger UI: http://localhost:8000/swagger/


## Project structure

```
PatientExcelReport_Django_API/
├── config/
│   ├── settings.py          
│   ├── urls.py              
│   └── wsgi.py
├── patients_excel_report/
│   ├── api/
│   │   ├── patient_views.py           
│   │   ├── patient_reader_report_views.py 
│   │   └── urls.py
│   └── services/
│       ├── care_validation_service.py     
│       ├── excel_reader_service.py        
│       ├── image_extractor.py             
│       ├── logging_service.py             
│       ├── signature_detection_service.py
│       └── text_similarity_service.py
├── logs/               
├── manage.py
├── requirements.txt
└── .env.example
```

---

## Dependency mapping

| .NET package | Python equivalent |
|---|---|
| ClosedXML / DocumentFormat.OpenXml | `openpyxl` |
| ExcelDataReader | `openpyxl` (streaming read) |
| SixLabors.ImageSharp | `Pillow` |
| NLog | Python `logging` + `TimedRotatingFileHandler` |
| Swashbuckle | `drf-spectacular` |
| Newtonsoft.Json | `json` |

---

## Notes

- **Upload limit:** 600 MB (set via `DATA_UPLOAD_MAX_MEMORY_SIZE` in settings).
- **Temp files** are written to `PATIENT_SIGNATURE_TEMP_DIR` (default `/tmp/patient_signature_files`) and deleted after each request.
- **SQL Server logging** via `SP_WriteApiExecutionLog` requires `pymssql` and a SQL Server database. It fails silently if unavailable (same behaviour as the .NET original).
- **OneDrive path access** (`CheckSignatureFilePath`) uses a plain `open()` check; Windows identity resolution is not available outside Windows.
- **Parallelism:** day-column processing uses `ThreadPoolExecutor` (mirrors `Parallel.ForEach`). Python's GIL limits CPU parallelism; for heavy workloads consider `ProcessPoolExecutor` or Celery tasks.
