# PatientsExcelReport API

A Django REST API that processes patient/caregiver care-timesheet Excel (`.xlsx`)
files, validates personal-care codes and mobile notes, and detects whether
patient/caregiver signature cells contain an actual signature image.

It is a Python/Django port of an original .NET service and is designed to run
as a **Windows Service** (via [Servy](https://github.com/aelassas/servy)) using
[Waitress](https://docs.pylonsproject.org/projects/waitress/) as the production
WSGI server.

---

## Features

- Upload and parse `.xlsx` timesheets (and `.txt` files containing base64-encoded `.xlsx` data)
- Detect caregiver/patient signature presence by analyzing the embedded image pixels
- Validate personal-care codes and "denied word" lists (configurable in `config/settings.py`)
- Full audit report output (`PatientReaderReport` endpoints) and a simpler legacy
  output (`Patient` endpoint)
- SQL Server (MSSQL) execution logging — every API call is recorded
  in an `ApiExecutionLog` table (auto-created on first use); **required** for this deployment
- Swagger / OpenAPI docs via `drf-spectacular`
- Rotating file logging to `logs/current.log`

---

## Tech Stack

| Component        | Version / Notes                          |
|-------------------|-------------------------------------------|
| Python            | 3.13.x                                     |
| Django            | 6.0.6                                      |
| Django REST Framework | 3.17.1                                |
| drf-spectacular   | 0.29.0 (Swagger UI)                        |
| openpyxl / Pillow | Excel parsing & image analysis             |
| pymssql           | Optional SQL Server execution logging      |
| Waitress          | Production WSGI server (used by the Windows service) |

---

## Project Structure

```
PatientExcelReport_Django_API/
├── config/                     # Django project settings, URLs, WSGI app
│   ├── settings.py
│   ├── urls.py
│   └── wsgi.py
├── patients_excel_report/
│   ├── api/
│   │   ├── patient_views.py            # Legacy /api/Patient/CheckSignatureFile
│   │   ├── patient_reader_report_views.py  # Full audit endpoints
│   │   └── urls.py
│   └── services/
│       ├── excel_reader_service.py     # Core Excel parsing/audit logic
│       ├── image_extractor.py          # Extracts images anchored to worksheet cells
│       ├── signature_detection_service.py
│       ├── care_validation_service.py  # Personal-care code / denied-word validation
│       ├── text_similarity_service.py
│       └── logging_service.py          # Optional MSSQL execution logging
├── logs/                       # Rotating log files (current.log)
├── manage.py
├── run_server_waitress.py      # Entry point used by the Windows service
└── requirements.txt
```

---

## API Endpoints

Base path: `/api/`

| Method | Endpoint                                              | Description |
|--------|--------------------------------------------------------|--------------|
| POST   | `/api/Patient/CheckSignatureFile`                       | Legacy endpoint. Upload a `.xlsx`/`.txt` file, returns a simple `{ sheetN: { date: "signed"/"unsigned" } }` map. |
| POST   | `/api/PatientReaderReport/CheckSignatureFilePath`        | Query param `filePath` — reads a file from a server-accessible path (local/OneDrive) and returns the full audit report. |
| POST   | `/api/PatientReaderReport/CheckSignatureFileUpload`      | `multipart/form-data` upload of a `.xlsx`/`.txt` file — returns the full audit report. |
| POST   | `/api/PatientReaderReport/CheckSignatureFromBase64String`| JSON body containing a base64 string + `filename` query param — returns the full audit report. |
| GET    | `/swagger/`                                             | Swagger UI |
---

## Configuration

All configuration is read from environment variables via `python-decouple`
(a `.env` file in the project root is also supported). None are strictly
required to run the service — sensible defaults are used.

| Variable                    | Default                                  | Description |
|------------------------------|-------------------------------------------|--------------|
| `SECRET_KEY`                  | `django-insecure-change-me-in-production` | Set a real secret key in production |
| `DEBUG`                       | `True`                                    | Set to `False` in production |
| `ALLOWED_HOSTS`                | `*`                                       | Comma-separated list of allowed hosts |
| `PATIENT_SIGNATURE_TEMP_DIR`   | `<project>/patient_signature_files`       | Temp folder used while processing uploads |
| `MAX_DEGREE_OF_PARALLELISM`    | `4`                                       | Parallelism for Excel processing |
| `SIMILARITY_THRESHOLD`         | `0.75`                                    | Text similarity threshold for note validation |
| `MSSQL_SERVER`                 | `localhost`                              | SQL Server host for execution logging |
| `MSSQL_PORT`                   | `1433`                                    | SQL Server port |
| `MSSQL_DATABASE`               | `PatientsExcelReportLogs`                 | Database used for the `ApiExecutionLog` table |
| `MSSQL_USER` / `MSSQL_PASSWORD`| (empty)                                   | Required only if `MSSQL_TRUSTED_CONNECTION=False` |
| `MSSQL_TRUSTED_CONNECTION`     | `True`                                    | Use Windows trusted connection instead of SQL auth |

> **SQL Server execution logging is required for this deployment.** A reachable
> MSSQL instance (matching `MSSQL_SERVER` / `MSSQL_DATABASE` / credentials)
> must be available before starting the service — every API call writes a row
> to the `ApiExecutionLog` table (auto-created on first use). Confirm
> connectivity and permissions on the target SQL Server before installing the
> Windows service.

The service listens on **port 5010** by default.

### SQL Server network configuration (TCP/IP, for use with SSMS)

`pymssql` connects over **TCP**, so the target SQL Server instance must have
the TCP/IP protocol enabled with a fixed port (this project assumes the
default `1433`, matching `MSSQL_PORT`). If you can already connect to the
instance from **SQL Server Management Studio (SSMS)** using its server name,
but the API can't reach it, this is usually the cause. Configure it as
follows:

1. Open **SQL Server Configuration Manager** on the machine hosting SQL Server.
2. Expand **SQL Server Network Configuration** → **Protocols for `<INSTANCE_NAME>`**
   (e.g. `Protocols for MSSQLSERVER`).
3. Right-click **TCP/IP** → **Enable** (if it isn't already).
4. Right-click **TCP/IP** again → **Properties**.
5. On the **IP Addresses** tab, scroll to the **`IPAll`** section at the bottom:
   - Set **TCP Dynamic Ports** to **blank/empty** (i.e. clear the `0` that SQL
     Server puts there by default — leaving it as `0` tells SQL Server to pick
     a *random* port on every restart, which is what we want to avoid).
   - Set **TCP Port** to a fixed value, e.g. **`1433`**, matching `MSSQL_PORT`
     in your `.env` file.
6. Click **OK** to save.
7. Go to **SQL Server Services** (left-hand pane), right-click your SQL
   Server instance (e.g. `SQL Server (MSSQLSERVER)`), and choose **Restart**
   for the TCP/IP changes to take effect.
8. Confirm the Windows Firewall allows inbound TCP traffic on the chosen port
   (`1433` by default) from the application server, if SQL Server is on a
   separate machine.

After the restart, verify connectivity from the application server, e.g.:

```powershell
Test-NetConnection -ComputerName localhost -Port 1433
```

A successful `TcpTestSucceeded : True` confirms the API will be able to reach
SQL Server for execution logging.

---

## Local Setup (development)

> Requires Python 3.13+ and (optionally) access to a SQL Server instance.

```powershell
# 1. Clone the repository
git clone https://github.com/kartik-batra-venture7/PatientExcelReport_Django_API.git
cd PatientExcelReport_Django_API

# 2. Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. (Optional) create a .env file to override defaults
copy NUL .env
notepad .env

# 5. Apply migrations (only needed for the default sqlite db used by Django internals)
python manage.py migrate

# 6. Run the development server
python manage.py runserver 0.0.0.0:5010
```

Swagger UI will be available at `http://localhost:5010/swagger/`.

### Run with Waitress (production-style, no Windows service yet)

This project uses `waitress-serve` (installed as part of `requirements.txt`)
directly as the WSGI entry point — no wrapper script is needed.

```powershell
.venv\Scripts\activate
waitress-serve --host=0.0.0.0 --port=5010 config.wsgi:application
```

This starts Waitress on `0.0.0.0:5010`, exactly as the Windows service will run it.

---

## Deploying as a Windows Service with Servy

[Servy](https://github.com/aelassas/servy) lets you wrap any executable as a
native Windows Service, with auto-start, auto-restart on failure, health
monitoring, and log rotation. This project ships a ready-to-import Servy
configuration: **`service_config.json`**.

### 1. Prepare the application on the target server

```powershell
cd "C:\Apps"
git clone https://github.com/kartik-batra-venture7/PatientExcelReport_Django_API.git
cd PatientExcelReport_Django_API

python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# Required: create a .env file with production settings, including the
# MSSQL_* variables — SQL Server execution logging is required for this
# deployment, so confirm the database is reachable before continuing.
notepad .env
```

Confirm the app runs correctly in the foreground first:

```powershell
.venv\Scripts\waitress-serve.exe --host=0.0.0.0 --port=5010 config.wsgi:application
```

Browse to `http://localhost:5010/swagger/` to confirm it's working, then stop
it with `Ctrl+C` before continuing.

### 2. Install / open Servy Manager

Download and install **Servy Manager** from the
[Servy releases page](https://github.com/aelassas/servy/releases), then launch
**Servy Manager** with administrator privileges (required to create Windows
services).

### 3. Create the service from `service_config.json`

The repository includes `service_config.json`, preconfigured for this app:

```json
{
  "Name": "CHS_Patient_Excel_Report_API",
  "DisplayName": "CHS_Patient_Excel_Report_API",
  "Description": "Service of Patient Excel Report Django API",
  "ExecutablePath": "C:\\Apps\\PatientExcelReport_Django_API\\.venv\\Scripts\\waitress-serve.exe",
  "StartupDirectory": "C:\\Apps\\PatientExcelReport_Django_API\\",
  "Parameters": "--host=0.0.0.0 --port=5010 config.wsgi:application",
  "StartupType": 3,
  "Priority": 2,
  "RecoveryAction": 1,
  "MaxRestartAttempts": 3,
  "StartTimeout": 10,
  "StopTimeout": 5
}
```

If Servy Manager supports **Import Configuration**, use it and point to this
file directly. Otherwise, open **Add Service** and enter the equivalent
values manually:

| Field                    | Value |
|----------------------------|-------|
| **Service Name**            | `CHS_Patient_Excel_Report_API` |
| **Display Name**            | `CHS_Patient_Excel_Report_API` |
| **Description**             | Service of Patient Excel Report Django API |
| **Executable Path**         | `<project>\.venv\Scripts\waitress-serve.exe` |
| **Parameters**               | `--host=0.0.0.0 --port=5010 config.wsgi:application` |
| **Startup Directory**         | `<project>\` (the repository root, so relative paths like `config.wsgi` resolve correctly) |
| **Startup Type**             | Automatic (value `3`) |
| **Priority**                  | Normal/Above-normal as needed (config uses `2`) |
| **Recovery Action**           | Restart the service on failure (config uses `1`), up to `MaxRestartAttempts: 3` |
| **Start/Stop Timeout**        | 10s start / 5s stop |
| **Run As**                    | An account with read/write access to the project folder (including the OneDrive-synced path) and access to the configured SQL Server instance |


You can enable additional Servy options as needed for production: size- or
date-based log rotation (`EnableSizeRotation` / `EnableDateRotation`), and
health monitoring (`EnableHealthMonitoring` with a `HeartbeatInterval`) to
have Servy detect and restart a hung process automatically.

### 4. Install and start the service

From Servy Manager, click **Install Service**, then **Start**.

### 5. Verify

- Check **Services.msc** — `CHS_Patient_Excel_Report_API` should show as **Running**.
- Browse to `http://<server-host>:5010/swagger/` from another machine on the
  network (ensure the Windows firewall allows inbound traffic on port 5010).
- Tail `logs\current.log` (application log) and any Servy-managed
  stdout/stderr logs to confirm clean startup with no errors — in particular,
  confirm there are no SQL Server connection errors, since execution logging
  is required.

### 6. Updating the deployed service

```powershell
# Stop the service from Servy Manager, or:
sc stop CHS_Patient_Excel_Report_API

cd "C:\Apps\PatientExcelReport_Django_API"
git pull
.venv\Scripts\activate
pip install -r requirements.txt

# Start the service again from Servy Manager, or:
sc start CHS_Patient_Excel_Report_API
```

---

## Logs

- **Application log**: `logs\current.log` (rotates daily, 7 days retained)
- **Servy stdout/stderr logs**: wherever configured in the Servy service entry
- **SQL execution log** (optional): `ApiExecutionLog` table in the configured
  MSSQL database, capturing request/response metadata, status codes, and
  processing time for each API call

---

## Notes

- Maximum upload size is 600 MB (`DATA_UPLOAD_MAX_MEMORY_SIZE` /
  `FILE_UPLOAD_MAX_MEMORY_SIZE` in `config/settings.py`), matching the
  original .NET service's limit.
- SQL Server execution logging is required for this deployment. Ensure the
  configured MSSQL instance is reachable and the `ApiExecutionLog` table can
  be created/written to before the service is relied upon in production; if
  `pymssql` cannot connect, the failure is recorded in `logs\current.log` but
  should be investigated immediately rather than left unresolved.
