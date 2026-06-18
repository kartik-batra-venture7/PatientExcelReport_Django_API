"""
LoggingService

Stores an audit record for every API call in SQL Server.
Table: ApiExecutionLog (auto-created on first run if it does not exist).

Columns stored:
  - Id                  BIGINT IDENTITY PK
  - ApiName             NVARCHAR(200)   endpoint name
  - FileName            NVARCHAR(500)   uploaded / referenced file name
  - StatusCode          INT             HTTP status code (200, 400, 500 …)
  - IsSuccess           BIT             1 = 2xx response
  - RequestMethod       NVARCHAR(10)    GET / POST …
  - RequestUrl          NVARCHAR(1000)  full request URL
  - RequestPayload      NVARCHAR(MAX)   serialised request body / params
  - ResponsePayload     NVARCHAR(MAX)   serialised response JSON
  - ErrorMessage        NVARCHAR(MAX)   exception message if failed
  - ProcessingTimeMs    INT             wall-clock ms for the call
  - SheetsProcessed     INT             number of Excel sheets processed
  - RowsGenerated       INT             number of ReportRow dicts returned
  - CreatedAt           DATETIME2       UTC timestamp
  - ClientIp            NVARCHAR(50)    remote IP address
"""

import json
import logging
import time
from datetime import datetime
from typing import Optional
from django.conf import settings
import os
import re
from urllib.parse import urlparse, unquote


def _extract_filename(raw: str) -> str:
    """
    Extracts just the filename from any path format:
      - Local Windows path  : C:\\Users\\kartik\\OneDrive\\report.xlsx
      - Local Unix path     : /home/user/files/report.xlsx
      - OneDrive sync path  : C:\\Users\\kartik\\OneDrive - Org\\report.xlsx
      - SharePoint URL      : https://company.sharepoint.com/.../report.xlsx
      - OneDrive live URL   : https://onedrive.live.com/?file=report.xlsx
      - OneDrive API path   : /drives/b!abc/root:/Timesheets/report.xlsx
      - URL-encoded name    : Lenawee%20Patient%20Signature.xlsx
      - Plain filename      : report.xlsx
      - Empty / None        : ""
    """
    if not raw:
        return ""

    raw = raw.strip()

    # Step 1 — handle URLs (SharePoint, OneDrive live, any https://)
    if raw.lower().startswith(("http://", "https://")):
        parsed = urlparse(raw)

        # OneDrive live URL carries filename in query param ?file=...
        # e.g. https://onedrive.live.com/?id=ABC&file=report.xlsx
        if "onedrive.live.com" in parsed.netloc:
            from urllib.parse import parse_qs
            params = parse_qs(parsed.query)
            for key in ("file", "name", "FileName"):
                if key in params:
                    return unquote(params[key][0])

        # SharePoint / generic URL — take last path segment
        path = unquote(parsed.path)          # decode %20 → space
        name = os.path.basename(path)
        if name:
            return name

    # Step 2 — URL-encoded plain filename (no http prefix, but has %20 etc.)
    if "%" in raw:
        raw = unquote(raw)

    # Step 3 — Windows or Unix local path (including OneDrive sync folder)
    # os.path.basename handles both / and \ separators on all platforms
    name = os.path.basename(raw.replace("\\", "/"))
    return name if name else raw

logger = logging.getLogger("patients_excel_report.logging_service")

# DDL – runs once on first use
_CREATE_TABLE_SQL = """
IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME = 'ApiExecutionLog'
)
BEGIN
    CREATE TABLE ApiExecutionLog (
        Id               BIGINT        IDENTITY(1,1) PRIMARY KEY,
        ApiName          NVARCHAR(200) NOT NULL,
        FileName         NVARCHAR(500) NULL,
        StatusCode       INT           NOT NULL,
        IsSuccess        BIT           NOT NULL,
        RequestMethod    NVARCHAR(10)  NULL,
        RequestUrl       NVARCHAR(1000) NULL,
        RequestPayload   NVARCHAR(MAX) NULL,
        ResponsePayload  NVARCHAR(MAX) NULL,
        ErrorMessage     NVARCHAR(MAX) NULL,
        ProcessingTimeMs INT           NULL,
        SheetsProcessed  INT           NULL,
        RowsGenerated    INT           NULL,
        CreatedAt        DATETIME2     NOT NULL DEFAULT GETUTCDATE(),
        ClientIp         NVARCHAR(50)  NULL
    );
END
"""

_INSERT_SQL = """
INSERT INTO ApiExecutionLog (
    ApiName, FileName, StatusCode, IsSuccess,
    RequestMethod, RequestUrl, RequestPayload, ResponsePayload,
    ErrorMessage, ProcessingTimeMs, SheetsProcessed, RowsGenerated,
    CreatedAt, ClientIp
) VALUES (
    %s, %s, %s, %s,
    %s, %s, CAST(%s AS NVARCHAR(MAX)), CAST(%s AS NVARCHAR(MAX)),
    CAST(%s AS NVARCHAR(MAX)), %s, %s, %s,
    %s, %s
)
"""


def _get_connection():
    """
    Open and return a pymssql connection using MSSQL_CONFIG from settings.
    Raises on failure — callers must handle.
    """
    import pymssql  # lazy import so missing package doesn't crash startup

    cfg = getattr(settings, "MSSQL_CONFIG", {})
    trusted = cfg.get("TRUSTED_CONNECTION", True)

    if trusted:
        conn = pymssql.connect(
            server=cfg.get("SERVER", "localhost"),
            port=cfg.get("PORT", "1433"),
            database=cfg.get("DATABASE", "PatientsExcelReportLogs"),
            as_dict=False,
        )
    else:
        conn = pymssql.connect(
            server=cfg.get("SERVER", "localhost"),
            port=cfg.get("PORT", "1433"),
            database=cfg.get("DATABASE", "PatientsExcelReportLogs"),
            user=cfg.get("USER", ""),
            password=cfg.get("PASSWORD", ""),
            as_dict=False,
        )
    return conn


def _ensure_table(conn) -> None:
    """Create ApiExecutionLog table if it does not already exist."""
    cursor = conn.cursor()
    cursor.execute(_CREATE_TABLE_SQL)
    conn.commit()
    cursor.close()


class LoggingService:
    """Writes API execution audit records to SQL Server."""

    _table_ensured: bool = False  # class-level flag; checked once per process

    def log(
        self,
        *,
        api_name: str,
        file_name: str = "",
        status_code: int,
        is_success: bool,
        request_method: str = "",
        request_url: str = "",
        request_payload: str = "",
        response_payload: str = "",
        error_message: str = "",
        processing_time_ms: int = 0,
        sheets_processed: int = 0,
        rows_generated: int = 0,
        client_ip: str = "",
        created_at: Optional[datetime] = None,
    ) -> bool:
        """
        Insert one audit row.  Never raises — returns False on any error.

        Usage:
            svc = LoggingService()
            svc.log(
                api_name="CheckSignatureFileUpload",
                file_name="march_timesheet.xlsx",
                status_code=200,
                is_success=True,
                request_method="POST",
                request_url="/api/PatientReaderReport/CheckSignatureFileUpload",
                response_payload=json.dumps(rows),
                processing_time_ms=412,
                sheets_processed=3,
                rows_generated=21,
                client_ip="192.168.1.10",
            )
        """
        if created_at is None:
            created_at = datetime.utcnow()

        # Truncate large payloads to avoid NVARCHAR(MAX) edge-cases with pymssql
        def _trunc(s: str, limit: int = 50_000) -> str:
            return s[:limit] if s and len(s) > limit else s

        try:
            conn = _get_connection()

            # Create table on first successful connection (once per process)
            if not LoggingService._table_ensured:
                _ensure_table(conn)
                LoggingService._table_ensured = True

            cursor = conn.cursor()
            cursor.execute(
                _INSERT_SQL,
                (
                    api_name,
                    _trunc(_extract_filename(file_name), 500),
                    status_code,
                    1 if is_success else 0,
                    request_method,
                    _trunc(request_url, 1000),
                    request_payload,
                    response_payload,
                    _trunc(error_message),
                    processing_time_ms,
                    sheets_processed,
                    rows_generated,
                    created_at,
                    client_ip,
                ),
            )
            conn.commit()
            cursor.close()
            conn.close()
            logger.debug(
                "Logged API call: %s | status=%s | rows=%s | ms=%s",
                api_name, status_code, rows_generated, processing_time_ms,
            )
            return True

        except ImportError:
            logger.warning("pymssql not installed — SQL logging disabled.")
            return False
        except Exception as exc:
            logger.error("SQL logging failed for %s: %s", api_name, exc)
            return False

    # ------------------------------------------------------------------
    # Convenience: backwards-compatible method used by older view code
    # ------------------------------------------------------------------
    def write_api_execution_log(
        self,
        api_name: str,
        file_name: str,
        request_payload: str,
        response_payload: str,
        status: int,
        is_success: bool,
        created_at: Optional[datetime] = None,
    ) -> bool:
        return self.log(
            api_name=api_name,
            file_name=_extract_filename(file_name),
            status_code=status,
            is_success=is_success,
            request_payload=request_payload,
            response_payload=response_payload,
            created_at=created_at,
        )