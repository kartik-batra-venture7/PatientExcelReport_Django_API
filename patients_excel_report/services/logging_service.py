"""
Port of LoggingService.cs

Writes API execution log entries to SQL Server via the stored procedure
SP_WriteApiExecutionLog.  Falls back silently on failure (same as .NET).

Connection string is read from Django settings:
  DATABASES['default']  or  settings.LOGGING_DB_CONNECTION_STRING
"""

import logging
from datetime import datetime
from typing import Optional

from django.conf import settings

logger = logging.getLogger("patients_excel_report.logging_service")


class LoggingService:
    """Writes API execution audit records to the SQL database."""

    def __init__(self):
        self._conn_str: Optional[str] = getattr(
            settings, "LOGGING_DB_CONNECTION_STRING", None
        )
        # Fall back to the default Django DB connection string if configured
        if not self._conn_str:
            db = settings.DATABASES.get("default", {})
            self._conn_str = db.get("NAME")  # may be a file path for SQLite

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
        """
        Call SP_WriteApiExecutionLog on SQL Server.

        Returns True on success, False on any error (never raises).
        """
        if created_at is None:
            created_at = datetime.utcnow()

        try:
            import pymssql  # type: ignore

            db_cfg = settings.DATABASES.get("default", {})
            conn = pymssql.connect(
                server=db_cfg.get("HOST", "(localdb)\\MSSQLLocalDB"),
                user=db_cfg.get("USER", ""),
                password=db_cfg.get("PASSWORD", ""),
                database=db_cfg.get("NAME", "Logging"),
            )
            cursor = conn.cursor()
            cursor.callproc(
                "SP_WriteApiExecutionLog",
                (
                    api_name,
                    file_name or None,
                    request_payload or None,
                    response_payload or None,
                    status,
                    is_success,
                ),
            )
            conn.commit()
            conn.close()
            return True

        except ImportError:
            # pymssql not installed or not SQL Server – log a warning and continue
            logger.warning(
                "pymssql not available; skipping DB execution log for %s", api_name
            )
            return False
        except Exception as exc:
            logger.error(
                "DB ERROR in write_api_execution_log | ApiName=%s | Status=%s | %s",
                api_name,
                status,
                exc,
            )
            return False
