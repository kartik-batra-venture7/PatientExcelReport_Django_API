"""
Port of PatientReaderReportController.cs

Endpoints (all return the full ReportRow audit list):

  POST /api/PatientReaderReport/CheckSignatureFilePath
    - Query param: filePath  (server-side local/OneDrive path)
    - Reads the file from disk, processes it, returns JSON list

  POST /api/PatientReaderReport/CheckSignatureFileUpload
    - multipart/form-data: file  (.xlsx or .txt with base64 content)
    - Processes the uploaded file, returns JSON list

  POST /api/PatientReaderReport/CheckSignatureFromBase64String
    - JSON body: base64 string
    - Query param: filename
    - Processes and returns JSON list
"""

import base64
import gc
import json
import logging
import os
import shutil
import time
import uuid

from django.conf import settings
from rest_framework import status
from rest_framework.parsers import MultiPartParser, JSONParser
from rest_framework.request import Request
from rest_framework.response import Response
from rest_framework.views import APIView
from drf_spectacular.utils import extend_schema, OpenApiParameter
from drf_spectacular.types import OpenApiTypes

from patients_excel_report.services.excel_reader_service import ExcelReaderService
from patients_excel_report.services.logging_service import LoggingService

logger = logging.getLogger("patients_excel_report.patient_reader_report_view")

_reader = ExcelReaderService()
_logging_service = LoggingService()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_temp_folder() -> str:
    folder = getattr(settings, "PATIENT_SIGNATURE_TEMP_DIR", "/tmp/PatientSignatureFiles")
    os.makedirs(folder, exist_ok=True)
    return folder


def _save_temp_file(data: bytes) -> str:
    folder = _get_temp_folder()
    path = os.path.join(folder, f"ps_{uuid.uuid4().hex}.xlsx")
    with open(path, "wb") as f:
        f.write(data)
    return path


def _save_temp_file_from_path(src_path: str) -> str:
    folder = _get_temp_folder()
    path = os.path.join(folder, f"ps_{uuid.uuid4().hex}.xlsx")
    shutil.copy2(src_path, path)
    return path


def _try_delete(path: str) -> None:
    for _ in range(3):
        try:
            if os.path.exists(path):
                os.remove(path)
            return
        except Exception:
            time.sleep(0.2)


def _strip_data_prefix(s: str) -> str:
    if s.lower().startswith("data:"):
        idx = s.find(",")
        if idx >= 0:
            return s[idx + 1:]
    return s


def _onedrive_file_exists(path: str) -> bool:
    """
    Mirrors OneDriveFileExists(): tries to open the file for reading.
    On non-Windows platforms there is no WindowsIdentity; we just check
    file accessibility.
    """
    try:
        with open(path, "rb"):
            return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Views
# ---------------------------------------------------------------------------

class CheckSignatureFilePathView(APIView):
    """
    POST /api/PatientReaderReport/CheckSignatureFilePath?filePath=<path>

    Reads the xlsx directly from a server-accessible path (e.g. a
    OneDrive-mounted folder) and returns the full audit JSON list.
    """

    @extend_schema(
        summary="Process Excel from server file path",
        description=(
            "Provide a server-side file path (e.g. a OneDrive-mounted .xlsx). "
            "Returns a list of audit rows covering all sheets and all 7 days."
        ),
        parameters=[
            OpenApiParameter(
                name="filePath",
                type=OpenApiTypes.STR,
                location=OpenApiParameter.QUERY,
                required=True,
                description="Absolute server-side path to the .xlsx file.",
            )
        ],
        responses={200: {"type": "array", "items": {"type": "object"}}},
    )
    def post(self, request: Request) -> Response:
        start_time = time.time()
        file_path = request.query_params.get("filePath", "").strip()
        logger.info("CheckSignatureFilePath called with path: %s", file_path)

        if not file_path:
            logger.debug("File path is empty.")
            return Response("File path is empty", status=status.HTTP_400_BAD_REQUEST)

        if not _onedrive_file_exists(file_path):
            logger.error("OneDrive file not accessible. Path=%s", file_path)
            return Response(
                f"OneDrive file not accessible. Path={file_path}",
                status=status.HTTP_404_NOT_FOUND,
            )

        try:
            temp = _save_temp_file_from_path(file_path)
        except Exception as exc:
            return Response(
                f"Temp file creation failed: {exc}",
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )

        try:
            logger.debug("Temp file saved: %s", temp)
            rows = _reader.process_file(temp)
            logger.debug("Excel file processed successfully.")
            elapsed_ms = round((time.time() - start_time) * 1000)
            _logging_service.log(
                api_name="CheckSignatureFilePath",
                file_name=file_path,
                status_code=200,
                is_success=True,
                request_method=request.method,
                request_url=request.build_absolute_uri(),
                response_payload=json.dumps(rows),
                processing_time_ms=elapsed_ms,
                sheets_processed=len({r.get("sheetName") for r in rows}),
                rows_generated=len(rows),
                client_ip=request.META.get("REMOTE_ADDR", ""),
            )
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as exc:
            elapsed_ms = round((time.time() - start_time) * 1000)
            _logging_service.log(
                api_name="CheckSignatureFilePath",
                file_name=file_path,
                status_code=400,
                is_success=False,
                request_method=request.method,
                request_url=request.build_absolute_uri(),
                error_message=str(exc),
                processing_time_ms=elapsed_ms,
                client_ip=request.META.get("REMOTE_ADDR", ""),
            )
            return Response(f"Error processing file: {exc}", status=status.HTTP_400_BAD_REQUEST)
        finally:
            _try_delete(temp)
            logger.debug("Temp file deleted: %s", temp)
            gc.collect()


class CheckSignatureFileUploadView(APIView):
    """
    POST /api/PatientReaderReport/CheckSignatureFileUpload

    Upload a .xlsx or a .txt file (containing base64 of the .xlsx).
    Returns the full audit JSON list.
    """

    parser_classes = [MultiPartParser]

    @extend_schema(
        summary="Process Excel from uploaded file",
        description=(
            "Upload a .xlsx file or a .txt file containing the base64-encoded "
            "contents of a .xlsx.  Returns a list of audit rows."
        ),
        request={
            "multipart/form-data": {
                "type": "object",
                "properties": {
                    "file": {"type": "string", "format": "binary"},
                },
                "required": ["file"],
            }
        },
        responses={200: {"type": "array", "items": {"type": "object"}}},
    )
    def post(self, request: Request) -> Response:
        start_time = time.time()
        file = request.FILES.get("file")
        logger.debug("CheckSignatureFileUpload called with file: %s", file)

        if file is None:
            logger.debug("No file uploaded.")
            return Response("No file uploaded", status=status.HTTP_400_BAD_REQUEST)

        fname = file.name or ""

        try:
            raw = file.read()
        except Exception as exc:
            return Response(
                f"Error reading uploaded file: {exc}",
                status=status.HTTP_400_BAD_REQUEST,
            )

        if fname.lower().endswith(".txt"):
            txt = raw.decode("utf-8", errors="ignore").strip()
            txt = _strip_data_prefix(txt)
            try:
                xlsx_bytes = base64.b64decode(txt)
            except Exception as exc:
                logger.error("Invalid base64 in txt file: %s", exc)
                return Response("Invalid base64 in txt", status=status.HTTP_400_BAD_REQUEST)
        elif fname.lower().endswith(".xlsx"):
            xlsx_bytes = raw
        else:
            logger.debug("Unsupported file type: %s", fname)
            return Response("Unsupported file", status=status.HTTP_400_BAD_REQUEST)

        temp = _save_temp_file(xlsx_bytes)
        try:
            logger.debug("Temp file saved: %s", temp)
            rows = _reader.process_file(temp)
            logger.debug("Excel file processed successfully.")
            elapsed_ms = round((time.time() - start_time) * 1000)
            _logging_service.log(
                api_name="CheckSignatureFileUpload",
                file_name=fname,
                status_code=200,
                is_success=True,
                request_method=request.method,
                request_url=request.build_absolute_uri(),
                response_payload=json.dumps(rows),
                processing_time_ms=elapsed_ms,
                sheets_processed=len({r.get("sheetName") for r in rows}),
                rows_generated=len(rows),
                client_ip=request.META.get("REMOTE_ADDR", ""),
            )
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as exc:
            logger.error("Error processing file: %s", exc)
            elapsed_ms = round((time.time() - start_time) * 1000)
            _logging_service.log(
                api_name="CheckSignatureFileUpload",
                file_name=fname,
                status_code=400,
                is_success=False,
                request_method=request.method,
                request_url=request.build_absolute_uri(),
                error_message=str(exc),
                processing_time_ms=elapsed_ms,
                client_ip=request.META.get("REMOTE_ADDR", ""),
            )
            return Response(f"Error processing file: {exc}", status=status.HTTP_400_BAD_REQUEST)
        finally:
            _try_delete(temp)
            logger.info("Temp file deleted: %s", temp)
            gc.collect()


class CheckSignatureFromBase64StringView(APIView):
    """
    POST /api/PatientReaderReport/CheckSignatureFromBase64String?filename=<name>

    Body: raw base64 string (JSON string)
    Returns the full audit JSON list.
    """

    parser_classes = [JSONParser]

    @extend_schema(
        summary="Process Excel from base64 string",
        description=(
            "POST a JSON string containing the base64-encoded .xlsx file. "
            "Provide the original filename via the `filename` query parameter."
        ),
        parameters=[
            OpenApiParameter(
                name="filename",
                type=OpenApiTypes.STR,
                location=OpenApiParameter.QUERY,
                required=False,
                description="Original filename for logging purposes.",
            )
        ],
        request={"application/json": {"type": "string"}},
        responses={200: {"type": "array", "items": {"type": "object"}}},
    )
    def post(self, request: Request) -> Response:
        start_time = time.time()
        filename = request.query_params.get("filename", "unknown.xlsx")
        logger.debug("CheckSignatureFromBase64String called with filename: %s", filename)

        file_b64 = request.data  # raw JSON string parsed by JSONParser
        if not file_b64:
            logger.debug("No Base64String string found.")
            return Response("No file uploaded", status=status.HTTP_400_BAD_REQUEST)

        # request.data may be the string itself or wrapped
        if isinstance(file_b64, dict):
            file_b64 = str(file_b64)

        try:
            xlsx_bytes = base64.b64decode(str(file_b64))
        except Exception as exc:
            logger.error("Invalid base64 string: %s", exc)
            return Response("Invalid base64 string", status=status.HTTP_400_BAD_REQUEST)

        temp = _save_temp_file(xlsx_bytes)
        try:
            logger.debug("Temp file saved: %s", temp)
            rows = _reader.process_file(temp)
            logger.debug("Excel file processed successfully.")
            elapsed_ms = round((time.time() - start_time) * 1000)
            _logging_service.log(
                api_name="CheckSignatureFromBase64String",
                file_name=filename,
                status_code=200,
                is_success=True,
                request_method=request.method,
                request_url=request.build_absolute_uri(),
                response_payload=json.dumps(rows),
                processing_time_ms=elapsed_ms,
                sheets_processed=len({r.get("sheetName") for r in rows}),
                rows_generated=len(rows),
                client_ip=request.META.get("REMOTE_ADDR", ""),
            )
            return Response(rows, status=status.HTTP_200_OK)
        except Exception as exc:
            logger.error("Error processing file: %s", exc)
            elapsed_ms = round((time.time() - start_time) * 1000)
            _logging_service.log(
                api_name="CheckSignatureFromBase64String",
                file_name=filename,
                status_code=400,
                is_success=False,
                request_method=request.method,
                request_url=request.build_absolute_uri(),
                error_message=str(exc),
                processing_time_ms=elapsed_ms,
                client_ip=request.META.get("REMOTE_ADDR", ""),
            )
            return Response(f"Error processing file: {exc}", status=status.HTTP_400_BAD_REQUEST)
        finally:
            _try_delete(temp)
            logger.info("Temp file deleted: %s", temp)
            gc.collect()
