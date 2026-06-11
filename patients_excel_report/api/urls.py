"""
API URL routing – mirrors the .NET controller routes exactly.
"""

from django.urls import path

from .patient_views import CheckSignatureFileView
from .patient_reader_report_views import (
    CheckSignatureFilePathView,
    CheckSignatureFileUploadView,
    CheckSignatureFromBase64StringView,
)

urlpatterns = [
    # PatientController  (legacy simple output)
    path(
        "Patient/CheckSignatureFile",
        CheckSignatureFileView.as_view(),
        name="patient-check-signature-file",
    ),

    # PatientReaderReportController  (full audit output)
    path(
        "PatientReaderReport/CheckSignatureFilePath",
        CheckSignatureFilePathView.as_view(),
        name="reader-check-signature-file-path",
    ),
    path(
        "PatientReaderReport/CheckSignatureFileUpload",
        CheckSignatureFileUploadView.as_view(),
        name="reader-check-signature-file-upload",
    ),
    path(
        "PatientReaderReport/CheckSignatureFromBase64String",
        CheckSignatureFromBase64StringView.as_view(),
        name="reader-check-signature-from-base64",
    ),
]
