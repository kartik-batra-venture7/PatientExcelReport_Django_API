"""
URL configuration for PatientsExcelReport API.

Endpoints:

  PatientController:
    POST /api/Patient/CheckSignatureFile

  PatientReaderReportController:
    POST /api/PatientReaderReport/CheckSignatureFilePath
    POST /api/PatientReaderReport/CheckSignatureFileUpload
    POST /api/PatientReaderReport/CheckSignatureFromBase64String

  Swagger UI:
    GET  /swagger/
    GET  /api/schema/
"""

from django.urls import path, include
from drf_spectacular.views import SpectacularAPIView, SpectacularSwaggerView

urlpatterns = [
    # OpenAPI schema + Swagger UI  (always enabled, same as original)
    path("api/schema/", SpectacularAPIView.as_view(), name="schema"),
    path(
        "swagger/",
        SpectacularSwaggerView.as_view(url_name="schema"),
        name="swagger-ui",
    ),
    # API endpoints
    path("api/", include("patients_excel_report.api.urls")),
]
