"""
Django settings for PatientsExcelReport API.
Compatible with Python 3.13.1 / Django 5.1.x
"""

import os
from pathlib import Path
from decouple import config
import re

BASE_DIR = Path(__file__).resolve().parent.parent

# ------------------------------------------------------------------
# Security
# ------------------------------------------------------------------
SECRET_KEY = config("SECRET_KEY", default="django-insecure-change-me-in-production")
DEBUG = config("DEBUG", default=True, cast=bool)
ALLOWED_HOSTS = config("ALLOWED_HOSTS", default="*").split(",")

# ------------------------------------------------------------------
# Application definition
# ------------------------------------------------------------------
INSTALLED_APPS = [
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.staticfiles",
    "rest_framework",
    "drf_spectacular",
    "patients_excel_report",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.middleware.common.CommonMiddleware",
]

ROOT_URLCONF = "config.urls"
WSGI_APPLICATION = "config.wsgi.application"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.template.context_processors.csrf",
                "django.template.context_processors.tz",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

STATIC_URL = "/static/"
STATIC_ROOT = BASE_DIR / "staticfiles"

# ------------------------------------------------------------------
# Database  (used only for API execution logging)
# ------------------------------------------------------------------
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}

# ---------------------------------------------------------------------------
# Denied-word lists
# ---------------------------------------------------------------------------
_DENIED_CLS_WORDS = [
    "hospital", "medication", "medicine", "sleep", "tv", "television",
    "doctor", "appointment", "hit", "attack", "police", "assault",
    "knife", "weapon", "nap", "asleep",
]

_DENIED_RESPITE_WORDS = [
    "Hospital", "Appointment", "Hit", "Attack", "Police",
    "Assault", "Knife", "Weapon", "doctor",
]

_DENIED_CLS_PATTERNS = [
    re.compile(rf"\b{re.escape(w)}\b", re.IGNORECASE) for w in _DENIED_CLS_WORDS
]
_DENIED_RESPITE_PATTERNS = [
    re.compile(rf"\b{re.escape(w)}\b", re.IGNORECASE) for w in _DENIED_RESPITE_WORDS
]

# Valid personal-care codes
_VALID_CODES: set[str] = {"01", "02", "03", "04", "05", "P", "p", "1", "2", "3", "4", "5"}
_VALID_RESPITE_PC_CODES: set[str] = {"01", "02", "03", "04", "05"}

# To use SQL Server set these in .env and switch ENGINE to "mssql" (requires django-mssql-backend or pymssql):
#
#   DB_ENGINE=mssql
#   DB_NAME=Logging
#   DB_HOST=(localdb)\MSSQLLocalDB
#   DB_USER=
#   DB_PASSWORD=
#   DB_TRUSTED_CONNECTION=yes

# ------------------------------------------------------------------
# Internationalization
# ------------------------------------------------------------------
LANGUAGE_CODE = "en-us"
TIME_ZONE = "UTC"
USE_I18N = False
USE_TZ = True

# ------------------------------------------------------------------
# REST Framework
# ------------------------------------------------------------------
REST_FRAMEWORK = {
    "DEFAULT_SCHEMA_CLASS": "drf_spectacular.openapi.AutoSchema",
    "DEFAULT_PARSER_CLASSES": [
        "rest_framework.parsers.MultiPartParser",
        "rest_framework.parsers.JSONParser",
    ],
}

# ------------------------------------------------------------------
# drf-spectacular (Swagger)
# ------------------------------------------------------------------
SPECTACULAR_SETTINGS = {
    "TITLE": "PatientsExcelReport API",
    "DESCRIPTION": (
        "Upload patient care Excel timesheets (.xlsx) and receive a structured "
        "audit report: caregiver/patient names, day/date, personal care code "
        "validation, mobile note validation, and signature presence."
    ),
    "VERSION": "1.0.0",
    "SERVE_INCLUDE_SCHEMA": False,
}

# ------------------------------------------------------------------
# File upload limits  (600 MB – matches .NET original)
# ------------------------------------------------------------------
DATA_UPLOAD_MAX_MEMORY_SIZE = 600 * 1024 * 1024   # 600 MB
FILE_UPLOAD_MAX_MEMORY_SIZE = 600 * 1024 * 1024   # 600 MB

# ------------------------------------------------------------------
# Temp folder for processing  (matches original C:\PatientSignatureFiles)
# ------------------------------------------------------------------
PATIENT_SIGNATURE_TEMP_DIR = config(
    "PATIENT_SIGNATURE_TEMP_DIR",
    default=str(BASE_DIR / "patient_signature_files"),
)

# ------------------------------------------------------------------
# Excel processing settings
# ------------------------------------------------------------------
EXCEL_PROCESSING = {
    "MAX_DEGREE_OF_PARALLELISM": config("MAX_DEGREE_OF_PARALLELISM", default=4, cast=int),
    "SIMILARITY_THRESHOLD": config("SIMILARITY_THRESHOLD", default=0.75, cast=float),
}

# ------------------------------------------------------------------
# Logging
# ------------------------------------------------------------------
LOGS_DIR = BASE_DIR / "logs"
LOGS_DIR.mkdir(exist_ok=True)

LOGGING = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "verbose": {
            "format": "{asctime} - {levelname} - {name} - {message} {exc_info}",
            "style": "{",
            "datefmt": "%Y-%m-%dT%H:%M:%SZ",
        },
    },
    "handlers": {
        "console": {
            "class": "logging.StreamHandler",
            "formatter": "verbose",
        },
        "file": {
            "class": "logging.handlers.TimedRotatingFileHandler",
            "filename": str(LOGS_DIR / "current.log"),
            "when": "midnight",
            "backupCount": 7,
            "encoding": "utf-8",
            "formatter": "verbose",
        },
    },
    "root": {
        "handlers": ["console", "file"],
        "level": "DEBUG" if DEBUG else "INFO",
    },
    "loggers": {
        "django": {
            "handlers": ["console", "file"],
            "level": "WARNING",
            "propagate": False,
        },
        "patients_excel_report": {
            "handlers": ["console", "file"],
            "level": "DEBUG",
            "propagate": False,
        },
    },
}

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"
