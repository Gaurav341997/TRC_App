import os
from pathlib import Path

from dotenv import load_dotenv
from saml2 import BINDING_HTTP_POST


# =========================================================
# BASE DIRECTORY AND ENVIRONMENT
# =========================================================

BASE_DIR = Path(__file__).resolve().parent.parent

load_dotenv(BASE_DIR / ".env")


# =========================================================
# SECURITY SETTINGS
# =========================================================

SECRET_KEY = os.getenv(
    "DJANGO_SECRET_KEY",
    "django-insecure-treasury-rate-console-dev-key-change-in-production",
)

DEBUG = os.getenv("DEBUG", "True").strip().lower() in (
    "true",
    "1",
    "yes",
)

ALLOWED_HOSTS = [
    host.strip()
    for host in os.getenv(
        "ALLOWED_HOSTS",
        "localhost,127.0.0.1,[::1],trc.auuat.bank.in",
    ).split(",")
    if host.strip()
]

for default_host in ("trc.auuat.bank.in", "localhost", "127.0.0.1", "[::1]"):
    if default_host not in ALLOWED_HOSTS:
        ALLOWED_HOSTS.append(default_host)


# =========================================================
# INSTALLED APPLICATIONS
# =========================================================

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",

    # SAML application
    "djangosaml2",

    # Treasury application
    "rates",
]


# =========================================================
# MIDDLEWARE
# =========================================================

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",

    # Required by djangosaml2 for SAML session handling
    "djangosaml2.middleware.SamlSessionMiddleware",

    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",

    # Request logging middleware
    "rates.middleware.RequestLoggingMiddleware",
]


# =========================================================
# URL AND TEMPLATE CONFIGURATION
# =========================================================

ROOT_URLCONF = "treasury.urls"

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
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]


# =========================================================
# WSGI CONFIGURATION
# =========================================================

WSGI_APPLICATION = "treasury.wsgi.application"


# =========================================================
# DATABASE CONFIGURATION
# =========================================================

# PostgreSQL is used when DB_DATABASE is specified (production/UAT).
# Falls back to SQLite for local development out of the box.
if os.getenv("DB_DATABASE"):
    DATABASES = {
        "default": {
            "ENGINE": os.getenv("DB_ENGINE", "django.db.backends.postgresql"),
            "NAME": os.getenv("DB_DATABASE"),
            "USER": os.getenv("DB_USER", ""),
            "PASSWORD": os.getenv("DB_PASSWORD", ""),
            "HOST": os.getenv("DB_HOST", "localhost"),
            "PORT": os.getenv("DB_PORT", "5432"),
        }
    }
else:
    DATABASES = {
        "default": {
            "ENGINE": "django.db.backends.sqlite3",
            "NAME": BASE_DIR / "db.sqlite3",
        }
    }


# =========================================================
# PASSWORD VALIDATION
# =========================================================

AUTH_PASSWORD_VALIDATORS = []


# =========================================================
# LANGUAGE AND TIME ZONE
# =========================================================

LANGUAGE_CODE = "en-us"

TIME_ZONE = "Asia/Kolkata"

USE_I18N = True

USE_TZ = True


# =========================================================
# STATIC AND MEDIA FILES
# =========================================================

STATIC_URL = "/static/"

STATIC_ROOT = BASE_DIR / "staticfiles"

MEDIA_URL = "/media/"

MEDIA_ROOT = BASE_DIR / "media"

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"


# =========================================================
# AUTHENTICATION BACKENDS
# =========================================================

AUTHENTICATION_BACKENDS = (
    # SAML authentication backend
    "djangosaml2.backends.Saml2Backend",

    # Standard Django authentication for local/admin login
    "django.contrib.auth.backends.ModelBackend",
)


# =========================================================
# LOGIN AND LOGOUT SETTINGS
# =========================================================

# If SAML IdP metadata exists, default to SAML login; otherwise standard Django login
SAML_METADATA_FILE = BASE_DIR / "saml" / "idp_metadata.xml"

LOGIN_URL = os.getenv(
    "LOGIN_URL",
    "/saml2/login/" if SAML_METADATA_FILE.exists() else "/accounts/login/",
)

LOGIN_REDIRECT_URL = "/"

LOGOUT_REDIRECT_URL = "/"


# =========================================================
# HTTPS AND REVERSE PROXY SETTINGS
# =========================================================

SESSION_COOKIE_SECURE = os.getenv(
    "SESSION_COOKIE_SECURE",
    "False" if DEBUG else "True",
).strip().lower() in ("true", "1", "yes")

CSRF_COOKIE_SECURE = os.getenv(
    "CSRF_COOKIE_SECURE",
    "False" if DEBUG else "True",
).strip().lower() in ("true", "1", "yes")

SECURE_PROXY_SSL_HEADER = (
    "HTTP_X_FORWARDED_PROTO",
    "https",
)

USE_X_FORWARDED_HOST = True

CSRF_TRUSTED_ORIGINS = [
    origin.strip()
    for origin in os.getenv(
        "CSRF_TRUSTED_ORIGINS",
        "https://trc.auuat.bank.in,http://localhost:8000,http://127.0.0.1:8000",
    ).split(",")
    if origin.strip()
]


# =========================================================
# SAML SESSION COOKIE SETTINGS
# =========================================================

SAML_SESSION_COOKIE_NAME = "saml_session"

SAML_SESSION_COOKIE_SECURE = SESSION_COOKIE_SECURE

SAML_SESSION_COOKIE_HTTPONLY = True

# Required because IdP sends SAMLResponse using cross-site HTTP POST.
SAML_SESSION_COOKIE_SAMESITE = "None" if SESSION_COOKIE_SECURE else "Lax"

# Silence CSP warning if django-csp is not installed
SAML_CSP_HANDLER = ""


# =========================================================
# SAML USER PROVISIONING
# =========================================================

SAML_CREATE_UNKNOWN_USER = False
SAML_USE_NAME_ID_AS_USERNAME = True

SAML_ATTRIBUTE_MAPPING = {
    "email": ("email",),
    "givenName": ("first_name",),
    "sn": ("last_name",),
}


# =========================================================
# SAML 2.0 SERVICE PROVIDER CONFIGURATION
# =========================================================

SAML_CONFIG = {
    # Treasury Rate Console Service Provider entity ID
    "entityid": os.getenv(
        "SAML_ENTITY_ID",
        "https://trc.auuat.bank.in/saml2/metadata/",
    ),

    # Linux xmlsec1 executable
    "xmlsec_binary": os.getenv("XMLSEC_BINARY", "/usr/bin/xmlsec1"),

    # Metadata received from Oracle Access Manager IdP
    "metadata": {
        "local": [
            str(SAML_METADATA_FILE),
        ] if SAML_METADATA_FILE.exists() else [],
    },

    "service": {
        "sp": {
            "name": "Treasury Rate Console",

            # Mandatory for IdP-initiated SAML authentication.
            # IdP can directly POST SAMLResponse without AuthnRequest.
            "allow_unsolicited": True,

            # IdP metadata says AuthnRequestsSigned=false.
            "authn_requests_signed": False,

            # Require signed response or assertion from IdP.
            "want_response_signed": False,
            "want_assertions_signed": False,

            "endpoints": {
                "assertion_consumer_service": [
                    (
                        "https://trc.auuat.bank.in/saml2/acs/",
                        BINDING_HTTP_POST,
                    ),
                ],
            },
        },
    },
}


# =========================================================
# LOGGING CONFIGURATION
# =========================================================

LOG_DIR = BASE_DIR / "logs"

LOG_DIR.mkdir(exist_ok=True)

LOGGING = {
    "version": 1,
    "disable_existing_loggers": False,

    "formatters": {
        "standard": {
            "format": (
                "%(asctime)s | %(levelname)s | "
                "%(name)s | %(message)s"
            ),
        },
    },

    "handlers": {
        "file": {
            "level": "INFO",
            "class": "logging.handlers.RotatingFileHandler",
            "filename": LOG_DIR / "application.log",
            "maxBytes": 10485760,
            "backupCount": 5,
            "formatter": "standard",
        },

        "console": {
            "class": "logging.StreamHandler",
            "formatter": "standard",
        },
    },

    "root": {
        "handlers": [
            "file",
            "console",
        ],
        "level": "INFO",
    },
}


# =========================================================
# AWS S3 CONFIGURATION
# =========================================================

AWS_ACCESS_KEY_ID = os.getenv(
    "AWS_ACCESS_KEY_ID",
    "",
).strip()

AWS_SECRET_ACCESS_KEY = os.getenv(
    "AWS_SECRET_ACCESS_KEY",
    "",
).strip()

AWS_REGION = os.getenv(
    "AWS_REGION",
    "ap-south-1",
).strip()

S3_BUCKET_NAME = os.getenv(
    "S3_BUCKET_NAME",
    "",
).strip()

S3_PREFIX = os.getenv(
    "S3_PREFIX",
    "Bulk_FD/",
).strip().strip("/")

# S3 Interface VPC Endpoint used by the on-prem TRC server
S3_ENDPOINT_URL = os.getenv(
    "S3_ENDPOINT_URL",
    "",
).strip()
