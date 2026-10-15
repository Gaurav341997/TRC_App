import json
import logging
import os
import subprocess
import sys
from datetime import datetime
from functools import wraps
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth import logout
from django.contrib.auth.decorators import login_required
from django.http import FileResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_GET, require_POST
from openpyxl import load_workbook

from .models import Product, Sheet, RateCardHistory
from .services import (
    sheet_to_dict,
    rebuild_sheet,
    validate_dict,
    export_csv,
    export_nominal_csv,
    stamp,
    eff_stamp,
    parse_csv_to_dict,
    upload_to_s3,
)


# ============================================================
# SECURITY CONFIGURATION
# ============================================================

TRC_APP_LABEL = "rates"

PERM_ACCESS_TRC = f"{TRC_APP_LABEL}.access_trc"
PERM_SAVE_DRAFT = f"{TRC_APP_LABEL}.save_trc_draft"
PERM_PUBLISH = f"{TRC_APP_LABEL}.publish_trc_rates"
PERM_UPLOAD = f"{TRC_APP_LABEL}.upload_trc_rates"
PERM_EDIT_PUBLISHED = f"{TRC_APP_LABEL}.edit_trc_published"

# Granular Module Permissions
PERM_ACCESS_RATE_CONSOLE = f"{TRC_APP_LABEL}.access_rate_console"
PERM_ACCESS_RETAIL_TD = f"{TRC_APP_LABEL}.access_retail_td"
PERM_ACCESS_BULK_TD = f"{TRC_APP_LABEL}.access_bulk_td"
PERM_ACCESS_NRE_NRO_TD = f"{TRC_APP_LABEL}.access_nre_nro_td"

# Maximum allowed CSV upload size = 1 MB
MAX_UPLOAD_SIZE = 1 * 1024 * 1024


# ============================================================
# RBAC DECORATOR
# ============================================================

def permission_required_json(permission):
    """
    Server-side Role-Based Access Control.

    Rules:
    1. User must be authenticated.
    2. Superuser is allowed.
    3. Staff/Admin is allowed.
    4. Normal user must have required permission.
    5. Otherwise HTTP 403 is returned.
    """

    def decorator(view_func):

        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):

            # Authentication check
            if not request.user.is_authenticated:
                return JsonResponse(
                    {
                        "ok": False,
                        "errors": ["Authentication required."],
                    },
                    status=401,
                )

            # Superuser access
            if request.user.is_superuser:
                return view_func(request, *args, **kwargs)

            # Existing staff/admin access
            if request.user.is_staff:
                return view_func(request, *args, **kwargs)

            # Explicit permission-based authorization
            if request.user.has_perm(permission):
                return view_func(request, *args, **kwargs)

            # User is authenticated but not authorized
            return JsonResponse(
                {
                    "ok": False,
                    "errors": [
                        "Forbidden: You are not authorized to perform this operation."
                    ],
                },
                status=403,
            )

        return _wrapped_view

    return decorator

def trc_page_access_required(view_func):
    """
    Protect HTML pages through SAML SSO.

    Anonymous users are redirected to the configured SAML login route.
    Authenticated users require TRC access permission.
    """

    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            login_url = settings.LOGIN_URL
            next_url = request.get_full_path()
            separator = "&" if "?" in login_url else "?"
            return redirect(f"{login_url}{separator}next={next_url}")

        if (
            request.user.is_superuser
            or request.user.is_staff
            or request.user.has_perm(PERM_ACCESS_TRC)
            or request.user.has_perm(PERM_ACCESS_RATE_CONSOLE)
            or request.user.has_perm(PERM_ACCESS_RETAIL_TD)
            or request.user.has_perm(PERM_ACCESS_BULK_TD)
            or request.user.has_perm(PERM_ACCESS_NRE_NRO_TD)
        ):
            return view_func(request, *args, **kwargs)

        return render(
            request,
            "rates/permission_denied.html",
            {
                "module_name": "Treasury Rate Modules",
                "is_admin": _is_trc_admin(request.user),
            },
            status=403,
        )

    return _wrapped_view


def user_has_module_access(user, module_key):
    """
    Check if a user has permission to access a specific module.

    Admins, superusers, and staff have access to all modules.
    For standard users, checks the specific module permission:
      - 'console': rates.access_rate_console (or legacy rates.access_trc)
      - 'retail': rates.access_retail_td (or legacy rates.access_trc)
      - 'bulk': rates.access_bulk_td (or legacy rates.access_trc)
      - 'nre_nro': rates.access_nre_nro_td (or legacy rates.access_trc)
    """
    if not user.is_authenticated:
        return False
    if user.is_superuser or user.is_staff:
        return True

    perm_map = {
        "console": PERM_ACCESS_RATE_CONSOLE,
        "retail": PERM_ACCESS_RETAIL_TD,
        "bulk": PERM_ACCESS_BULK_TD,
        "nre_nro": PERM_ACCESS_NRE_NRO_TD,
    }
    required_perm = perm_map.get(module_key)
    if not required_perm:
        return False

    return user.has_perm(required_perm) or user.has_perm(PERM_ACCESS_TRC)


def module_access_required(module_key, module_label):
    """
    Protect HTML module pages with granular per-module permissions.
    """
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                login_url = settings.LOGIN_URL
                next_url = request.get_full_path()
                separator = "&" if "?" in login_url else "?"
                return redirect(f"{login_url}{separator}next={next_url}")

            if user_has_module_access(request.user, module_key):
                return view_func(request, *args, **kwargs)

            return render(
                request,
                "rates/permission_denied.html",
                {
                    "module_name": module_label,
                    "is_admin": _is_trc_admin(request.user),
                },
                status=403,
            )

        return _wrapped_view

    return decorator


def module_api_access_required(module_key):
    """
    Protect API/download endpoints with granular per-module permissions.
    """
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return JsonResponse(
                    {"ok": False, "errors": ["Authentication required."]},
                    status=401,
                )
            if user_has_module_access(request.user, module_key):
                return view_func(request, *args, **kwargs)

            return JsonResponse(
                {
                    "ok": False,
                    "errors": [
                        f"Forbidden: You do not have permission for the {module_key} module."
                    ],
                },
                status=403,
            )

        return _wrapped_view

    return decorator


# ============================================================
# ROLE-SPECIFIC DECORATORS
# ============================================================

def trc_access_required(view_func):
    """
    Permission required for dashboard, rates API and history.
    """
    return permission_required_json(PERM_ACCESS_TRC)(view_func)


def trc_draft_required(view_func):
    """
    Permission required for saving working draft.
    """
    return permission_required_json(PERM_SAVE_DRAFT)(view_func)


def trc_publish_required(view_func):
    """
    Permission required for publishing rates.
    """
    return permission_required_json(PERM_PUBLISH)(view_func)


def trc_upload_required(view_func):
    """
    Permission required for uploading published CSV.
    """
    return permission_required_json(PERM_UPLOAD)(view_func)


def trc_edit_published_required(view_func):
    """
    Permission required for editing published rates.
    """
    return permission_required_json(PERM_EDIT_PUBLISHED)(view_func)


# ============================================================
# DATE FORMATTER HELPER (Converts any date to YYYY-MM-DD)
# ============================================================

def format_to_yyyy_mm_dd(d_str):
    """Converts dates like '25-Sep-26' or '25-Sep-2026' to '2026-09-25' (YYYY-MM-DD)."""
    if not d_str:
        return d_str

    months = {
        "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
        "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12"
    }

    # Already YYYY-MM-DD format
    if "-" in d_str and len(d_str.split("-")[0]) == 4:
        return d_str

    parts = d_str.split("-")
    if len(parts) == 3:
        dd = parts[0].zfill(2)
        m_part = parts[1].lower()[:3]
        mm = months.get(m_part, parts[1].zfill(2))
        yy = parts[2]
        if len(yy) == 2:
            yy = "20" + yy
        return f"{yy}-{mm}-{dd}"

    return d_str


# ============================================================
# SSO LOGOUT
# ============================================================

def sso_logout(request):
    """
    Clear the Django session and redirect to the UAT SSO logout endpoint.
    """
    logout(request)
    return redirect("https://ssouat.aubankuat.in/oam/server/logout")

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def save_history_snapshot(product, sheet_dict):
    """
    Save/update historical rate-card snapshot.
    """

    if not product or not sheet_dict:
        return

    eff_date = sheet_dict.get("effectiveDate") or "Default"

    RateCardHistory.objects.update_or_create(
        product=product,
        effective_date=eff_date,
        defaults={
            "snapshot": sheet_dict,
        },
    )


def _sheet(slug, kind):
    """
    Get required Sheet object for a product.
    """

    product = get_object_or_404(
        Product,
        slug=slug,
    )

    return get_object_or_404(
        Sheet,
        product=product,
        kind=kind,
    )


def _is_trc_admin(user):
    """
    Return True when the current user has TRC administrative access.
    """
    return (
        user.is_superuser
        or user.is_staff
        or user.has_perm(PERM_PUBLISH)
        or user.has_perm(PERM_UPLOAD)
        or user.has_perm(PERM_EDIT_PUBLISHED)
    )


# ============================================================
# APPLICATION HOME PAGE
# ============================================================


@trc_page_access_required
def home(request):
    """
    Main TRC landing page displayed after successful SSO login.
    """
    return render(
        request,
        "rates/home.html",
        {
            "is_admin": _is_trc_admin(request.user),
            "can_access_console": user_has_module_access(request.user, "console"),
            "can_access_retail": user_has_module_access(request.user, "retail"),
            "can_access_bulk": user_has_module_access(request.user, "bulk"),
            "can_access_nre_nro": user_has_module_access(request.user, "nre_nro"),
        },
    )

def _format_retail_excel_value(
    value,
    row_index,
    column_index,
):
    """
    Format Retail workbook values for dashboard display.

    Rules applied:
    - Header rows remain unchanged.
    - Tenure column (Column 1) remains unchanged.
    - Numerical bank rate cells are formatted as:
      * 7 -> 7.00%
      * 7.25 -> 7.25%
    - Blank or empty values are rendered as '-'
    """
    if value is None:
        return "-"

    # Header rows (Row 0 and Row 1) must retain original labels
    if row_index < 2:
        text = str(value).strip()
        return text or "-"

    # Tenure column (Column 0 / index 0)
    if column_index == 0:
        text = str(value).strip()
        return text or "-"

    # Rate columns (Column 1 and beyond)
    if isinstance(value, (int, float)):
        # Excel rate stored as decimal, e.g., 0.075 for 7.50%
        if 0 < value < 1:
            rate_percentage = value * 100
        else:
            rate_percentage = value

        return f"{rate_percentage:.2f}%"

    # String values that represent numbers, e.g., '7.25' or '7.25%'
    if isinstance(value, str):
        cleaned_text = value.strip().replace("%", "")

        try:
            numeric_rate = float(cleaned_text)

            if 0 < numeric_rate < 1:
                numeric_rate = numeric_rate * 100

            return f"{numeric_rate:.2f}%"

        except ValueError:
            return value.strip() or "-"

    text = str(value).strip()

    return text or "-"


def _bulk_workbook_path():
    """
    Return the absolute path of the generated Bulk workbook.
    """
    return (
        Path(settings.BASE_DIR)
        / "rates"
        / "comparison_engine"
        / BULK_WORKBOOK_NAME
    )


def _nre_nro_workbook_path():
    """
    Return the absolute path of the generated NRE/NRO workbook.
    """
    return (
        Path(settings.BASE_DIR)
        / "rates"
        / "comparison_engine"
        / NRE_NRO_WORKBOOK_NAME
    )


# ============================================================
# RETAIL COMPARISON HELPERS
# ============================================================

# ============================================================
# NRE/NRO COMPARISON CONFIGURATION
# ============================================================

NRE_NRO_WORKBOOK_NAME = "AU_NRE_NRO_Rate_Comparison.xlsx"

NRE_NRO_SHEET_CONFIG = [
    {
        "key": "comparison",
        "label": "NRE/NRO Comparison",
        "sheet_name": "NRE NRO Comparison",
    },
    {
        "key": "peak_rates",
        "label": "Peak Rates",
        "sheet_name": "NRE NRO Peak Rates",
    },
]

BULK_WORKBOOK_NAME = "AU_Bulk_TD_Rate_Comparison.xlsx"

BULK_SHEET_CONFIG = [
    {
        "key": "three_to_ten",
        "label": "3 Cr to 10 Cr",
        "sheet_name": "3-10 Cr",
    },
    {
        "key": "ten_and_above",
        "label": "10 Cr and Above",
        "sheet_name": "10 Cr and Above",
    },
]
RETAIL_WORKBOOK_NAME = "AU_FD_Rate_Comparison.xlsx"

RETAIL_SHEET_CONFIG = [
    {
        "key": "regular",
        "label": "Regular Comparison",
        "sheet_name": "Regular Comparison",
    },
    {
        "key": "senior",
        "label": "Senior Citizen Comparison",
        "sheet_name": "Senior Citizen Comparison",
    },
    {
        "key": "regular_peak",
        "label": "Regular Peak Rates",
        "sheet_name": "Regular Peak Rates",
    },
    {
        "key": "senior_peak",
        "label": "Senior Peak Rates",
        "sheet_name": "Senior Peak Rates",
    },
]


def _retail_workbook_path():
    """
    Return the absolute path of the generated Retail workbook.
    """
    return (
        Path(settings.BASE_DIR)
        / "rates"
        / "comparison_engine"
        / RETAIL_WORKBOOK_NAME
    )


def _clean_excel_value(value):
    """
    Convert raw Excel values into cleaned dashboard strings.
    """
    if value is None:
        return ""

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.4f}".rstrip("0").rstrip(".")

    text = str(value).strip()

    if text.endswith(".0"):
        return text[:-2]

    return text


def _to_numeric_rate(value):
    """
    Convert rate cell values into float numbers for comparison.
    """
    if value is None:
        return None

    if isinstance(value, (int, float)):
        numeric_value = float(value)
        if 0 < numeric_value < 1:
            return round(numeric_value * 100, 4)
        return round(numeric_value, 4)

    text = str(value).strip().replace("%", "")

    if not text or text == "-":
        return None

    try:
        numeric_value = float(text)
        if 0 < numeric_value < 1:
            return round(numeric_value * 100, 4)
        return round(numeric_value, 4)
    except ValueError:
        return None


def _trim_excel_matrix(matrix):
    """
    Remove completely empty trailing rows and columns from matrix.
    """
    if not matrix:
        return []

    # Strip empty bottom rows
    last_valid_row = -1
    for row_index in range(len(matrix) - 1, -1, -1):
        row_has_value = any(
            str(cell).strip() not in ("", "None", "-")
            for cell in matrix[row_index]
        )
        if row_has_value:
            last_valid_row = row_index
            break

    if last_valid_row == -1:
        return []

    trimmed_matrix = matrix[: last_valid_row + 1]

    # Strip empty trailing columns
    max_column_count = max(len(row) for row in trimmed_matrix)
    last_valid_col = -1

    for col_index in range(max_column_count - 1, -1, -1):
        col_has_value = False
        for row in trimmed_matrix:
            if col_index < len(row):
                cell_value = str(row[col_index]).strip()
                if cell_value not in ("", "None", "-"):
                    col_has_value = True
                    break

        if col_has_value:
            last_valid_col = col_index
            break

    if last_valid_col == -1:
        return []

    # Final matrix normalisation
    normalized_matrix = []
    for row in trimmed_matrix:
        row_slice = row[: last_valid_col + 1]
        padded_row = row_slice + [""] * (
            (last_valid_col + 1) - len(row_slice)
        )
        normalized_matrix.append(padded_row)

    return normalized_matrix


def _remove_dashboard_note_rows(matrix, stop_markers=None):
    """
    Remove footer notes, legends, and colour logic rows.
    """
    if not matrix:
        return []

    default_stop_markers = [
        "Colour logic",
        "Color logic",
        "AU Bank rate is higher",
        "AU Bank rate is lower",
        "Rates are equal",
        "Rate category",
        "Tenure rule",
        "Note",
        "Notes",
    ]

    effective_markers = stop_markers or default_stop_markers

    cleaned_matrix = []

    for row in matrix:
        first_cell_text = ""
        second_cell_text = ""

        if len(row) > 0 and row[0] is not None:
            first_cell_text = str(row[0]).strip().lower()

        if len(row) > 1 and row[1] is not None:
            second_cell_text = str(row[1]).strip().lower()

        combined_lead_text = f"{first_cell_text} {second_cell_text}".strip()

        should_stop = any(
            marker.lower() in combined_lead_text
            for marker in effective_markers
        )

        if should_stop:
            break

        cleaned_matrix.append(row)

    return cleaned_matrix


def _build_retail_sheet_data(
    worksheet,
    header_row_count=2,
    stop_markers=None,
):
    """
    Parse worksheet rows and prepare display-ready JSON structure.
    """
    raw_matrix = []

    for row in worksheet.iter_rows(values_only=True):
        raw_matrix.append(list(row))

    # Strip empty rows and trailing empty columns
    clean_matrix = _trim_excel_matrix(raw_matrix)

    # Strip footer notes and legend blocks
    clean_matrix = _remove_dashboard_note_rows(
        clean_matrix,
        stop_markers=stop_markers,
    )

    if not clean_matrix:
        return {
            "rows": [],
            "row_count": 0,
            "column_count": 0,
        }

    row_count = len(clean_matrix)
    column_count = max(len(row) for row in clean_matrix)

    processed_rows = []

    au_col_idx = 1  # Standard AU Bank rate column position

    for row_idx, row in enumerate(clean_matrix):
        processed_cells = []

        is_header = row_idx < header_row_count

        au_rate_value = None
        if not is_header and len(row) > au_col_idx:
            au_rate_value = _to_numeric_rate(row[au_col_idx])

        for col_idx in range(column_count):
            raw_value = row[col_idx] if col_idx < len(row) else None

            display_value = _format_retail_excel_value(
                raw_value,
                row_idx,
                col_idx,
            )

            cell_data = {
                "value": display_value,
                "is_header": is_header,
                "is_tenure": (col_idx == 0 and not is_header),
                "is_au_bank": (col_idx == au_col_idx and not is_header),
                "status": "normal",
            }

            # Comparison highlights for competitor columns
            if (
                not is_header
                and col_idx > au_col_idx
                and au_rate_value is not None
            ):
                comp_rate = _to_numeric_rate(raw_value)

                if comp_rate is not None:
                    if au_rate_value > comp_rate:
                        cell_data["status"] = "higher"
                    elif au_rate_value < comp_rate:
                        cell_data["status"] = "lower"
                    else:
                        cell_data["status"] = "equal"

            processed_cells.append(cell_data)

        processed_rows.append(
            {
                "is_header": is_header,
                "cells": processed_cells,
            }
        )

    return {
        "rows": processed_rows,
        "row_count": row_count,
        "column_count": column_count,
    }


def _load_retail_dashboard_data():
    """
    Load the configured Retail workbook sheets.

    The workbook is opened in read-only mode and is never modified.
    This function always returns a dictionary.
    """
    workbook_path = _retail_workbook_path()

    result = {
        "available": False,
        "error": None,
        "sheets": [],
        "updated_at": None,
        "size": None,
    }

    if not workbook_path.exists():
        result["error"] = (
            "Retail comparison workbook has not been generated yet."
        )
        return result

    workbook = None

    try:
        workbook_stats = workbook_path.stat()

        result["updated_at"] = timezone.localtime(
            datetime.fromtimestamp(
                workbook_stats.st_mtime,
                tz=timezone.get_current_timezone(),
            )
        )

        result["size"] = workbook_stats.st_size

        workbook = load_workbook(
            filename=workbook_path,
            read_only=True,
            data_only=True,
        )

        for sheet_config in RETAIL_SHEET_CONFIG:
            configured_sheet_name = sheet_config["sheet_name"]

            if configured_sheet_name not in workbook.sheetnames:
                result["sheets"].append(
                    {
                        "key": sheet_config["key"],
                        "label": sheet_config["label"],
                        "sheet_name": configured_sheet_name,
                        "available": False,
                        "error": (
                            f"Sheet '{configured_sheet_name}' "
                            "was not found in the workbook."
                        ),
                        "rows": [],
                        "row_count": 0,
                        "column_count": 0,
                    }
                )
                continue

            worksheet = workbook[configured_sheet_name]

            sheet_data = _build_retail_sheet_data(
                worksheet=worksheet,
                header_row_count=2,
                stop_markers=[
                    "Colour logic",
                    "Color logic",
                    "AU Bank rate is higher",
                    "AU Bank rate is lower",
                    "Rates are equal",
                    "Rate category",
                    "Tenure rule",
                ],
            )

            result["sheets"].append(
                {
                    "key": sheet_config["key"],
                    "label": sheet_config["label"],
                    "sheet_name": configured_sheet_name,
                    "available": True,
                    "error": None,
                    "rows": sheet_data["rows"],
                    "row_count": sheet_data["row_count"],
                    "column_count": sheet_data["column_count"],
                }
            )

        result["available"] = True
        result["error"] = None

    except Exception as exc:
        result["available"] = False
        result["error"] = (
            "Unable to read the Retail comparison workbook. "
            f"Reason: {exc}"
        )
        result["sheets"] = []

    finally:
        if workbook is not None:
            workbook.close()

    return result


# ============================================================
# DASHBOARD
# ============================================================

@module_access_required("console", "Rate Console")
def dashboard(request):
    """
    Main TRC dashboard.
    """

    data = {}

    ordered_slugs = [
        "fin_callable",
        "callable",
        "noncallable",
        "ib_callable",
        "ib_noncallable",
    ]

    products = list(Product.objects.all())

    products.sort(
        key=lambda p: (
            ordered_slugs.index(p.slug)
            if p.slug in ordered_slugs
            else 99
        )
    )

    for product in products:

        published_sheet = product.sheets.filter(
            kind=Sheet.PUBLISHED
        ).first()

        working_sheet = product.sheets.filter(
            kind=Sheet.WORKING
        ).first()

        if published_sheet:
            published_dict = sheet_to_dict(published_sheet)
        else:
            published_dict = {}

        if working_sheet:
            working_dict = sheet_to_dict(working_sheet)
        else:
            working_dict = {}

        data[product.slug] = {
            "published": published_dict,
            "working": working_dict,
        }

        if not product.history.exists() and published_dict:
            save_history_snapshot(
                product,
                published_dict,
            )

    return render(
        request,
        "rates/dashboard.html",
        {
            "data_json": json.dumps(data),
            "is_admin": _is_trc_admin(request.user),
        },
    )


# ============================================================
# WORKING DRAFT
# ============================================================

@trc_draft_required
@require_POST
def save_draft(request, slug):
    """
    Save working draft. Persist the working draft for Admin review.
    """

    try:
        payload = json.loads(request.body or "{}")

    except json.JSONDecodeError:
        return JsonResponse(
            {
                "ok": False,
                "errors": ["Invalid JSON payload."],
            },
            status=400,
        )

    sheet_data = payload.get("sheet", {})

    errors = validate_dict(sheet_data)

    if errors:
        return JsonResponse(
            {
                "ok": False,
                "errors": errors,
            },
            status=400,
        )

    working_sheet = _sheet(
        slug,
        Sheet.WORKING,
    )

    rebuild_sheet(
        working_sheet,
        sheet_data,
    )

    updated_working_sheet = _sheet(
        slug,
        Sheet.WORKING,
    )

    return JsonResponse(
        {
            "ok": True,
            "message": "File saved - Draft sent to Admin",
            "working": sheet_to_dict(updated_working_sheet),
        }
    )


# ============================================================
# PUBLISH RATES
# ============================================================

@trc_publish_required
@require_POST
def publish(request, slug):
    """
    Publish rates.

    Restricted to staff/admin or user with publish permission.
    Validates the working draft, copies it to published rates,
    saves historical snapshot, generates BOTH CSV files (Quarterly + Nominal Monthly),
    saves local copies to server /Input directory (cleaning same-day old files for this product),
    and uploads BOTH CSV files to Amazon S3.
    """

    try:
        payload = json.loads(request.body or "{}")

    except json.JSONDecodeError:
        return JsonResponse(
            {
                "ok": False,
                "errors": ["Invalid JSON payload."],
            },
            status=400,
        )

    try:
        sheet_data = payload.get("sheet", {})

        errors = validate_dict(sheet_data)

        if errors:
            return JsonResponse(
                {
                    "ok": False,
                    "errors": errors,
                },
                status=400,
            )

        product = get_object_or_404(
            Product,
            slug=slug,
        )

        # Update working sheet
        rebuild_sheet(
            _sheet(
                slug,
                Sheet.WORKING,
            ),
            sheet_data,
        )

        # Update published sheet
        rebuild_sheet(
            _sheet(
                slug,
                Sheet.PUBLISHED,
            ),
            sheet_data,
        )

        published_sheet = _sheet(
            slug,
            Sheet.PUBLISHED,
        )

        published_dict = sheet_to_dict(
            published_sheet
        )

        working_sheet = _sheet(
            slug,
            Sheet.WORKING,
        )

        working_dict = sheet_to_dict(
            working_sheet
        )

        # Save historical snapshot
        save_history_snapshot(
            product,
            published_dict,
        )

        # Format EffectiveDate to YYYY-MM-DD (e.g. 2026-09-25) for BOTH exports
        d_export = dict(sheet_data)
        d_export["effectiveDate"] = format_to_yyyy_mm_dd(sheet_data.get("effectiveDate"))

        clean_label = (
            sheet_data
            .get("productLabel", "rates")
            .replace(" ", "_")
            .replace("/", "_")
        )

        timestamp_str = eff_stamp(sheet_data.get("effectiveDate"))

        # Generate CSV 1 (Quarterly CSV - EffectiveDate: YYYY-MM-DD)
        content1 = export_csv(d_export)

        filename1 = (
            f"{clean_label}_rates_"
            f"{timestamp_str}.csv"
        )

        # Generate CSV 2 (Nominal Monthly CSV - EffectiveDate: YYYY-MM-DD)
        content2 = export_nominal_csv(d_export)

        filename2 = (
            f"{clean_label}_Nominal_Monthly_Rates_"
            f"{timestamp_str}.csv"
        )

        # Save local server copy (Delete same-day older files ONLY for THIS product)
        SERVER_EXPORT_DIR = getattr(settings, "SERVER_EXPORT_DIR", "/Input")

        try:
            os.makedirs(
                SERVER_EXPORT_DIR,
                exist_ok=True,
            )

            # Extract date part (YYYYMMDD) from timestamp_str
            date_part = timestamp_str.split("_")[0]

            # Delete ONLY same-day old files for THIS specific product (prevents replacing other products)
            for existing_file in os.listdir(SERVER_EXPORT_DIR):
                is_same_product = (
                    existing_file.startswith(f"{clean_label}_rates_") or
                    existing_file.startswith(f"{clean_label}_Nominal_Monthly_Rates_")
                )
                if is_same_product and f"_{date_part}_" in existing_file:
                    old_file_path = os.path.join(SERVER_EXPORT_DIR, existing_file)
                    if os.path.isfile(old_file_path):
                        os.remove(old_file_path)
                        print(f"[REPLACE SUCCESS] Removed same-day old file: {existing_file}")

            # Save new latest files
            file_path1 = os.path.join(SERVER_EXPORT_DIR, filename1)
            file_path2 = os.path.join(SERVER_EXPORT_DIR, filename2)

            with open(file_path1, "w", encoding="utf-8") as f:
                f.write(content1)

            with open(file_path2, "w", encoding="utf-8") as f:
                f.write(content2)

            print(f"[LOCAL EXPORT SUCCESS] Saved latest {filename1} & {filename2} to {SERVER_EXPORT_DIR}")

        except Exception as exc:
            print(f"[LOCAL EXPORT ERROR] Local file replace failed: {exc}")

        # Upload BOTH generated CSV files to S3
        s3_success1 = False
        s3_success2 = False

        try:
            s3_success1 = upload_to_s3(filename1, content1)
            s3_success2 = upload_to_s3(filename2, content2)

            if s3_success1 and s3_success2:
                print(f"[S3 SUCCESS] Both files uploaded to S3: {filename1}, {filename2}")
            else:
                print(f"[S3 WARNING] S3 Upload status - File1: {s3_success1}, File2: {s3_success2}")

        except Exception as exc:
            print(f"[S3 WARNING] S3 upload failed: {exc}")

        return JsonResponse(
            {
                "ok": True,
                "filename": filename1,
                "content": content1,
                "nom_filename": filename2,
                "nom_content": content2,
                "s3_uploaded": (s3_success1 and s3_success2),
                "published": published_dict,
                "working": working_dict,
            }
        )

    except Exception as exc:
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    str(exc) or "Unable to publish rates."
                ],
            },
            status=500,
        )


# ============================================================
# CSV UPLOAD
# ============================================================

@trc_upload_required
@require_POST
def upload_published(request, slug):
    """
    Upload published rate CSV.

    Security:
    - Authentication
    - RBAC
    - POST only
    - Maximum 1 MB file size
    - CSV extension validation
    """

    # --------------------------------------------------------
    # FILE PRESENCE
    # --------------------------------------------------------

    if "file" not in request.FILES:
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "No CSV file was uploaded."
                ],
            },
            status=400,
        )

    csv_file = request.FILES["file"]

    # --------------------------------------------------------
    # FILE SIZE
    # --------------------------------------------------------

    if csv_file.size > MAX_UPLOAD_SIZE:
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "File size exceeds the maximum allowed limit of 1 MB."
                ],
            },
            status=400,
        )

    if csv_file.size == 0:
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "Uploaded file is empty."
                ],
            },
            status=400,
        )

    # --------------------------------------------------------
    # FILE EXTENSION
    # --------------------------------------------------------

    filename = csv_file.name or ""

    if not filename.lower().endswith(".csv"):
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "Only CSV files are allowed."
                ],
            },
            status=400,
        )

    # --------------------------------------------------------
    # PROCESS CSV
    # --------------------------------------------------------

    try:
        csv_bytes = csv_file.read()

        try:
            csv_text = csv_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            csv_text = csv_bytes.decode("utf-8-sig", errors="ignore")

        product = get_object_or_404(
            Product,
            slug=slug,
        )

        sheet_data = parse_csv_to_dict(
            csv_text,
            product.slug,
            product.label,
        )

        errors = validate_dict(
            sheet_data
        )

        if errors:
            return JsonResponse(
                {
                    "ok": False,
                    "errors": errors,
                },
                status=400,
            )

        # Update published sheet
        rebuild_sheet(
            _sheet(
                slug,
                Sheet.PUBLISHED,
            ),
            sheet_data,
        )

        # Update working sheet
        rebuild_sheet(
            _sheet(
                slug,
                Sheet.WORKING,
            ),
            sheet_data,
        )

        published_dict = sheet_to_dict(
            _sheet(
                slug,
                Sheet.PUBLISHED,
            )
        )

        working_dict = sheet_to_dict(
            _sheet(
                slug,
                Sheet.WORKING,
            )
        )

        # Maintain history
        save_history_snapshot(
            product,
            published_dict,
        )

        return JsonResponse(
            {
                "ok": True,
                "published": published_dict,
                "working": working_dict,
            }
        )

    except Exception as exc:
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    str(exc) or "Unable to process the uploaded CSV file."
                ],
            },
            status=400,
        )


# ============================================================
# SAVE PUBLISHED RATES
# ============================================================

@trc_edit_published_required
@require_POST
def save_published(request, slug):
    """
    Save manual changes to published rates.
    """

    try:
        payload = json.loads(request.body or "{}")

    except json.JSONDecodeError:
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "Invalid JSON payload."
                ],
            },
            status=400,
        )

    sheet_data = payload.get(
        "sheet",
        {},
    )

    errors = validate_dict(
        sheet_data
    )

    if errors:
        return JsonResponse(
            {
                "ok": False,
                "errors": errors,
            },
            status=400,
        )

    product = get_object_or_404(
        Product,
        slug=slug,
    )

    rebuild_sheet(
        _sheet(
            slug,
            Sheet.PUBLISHED,
        ),
        sheet_data,
    )

    published_dict = sheet_to_dict(
        _sheet(
            slug,
            Sheet.PUBLISHED,
        )
    )

    save_history_snapshot(
        product,
        published_dict,
    )

    return JsonResponse(
        {
            "ok": True,
            "published": published_dict,
        }
    )


# ============================================================
# RATES JSON API
# ============================================================

@trc_access_required
def get_rates_api(request):
    """
    Return latest published and working rates.

    Restricted to authorized TRC users.
    """

    data = {}

    products = Product.objects.all().order_by("id")

    for product in products:

        published_sheet = product.sheets.filter(
            kind=Sheet.PUBLISHED
        ).first()

        working_sheet = product.sheets.filter(
            kind=Sheet.WORKING
        ).first()

        if published_sheet:
            published_dict = sheet_to_dict(
                published_sheet
            )
        else:
            published_dict = {}

        if working_sheet:
            working_dict = sheet_to_dict(
                working_sheet
            )
        else:
            working_dict = {}

        data[product.slug] = {
            "published": published_dict,
            "working": working_dict,
        }

    return JsonResponse(
        {
            "ok": True,
            "data": data,
        }
    )


# ============================================================
# HISTORICAL RATE CARDS
# ============================================================

@trc_access_required
def get_history_list(request, slug):
    """
    Return historical effective dates for a product.

    Restricted to authorized TRC users.
    """

    product = get_object_or_404(
        Product,
        slug=slug,
    )

    histories = product.history.all()

    if not histories.exists():

        published_sheet = product.sheets.filter(
            kind=Sheet.PUBLISHED
        ).first()

        if published_sheet:

            published_dict = sheet_to_dict(
                published_sheet
            )

            save_history_snapshot(
                product,
                published_dict,
            )

            histories = product.history.all()

    data = [
        {
            "id": history.id,
            "effective_date": history.effective_date,
            "updated_at": history.created_at.strftime(
                "%d-%b-%Y %H:%M"
            ),
        }
        for history in histories
    ]

    return JsonResponse(
        {
            "ok": True,
            "history": data,
        }
    )


@trc_access_required
def get_history_detail(request, history_id):
    """
    Return historical rate-card snapshot.

    Restricted to authorized TRC users.
    """

    history = get_object_or_404(
        RateCardHistory,
        id=history_id,
    )

    return JsonResponse(
        {
            "ok": True,
            "snapshot": history.snapshot,
            "effective_date": history.effective_date,
        }
    )


# ============================================================
# RATE COMPARISON VIEWS
# ============================================================

@module_access_required("retail", "Retail TD Comparison")
def retail_comparison(request):
    """
    Display the Retail comparison workbook as a TRC dashboard.
    """
    retail_data = _load_retail_dashboard_data()

    return render(
        request,
        "rates/comparison/retail.html",
        {
            "is_admin": _is_trc_admin(request.user),
            "workbook_available": retail_data["available"],
            "workbook_error": retail_data["error"],
            "workbook_updated_at": retail_data["updated_at"],
            "workbook_size": retail_data["size"],
            "retail_sheets": retail_data["sheets"],
        },
    )


def _load_comparison_dashboard_data(
    workbook_path,
    sheet_config,
    module_label,
):
    """
    Load configured workbook sheets for comparison dashboards.

    Supported dashboards:
    - Retail TD
    - Bulk TD
    - NRE/NRO TD

    Bulk worksheets use four header rows.
    Retail and NRE/NRO worksheets use two header rows.

    This function always returns a dictionary.
    """
    result = {
        "available": False,
        "error": None,
        "sheets": [],
        "updated_at": None,
        "size": None,
    }

    # --------------------------------------------------------
    # WORKBOOK EXISTENCE CHECK
    # --------------------------------------------------------

    if not workbook_path.exists():
        result["error"] = (
            f"{module_label} comparison workbook "
            "has not been generated yet."
        )
        return result

    workbook = None

    try:
        # ----------------------------------------------------
        # WORKBOOK FILE INFORMATION
        # ----------------------------------------------------

        workbook_stats = workbook_path.stat()

        result["updated_at"] = timezone.localtime(
            datetime.fromtimestamp(
                workbook_stats.st_mtime,
                tz=timezone.get_current_timezone(),
            )
        )

        result["size"] = workbook_stats.st_size

        # ----------------------------------------------------
        # OPEN WORKBOOK
        # ----------------------------------------------------

        workbook = load_workbook(
            filename=workbook_path,
            read_only=True,
            data_only=True,
        )

        # ----------------------------------------------------
        # HEADER ROW CONFIGURATION
        # ----------------------------------------------------

        if module_label == "Bulk TD":
            header_row_count = 4
        else:
            header_row_count = 2

        # ----------------------------------------------------
        # ITERATE CONFIGURED SHEETS
        # ----------------------------------------------------

        for config_entry in sheet_config:
            sheet_name = config_entry["sheet_name"]

            if sheet_name not in workbook.sheetnames:
                result["sheets"].append(
                    {
                        "key": config_entry["key"],
                        "label": config_entry["label"],
                        "sheet_name": sheet_name,
                        "available": False,
                        "error": (
                            f"Sheet '{sheet_name}' was not "
                            "found in the workbook."
                        ),
                        "rows": [],
                        "row_count": 0,
                        "column_count": 0,
                    }
                )
                continue

            worksheet = workbook[sheet_name]

            sheet_data = _build_retail_sheet_data(
                worksheet=worksheet,
                header_row_count=header_row_count,
                stop_markers=[
                    "Colour logic",
                    "Color logic",
                    "AU Bank rate is higher",
                    "AU Bank rate is lower",
                    "Rates are equal",
                    "Rate category",
                    "Tenure rule",
                    "Note",
                    "Notes",
                ],
            )

            # ------------------------------------------------
            # BULK UI COLUMN COMPACTION
            # ------------------------------------------------
            # Bulk workbooks contain blank separator columns between
            # bank groups. The Bulk dashboard template expects a compact
            # matrix containing exactly:
            #
            # 1 tenure column + 28 rate columns = 29 UI columns.
            #
            # Excel columns retained:
            # A:I   -> Tenure and Public Banks
            # K:R   -> Large Private Banks
            # T:AC  -> Mid Private and Other Banks
            # AF:AG -> AU Bank Callable and Non-Callable
            #
            # This is intentionally Bulk-only so Retail and NRE/NRO
            # retain their existing workbook layouts and behaviour.
            if module_label == "Bulk TD":
                bulk_ui_column_indexes = [
                    0,                  # Excel A: Tenure
                    *range(1, 9),       # Excel B:I
                    *range(10, 18),     # Excel K:R
                    *range(19, 29),     # Excel T:AC
                    31,                 # Excel AF: AU Callable
                    32,                 # Excel AG: AU Non-Callable
                ]

                compact_rows = []

                for dashboard_row in sheet_data.get("rows", []):
                    source_cells = dashboard_row.get("cells", [])

                    compact_cells = [
                        source_cells[column_index]
                        for column_index in bulk_ui_column_indexes
                        if column_index < len(source_cells)
                    ]

                    compact_row = dict(dashboard_row)
                    compact_row["cells"] = compact_cells
                    compact_rows.append(compact_row)

                sheet_data["rows"] = compact_rows
                sheet_data["column_count"] = len(
                    bulk_ui_column_indexes
                )

            result["sheets"].append(
                {
                    "key": config_entry["key"],
                    "label": config_entry["label"],
                    "sheet_name": sheet_name,
                    "available": True,
                    "error": None,
                    "rows": sheet_data["rows"],
                    "row_count": sheet_data["row_count"],
                    "column_count": sheet_data["column_count"],
                }
            )

        result["available"] = True
        result["error"] = None

    except Exception as exc:
        result["available"] = False
        result["error"] = (
            f"Unable to read the {module_label} "
            f"comparison workbook. Reason: {exc}"
        )

        result["sheets"] = []

    finally:
        if workbook is not None:
            workbook.close()

    return result


@module_access_required("bulk", "Bulk TD Comparison")
def bulk_comparison(request):
    """
    Display the Bulk comparison workbook as a TRC dashboard.
    """
    bulk_data = _load_comparison_dashboard_data(
        workbook_path=_bulk_workbook_path(),
        sheet_config=BULK_SHEET_CONFIG,
        module_label="Bulk TD",
    )

    return render(
        request,
        "rates/comparison/bulk.html",
        {
            "is_admin": _is_trc_admin(request.user),
            "workbook_available": bulk_data.get(
                "available",
                False,
            ),
            "workbook_error": bulk_data.get("error"),
            "workbook_updated_at": bulk_data.get(
                "updated_at"
            ),
            "workbook_size": bulk_data.get("size"),
            "bulk_sheets": bulk_data.get(
                "sheets",
                [],
            ),
        },
    )


@module_access_required("nre_nro", "NRE/NRO TD Comparison")
def nre_nro_comparison(request):
    """
    Display the NRE/NRO comparison workbook as a TRC dashboard.
    """
    nre_nro_data = _load_comparison_dashboard_data(
        workbook_path=_nre_nro_workbook_path(),
        sheet_config=NRE_NRO_SHEET_CONFIG,
        module_label="NRE/NRO TD",
    )

    return render(
        request,
        "rates/comparison/nre_nro.html",
        {
            "is_admin": _is_trc_admin(request.user),
            "workbook_available": nre_nro_data.get(
                "available",
                False,
            ),
            "workbook_error": nre_nro_data.get(
                "error",
            ),
            "workbook_updated_at": nre_nro_data.get(
                "updated_at",
            ),
            "workbook_size": nre_nro_data.get(
                "size",
            ),
            "nre_nro_sheets": nre_nro_data.get(
                "sheets",
                [],
            ),
        },
    )


@module_api_access_required("bulk")
def download_bulk_comparison(request):
    """
    Download the latest generated Bulk comparison workbook.
    """
    workbook_path = _bulk_workbook_path()

    if not workbook_path.exists():
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "Bulk comparison workbook has not "
                    "been generated yet."
                ],
            },
            status=404,
        )

    return FileResponse(
        workbook_path.open("rb"),
        as_attachment=True,
        filename=workbook_path.name,
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


@module_api_access_required("nre_nro")
def download_nre_nro_comparison(request):
    """
    Download the latest generated NRE/NRO comparison workbook.
    """
    workbook_path = _nre_nro_workbook_path()

    if not workbook_path.exists():
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "NRE/NRO comparison workbook has not "
                    "been generated yet."
                ],
            },
            status=404,
        )

    return FileResponse(
        workbook_path.open("rb"),
        as_attachment=True,
        filename=workbook_path.name,
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# ============================================================
# RETAIL COMPARISON APIs
# ============================================================

@trc_access_required
def retail_comparison_status(request):
    """
    Return the current Retail workbook status.
    """
    workbook_path = _retail_workbook_path()

    if not workbook_path.exists():
        return JsonResponse(
            {
                "ok": True,
                "available": False,
                "message": "Retail comparison workbook has not been generated yet.",
            }
        )

    file_stats = workbook_path.stat()

    return JsonResponse(
        {
            "ok": True,
            "available": True,
            "filename": workbook_path.name,
            "size": file_stats.st_size,
            "updated_at": file_stats.st_mtime,
        }
    )


@module_api_access_required("retail")
def download_retail_comparison(request):
    """
    Download the latest generated Retail comparison workbook.
    """
    workbook_path = _retail_workbook_path()

    if not workbook_path.exists():
        return JsonResponse(
            {
                "ok": False,
                "errors": [
                    "Retail comparison workbook has not been generated yet."
                ],
            },
            status=404,
        )

    return FileResponse(
        workbook_path.open("rb"),
        as_attachment=True,
        filename=workbook_path.name,
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

# ============================================================
# MANUAL RATE REFRESH
# ============================================================

logger = logging.getLogger(__name__)


@trc_access_required
@require_GET
def refresh_all_rates(request):
    """
    Run Retail, Bulk and NRE/NRO comparison scripts.

    The user is redirected back to the module from which refresh was started.
    Each script runs in a separate process so one failure does not crash Django.
    """
    requested_page = request.GET.get("next", "home")
    allowed_redirects = {
        "retail": "retail_comparison",
        "bulk": "bulk_comparison",
        "nre_nro": "nre_nro_comparison",
        "home": "home",
    }
    redirect_url_name = allowed_redirects.get(requested_page, "home")
    project_dir = Path(settings.BASE_DIR)

    scrapers = [
        (
            "Retail",
            project_dir / "rates" / "comparison_engine" / "retail_main.py",
        ),
        (
            "Bulk",
            project_dir / "rates" / "comparison_engine" / "bulk_main.py",
        ),
        (
            "NRE/NRO",
            project_dir / "rates" / "comparison_engine" / "nre_nro_main.py",
        ),
    ]

    successful_scrapers = []
    failed_scrapers = []

    for scraper_name, scraper_path in scrapers:
        if not scraper_path.exists():
            logger.error("%s script not found: %s", scraper_name, scraper_path)
            failed_scrapers.append(scraper_name)
            continue

        try:
            process = subprocess.run(
                [sys.executable, str(scraper_path)],
                cwd=str(project_dir),
                capture_output=True,
                text=True,
                timeout=900,
                check=False,
            )

            if process.returncode == 0:
                successful_scrapers.append(scraper_name)
                logger.info("%s rates refreshed successfully.", scraper_name)
            else:
                failed_scrapers.append(scraper_name)
                logger.error(
                    "%s refresh failed.\nSTDOUT:\n%s\nSTDERR:\n%s",
                    scraper_name,
                    process.stdout,
                    process.stderr,
                )

        except subprocess.TimeoutExpired:
            failed_scrapers.append(scraper_name)
            logger.exception("%s scraper timed out.", scraper_name)
        except Exception:
            failed_scrapers.append(scraper_name)
            logger.exception("%s scraper execution failed.", scraper_name)

    if not failed_scrapers:
        messages.success(request, "Rates refreshed successfully.")
    elif successful_scrapers:
        messages.warning(
            request,
            "Rates were partially refreshed. Please check the application logs.",
        )
    else:
        messages.error(
            request,
            "Rates could not be refreshed. Please check the application logs.",
        )

    return redirect(redirect_url_name)

