import csv
import logging
import re
from datetime import datetime

import boto3
import copy
from django.conf import settings

from django.conf import settings
from django.db import transaction

from botocore.config import Config
from botocore.exceptions import (
    ClientError,
    ConnectTimeoutError,
    ReadTimeoutError,
    EndpointConnectionError,
)

from .models import Group, Slab, Tenure, Rate


logger = logging.getLogger(__name__)


# =========================================================
# EXPORT CONFIGURATION
# =========================================================

EXPORT_COLS = [
    "EffectiveDate",
    "StartMonths",
    "StartDays",
    "AmountFrom",
    "AmountTo",
    "InterestIndexRate",
    "InterestVarianceRate",
    "PenaltyIndexRate",
    "PenaltyVarianceRate",
]

INTEREST_VARIANCE = "0"
PENALTY_INDEX = "1%"
PENALTY_VARIANCE = "0"
DELIMITER = ","


# =========================================================
# HELPERS
# =========================================================

def parse_amt(v):
    digits = re.sub(
        r"[^\d.]",
        "",
        str(v if v is not None else "")
    )

    try:
        return float(digits)
    except ValueError:
        return None


def to_int(v):
    a = parse_amt(v)
    return int(a) if a is not None else None


def fmt_int(v):
    a = parse_amt(v)
    return "" if a is None else str(int(a))


def fmt_rate(v):
    try:
        return "" if v in ("", None) else ("%g%%" % float(v))
    except (ValueError, TypeError):
        return ""


def amt_str(n):
    return "" if n is None else str(n)


def stamp():
    return datetime.now().strftime("%Y%m%d_%H%M")


def eff_stamp(eff_date_str=None):
    time_part = datetime.now().strftime("%H%M")

    if eff_date_str:
        s = str(eff_date_str).strip()

        for fmt in (
            "%d-%b-%y",
            "%d-%b-%Y",
            "%Y-%m-%d",
        ):
            try:
                dt = datetime.strptime(s, fmt)

                return (
                    f"{dt.strftime('%Y%m%d')}_"
                    f"{time_part}"
                )

            except ValueError:
                pass

        digits = re.sub(
            r"[^\d]",
            "",
            s
        )

        if len(digits) == 8:
            return f"{digits}_{time_part}"

    return datetime.now().strftime(
        "%Y%m%d_%H%M"
    )


# =========================================================
# SERIALIZE DB SHEET
# =========================================================

def sheet_to_dict(sheet):

    groups = list(
        sheet.groups.order_by(
            "order",
            "id"
        )
    )

    slabs = list(
        sheet.slabs.order_by(
            "order",
            "id"
        )
    )

    tenures = list(
        sheet.tenures.order_by(
            "order",
            "id"
        )
    )

    group_id_map = {
        g.id: f"g{i + 1}"
        for i, g in enumerate(groups)
    }

    s_idx_map = {
        s.id: i
        for i, s in enumerate(slabs)
    }

    t_idx_map = {
        t.id: i
        for i, t in enumerate(tenures)
    }

    rates = sheet.rates.all()

    rates_dict = {}

    for r in rates:

        ti = t_idx_map.get(
            r.tenure_id
        )

        si = s_idx_map.get(
            r.slab_id
        )

        if (
            ti is not None
            and si is not None
        ):
            rates_dict[
                f"idx_{ti}_{si}"
            ] = r.value

        rates_dict[
            f"t{r.tenure_id}|s{r.slab_id}"
        ] = r.value

    slabs_serialized = []

    for i, s in enumerate(slabs):

        at = s.amt_to or 0

        if (
            at <= 100000000
            and len(groups) >= 1
        ):
            grp_id = group_id_map.get(
                groups[0].id,
                "g1"
            )

        elif (
            at <= 500000000
            and len(groups) >= 2
        ):
            grp_id = group_id_map.get(
                groups[1].id,
                "g2"
            )

        elif (
            at <= 1050000000
            and len(groups) >= 3
        ):
            grp_id = group_id_map.get(
                groups[2].id,
                "g3"
            )

        elif len(groups) >= 4:
            grp_id = group_id_map.get(
                groups[3].id,
                "g4"
            )

        else:
            grp_id = group_id_map.get(
                s.group_id,
                "g1"
            )

        slabs_serialized.append({
            "id": f"s{s.id}",
            "groupId": grp_id,
            "label": s.label,
            "amtFrom": amt_str(s.amt_from),
            "amtTo": amt_str(s.amt_to),
            "tint": s.tint,
            "sub": s.sub,
        })

    return {
        "product": sheet.product.slug,
        "kind": sheet.kind,
        "effectiveDate": sheet.effective_date,
        "productLabel": sheet.product.label,

        "groups": [
            {
                "id": f"g{i + 1}",
                "label": g.label,
            }
            for i, g in enumerate(groups)
        ],

        "slabs": slabs_serialized,

        "tenures": [
            {
                "id": f"t{t.id}",
                "fromMonth": t.from_month,
                "toDays": t.to_days,
            }
            for t in tenures
        ],

        "rates": rates_dict,
    }


# =========================================================
# REBUILD DB SHEET
# =========================================================

@transaction.atomic
def rebuild_sheet(sheet, d):

    sheet.effective_date = d.get(
        "effectiveDate",
        sheet.effective_date
    )

    sheet.save()

    sheet.rates.all().delete()
    sheet.slabs.all().delete()
    sheet.tenures.all().delete()
    sheet.groups.all().delete()

    gmap = {}
    g_list = []

    for i, g in enumerate(
        d.get("groups", [])
    ):
        grp_obj = Group.objects.create(
            sheet=sheet,
            label=g.get(
                "label",
                ""
            ),
            order=i
        )

        gmap[g["id"]] = grp_obj
        gmap[f"g{i + 1}"] = grp_obj

        g_list.append(grp_obj)

    fallback_group = (
        g_list[0]
        if g_list
        else Group.objects.create(
            sheet=sheet,
            label="3 -10 crs",
            order=0
        )
    )

    smap = {}

    for i, s in enumerate(
        d.get("slabs", [])
    ):

        at = (
            to_int(
                s.get("amtTo")
            )
            or 0
        )

        grp = None

        if (
            at <= 100000000
            and len(g_list) >= 1
        ):
            grp = g_list[0]

        elif (
            at <= 500000000
            and len(g_list) >= 2
        ):
            grp = g_list[1]

        elif (
            at <= 1050000000
            and len(g_list) >= 3
        ):
            grp = g_list[2]

        elif len(g_list) >= 4:
            grp = g_list[3]

        if grp is None:

            grp = (
                gmap.get(
                    s.get("groupId")
                )
                or fallback_group
            )

        smap[s["id"]] = (
            Slab.objects.create(
                sheet=sheet,
                group=grp,
                label=s.get(
                    "label",
                    ""
                ),
                amt_from=to_int(
                    s.get("amtFrom")
                ),
                amt_to=to_int(
                    s.get("amtTo")
                ),
                tint=bool(
                    s.get("tint")
                ),
                sub=s.get(
                    "sub",
                    "Callable"
                ),
                order=i,
            )
        )

    tmap = {}
    ordered_tenures = []

    for i, t in enumerate(
        d.get("tenures", [])
    ):

        ten_obj = (
            Tenure.objects.create(
                sheet=sheet,

                from_month=int(
                    float(
                        t.get(
                            "fromMonth"
                        )
                        or 0
                    )
                ),

                to_days=int(
                    float(
                        t.get(
                            "toDays"
                        )
                        or 0
                    )
                ),

                order=i,
            )
        )

        tmap[t["id"]] = ten_obj

        ordered_tenures.append(
            ten_obj
        )

    s_list = list(
        sheet.slabs.order_by(
            "order",
            "id"
        )
    )

    bulk = []
    seen = set()

    for key, val in (
        d.get(
            "rates",
            {}
        ).items()
    ):

        tt = None
        ss = None

        if key.startswith("idx_"):

            parts = key.split("_")

            if len(parts) == 3:

                try:
                    ti = int(parts[1])
                    si = int(parts[2])

                    if (
                        ti < len(
                            ordered_tenures
                        )
                        and si < len(
                            s_list
                        )
                    ):
                        tt = ordered_tenures[ti]
                        ss = s_list[si]

                except ValueError:
                    pass

        elif "|" in key:

            tid, sid = key.split(
                "|",
                1
            )

            tt = tmap.get(tid)
            ss = smap.get(sid)

        if tt and ss:

            pair = (
                tt.id,
                ss.id
            )

            if pair not in seen:

                seen.add(pair)

                bulk.append(
                    Rate(
                        sheet=sheet,
                        tenure=tt,
                        slab=ss,
                        value=str(val)
                    )
                )

    Rate.objects.bulk_create(
        bulk
    )


# =========================================================
# VALIDATION
# =========================================================

def validate_dict(d):

    errors = []

    slabs = d.get(
        "slabs",
        []
    )

    if slabs:

        f_first = parse_amt(
            slabs[0].get(
                "amtFrom"
            )
        )

        if (
            f_first is not None
            and f_first != 30000000
        ):
            errors.append(
                "First slab Amt From must be "
                "fixed to 3 Cr (3,00,00,000)."
            )

        t_last = parse_amt(
            slabs[-1].get(
                "amtTo"
            )
        )

        if (
            t_last is not None
            and t_last != 99999999999
        ):
            errors.append(
                "Last slab Amt To must be fixed "
                "to 99,99,99,99,999."
            )

    intervals = []

    for s in slabs:

        label = (
            s.get("label")
            or "(unnamed)"
        )

        f = parse_amt(
            s.get("amtFrom")
        )

        t = parse_amt(
            s.get("amtTo")
        )

        if (
            f is None
            or t is None
        ):
            errors.append(
                f'Slab "{label}" is missing an amount.'
            )

        elif f < 30000000:

            errors.append(
                f'Slab "{label}": '
                "Amt From cannot be below "
                "3 Cr (3,00,00,000)."
            )

        elif t > 99999999999:

            errors.append(
                f'Slab "{label}": '
                "Amt To cannot exceed "
                "99,99,99,99,999."
            )

        elif f > t:

            errors.append(
                f'Slab "{label}": '
                "Amt From is greater than Amt To."
            )

        else:

            intervals.append(
                (f, t, label)
            )

    intervals.sort(
        key=lambda x: x[0]
    )

    for i in range(
        1,
        len(intervals)
    ):

        if (
            intervals[i][0]
            <= intervals[i - 1][1]
        ):
            errors.append(
                f'Amounts overlap: '
                f'"{intervals[i - 1][2]}" and '
                f'"{intervals[i][2]}" '
                "cover the same range."
            )

    tens = d.get(
        "tenures",
        []
    )

    for i, t in enumerate(tens):

        am = int(
            t.get("fromMonth")
            or 0
        )

        ad = int(
            t.get("toDays")
            or 0
        )

        if (
            am == 0
            and ad < 7
        ):
            errors.append(
                f"Tenure row {i + 1} "
                f"({am}m {ad}d): "
                "Duration cannot be below 7 days."
            )

        elif (
            am > 120
            or (
                am == 120
                and ad > 0
            )
        ):
            errors.append(
                f"Tenure row {i + 1} "
                f"({am}m {ad}d): "
                "Duration cannot exceed 120 months."
            )

    for i in range(
        1,
        len(tens)
    ):

        a = tens[i - 1]
        b = tens[i]

        am = int(
            a.get("fromMonth")
            or 0
        )

        ad = int(
            a.get("toDays")
            or 0
        )

        bm = int(
            b.get("fromMonth")
            or 0
        )

        bd = int(
            b.get("toDays")
            or 0
        )

        if not (
            bm > am
            or (
                bm == am
                and bd > ad
            )
        ):

            errors.append(
                f"Tenure not ascending: "
                f"{bm}m {bd}d must come "
                f"after {am}m {ad}d."
            )

    rates = d.get(
        "rates",
        {}
    )

    for key, val in rates.items():

        if val in ("", None):
            continue

        try:
            r_float = float(val)

        except ValueError:

            errors.append(
                f'Rate "{val}" is not a valid number.'
            )

            continue

        if (
            r_float < 2
            or r_float > 10
        ):
            errors.append(
                f'Rate "{val}%" must be '
                "between 2% and 10%."
            )

        if "." in str(val):

            decimals = (
                str(val)
                .split(
                    ".",
                    1
                )[1]
            )

            if len(decimals) > 3:

                errors.append(
                    f'Rate "{val}%" must have '
                    "at most 3 decimal places."
                )

    return errors


# ---------- flat file in the shared format ----------
def format_eff_date_csv(eff_date_str):
    """
    Format date string strictly to yyyy-mm-dd (e.g. 2026-09-25) for CSV rows
    """
    if not eff_date_str:
        return ""
    s = str(eff_date_str).strip()
    # Try parsing with .title() so "25-sep-26", "25-SEP-26", "25-Sep-26" all parse correctly
    for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%d-%B-%Y"):
        try:
            dt = datetime.strptime(s.title(), fmt)
            return dt.strftime("%Y-%m-%d")  # Format: yyyy-mm-dd
        except ValueError:
            pass
    return s
 


def export_csv(d):
    print("DEBUG RAW DATE:", repr(d.get("effectiveDate")))
    date = format_eff_date_csv(d.get("effectiveDate", ""))
    print("DEBUG FORMATTED DATE:", repr(date))

    tenures = sorted(d.get("tenures", []), key=lambda t: (int(float(t.get("fromMonth") or 0)), int(float(t.get("toDays") or 0))))
    slabs = sorted(d.get("slabs", []), key=lambda s: (parse_amt(s.get("amtFrom")) or 0))
    rates = d.get("rates", {})
    lines = [DELIMITER.join(EXPORT_COLS)]
    for t in tenures:
        for s in slabs:
            rate = fmt_rate(rates.get(f"{t['id']}|{s['id']}", ""))
            lines.append(DELIMITER.join([
                date, str(t.get("fromMonth")), str(t.get("toDays")),
                fmt_int(s.get("amtFrom")), fmt_int(s.get("amtTo")),
                rate, INTEREST_VARIANCE, PENALTY_INDEX, PENALTY_VARIANCE,
            ]))
    return "\n".join(lines)
 

# =========================================================
# MATRIX CSV PARSER
# =========================================================

def parse_matrix_csv_to_dict(
    lines,
    product_slug,
    product_label
):

    reader = csv.reader(lines)

    rows = list(reader)

    if not rows:
        raise ValueError(
            "Uploaded CSV file is empty."
        )

    eff_date = "06-Aug-26"

    amt_from_row = None
    amt_to_row = None
    label_row = None
    sub_row = None
    data_start_idx = None

    for idx, r in enumerate(rows):

        if not r:
            continue

        r_clean = [
            c.strip()
            for c in r
        ]

        first_col = (
            r_clean[0].lower()
            if len(r_clean) > 0
            else ""
        )

        second_col = (
            r_clean[1].lower()
            if len(r_clean) > 1
            else ""
        )

        if (
            "effective date" in first_col
            and len(r_clean) > 1
            and r_clean[1]
        ):
            eff_date = r_clean[1]

        elif "amt slab" in second_col:
            label_row = r_clean

        elif "amt from" in second_col:
            amt_from_row = r_clean

        elif "amt to" in second_col:
            amt_to_row = r_clean

        elif "product type" in second_col:
            sub_row = r_clean

        elif "tenure from month" in first_col:
            data_start_idx = idx + 1

    if (
        amt_from_row is None
        or amt_to_row is None
        or data_start_idx is None
    ):
        raise ValueError(
            "Unrecognized CSV format. "
            "Please upload either a 9-column "
            "Flat CSV or a Working Draft Matrix CSV."
        )

    sub_default = (
        "Callable"
        if (
            "callable" in product_slug.lower()
            and "non" not in product_slug.lower()
        )
        else "Non-Callable"
    )

    num_cols = max(
        len(amt_from_row),
        len(amt_to_row)
    )

    parsed_slabs = []

    for c in range(
        2,
        num_cols
    ):

        af_val = (
            to_int(
                amt_from_row[c]
            )
            if c < len(amt_from_row)
            else None
        )

        at_val = (
            to_int(
                amt_to_row[c]
            )
            if c < len(amt_to_row)
            else None
        )

        if (
            af_val is None
            or at_val is None
        ):
            continue

        lbl_val = (
            label_row[c]
            if (
                label_row
                and c < len(label_row)
                and label_row[c]
            )
            else f"{af_val} to {at_val}"
        )

        sub_val = (
            sub_row[c]
            if (
                sub_row
                and c < len(sub_row)
                and sub_row[c]
            )
            else sub_default
        )

        if at_val <= 100000000:
            grp_id = "g1"

        elif at_val <= 500000000:
            grp_id = "g2"

        elif at_val <= 1050000000:
            grp_id = "g3"

        else:
            grp_id = "g4"

        parsed_slabs.append({
            "id": (
                f"s{len(parsed_slabs) + 1}"
            ),

            "groupId": grp_id,

            "label": lbl_val,

            "amtFrom": str(af_val),

            "amtTo": str(at_val),

            "tint": (
                len(parsed_slabs) % 2 == 1
            ),

            "sub": sub_val,
        })

    parsed_tenures = []
    rates = {}

    for row_idx in range(
        data_start_idx,
        len(rows)
    ):

        r_clean = [
            c.strip()
            for c in rows[row_idx]
        ]

        if (
            not r_clean
            or len(r_clean) < 2
        ):
            continue

        try:

            m = int(
                float(
                    r_clean[0]
                )
            )

            d = int(
                float(
                    r_clean[1]
                )
            )

        except ValueError:
            continue

        t_id = (
            f"t{len(parsed_tenures) + 1}"
        )

        parsed_tenures.append({
            "id": t_id,
            "fromMonth": m,
            "toDays": d,
        })

        for s_idx, s_obj in enumerate(
            parsed_slabs
        ):

            col_idx = (
                s_idx + 2
            )

            r_val = ""

            if col_idx < len(r_clean):

                r_val = (
                    r_clean[col_idx]
                    .rstrip("%")
                    .strip()
                )

            rates[
                f"{t_id}|{s_obj['id']}"
            ] = r_val

            rates[
                f"idx_{len(parsed_tenures) - 1}_{s_idx}"
            ] = r_val

    groups = [
        {
            "id": "g1",
            "label": "3 -10 crs"
        },
        {
            "id": "g2",
            "label": "10-50crs"
        },
        {
            "id": "g3",
            "label": "50-105crs"
        },
        {
            "id": "g4",
            "label": "Greater than 105"
        },
    ]

    return {
        "product": product_slug,
        "kind": "published",
        "effectiveDate": eff_date,
        "productLabel": product_label,
        "groups": groups,
        "slabs": parsed_slabs,
        "tenures": parsed_tenures,
        "rates": rates,
    }


# =========================================================
# FLAT/MATRIX CSV PARSER
# =========================================================

def parse_csv_to_dict(
    csv_text,
    product_slug,
    product_label
):

    lines = [
        line.strip()
        for line in csv_text.splitlines()
        if line.strip()
    ]

    if not lines:

        raise ValueError(
            "Uploaded CSV file is empty."
        )

    reader = csv.reader(lines)

    header = next(reader)

    col_map = {
        col.strip(): i
        for i, col in enumerate(header)
    }

    req_cols = [
        "EffectiveDate",
        "StartMonths",
        "StartDays",
        "AmountFrom",
        "AmountTo",
        "InterestIndexRate",
    ]

    is_flat = all(
        c in col_map
        for c in req_cols
    )

    if not is_flat:

        return parse_matrix_csv_to_dict(
            lines,
            product_slug,
            product_label
        )

    eff_date = "06-Aug-26"

    tenure_set = set()
    slab_set = set()
    raw_rates = {}

    for row in reader:

        if (
            not row
            or len(row)
            <= max(
                col_map.values()
            )
        ):
            continue

        date_val = (
            row[
                col_map[
                    "EffectiveDate"
                ]
            ]
            .strip()
        )

        if date_val:
            eff_date = date_val

        try:

            m = int(
                float(
                    row[
                        col_map[
                            "StartMonths"
                        ]
                    ]
                )
            )

            d = int(
                float(
                    row[
                        col_map[
                            "StartDays"
                        ]
                    ]
                )
            )

        except ValueError:
            continue

        af = to_int(
            row[
                col_map[
                    "AmountFrom"
                ]
            ]
        )

        at = to_int(
            row[
                col_map[
                    "AmountTo"
                ]
            ]
        )

        if (
            af is None
            or at is None
        ):
            continue

        rate_val = (
            row[
                col_map[
                    "InterestIndexRate"
                ]
            ]
            .strip()
            .rstrip("%")
        )

        tenure_key = (
            m,
            d
        )

        slab_key = (
            af,
            at
        )

        tenure_set.add(
            tenure_key
        )

        slab_set.add(
            slab_key
        )

        raw_rates[
            (
                tenure_key,
                slab_key
            )
        ] = rate_val

    sorted_tenures = sorted(
        list(tenure_set),
        key=lambda x: (
            x[0],
            x[1]
        )
    )

    sorted_slabs = sorted(
        list(slab_set),
        key=lambda x: x[0]
    )

    groups = [
        {
            "id": "g1",
            "label": "3 -10 crs"
        },
        {
            "id": "g2",
            "label": "10-50crs"
        },
        {
            "id": "g3",
            "label": "50-105crs"
        },
        {
            "id": "g4",
            "label": "Greater than 105"
        },
    ]

    sub = (
        "Callable"
        if (
            "callable"
            in product_slug.lower()
            and "non"
            not in product_slug.lower()
        )
        else "Non-Callable"
    )

    slabs = []

    for i, (
        af,
        at
    ) in enumerate(
        sorted_slabs
    ):

        if at <= 100000000:
            grp_id = "g1"

        elif at <= 500000000:
            grp_id = "g2"

        elif at <= 1050000000:
            grp_id = "g3"

        else:
            grp_id = "g4"

        lbl = (
            f"{af} to {at}"
        )

        if (
            af == 30000000
            and at == 32499999
        ):
            lbl = "3cr to <3.25cr"

        elif (
            af == 32500000
            and at == 34999999
        ):
            lbl = "3.25cr to <3.50cr"

        elif (
            af == 35000000
            and at == 100000000
        ):
            lbl = "3.50cr to 10cr"

        elif (
            af == 100000001
            and at == 150000000
        ):
            lbl = "Above 10cr to 15cr"

        elif (
            af == 150000001
            and at == 200000000
        ):
            lbl = "Above 15 cr to 20 cr"

        elif (
            af == 200000001
            and at == 250000000
        ):
            lbl = "Above 20 cr to 25 cr"

        elif (
            af == 250000001
            and at == 350000000
        ):
            lbl = "Above 25 cr to 35 cr"

        elif (
            af == 350000001
            and at == 500000000
        ):
            lbl = "Above 35 cr to 50 cr"

        elif (
            af == 500000001
            and at == 700000000
        ):
            lbl = "Above 50 cr to 70 cr"

        elif (
            af == 700000001
            and at == 950000000
        ):
            lbl = "Above 70 cr to 95 cr"

        elif (
            af == 950000001
            and at == 1050000000
        ):
            lbl = "Above 95 cr to 105 cr"

        elif (
            af == 1050000001
            and at == 1500000000
        ):
            lbl = "Above 105 cr to 150 cr"

        elif (
            af == 1500000001
            and at == 2000000000
        ):
            lbl = "Above 150 cr to 200 cr"

        elif af >= 2000000001:
            lbl = "Above 200 cr"

        tint = (
            i % 2 == 1
        )

        slabs.append({
            "id": f"s{i + 1}",
            "groupId": grp_id,
            "label": lbl,
            "amtFrom": str(af),
            "amtTo": str(at),
            "tint": tint,
            "sub": sub,
        })

    tenures = []

    for i, (
        m,
        d
    ) in enumerate(
        sorted_tenures
    ):

        tenures.append({
            "id": f"t{i + 1}",
            "fromMonth": m,
            "toDays": d,
        })

    rates = {}

    for t_idx, t_key in enumerate(
        sorted_tenures
    ):

        for s_idx, s_key in enumerate(
            sorted_slabs
        ):

            r_val = raw_rates.get(
                (
                    t_key,
                    s_key
                ),
                ""
            )

            rates[
                f"t{t_idx + 1}|s{s_idx + 1}"
            ] = r_val

            rates[
                f"idx_{t_idx}_{s_idx}"
            ] = r_val

    return {
        "product": product_slug,
        "kind": "published",
        "effectiveDate": eff_date,
        "productLabel": product_label,
        "groups": groups,
        "slabs": slabs,
        "tenures": tenures,
        "rates": rates,
    }


# =========================================================
# AMAZON S3 UPLOAD
# =========================================================

def upload_to_s3(
    filename,
    file_content
):
    """
    Upload a generated TRC CSV file to Amazon S3.

    Configuration is read from Django settings.

    Expected destination:

        s3://<bucket>/<prefix>/<filename>

    The on-premises TRC server connects through
    the configured Amazon S3 Interface VPC Endpoint.
    """

    # -----------------------------------------------------
    # Read configuration
    # -----------------------------------------------------

    aws_access_key = getattr(
        settings,
        "AWS_ACCESS_KEY_ID",
        ""
    ).strip()

    aws_secret_key = getattr(
        settings,
        "AWS_SECRET_ACCESS_KEY",
        ""
    ).strip()

    region = getattr(
        settings,
        "AWS_REGION",
        "ap-south-1"
    ).strip()

    bucket = getattr(
        settings,
        "S3_BUCKET_NAME",
        ""
    ).strip()

    prefix = getattr(
        settings,
        "S3_PREFIX",
        "Bulk_FD"
    ).strip().strip("/")

    endpoint_url = getattr(
        settings,
        "S3_ENDPOINT_URL",
        ""
    ).strip()


    # -----------------------------------------------------
    # Validate configuration
    # -----------------------------------------------------

    if not aws_access_key:

        logger.error(
            "[S3] AWS_ACCESS_KEY_ID "
            "is not configured."
        )

        return False


    if not aws_secret_key:

        logger.error(
            "[S3] AWS_SECRET_ACCESS_KEY "
            "is not configured."
        )

        return False


    if not region:

        logger.error(
            "[S3] AWS_REGION "
            "is not configured."
        )

        return False


    if not bucket:

        logger.error(
            "[S3] S3_BUCKET_NAME "
            "is not configured."
        )

        return False


    if not endpoint_url:

        logger.error(
            "[S3] S3_ENDPOINT_URL "
            "is not configured."
        )

        return False


    # -----------------------------------------------------
    # Build destination key
    # -----------------------------------------------------

    if prefix:

        object_key = (
            f"{prefix}/{filename}"
        )

    else:

        object_key = filename


    # -----------------------------------------------------
    # Convert content to bytes
    # -----------------------------------------------------

    if isinstance(
        file_content,
        str
    ):

        body = file_content.encode(
            "utf-8"
        )

    elif isinstance(
        file_content,
        bytes
    ):

        body = file_content

    else:

        body = str(
            file_content
        ).encode(
            "utf-8"
        )


    # -----------------------------------------------------
    # Configure AWS SDK
    # -----------------------------------------------------

    config = Config(
        region_name=region,

        connect_timeout=10,

        read_timeout=30,

        retries={
            "max_attempts": 2,
            "mode": "standard",
        },
    )


    try:

        logger.info(
            "[S3] Starting upload. "
            "Bucket=%s Key=%s",
            bucket,
            object_key,
        )


        # -------------------------------------------------
        # Create S3 client
        # -------------------------------------------------

        s3_client = boto3.client(
            "s3",

            aws_access_key_id=(
                aws_access_key
            ),

            aws_secret_access_key=(
                aws_secret_key
            ),

            region_name=region,

            endpoint_url=(
                endpoint_url
            ),

            config=config,
        )


        # -------------------------------------------------
        # Upload object
        # -------------------------------------------------

        response = (
            s3_client.put_object(
                Bucket=bucket,

                Key=object_key,

                Body=body,

                ContentType=(
                    "text/csv"
                ),
            )
        )


        # -------------------------------------------------
        # Validate response
        # -------------------------------------------------

        status_code = (
            response
            .get(
                "ResponseMetadata",
                {}
            )
            .get(
                "HTTPStatusCode"
            )
        )


        if status_code == 200:

            logger.info(
                "[S3 SUCCESS] "
                "Successfully uploaded %s "
                "to s3://%s/%s",
                filename,
                bucket,
                object_key,
            )

            return True


        logger.error(
            "[S3 ERROR] "
            "Unexpected HTTP status %s "
            "while uploading %s",
            status_code,
            filename,
        )

        return False


    # -----------------------------------------------------
    # AWS / IAM / BUCKET POLICY ERROR
    # -----------------------------------------------------

    except ClientError as exc:

        error = (
            exc.response.get(
                "Error",
                {}
            )
        )

        logger.error(
            "[S3 ERROR] "
            "AWS rejected upload of %s. "
            "Code=%s Message=%s",
            filename,
            error.get("Code"),
            error.get("Message"),
        )

        return False


    # -----------------------------------------------------
    # CONNECTION TIMEOUT
    # -----------------------------------------------------

    except ConnectTimeoutError as exc:

        logger.error(
            "[S3 ERROR] "
            "Connection timeout while "
            "uploading %s: %s",
            filename,
            exc,
        )

        return False


    # -----------------------------------------------------
    # READ TIMEOUT
    # -----------------------------------------------------

    except ReadTimeoutError as exc:

        logger.error(
            "[S3 ERROR] "
            "Read timeout while "
            "uploading %s: %s",
            filename,
            exc,
        )

        return False


    # -----------------------------------------------------
    # ENDPOINT ERROR
    # -----------------------------------------------------

    except EndpointConnectionError as exc:

        logger.error(
            "[S3 ERROR] "
            "Unable to connect to "
            "S3 endpoint while "
            "uploading %s: %s",
            filename,
            exc,
        )

        return False


    # -----------------------------------------------------
    # UNEXPECTED ERROR
    # -----------------------------------------------------

    except Exception as exc:

        logger.exception(
            "[S3 ERROR] "
            "Unexpected error while "
            "uploading %s: %s",
            filename,
            exc,
        )

        return False


# =========================================================
# NOMINAL MONTHLY CSV EXPORTER
# =========================================================

# NOMINAL MONTHLY CSV EXPORTER
# =========================================================
 
def convert_to_nominal_monthly_dict(d):
    d_nom = copy.deepcopy(d)
    new_rates = {}

    old_tenures = d_nom.get("tenures", [])
    old_rates = d_nom.get("rates", {})
    slabs = d_nom.get("slabs", [])

    def to_nominal_monthly(r):
        if not r:
            return ""
        try:
            num = float(str(r).replace("%", "").strip())
            if num <= 0:
                return str(r)
            monthly = 12 * ((1 + num / 400.0) ** (1.0 / 3.0) - 1.0) * 100.0
            return f"{monthly:.2f}"
        except (ValueError, TypeError):
            return str(r)

    # Exact same tenures as shown in UI
    for t_idx, t in enumerate(old_tenures):
        for s_idx, s in enumerate(slabs):
            v = old_rates.get(f"idx_{t_idx}_{s_idx}") or old_rates.get(f"{t['id']}|{s['id']}", "")
            nom_v = to_nominal_monthly(v)
            new_rates[f"{t['id']}|{s['id']}"] = nom_v
            new_rates[f"idx_{t_idx}_{s_idx}"] = nom_v

    d_nom["rates"] = new_rates
    d_nom["effectiveDate"] = format_eff_date_csv(d.get("effectiveDate", ""))
    return d_nom
 

 
 
def export_nominal_csv(d):
    d_nom = convert_to_nominal_monthly_dict(d)
    return export_csv(d_nom)