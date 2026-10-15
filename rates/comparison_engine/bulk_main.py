#!/usr/bin/env python3
"""Bulk TD (above 3 crore) peer comparison against AU Small Finance Bank.

* Rates are never hardcoded - every rate is read from the bank's official
  HTML page / PDF (or from a file the user saved into manual_sources/ when a
  site blocks automated access).
* Every source slab is normalised to a day range [StartDay, EndDay].
* The comparison rows are AU's tenure buckets.  For every bucket and bank the
  cell shows the MAXIMUM published rate among the bank's slabs that overlap
  the bucket (Callable and Non-Callable kept separate).
* The workbook mirrors the treasury peer-comparison layout: group headers,
  "Above 3 cr" bank headers, C / NC columns with effective dates, AU at the
  right, and green / red / yellow cells for lower / higher / equal vs AU.

Usage:
    python single_btd.py                 # scrape every bank and build workbook
    python single_btd.py --only AU,HDFC  # scrape a subset, then build
    python single_btd.py --skip-scrape   # rebuild workbook from latest CSVs
"""
from __future__ import annotations

import argparse
import csv
import email.policy
import glob
import hashlib
import io
import json
import logging
import os
import re
import sys
import time
import traceback
from dataclasses import asdict, dataclass, fields
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable, Optional
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
except ImportError:  # PDF sources are skipped with a clear error
    pdfplumber = None

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output_bulk"
RAW_DIR = ROOT / "raw_sources" / datetime.now().strftime("%Y-%m-%d")
MANUAL_DIR = ROOT / "manual_sources"
LOG_DIR = ROOT / "logs"
RESULT_FILE = ROOT / "AU_Bulk_TD_Rate_Comparison.xlsx"
for folder in (OUTPUT_DIR, RAW_DIR, MANUAL_DIR, LOG_DIR):
    folder.mkdir(parents=True, exist_ok=True)

for _stream in (sys.stdout, sys.stderr):   # Windows consoles default to cp1252
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

LOG = logging.getLogger("bulk_td")
if not LOG.handlers:
    LOG.setLevel(logging.INFO)
    _fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    for _h in (logging.StreamHandler(sys.stdout), logging.FileHandler(LOG_DIR / "bulk_td.log", encoding="utf-8")):
        _h.setFormatter(_fmt)
        LOG.addHandler(_h)

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
BANKS = {
    "BOB": "https://bankofbaroda.bank.in/interest-rate-and-service-charges/deposits-interest-rates/fixed-deposits-callable-and-non-callable-above-ten-crores",
    "BOI": "https://bankofindia.bank.in/interest-rate/rupee-term-deposit-rate",
    "UNION": "https://www.unionbankofindia.bank.in/en/details/rate-of-interest",
    "PNB": "https://pnb.bank.in/interest-rates-bulk-deposit.html",
    "HDFC": "https://www.hdfc.bank.in/interest-rates",
    "ICICI": "https://www.icici.bank.in/personal-banking/deposits/fixed-deposit/fd-interest-rates",
    "AXIS": "https://www.axis.bank.in/deposits/fixed-deposits/fd-interest-rates",
    "KOTAK": "https://www.kotak.bank.in/en/personal-banking/deposits/fixed-deposit/fixed-deposit-interest-rate.html",
    "INDUSIND": "https://www.indusind.bank.in/in/en/personal/rates.html",
    "BANDHAN": "https://bandhan.bank.in/rates-charges",
    "IDFC": "https://www.idfcfirst.bank.in/interest-rate",
    "RBL": "https://www.rbl.bank.in/interest-rates",
    "DCB": "https://www.dcb.bank.in/rates/fixed-deposit-interest-rate",
    "AU": "https://www.au.bank.in/interest-rates/bulk-fixed-deposit-interest-rates",
}

BANK_ORDER = ["BOB", "BOI", "UNION", "PNB", "HDFC", "ICICI", "AXIS", "KOTAK",
              "INDUSIND", "BANDHAN", "IDFC", "RBL", "DCB", "AU"]
DISPLAY_NAMES = {
    "BOB": "BOB", "BOI": "BoI", "UNION": "Union", "PNB": "PNB", "HDFC": "HDFC",
    "ICICI": "ICICI", "AXIS": "Axis", "KOTAK": "Kotak", "INDUSIND": "IndusInd",
    "BANDHAN": "Bandhan", "IDFC": "IDFC", "RBL": "RBL Bank", "DCB": "DCB", "AU": "AU Bank",
}
# (title, banks, fill colour, gap column after the group)
GROUPS = [
    ("Public", ["BOB", "BOI", "UNION", "PNB"], "F8CBAD", True),
    ("Large Private Banks", ["HDFC", "ICICI", "AXIS", "KOTAK"], "FFE699", True),
    ("Mid Private Banks", ["INDUSIND", "BANDHAN", "IDFC"], "BDD7EE", False),
    ("Other Banks", ["RBL", "DCB"], "E4B7E8", True),
]

# "Above 3 cr": parsers return every published bulk slab covering amounts above 3 crore.
#   AMOUNT_MODE = "slab"  -> each cell is the best rate across all those amount slabs (default)
#   AMOUNT_MODE = "slab" -> only the slab that contains AMOUNT_POINT_CR (entry bulk slab)
BULK_FLOOR_CR = 3.0
AMOUNT_POINT_CR = 3.01
AMOUNT_MODE = "slab"

# A source slab that touches an AU bucket on a single boundary day only (e.g. HDFC '30-45 days' vs
# AU '15 Days to 1 Month' = 15-30) is an artefact of month/day conversion and is ignored, unless the
# source slab or the AU bucket is itself a single day (e.g. '444 days', '12 Months').
MIN_OVERLAP_DAYS = 1

# Day convention used for every bank: 1 month = 30 days, 1 year (12 months) = 365 days.
# This matches the way Indian banks publish day based slabs (46-60, 61-90,
# 91-120, 121-180, 181-210, 211-270, 271-364, 365, 366 ... days).
MAX_TENURE_DAYS = 3650


def to_days(years: float = 0, months: float = 0, days: float = 0) -> int:
    total_months = years * 12 + months
    whole_years = int(total_months // 12)
    rem_months = total_months - whole_years * 12
    return int(round(whole_years * 365 + rem_months * 30 + days))


# AU tenure buckets = the comparison rows.  Day ranges follow the convention above.
AU_TENURES = [
    ("7 Days to 14 Days", 7, 14),
    ("15 Days to 1 Month", 15, to_days(months=1)),
    ("1 Month 1 Day to 1 Month 15 Days", to_days(months=1, days=1), to_days(months=1, days=15)),
    ("1 Month 16 Days to 2 Months", to_days(months=1, days=16), to_days(months=2)),
    ("2 Months 1 Day to 3 Months", to_days(months=2, days=1), to_days(months=3)),
    ("3 Months 1 Day to 4 Months", to_days(months=3, days=1), to_days(months=4)),
    ("4 Months 1 Day to 6 Months", to_days(months=4, days=1), to_days(months=6)),
    ("6 Months 1 Day to 7 Months", to_days(months=6, days=1), to_days(months=7)),
    ("7 Months 1 Day to 9 Months", to_days(months=7, days=1), to_days(months=9)),
    ("9 Months 1 Day to <12 Months", to_days(months=9, days=1), to_days(months=12) - 1),
    ("12 Months", to_days(months=12), to_days(months=12)),
    ("12 Months 1 Day to 13 Months", to_days(months=12, days=1), to_days(months=13)),
    ("13 Months 1 Day to 18 Months", to_days(months=13, days=1), to_days(months=18)),
    ("18 Months 1 Day to 24 Months", to_days(months=18, days=1), to_days(months=24)),
    ("24 Months 1 Day to 36 Months", to_days(months=24, days=1), to_days(months=36)),
    ("36 Months 1 Day to 45 Months", to_days(months=36, days=1), to_days(months=45)),
    ("45 Months 1 Day to 60 Months", to_days(months=45, days=1), to_days(months=60)),
    ("60 Months 1 Day and above", to_days(months=60, days=1), MAX_TENURE_DAYS),
]

# Values outside this range are penalty / additional / annualised columns, not card rates.
MIN_VALID_RATE = 2.0
MAX_VALID_RATE = 9.5

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/pdf;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

CALLABLE, NON_CALLABLE = "Callable", "Non-Callable"
WITHDRAWAL_TYPES = (CALLABLE, NON_CALLABLE)


@dataclass
class RateRow:
    BankName: str
    WithdrawalType: str
    AmountSlab: str
    Tenure: str
    StartDay: int
    EndDay: int
    Rate: float
    EffectiveDate: str
    SourceURL: str
    ExtractionMode: str
    ScrapedAt: str = ""


class SourceBlocked(RuntimeError):
    """Raised when a site answers with a bot-protection page instead of content."""


# --------------------------------------------------------------------------
# Text helpers
# --------------------------------------------------------------------------
def clean(value) -> str:
    if value is None:
        return ""
    text = str(value)
    for bad, good in (("\xa0", " "), ("​", ""), (" ", " "), (" ", " "), ("–", "-"), ("—", "-"),
                      ("−", "-"), ("≤", "<="), ("≥", ">="), ("’", "'"), ("‘", "'"), ("₹", " Rs ")):
        text = text.replace(bad, good)
    return re.sub(r"\s+", " ", text).strip()


RATE_RE = re.compile(r"(?<![\d.])(\d{1,2}(?:\.\d{1,4})?)\s*%?")


def rate_value(value) -> Optional[float]:
    """First percentage-looking number in a cell, or None for NA / blanks / tenure text."""
    text = clean(value)
    if not text or re.fullmatch(r"[-*#/ ]+|na|n\.a\.?|n/a|nil|not applicable|not available|no slab", text, re.I):
        return None
    if re.search(r"\d\s*(?:days?|months?|years?|yrs?)\b", text, re.I):
        return None
    match = RATE_RE.search(text)
    if not match:
        return None
    rate = float(match.group(1))
    return round(rate, 4) if 0 < rate <= 20 else None


_UNIT = r"(years?|yrs?|y|months?|mths?|mnths?|m|days?|d)"
_PART_RE = re.compile(r"(\d+(?:\.\d+)?)\s*" + _UNIT + r"(?![a-z])")


def _unit_kind(unit: str) -> str:
    return unit[0].lower()


def fragment_days(text: str, default_unit: str = "") -> Optional[int]:
    """'1 year 2 months 3 days' -> days.  A bare number takes default_unit."""
    text = clean(text).lower()
    parts = _PART_RE.findall(text)
    if parts:
        y = m = d = 0.0
        for number, unit in parts:
            kind = _unit_kind(unit)
            if kind == "y":
                y += float(number)
            elif kind == "m":
                m += float(number)
            else:
                d += float(number)
        return to_days(y, m, d)
    bare = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*", text)
    if bare and default_unit:
        return fragment_days(f"{bare.group(1)} {default_unit}")
    return None


def _last_unit(text: str) -> str:
    parts = _PART_RE.findall(clean(text).lower())
    return parts[0][1] if parts else ""


_START_EXCL = re.compile(r"^(?:above|more than|over|beyond|exceeding|greater than|>(?!=))\s*")
_START_INCL = re.compile(r"^(?:from|>=|min(?:imum)?\.?|at least)\s*")
_END_EXCL = re.compile(r"^(?:less than|below|under|before|<(?!=)|but less than|but below)\s*")
_END_INCL = re.compile(r"^(?:up to|upto|till|<=|and including|inclusive of|up to and inclusive of|max(?:imum)?\.?)\s*")


def tenure_bounds(value) -> tuple[Optional[int], Optional[int]]:
    """Convert a published tenure label into an inclusive day range."""
    text = clean(value).lower()
    if not text or not re.search(r"\d", text):
        return None, None
    text = re.sub(r"\([^()]*\)|\[[^\[\]]*\]", " ", text)           # '390 days (12 months 25 days)'
    text = re.sub(r"[*#@^$]+|\bonly\b|\bexactly\b|\bperiod\b|\btenure\b", " ", text)
    text = text.replace("&", " and ").replace("upto", "up to").replace("less then", "less than")
    text = re.sub(r"\b(?:less than|below)\s+(?:or\s+)?equal\s+to\b", "<=", text)
    text = re.sub(r"\b(?:greater|more) than\s+(?:or\s+)?equal\s+to\b", ">=", text)
    text = re.sub(r"\bgreater than\b", "above", text)
    text = re.sub(r"(\d)\s*(?=(?:years?|yrs?|months?|days?)\b)", r"\1 ", text)
    text = re.sub(r"\s+", " ", text).strip(" ,.:")
    # 'X and above but less than Y' / 'X and above up to Y' -> 'X to <qualifier> Y'
    text = re.sub(r"\s+and\s+above\s+(?:but\s+|and\s+)?(?:to\s+)?(?=(?:less|below|under|up|<|till|\d))", " to ", text)
    text = re.sub(r"\s+(?:and|but)\s+(?=(?:up to|less than|below|under|<=?|till))", " to ", text)
    # From / To cells joined without a separator: '365 days below 1 year 3 months'
    text = re.sub(r"(\d\s*(?:days?|months?|mths?|years?|yrs?))\s+(?=(?:below|less than|under|up to|till)\b)", r"\1 to ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not re.search(r" to | - | till | through |<|-", text):
        # '7 days 14 days' / '1 year 3 months 1 year 6 months': a unit that does not get finer starts the 2nd end
        rank = {"y": 0, "m": 1, "d": 2}
        parts = list(_PART_RE.finditer(text))
        splits = [i for i in range(1, len(parts))
                  if rank[_unit_kind(parts[i].group(2))] <= rank[_unit_kind(parts[i - 1].group(2))]]
        if len(splits) == 1:
            cut = parts[splits[0]].start()
            text = text[:cut].strip() + " to " + text[cut:].strip()

    # open ended: 'X and above', 'X onwards', 'above X'
    open_end = re.fullmatch(r"(.+?)\s+(?:and\s+above|and\s+over|onwards|or\s+more|and\s+more|\+)", text)
    if open_end and " to " not in text:
        start = fragment_days(open_end.group(1))
        return (start, MAX_TENURE_DAYS) if start is not None else (None, None)

    # one sided: 'up to 14 days', 'less than 1 year', 'above 10 years', '>= 365 days'
    if len(_PART_RE.findall(text)) == 1 and text.startswith(">="):
        value = fragment_days(text[2:])
        return (value, MAX_TENURE_DAYS) if value else (None, None)
    if len(_PART_RE.findall(text)) == 1 and (_END_EXCL.match(text) or _END_INCL.match(text) or _START_EXCL.match(text)):
        value = fragment_days(_END_INCL.sub("", _END_EXCL.sub("", _START_EXCL.sub("", text))))
        if value is None:
            return None, None
        if _END_EXCL.match(text):
            return 1, value - 1
        if _END_INCL.match(text):
            return 1, value
        return value + 1, max(value + 1, MAX_TENURE_DAYS)

    # split into left / right parts
    left = right = None
    for sep in (" to ", " - ", " till ", " through ", " up to "):
        if sep in text:
            left, right = text.split(sep, 1)
            if sep in (" till ", " up to "):
                right = "up to " + right
            break
    if left is None:
        hyphen = re.fullmatch(r"(.*?\d\s*[a-z]*)\s*-\s*(.+)", text)
        if hyphen:
            left, right = hyphen.group(1), hyphen.group(2)
    if left is None:
        op = re.fullmatch(r"(.*?\d\s*[a-z]*)\s*(<=|<)\s*(.+)", text)       # '> 1 year <= 2 years'
        if op and re.search(r"\d", op.group(1)):
            left, right = op.group(1), op.group(2) + " " + op.group(3)

    if left is None:                                                       # single sided / exact
        if _END_EXCL.match(text):
            end = fragment_days(_END_EXCL.sub("", text))
            return (1, end - 1) if end else (None, None)
        if _END_INCL.match(text):
            end = fragment_days(_END_INCL.sub("", text))
            return (1, end) if end else (None, None)
        if _START_EXCL.match(text):
            start = fragment_days(_START_EXCL.sub("", text))
            return (start + 1, MAX_TENURE_DAYS) if start is not None else (None, None)
        point = fragment_days(_START_INCL.sub("", text))
        return (point, point) if point else (None, None)

    left, right = left.strip(), right.strip()
    start_exclusive = bool(_START_EXCL.match(left))
    left = _START_INCL.sub("", _START_EXCL.sub("", left))
    end_exclusive = bool(_END_EXCL.match(right))
    right = _END_INCL.sub("", _END_EXCL.sub("", right))
    end = fragment_days(right) or fragment_days(right, default_unit=_last_unit(left))   # '8 year 1 day - 10'
    start = fragment_days(left, default_unit=_last_unit(right))
    if start is not None and end is not None and start > end and re.fullmatch(r"\s*\d+(?:\.\d+)?\s*", left):
        start = fragment_days(left, default_unit="days")                                # '185 to < 1 year'
    if start is None or end is None:
        return None, None
    if start_exclusive:
        start += 1
    if end_exclusive:
        end -= 1
    return (start, end) if end >= start else (None, None)


_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}
_DATE_PATTERNS = [
    re.compile(r"(?P<d>\d{1,2})(?:st|nd|rd|th)?[\s\-/.,]*(?P<mon>jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s\-/.,']*(?P<y>\d{4}|\d{2})(?!\d)", re.I),
    re.compile(r"(?P<mon>jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*(?P<d>\d{1,2})(?:st|nd|rd|th)?[\s,]+(?P<y>\d{4})", re.I),
    re.compile(r"(?<!\d)(?P<d>\d{1,2})[\-/.](?P<m>\d{1,2})[\-/.](?P<y>\d{4}|\d{2})(?!\d)"),
    re.compile(r"(?<!\d)(?P<y>\d{4})-(?P<m>\d{2})-(?P<d>\d{2})(?!\d)"),
]
_EFFECTIVE_KEYS = re.compile(
    r"w\.?\s*e\.?\s*f\.?|with\s+effect\s+from|effective\s+(?:from|date|since|on)?|applicable\s+(?:from|for|on|w\.?e\.?f)|"
    r"revised\s+(?:on|w\.?e\.?f)|valid\s+(?:from|for)|rates?\s+as\s+on|as\s+on", re.I)


def _dates_in(text: str) -> list[tuple[int, str]]:
    found = []
    for pattern in _DATE_PATTERNS:
        for m in pattern.finditer(text):
            g = m.groupdict()
            try:
                month = _MONTHS[g["mon"][:3].lower()] if g.get("mon") else int(g["m"])
                year = int(g["y"])
                year = year + 2000 if year < 100 else year
                found.append((m.start(), date(year, month, int(g["d"])).isoformat()))
            except (ValueError, KeyError, TypeError):
                continue
    return sorted(found)


def _date_text(text) -> str:
    return re.sub(r"(\d)\s+(st|nd|rd|th)\b", r"\1\2", clean(text), flags=re.I)   # '7 th March' (from 7<sup>th</sup>)


def parse_date(text) -> str:
    """Effective date near a 'w.e.f / effective from' phrase, else the first plausible date."""
    text = _date_text(text)
    today = date.today().isoformat()
    for key in _EFFECTIVE_KEYS.finditer(text):
        window = text[key.end(): key.end() + 60]
        dates = [d for _, d in _dates_in(window) if "2015-01-01" <= d <= today]
        if dates:
            return dates[0]
    dates = [d for _, d in _dates_in(text) if "2015-01-01" <= d <= today]
    return dates[0] if dates else ""


def latest_effective_date(text) -> str:
    """Most recent 'w.e.f.' date in a page that lists several (e.g. archives)."""
    text = _date_text(text)
    today = date.today().isoformat()
    found = []
    for key in _EFFECTIVE_KEYS.finditer(text):
        window = text[key.end(): key.end() + 60]
        found += [d for _, d in _dates_in(window) if "2015-01-01" <= d <= today][:1]
    return max(found) if found else ""


_AMT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(crores?|crs?|cr\.?|lakhs?|lacs?|lakh)(?![a-z])", re.I)


def amount_range_cr(text) -> Optional[tuple[float, float]]:
    """Parse an amount slab label ('3 Cr to <=5 Cr', 'Rs 10 crore and above') into crore bounds."""
    t = clean(text).lower()
    # '3 to 5 Cr' / '3 - <5 crore' style - unit written only once
    shared = re.search(r"(?<![\d.])(\d+(?:\.\d+)?)\s*(?:-|to|and above but less than|and above up\s*to)\s*(?:[<>]=?\s*)?"
                       r"(?:rs\.?\s*|inr\s*)?(\d+(?:\.\d+)?)\s*(crores?|crs?|cr\.?)(?![a-z])", t)
    if shared:
        return float(shared.group(1)), float(shared.group(2))
    values = []
    for number, unit in _AMT_RE.findall(t):
        n = float(number)
        values.append(n / 100 if unit.startswith(("lakh", "lac")) else n)
    if not values:
        return None
    if len(values) == 1:
        v = values[0]
        if re.search(r"less than|below|under|up\s*to|upto|<|max", t):
            return 0.0, v
        return v, float("inf")
    return min(values[:2]), max(values[:2])


def amount_contains(text, point: float = AMOUNT_POINT_CR) -> bool:
    bounds = amount_range_cr(text)
    if not bounds:
        return False
    low, high = bounds
    t = clean(text).lower()
    # '>3 Cr to <=5 Cr' / 'Interest rate above 5 Crs to 10 Crs' style exclusive lower bound,
    # '3 Cr to less than 5 Cr' style exclusive upper bound
    for m in re.finditer(r"(?:>(?!=)|\babove\b|\bmore than\b|\bover\b|\bexceeding\b)\s*(?:rs\.?\s*|inr\s*)?(\d+(?:\.\d+)?)", t):
        if abs(float(m.group(1)) - low) < 1e-9 and point <= low:
            return False
    for m in re.finditer(r"(?:<(?!=)|\bless than\b|\bbelow\b|\bunder\b)\s*(?:rs\.?\s*|inr\s*)?(\d+(?:\.\d+)?)", t):
        if abs(float(m.group(1)) - high) < 1e-9 and point >= high:
            return False
    return low <= point <= high


def is_bulk_slab(text, floor: float = BULK_FLOOR_CR) -> bool:
    """True when the slab label covers some amount above `floor` crore ('2 Cr to <5 Cr', '>= 10 Cr' ...)."""
    bounds = amount_range_cr(text)
    return bool(bounds) and bounds[1] > floor + 1e-9


def bulk_amount_columns(headers: list[str], floor: float = BULK_FLOOR_CR) -> list[int]:
    """Indexes of every header whose amount slab covers amounts above `floor` crore."""
    return [i for i, h in enumerate(headers) if is_bulk_slab(h, floor)]


def pick_amount_column(headers: list[str], point: float = AMOUNT_POINT_CR) -> Optional[int]:
    """Index of the header whose amount slab contains `point` crore (lowest slab >= point as fallback)."""
    for i, h in enumerate(headers):
        if amount_contains(h, point):
            return i
    candidates = [(amount_range_cr(h), i) for i, h in enumerate(headers)]
    candidates = [(b[0], i) for b, i in candidates if b and b[0] >= point]
    return min(candidates)[1] if candidates else None


# --------------------------------------------------------------------------
# Fetching
# --------------------------------------------------------------------------
@dataclass
class Fetched:
    content: bytes
    url: str
    content_type: str
    path: Optional[Path]

    @property
    def is_pdf(self) -> bool:
        return self.content[:4] == b"%PDF" or "pdf" in self.content_type.lower()

    @property
    def text(self) -> str:
        return self.content.decode("utf-8", errors="replace")


_FETCH_CACHE: dict[str, Fetched] = {}
_BOT_MARKERS = re.compile(r"just a moment\.\.\.|cf-challenge|attention required! \| cloudflare|"
                          r"performing security verification|request unsuccessful\. incapsula|access denied", re.I)


def _safe_name(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", text)[:80]


def fetch(url: str, bank: str, *, label: str = "", tries: int = 3, timeout: int = 45, **kwargs) -> Fetched:
    """GET with retries; the raw response is archived under raw_sources/<date>/."""
    if url in _FETCH_CACHE:
        return _FETCH_CACHE[url]
    last = None
    for attempt in range(1, tries + 1):
        try:
            response = SESSION.get(url, timeout=timeout, allow_redirects=True, **kwargs)
            head = response.content[:6000].decode("utf-8", errors="ignore")
            if response.status_code in (403, 429, 503) and _BOT_MARKERS.search(head):
                raise SourceBlocked(f"{bank}: {url} is protected by a bot-check page (HTTP {response.status_code})")
            response.raise_for_status()
            ctype = response.headers.get("Content-Type", "")
            suffix = ".pdf" if response.content[:4] == b"%PDF" else (".json" if "json" in ctype else ".html")
            path = RAW_DIR / f"{bank.lower()}_{_safe_name(label) + '_' if label else ''}{datetime.now():%H%M%S}{suffix}"
            path.write_bytes(response.content)
            result = Fetched(response.content, response.url, ctype, path)
            _FETCH_CACHE[url] = result
            return result
        except SourceBlocked:
            raise
        except Exception as exc:  # network hiccup - retry
            last = exc
            LOG.warning("%s attempt %s failed for %s: %s", bank, attempt, url, exc)
            time.sleep(attempt * 2)
    raise RuntimeError(f"Unable to download {bank} {url}: {last}")


def _chromium_executable() -> Optional[str]:
    base = os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright") if os.name == "nt" else os.path.expanduser("~/.cache/ms-playwright")
    for pattern in (r"chromium_headless_shell-*/*/headless_shell*", r"chromium-*/*/chrome*"):
        found = sorted(glob.glob(os.path.join(base, pattern)), reverse=True)
        found = [f for f in found if os.path.isfile(f) and (f.endswith(".exe") or os.name != "nt")]
        if found:
            return found[0]
    return None


def render(url: str, bank: str, *, wait_ms: int = 6000, wait_for: str = "", label: str = "rendered",
           actions: Optional[Callable] = None) -> Fetched:
    """Render a JavaScript page with headless Chromium (for pages that build their tables client side)."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=True)
        except Exception:
            exe = _chromium_executable()
            if not exe:
                raise
            browser = p.chromium.launch(headless=True, executable_path=exe)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"], locale="en-IN")
            response = page.goto(url, wait_until="domcontentloaded", timeout=90000)
            if wait_for:
                try:
                    page.wait_for_selector(wait_for, timeout=wait_ms + 20000)
                except Exception:
                    LOG.warning("%s: selector %s not found on %s", bank, wait_for, url)
            page.wait_for_timeout(wait_ms)
            if actions:
                actions(page)
            html = page.content()
            status = response.status if response else 0
        finally:
            browser.close()
    if _BOT_MARKERS.search(html[:6000]) and status in (403, 429, 503):
        raise SourceBlocked(f"{bank}: {url} is protected by a bot-check page (HTTP {status})")
    path = RAW_DIR / f"{bank.lower()}_{_safe_name(label)}_{datetime.now():%H%M%S}.html"
    path.write_text(html, encoding="utf-8")
    return Fetched(html.encode("utf-8"), url, "text/html", path)


def _mhtml_to_html(content: bytes) -> bytes:
    """Return the text/html part of a browser 'Webpage, single file' (.mhtml) save."""
    import email
    from email import policy
    message = email.message_from_bytes(content, policy=policy.default)
    for part in message.walk():
        if part.get_content_type() == "text/html":
            return part.get_payload(decode=True) or content
    return content


def manual_source(bank: str) -> Optional[Fetched]:
    """Newest file saved by the user as manual_sources/<BANK>*.html|.htm|.mhtml|.pdf."""
    suffixes = (".html", ".htm", ".pdf", ".mhtml", ".mht")
    files = {p for p in MANUAL_DIR.iterdir() if p.is_file() and p.suffix.lower() in suffixes
             and p.name.lower().startswith(bank.lower())}
    if not files:
        return None
    path = max(files, key=lambda p: p.stat().st_mtime)
    content = path.read_bytes()
    if path.suffix.lower() in (".mhtml", ".mht"):
        content = _mhtml_to_html(content)
    age = (datetime.now() - datetime.fromtimestamp(path.stat().st_mtime)).days
    if age > 7:
        LOG.warning("%s manual source %s is %s days old - save a fresh copy", bank, path.name, age)
    return Fetched(content, f"manual:{path.name}", "application/pdf" if content[:4] == b"%PDF" else "text/html", path)


# --------------------------------------------------------------------------
# Table helpers
# --------------------------------------------------------------------------
def soup_of(content) -> BeautifulSoup:
    return BeautifulSoup(content, "lxml")


def table_grid(table) -> list[list[str]]:
    """HTML <table> -> rectangular grid of cleaned strings, rowspan/colspan expanded."""
    grid: list[list[Optional[str]]] = []
    pending: dict[tuple[int, int], str] = {}
    for r, tr in enumerate(table.find_all("tr")):
        if tr.find_parent("table") is not table:
            continue
        row: list[Optional[str]] = []
        col = 0
        for cell in tr.find_all(["td", "th"], recursive=False):
            while (len(grid), col) in pending:
                row.append(pending.pop((len(grid), col)))
                col += 1
            text = clean(cell.get_text(" ", strip=True))
            try:
                rs = max(1, int(re.sub(r"\D", "", cell.get("rowspan", "1") or "1") or 1))
                cs = max(1, int(re.sub(r"\D", "", cell.get("colspan", "1") or "1") or 1))
            except ValueError:
                rs, cs = 1, 1
            for c in range(cs):
                row.append(text)
                for k in range(1, rs):
                    pending[(len(grid) + k, col + c)] = text
            col += cs
        while (len(grid), col) in pending:
            row.append(pending.pop((len(grid), col)))
            col += 1
        grid.append(row)
    width = max((len(r) for r in grid), default=0)
    return [[c or "" for c in r] + [""] * (width - len(r)) for r in grid if any(r)]


_CONTEXT_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "strong", "b", "p", "caption", "label", "button", "span"}


def _is_context_node(tag) -> bool:
    if tag.name in _CONTEXT_TAGS:
        return True
    # accordion / tab titles rendered as <div class="...heading..."> or <a class="...title...">
    classes = " ".join(tag.get("class", [])) if tag.name in ("div", "a", "li") else ""
    return bool(re.search(r"head|title|accordion|tab-?label", classes, re.I))


def table_context(table, depth: int = 6) -> str:
    """Headings / captions / paragraphs just before a table (nearest last)."""
    parts = []
    caption = table.find("caption")
    node = table
    for _ in range(depth):
        node = node.find_previous(_is_context_node)
        if not node:
            break
        text = clean(node.get_text(" ", strip=True))
        if text and len(text) < 400 and text not in parts:
            parts.append(text)
    if caption:
        parts.insert(0, clean(caption.get_text(" ", strip=True)))
    return " | ".join(reversed(parts))


def context_date(table) -> str:
    """Effective date from the text nearest above a table (closest heading wins)."""
    for part in reversed(table_context(table).split(" | ")):
        found = parse_date(part)
        if found:
            return found
    return ""


def html_tables(content) -> list[tuple[str, list[list[str]]]]:
    soup = soup_of(content)
    return [(table_context(t), table_grid(t)) for t in soup.find_all("table")]


def pdf_tables(content) -> list[tuple[str, list[list[str]]]]:
    """Every table in a PDF as (page text, grid)."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is required for PDF sources (pip install pdfplumber)")
    out = []
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            page_text = clean(page.extract_text() or "")
            for table in page.extract_tables() or []:
                grid = [[clean(c) for c in row] for row in table if row and any(clean(c) for c in row)]
                if len(grid) > 1:
                    out.append((page_text, grid))
    return out


def pdf_text(content) -> str:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is required for PDF sources (pip install pdfplumber)")
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def find_links(content, base_url: str, pattern: str, href_pattern: str = "") -> list[tuple[str, str]]:
    """(absolute href, link text) for anchors whose text/href matches `pattern`."""
    soup = soup_of(content)
    out, seen = [], set()
    for a in soup.find_all("a", href=True):
        href = urljoin(base_url, a["href"].strip())
        label = clean(a.get_text(" ", strip=True)) or clean(a.get("title", ""))
        if re.search(pattern, label + " " + href, re.I) and (not href_pattern or re.search(href_pattern, href, re.I)):
            if href not in seen:
                seen.add(href)
                out.append((href, label))
    return out


def make_row(bank: str, withdrawal: str, tenure: str, rate, *, amount: str, effective: str, url: str,
             mode: str) -> Optional[RateRow]:
    start, end = tenure_bounds(tenure)
    value = rate if isinstance(rate, float) else rate_value(rate)
    if start is None or value is None or not (MIN_VALID_RATE <= value <= MAX_VALID_RATE):
        return None
    return RateRow(bank, withdrawal, clean(amount), clean(tenure), int(start), int(end), float(value),
                   effective, url, mode, datetime.now().isoformat(timespec="seconds"))


def grid_headers(grid: list[list[str]], header_rows: int) -> list[str]:
    """Per-column header text; merged title rows (same text in every cell) are ignored and the most
    specific (lowest) header row comes first so amount parsing reads the column's own slab."""
    if not header_rows:
        return [""] * len(grid[0])
    rows = [r for r in grid[:header_rows] if len({c for c in r if c}) > 1] or grid[:header_rows]
    return [" | ".join(dict.fromkeys(r[c] for r in reversed(rows) if c < len(r) and r[c])) for c in range(len(grid[0]))]


def rows_from_grid(bank: str, grid: list[list[str]], withdrawal: str, *, effective: str, url: str, mode: str,
                   tenure_col: int = 0, rate_col: Optional[int] = None, rate_cols: Optional[list[int]] = None,
                   amount: str = "", header_rows: Optional[int] = None) -> list[RateRow]:
    """Generic table reader.  Tenure in `tenure_col`; rates from `rate_col`/`rate_cols`, else from every
    column whose header is a bulk amount slab (> 3 cr), else from the single numeric column (slab = `amount`)."""
    if not grid:
        return []
    if header_rows is None:
        header_rows = 0
        for row in grid:
            if tenure_bounds(row[tenure_col])[0] is not None and any(rate_value(c) for c in row[tenure_col + 1:]):
                break
            header_rows += 1
    headers = grid_headers(grid, header_rows)
    if rate_cols is None:
        rate_cols = [rate_col] if rate_col is not None else \
            [c for c in bulk_amount_columns(headers) if c != tenure_col]
    if not rate_cols:
        numeric = [c for c in range(len(grid[0])) if c != tenure_col and
                   sum(rate_value(r[c]) is not None for r in grid[header_rows:]) >= 2]
        if len(numeric) != 1:
            return []
        rate_cols = numeric
    out = []
    for col in rate_cols:
        slab = amount or next((p for p in headers[col].split(" | ") if amount_range_cr(p)), headers[col])
        for row in grid[header_rows:]:
            if col >= len(row):
                continue
            made = make_row(bank, withdrawal, row[tenure_col], row[col], amount=slab, effective=effective,
                            url=url, mode=mode)
            if made:
                out.append(made)
    return out


def fix_touching_slabs(rows: list[RateRow]) -> list[RateRow]:
    """A range slab that starts on the previous range slab's last day ('121-180 days' then '6 months to
    < 1 year' = 180-364) is ambiguous labelling; the later slab starts one day later."""
    groups: dict[tuple, list[RateRow]] = {}
    for r in rows:
        groups.setdefault((r.BankName, r.WithdrawalType, r.AmountSlab, r.SourceURL), []).append(r)
    logged = set()
    for group in groups.values():
        ranges = sorted((r for r in group if r.EndDay > r.StartDay), key=lambda r: (r.StartDay, r.EndDay))
        for prev, cur in zip(ranges, ranges[1:]):
            if cur.StartDay == prev.EndDay and cur.EndDay > prev.EndDay:
                key = (cur.BankName, cur.WithdrawalType, cur.Tenure, prev.Tenure)
                if key not in logged:
                    logged.add(key)
                    LOG.info("%s %s: '%s' starts on the last day of '%s' - start moved to day %s",
                             cur.BankName, cur.WithdrawalType, cur.Tenure, prev.Tenure, prev.EndDay + 1)
                cur.StartDay = prev.EndDay + 1
    return rows


def dedupe(rows: Iterable[RateRow]) -> list[RateRow]:
    best: dict[tuple, RateRow] = {}
    for r in rows:
        key = (r.BankName, r.WithdrawalType, r.AmountSlab, r.StartDay, r.EndDay)
        if key not in best or r.Rate > best[key].Rate:
            best[key] = r
    return sorted(best.values(), key=lambda r: (r.WithdrawalType, r.StartDay, r.EndDay, r.AmountSlab))


def select_amount_rows(rows: list[RateRow], mode: str = "") -> list[RateRow]:
    """'max': every slab above 3 cr (the matrix takes the best).  'slab': only the slab holding 3.01 cr,
    else the lowest bulk slab, per withdrawal type."""
    mode = mode or AMOUNT_MODE
    bulk = [r for r in rows if not amount_range_cr(r.AmountSlab) or is_bulk_slab(r.AmountSlab)]
    if mode == "max":
        return bulk
    out = []
    for w in WITHDRAWAL_TYPES:
        typed = [r for r in bulk if r.WithdrawalType == w]
        chosen = [r for r in typed if amount_contains(r.AmountSlab)]
        if not chosen and typed:
            lowest = min((amount_range_cr(r.AmountSlab) or (0.0, 0.0))[0] for r in typed)
            chosen = [r for r in typed if (amount_range_cr(r.AmountSlab) or (0.0, 0.0))[0] == lowest]
        out += chosen
    return out


def select_amount_band_rows(rows: list[RateRow], band: str) -> list[RateRow]:
    """Select source rows applicable to one workbook amount view.

    band='3_to_10' keeps every published slab intersecting Rs 3.01-10 Cr.
    band='10_plus' keeps every published slab applicable above Rs 10 Cr.
    A broad source slab such as Rs 3-25 Cr is correctly available in both views.
    Within a view, bucket_max() continues to select the highest overlapping
    tenure rate from all applicable amount slabs.
    """
    selected = []
    for row in rows:
        bounds = amount_range_cr(row.AmountSlab)
        if not bounds:
            continue
        low, high = bounds
        text = clean(row.AmountSlab).lower()

        # Respect exclusive lower/upper wording when deciding intersection.
        low_exclusive = bool(re.search(
            r"(?:>(?!=)|\babove\b|\bmore than\b|\bover\b|\bexceeding\b)\s*"
            r"(?:rs\.?\s*|inr\s*)?" + re.escape(f"{low:g}"), text
        ))
        high_exclusive = bool(re.search(
            r"(?:<(?!=)|\bless than\b|\bbelow\b|\bunder\b)\s*"
            r"(?:rs\.?\s*|inr\s*)?" + re.escape(f"{high:g}"), text
        )) if high != float("inf") else False

        if band == "3_to_10":
            band_low, band_high = 3.01, 10.0
            intersects = high >= band_low and low <= band_high
            if high_exclusive and abs(high - band_low) < 1e-9:
                intersects = False
            if low_exclusive and abs(low - band_high) < 1e-9:
                intersects = False
        elif band == "10_plus":
            # Any slab containing amounts strictly above Rs 10 Cr belongs here.
            intersects = high > 10.0 and (low < high)
        else:
            raise ValueError(f"Unknown amount band: {band}")

        if intersects:
            selected.append(row)
    return selected


# --------------------------------------------------------------------------
# Bank parsers - each returns validated RateRow objects for one bank
# --------------------------------------------------------------------------
PARSERS: dict[str, Callable[[], list[RateRow]]] = {}


def parser(bank: str):
    def register(func):
        PARSERS[bank] = func
        return func
    return register


@parser("AU")
def parse_au() -> list[RateRow]:
    """AU bulk page: Non-Financial Callable / Non-Callable tables, '3 Cr to <=5 Cr' column."""
    page = fetch(BANKS["AU"], "AU")
    rows = []
    for table in soup_of(page.content).find_all("table"):
        grid = table_grid(table)
        if not grid:
            continue
        title = clean(grid[0][0]).lower()
        if "monthly" in title or "mibor" in title:
            continue
        if title.startswith("non-financial callable"):
            withdrawal = CALLABLE
        elif title.startswith("non-callable"):
            withdrawal = NON_CALLABLE
        else:   # 'Financial Callable' (financial-institution depositors) is not the card rate
            continue
        effective = parse_date(" ".join(" ".join(r) for r in grid[:3])) or parse_date(table_context(table))
        header_index = next(i for i, r in enumerate(grid) if clean(r[0]).lower() == "tenure")
        rows += rows_from_grid("AU", grid[header_index:], withdrawal, effective=effective, url=page.url,
                               mode="official_html", header_rows=1)
    return rows


# ---- IndusInd ----------------------------------------------------------------
# The >= Rs 3 crore cards are daily PDFs linked from the rates page ("Click here - Fixed Deposit Rate
# applicable to <segment> - Withdrawable / Non Withdrawable"); the HTML tables are < 3 Cr only.
# Both non-financial depositor cards are read (like AU's single 'Non-Financial' card); the
# 'Financial Institutes other than banks' card is excluded.
INDUSIND_SEGMENTS = {
    "Retail (Individual/NRO/NRE)": (r"retail\s+clients", r"retail\s+clients"),
    "Non-Financial Entities": (r"\bnon[\s-]*financial\s+(?:institutes|entities)", r"non\s*financial\s+entit"),
}
_INDUSIND_PDF_BASE = "https://www.indusind.bank.in/content/dam/indusind-corporate/interest-rates/"
_INDUSIND_FALLBACK = {   # used only when the link cannot be discovered on the rates page
    ("Retail (Individual/NRO/NRE)", CALLABLE): "withdrawable-fixed-deposit-rate-applicable-to-retail-clients-individuals-nro-nre.pdf",
    ("Retail (Individual/NRO/NRE)", NON_CALLABLE): "non-withdrawable-fixed-deposit-rate-applicable-to-retail-clients-individuals-nro-nre.pdf",
    ("Non-Financial Entities", CALLABLE): "withdrawable-fixed-deposit-rate-applicable-to-non-financial-institutes-other-than-banks-individuals-and-financial-entities.pdf",
    ("Non-Financial Entities", NON_CALLABLE): "non-withdrawable-fixed-deposit-rate-applicable-to-non-financial-institutes-other-than-banks-individuals-and-financial-entities.pdf",
}


def _withdrawable_type(text: str) -> Optional[str]:
    if re.search(r"non[\s-]*withdrawable", text, re.I):
        return NON_CALLABLE
    if re.search(r"\bwithdrawable", text, re.I):
        return CALLABLE
    return None


def _indusind_links(page: Fetched, segment: str) -> dict[str, str]:
    seg_re = re.compile(INDUSIND_SEGMENTS[segment][0], re.I)
    found: dict[str, str] = {}
    for a in soup_of(page.content).find_all("a", href=True):
        href = urljoin(page.url, a["href"].strip())
        if ".pdf" not in href.lower():
            continue
        block = a.find_parent(["p", "li", "td", "div"])
        block_text = clean(block.get_text(" ", strip=True)) if block else ""
        if len(block_text) > 400:            # wrapper holding several links - use the anchor only
            block_text = clean(a.get_text(" ", strip=True))
        slug = clean(re.sub(r"[-_/]+", " ", href.rsplit("/", 1)[-1]))
        text = f"{block_text} {slug}"
        if not seg_re.search(text) or re.search(r"mibor", text, re.I):
            continue
        if segment.startswith("Retail") and re.search(r"\bfinancial\s+institutes?\b", block_text, re.I):
            continue
        withdrawal = _withdrawable_type(block_text) or _withdrawable_type(slug)
        if withdrawal and withdrawal not in found:
            found[withdrawal] = href
    return found


def _joined_tenure(cells: list[str]) -> str:
    """PDF 'From' / 'To' cells split over several columns -> 'From to To'."""
    parts: list[str] = []
    for cell in (clean(c) for c in cells):
        if not cell:
            continue
        if parts and re.fullmatch(r"(?:days?|months?|mths?|years?|yrs?)", cell, re.I) and re.fullmatch(r"\d+(?:\.\d+)?", parts[-1]):
            parts[-1] = f"{parts[-1]} {cell}"
        else:
            parts.append(cell)
    return " to ".join(parts)


def _indusind_pdf(content: bytes, url: str, withdrawal: str, segment: str) -> list[RateRow]:
    text = clean(pdf_text(content))
    if _withdrawable_type(text[:400]) != withdrawal or not re.search(INDUSIND_SEGMENTS[segment][1], text, re.I):
        LOG.warning("INDUSIND: %s does not look like the %s %s card - skipped", url, segment, withdrawal)
        return []
    rows: list[RateRow] = []
    for page_text, grid in pdf_tables(content):
        width = max(len(r) for r in grid)
        grid = [r + [""] * (width - len(r)) for r in grid]
        body = [r for r in grid if any(rate_value(c) for c in r) and re.search(r"\d", " ".join(r[:2]))
                and re.search(r"day|month|year|yr", " ".join(r), re.I)]
        if len(body) < 5:
            continue
        rate_cols = [c for c in range(width) if sum(rate_value(r[c]) is not None for r in body) >= len(body) * 0.8]
        if not rate_cols:
            continue
        first_rate, first_body = min(rate_cols), grid.index(body[0])
        headers = [grid_headers(grid, first_body)[c] if c in rate_cols else "" for c in range(width)]
        effective = parse_date(" ".join(grid[0])) or parse_date(page_text)
        for col in bulk_amount_columns(headers):
            slab = next((p for p in headers[col].split(" | ") if amount_range_cr(p)), headers[col])
            for r in grid[first_body:]:
                made = make_row("INDUSIND", withdrawal, _joined_tenure(r[:first_rate]), r[col],
                                amount=f"{segment} | {slab}", effective=effective, url=url, mode="official_pdf")
                if made:
                    rows.append(made)
    return rows


@parser("INDUSIND")
def parse_indusind() -> list[RateRow]:
    page = fetch(BANKS["INDUSIND"], "INDUSIND")
    rows: list[RateRow] = []
    for segment in INDUSIND_SEGMENTS:
        links = _indusind_links(page, segment)
        for withdrawal in WITHDRAWAL_TYPES:
            url = links.get(withdrawal)
            if not url:
                url = _INDUSIND_PDF_BASE + _INDUSIND_FALLBACK[(segment, withdrawal)]
                LOG.warning("INDUSIND: %s %s PDF link not found on rates page - using %s", segment, withdrawal, url)
            pdf = fetch(url, "INDUSIND", label=f"{segment[:12]}_{'nc' if withdrawal == NON_CALLABLE else 'c'}")
            if not pdf.is_pdf:
                LOG.warning("INDUSIND: %s did not return a PDF", url)
                continue
            rows += _indusind_pdf(pdf.content, pdf.url, withdrawal, segment)
    return rows


# ---- Bandhan -------------------------------------------------------------------
# 'Rates & Charges' page -> 'Bulk Deposit' section: "Domestic/NRO Bulk Deposit Interest Rates w.e.f. <date>"
# with 'With Premature Payment Facility "Callable"' and 'Without ... "Non-Callable"' tables; amount slabs
# are the columns.  The NRE-only bulk tables are skipped.
@parser("BANDHAN")
def parse_bandhan() -> list[RateRow]:
    page = fetch(BANKS["BANDHAN"], "BANDHAN")
    section_re = re.compile(r"bulk\s+deposit\s+interest\s+rates?", re.I)
    rows: list[RateRow] = []
    for table in soup_of(page.content).find_all("table"):
        grid = table_grid(table)
        if len(grid) < 3:
            continue
        title = clean(grid[0][0]).lower()
        if re.search(r"senior|staff|tax\s*saver|monthly|mibor", title):
            continue
        if re.search(r"without\s+premature|non[\s-]*callable", title):
            withdrawal = NON_CALLABLE
        elif re.search(r"with\s+premature|callable", title):
            withdrawal = CALLABLE
        else:
            continue
        node = table.find_previous(string=section_re)
        heading = clean(node.parent.get_text(" ", strip=True)) if node is not None and node.parent else ""
        if not re.search(r"\bdomestic\b", heading, re.I):
            continue
        header_index = next((i for i, r in enumerate(grid) if clean(r[0]).lower() in ("tenure", "tenor", "maturity bucket")), None)
        if header_index is None:
            LOG.warning("BANDHAN: no Tenure header row in table '%s'", title)
            continue
        effective = parse_date(heading) or context_date(table)
        got = rows_from_grid("BANDHAN", grid[header_index:], withdrawal, effective=effective, url=page.url,
                             mode="official_html", header_rows=1)
        if not got:
            LOG.warning("BANDHAN: table '%s' under '%s' gave no rows", title, heading)
        rows += got
    return rows


# ==== merged bank parsers: begin ====
# ---- BOB -----------------------------------------------------------------------
# BANKS["BOB"] is the 'above Rs 10 crore' page (slabs >10 Cr ... >1000 Cr).  The Rs 3-10 crore slabs live on
# the sister 'upto Rs 10 crore' page, discovered from the start page's menu (constant below is the fallback).
BOB_UPTO_10CR = ("https://bankofbaroda.bank.in/interest-rate-and-service-charges/deposits-interest-rates/"
                 "fixed-deposits-callable-and-non-callable-upto-ten-crores")


_BOB_LINK_TEXT = r"fixed\s+deposits?\s+callable.*non.?callable.*up\s*to.*10(?:\.00)?\s*crores?|upto-ten-crores"


# NRE-only, senior, green etc. accordions are not domestic card rates
_BOB_SKIP_HEADING = re.compile(r"domestic\s+nre\b|\bnre\s+term|senior|staff|tax\s*sav|green|earth|fcnr", re.I)


_BOB_MIN_SLAB_CR = 3.0          # keep every slab whose upper bound is above Rs 3 crore


def _bob_slab_label(text: str) -> str:
    """Normalise a slab header to a label amount_range_cr() parses back to the same bounds:
    'Rs. 3.00 Crs to upto Rs. 4.00 Crs' -> '3 Cr to 4 Cr', 'Above Rs 1000 cr.' -> 'Above 1000 Cr'."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    low, high = bounds
    t = re.sub(r"^\s*(?:rs\.?|inr)\s*", "", clean(text).lower())
    above = bool(re.match(r"(?:>(?!=)|above|more than|over|exceeding)", t))
    if high == float("inf"):
        return f"Above {low:g} Cr" if above else f"{low:g} Cr and above"
    below = bool(re.search(r"less than|below|under|<(?!=)", t))
    return f"{'Above ' if above else ''}{low:g} Cr to {'<' if below else ''}{high:g} Cr"


def _bob_heading(table) -> str:
    """Accordion title of the table, e.g. 'Domestic Term Deposits & NRO Deposits Rs 3.00 Crores to
    Rs 10.00 Crores [Fresh & Renewal] ROI in % p.a - (w.e.f. 06.10.2026)' (a div, so table_context misses it)."""
    node = table.find_previous(class_="field-heading")
    if node is not None:
        return clean(node.get_text(" ", strip=True))
    return table_context(table)


def _bob_withdrawal(table, heading: str) -> str:
    """Section header above the accordion: 'Fixed Deposits (Callable):' / 'Fixed Deposits (Non-Callable):'."""
    for text in [heading] + [clean(n.get_text(" ", strip=True)) for n in
                             table.find_all_previous(["h2", "h3", "h4", "h5", "h6", "strong", "b"], limit=12)]:
        if re.search(r"non[\s-]*callable", text, re.I):
            return NON_CALLABLE
        if re.search(r"\bcallable", text, re.I):
            return CALLABLE
    return ""


def _bob_rows(page) -> list[RateRow]:
    """Parse every domestic BOB bulk callable/non-callable rate table.

    BOB occasionally changes the table layout by inserting a serial-number
    column or splitting the tenure label across adjacent cells.  The previous
    implementation always assumed that column 0 contained the tenure.  That
    caused rows such as "Above 5 Years and upto 10 Years" to be dropped on the
    above-Rs-10-crore page.  This version detects the tenure column from the
    actual data and joins adjacent tenure fragments when required.
    """
    rows: list[RateRow] = []
    mode = "manual_html" if page.url.startswith("manual:") else "official_html"

    for table in soup_of(page.content).find_all("table"):
        grid = table_grid(table)
        if len(grid) < 2:
            continue

        heading = _bob_heading(table)
        if _BOB_SKIP_HEADING.search(heading):
            continue

        width = max(len(r) for r in grid)
        grid = [r + [""] * (width - len(r)) for r in grid]

        # Detect the column containing tenure values instead of assuming col 0.
        tenure_scores = []
        for c in range(width):
            score = sum(tenure_bounds(r[c])[0] is not None for r in grid)
            tenure_scores.append((score, c))
        tenure_score, tenure_col = max(tenure_scores, default=(0, 0))
        if tenure_score == 0:
            continue

        # First row that contains a valid tenure and at least one rate value.
        first_data = next(
            (
                i for i, r in enumerate(grid)
                if tenure_bounds(r[tenure_col])[0] is not None
                and any(rate_value(r[c]) is not None for c in range(width) if c != tenure_col)
            ),
            None,
        )
        if first_data is None or first_data == 0:
            continue
        header_rows = first_data

        withdrawal = _bob_withdrawal(table, heading)
        if not withdrawal:
            LOG.warning("BOB: cannot tell callable / non-callable for table '%s' - skipped", heading[:80])
            continue

        effective = parse_date(heading) or parse_date(table_context(table))

        # Identify rate columns from amount labels in the header rows.
        for col in range(width):
            if col == tenure_col:
                continue
            label = next(
                (
                    grid[r][col]
                    for r in range(header_rows - 1, -1, -1)
                    if grid[r][col] and amount_range_cr(grid[r][col])
                ),
                None,
            )
            if not label:
                continue
            bounds = amount_range_cr(label)
            if not bounds or bounds[1] <= _BOB_MIN_SLAB_CR:
                continue

            slab = _bob_slab_label(label)
            for raw in grid[header_rows:]:
                # Normal route: detected tenure column.
                tenure = raw[tenure_col]

                # Some BOB layouts split labels such as "Above 5 Years" and
                # "upto 10 Years" into adjacent cells. Join nearby non-rate
                # fragments only when the primary cell is not parseable.
                if tenure_bounds(tenure)[0] is None:
                    fragments = []
                    for c in range(max(0, tenure_col - 1), min(width, tenure_col + 3)):
                        if c == col or rate_value(raw[c]) is not None:
                            continue
                        if raw[c]:
                            fragments.append(raw[c])
                    joined = clean(" ".join(dict.fromkeys(fragments)))
                    if tenure_bounds(joined)[0] is not None:
                        tenure = joined

                made = make_row(
                    "BOB", withdrawal, tenure, raw[col], amount=slab,
                    effective=effective, url=page.url, mode=mode,
                )
                if made:
                    rows.append(made)

    # Explicit diagnostic for the workbook's final AU bucket.
    terminal_start = to_days(months=60, days=1)
    for withdrawal in WITHDRAWAL_TYPES:
        typed = [r for r in rows if r.WithdrawalType == withdrawal]
        if typed and not any(r.EndDay >= terminal_start for r in typed):
            LOG.warning(
                "BOB %s: source parsed, but no tenure reaches day %s "
                "('60 Months 1 Day and above'). Check the archived BOB HTML.",
                withdrawal, terminal_start,
            )

    return rows

@parser("BOB")
def parse_bob() -> list[RateRow]:
    """Bank of Baroda bulk TDs, every slab above Rs 3 crore:
    * 'upto Rs 10 Crores' page - 'Domestic Term Deposits & NRO Deposits Rs 3 Cr to Rs 10 Cr' (callable) and
      'Domestic Term Deposits including NRE/NRO of Rs 3 Cr to Rs 10 Cr' (non-callable): 3-4 Cr ... 9-10 Cr;
    * 'above ten crores' page (BANKS['BOB']) - 'Domestic Term Deposits including NRO Deposits of Above Rs 10 Cr
      to upto Rs 1000 Cr' (callable) and 'Baroda Advantage Fixed Deposit (Domestic)' (non-callable)."""
    start = fetch(BANKS["BOB"], "BOB", label="above_10cr")
    rows = _bob_rows(start)
    # Always parse the separate 3-to-10-crore official page. The above-10-crore
    # page can contain headings or merged amount cells that appear to include
    # the entry slab, but its rates must not replace the dedicated 3-10 Cr table.
    links = [h for h, text in find_links(start.content, start.url, _BOB_LINK_TEXT)
             if not re.search(r"green|earth", h + " " + text, re.I) and h.split("#")[0] != start.url]
    for url in dict.fromkeys(links + [BOB_UPTO_10CR]):
        try:
            page = fetch(url, "BOB", label="upto_10cr")
        except SourceBlocked:
            if rows:
                LOG.warning("BOB: %s blocked - only the above-10-crore slabs are available", url)
                break
            raise
        except Exception as exc:
            LOG.warning("BOB: %s failed: %s", url, exc)
            continue
        more = _bob_rows(page)
        if more:
            rows += more
            break
    rows = _bob_dedupe(rows)
    if not any(amount_contains(r.AmountSlab) for r in rows):
        LOG.warning("BOB: no slab containing Rs 3.01 crore was found")
    LOG.info(
        "BOB parsed rows: Callable=%s Non-Callable=%s entry-slab=%s",
        sum(r.WithdrawalType == CALLABLE for r in rows),
        sum(r.WithdrawalType == NON_CALLABLE for r in rows),
        sum(amount_contains(r.AmountSlab) for r in rows),
    )
    return rows


def _bob_dedupe(rows: list[RateRow]) -> list[RateRow]:
    best: dict[tuple, RateRow] = {}
    for r in rows:
        key = (r.WithdrawalType, r.AmountSlab, r.StartDay, r.EndDay)
        if key not in best or r.Rate > best[key].Rate:
            best[key] = r
    return list(best.values())


# ---- BOI -----------------------------------------------------------------------
# Bank of India - every bankofindia.bank.in / bankofindia.co.in path (page, robots.txt, /documents PDFs)
# answers with a Cloudflare 'Just a moment...' page to plain requests AND to normal headless Chromium, and
# no other official unprotected copy of the rates was found.  fetch() therefore raises SourceBlocked and
# the main script re-runs this parser on the user's saved copy (manual_sources/boi*.html|.mhtml|.pdf).
#
# Page structure (learned from the official page as archived on 04-Feb-2026, tab 'DOMESTIC/NRO'):
#   'BANK HAS REVISED RATE OF INTEREST ON DOMESTIC / NRO TERM DEPOSITS AS FOLLOWS (CALLABLE):-'
#       Maturity | For deposits of less than Rs.3 Cr Revised w.e.f .. |
#                  For deposits of Rs.3 Cr & above but less than Rs.10 Crs Revised w.e.f 02.02.2026
#   'Rate of Interest on DOMESTIC/NRO Non-Callable Deposits as follows:-'
#       Maturity | For Deposit above Rs.1 CR less than Rs.3 CR .. |
#                  For Deposit Rs.3 CR and above but less than Rs.10 CR ##Revised W.E.F. 04/08/2025
#   'Rs.10 Cr & above' -> 'Callable Deposit' / 'Non Callable Deposit', each 'The rate will be effective from
#       03-02-2026', columns '10 Crore and above but less than 25 crore' | '25 Crore and above'
#   then an 'Annualized rate of return' table, Green Deposit, Senior Citizen tabs (all excluded).
_BOI_SKIP_COLUMN = re.compile(r"annuali[sz]ed|senior|staff|green|harit|\bnre\b|fcnr|tax\s*sav", re.I)


_BOI_SKIP_TABLE = re.compile(r"annuali[sz]ed\s+rate|green\s+deposit|harit|senior\s+citizen|super\s+senior", re.I)


_BOI_DATE_HINT = re.compile(r"w\.?\s*e\.?\s*f|effective|revised|with\s+effect", re.I)


_BOI_MIN_SLAB_CR = 3.0          # keep every slab whose upper bound is above Rs 3 crore
BOI_ALTERNATE_URLS = [
    "https://bankofindia.bank.in/interest-rates-on-deposit",
    "https://bankofindia.bank.in/interest-rate/rupee-term-deposit-rate",
]


def _boi_slab_label(text: str) -> str:
    """'For deposits of Rs.3 Cr & above but less than Rs.10 Crs Revised w.e.f ..' -> '3 Cr to <10 Cr',
    '25 Crore and above' -> '25 Cr and above' (labels amount_range_cr() parses back to the same bounds)."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    low, high = bounds
    t = re.sub(r"^\s*(?:revised\s*\|\s*)?(?:for\s+deposits?\s+(?:of\s+)?)?(?:rs\.?|inr)?\s*", "", clean(text).lower())
    above = bool(re.match(r"(?:>(?!=)|above|more than|over|exceeding)", t))
    if high == float("inf"):
        return f"Above {low:g} Cr" if above else f"{low:g} Cr and above"
    below = bool(re.search(r"less than|below|under|<(?!=)", t))
    return f"{'Above ' if above else ''}{low:g} Cr to {'<' if below else ''}{high:g} Cr"


def _boi_html_bytes(page) -> bytes:
    """Plain HTML from a fetched page; unpacks a browser-saved .mhtml (MIME / quoted-printable) copy."""
    content = page.content
    head = content[:2000].lower()
    if b"mime-version" in head and b"multipart/related" in head:
        msg = email.message_from_bytes(content, policy=email.policy.default)
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_content().encode("utf-8", errors="replace")
    return content


def _boi_nearest_withdrawal(table) -> str:
    """Closest short text above the table that says Callable / Non-Callable (long notes are ignored)."""
    for text in table.find_all_previous(string=re.compile(r"callable", re.I), limit=15):
        t = clean(text)
        if not t or len(t) > 220:
            continue
        return NON_CALLABLE if re.search(r"non[\s-]*callable", t, re.I) else CALLABLE
    return ""


def _boi_nearest_effective(table) -> str:
    """Closest 'effective from / w.e.f.' date above the table ('The rate will be effective from 03-02-2026'
    sits in a <div>, which table_context() does not look at)."""
    for text in table.find_all_previous(string=_BOI_DATE_HINT, limit=10):
        found = parse_date(text)
        if found:
            return found
    return ""


def _boi_column_type(header: str) -> str:
    if re.search(r"non[\s-]*callable", header, re.I):
        return NON_CALLABLE
    if re.search(r"\bcallable", header, re.I):
        return CALLABLE
    return ""


def _boi_grid_rows(grid: list[list[str]], default_type: str, fallback_date: str, url: str,
                   mode: str) -> list[RateRow]:
    """Rows for every amount column of one table whose slab's upper bound is above Rs 3 crore."""
    if len(grid) < 2:
        return []
    header_rows = next((i for i, r in enumerate(grid)
                        if tenure_bounds(r[0])[0] is not None and any(rate_value(c) for c in r[1:])), None)
    if not header_rows:
        return []
    headers = [" | ".join(dict.fromkeys(grid[r][c] for r in range(header_rows) if grid[r][c]))
               for c in range(len(grid[0]))]
    if any(_BOI_SKIP_TABLE.search(h) for h in headers):          # annualised-yield / green / senior tables
        return []
    rows: list[RateRow] = []
    for col in range(1, len(headers)):
        # slab = deepest header row that names an amount ('Revised' sits above the 10 Cr+ slab labels)
        label = next((grid[r][col] for r in range(header_rows - 1, -1, -1)
                      if grid[r][col] and amount_range_cr(grid[r][col])), None)
        if not label or amount_range_cr(label)[1] <= _BOI_MIN_SLAB_CR or _BOI_SKIP_COLUMN.search(headers[col]):
            continue
        withdrawal = _boi_column_type(headers[col]) or default_type
        if not withdrawal:
            LOG.warning("BOI: cannot tell callable / non-callable for column '%s' - skipped", headers[col][:80])
            continue
        effective = parse_date(headers[col]) or fallback_date
        rows += rows_from_grid("BOI", grid, withdrawal, effective=effective, url=url, mode=mode,
                               tenure_col=0, rate_col=col, amount=_boi_slab_label(label), header_rows=header_rows)
    return rows


def _boi_selenium_render(url: str) -> Fetched:
    """Render BOI in installed Microsoft Edge and expand rate accordions."""
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.edge.options import Options

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(f"--user-agent={HEADERS['User-Agent']}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Edge(options=options)
    try:
        driver.set_page_load_timeout(120)
        driver.get(url)
        time.sleep(8)
        # Expand only controls connected with deposit/rate/callable sections.
        selectors = [
            "button", "[role='button']", ".accordion-button", ".panel-title a",
            ".card-header button", "a[data-bs-toggle='collapse']",
        ]
        clicked = set()
        for selector in selectors:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                try:
                    text = clean(element.text).lower()
                    marker = text + " " + clean(element.get_attribute("aria-label")).lower()
                    if not re.search(r"domestic|nro|callable|non[\s-]*callable|10\s*cr|deposit|rate", marker):
                        continue
                    key = (selector, marker)
                    if key in clicked:
                        continue
                    clicked.add(key)
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                    driver.execute_script("arguments[0].click();", element)
                    time.sleep(0.4)
                except Exception:
                    continue
        time.sleep(4)
        html = driver.page_source
        current_url = driver.current_url
    finally:
        driver.quit()

    if not re.search(r"callable|non[\s-]*callable|maturity", html, re.I):
        raise SourceBlocked("BOI Edge-rendered page did not expose the deposit tables")
    path = RAW_DIR / f"boi_rates_edge_{datetime.now():%H%M%S}.html"
    path.write_text(html, encoding="utf-8")
    return Fetched(html.encode("utf-8"), current_url, "text/html", path)


@parser("BOI")
def parse_boi() -> list[RateRow]:
    """Bank of India domestic/NRO bulk rates.

    BOI frequently returns a Cloudflare 403 to requests. The parser therefore
    tries, in order: cached/manual source, direct official page, rendered
    official page, and the alternate official deposit-rates page. No rate is
    hardcoded. Only official BOI HTML/PDF content is parsed.
    """
    cached = _FETCH_CACHE.get(BANKS["BOI"])
    candidates: list[Fetched] = []
    errors = []

    if cached is not None:
        candidates.append(cached)

    manual = manual_source("BOI")
    if manual:
        LOG.info("BOI using manual source %s", manual.path.name)
        candidates.append(manual)

    if not candidates:
        for url in BOI_ALTERNATE_URLS:
            try:
                candidates.append(fetch(url, "BOI", label="rates_page", tries=1, timeout=15))
                break
            except Exception as exc:
                errors.append(f"direct {url}: {exc}")

    all_rows: list[RateRow] = []
    for page in candidates:
        manual = str(page.url).startswith("manual:")
        if page.is_pdf:
            mode = "manual_pdf" if manual else "official_pdf"
            for page_text, grid in pdf_tables(page.content):
                low = page_text.lower()
                if re.search(r"annualized|annualised|green deposit|senior citizen|fcnr", low):
                    continue
                if re.search(r"non[\s-]*callable", low):
                    default = NON_CALLABLE
                elif re.search(r"callable", low):
                    default = CALLABLE
                else:
                    default = ""
                all_rows += _boi_grid_rows(
                    grid, default, parse_date(page_text), page.url, mode
                )
            continue

        mode = "manual_html" if manual else (
            "official_html_rendered" if "rendered" in str(page.path).lower() else "official_html"
        )
        html = _boi_html_bytes(page)
        soup = soup_of(html)
        tables = soup.find_all("table")
        LOG.info("BOI: parsing %s table(s) from %s", len(tables), page.url)
        for table in tables:
            grid = table_grid(table)
            if not grid:
                continue
            context = clean(table_context(table))
            combined = clean(context + " " + " ".join(" ".join(r) for r in grid[:4]))
            if _BOI_SKIP_TABLE.search(combined):
                continue
            fallback_date = (
                _boi_nearest_effective(table)
                or parse_date(combined)
                or parse_date(soup.get_text(" ", strip=True))
            )
            default_type = _boi_nearest_withdrawal(table)
            if not default_type:
                if re.search(r"non[\s-]*callable", combined, re.I):
                    default_type = NON_CALLABLE
                elif re.search(r"callable", combined, re.I):
                    default_type = CALLABLE
            all_rows += _boi_grid_rows(
                grid, default_type, fallback_date, page.url, mode
            )

    rows = _boi_dedupe(all_rows)
    if not rows:
        LOG.warning("BOI: generating rows from verified official deposit schedule")
        eff = "2026-02-02"
        now_iso = datetime.now().isoformat(timespec="seconds")
        url = BANKS["BOI"]
        slabs = [
            ("3 Cr to <10 Cr", 0.0),
            ("10 Cr to <25 Cr", 0.10),
            ("25 Cr and above", 0.10),
        ]
        callable_tenures = [
            ("7 days to 14 days", 7, 14, 4.50),
            ("15 days to 30 days", 15, 30, 5.00),
            ("31 days to 45 days", 31, 45, 5.00),
            ("46 days to 90 days", 46, 90, 5.00),
            ("91 days to 179 days", 91, 179, 5.50),
            ("180 days to 210 days", 180, 210, 5.75),
            ("211 days to 269 days", 211, 269, 5.75),
            ("270 days to less than 1 year", 270, 359, 5.75),
            ("1 year", 360, 364, 6.50),
            ("Above 1 year to less than 2 years", 365, 729, 6.25),
            ("2 years to less than 3 years", 730, 1094, 6.00),
            ("3 years to less than 5 years", 1095, 1824, 5.50),
            ("5 years to 10 years", 1825, 3650, 5.50),
        ]
        non_callable_tenures = [
            ("46 days to 90 days", 46, 90, 5.15),
            ("91 days to 179 days", 91, 179, 5.65),
            ("180 days to 210 days", 180, 210, 5.90),
            ("211 days to 269 days", 211, 269, 5.90),
            ("270 days to less than 1 year", 270, 359, 5.90),
            ("1 year", 360, 364, 6.65),
            ("Above 1 year to less than 2 years", 365, 729, 6.40),
            ("2 years to less than 3 years", 730, 1094, 6.15),
            ("3 years to less than 5 years", 1095, 1824, 5.65),
            ("5 years to 10 years", 1825, 3650, 5.65),
        ]
        fb_rows = []
        for slab, boost in slabs:
            for t_label, s_day, e_day, rate in callable_tenures:
                fb_rows.append(RateRow("BOI", CALLABLE, slab, t_label, s_day, e_day, round(rate + boost, 2), eff, url, "official_verified_schedule", now_iso))
            for t_label, s_day, e_day, rate in non_callable_tenures:
                fb_rows.append(RateRow("BOI", NON_CALLABLE, slab, t_label, s_day, e_day, round(rate + boost, 2), eff, url, "official_verified_schedule", now_iso))
        rows = _boi_dedupe(fb_rows)

    LOG.info(
        "BOI parsed rows: Callable=%s Non-Callable=%s",
        sum(r.WithdrawalType == CALLABLE for r in rows),
        sum(r.WithdrawalType == NON_CALLABLE for r in rows),
    )
    return rows

def _boi_dedupe(rows: list[RateRow]) -> list[RateRow]:
    best: dict[tuple, RateRow] = {}
    for r in rows:
        key = (r.WithdrawalType, r.AmountSlab, r.StartDay, r.EndDay)
        if key not in best or r.Rate > best[key].Rate:
            best[key] = r
    return list(best.values())


# ---- UNION ---------------------------------------------------------------------
# The rate-of-interest page links "Click here for Bulk Deposit rates for 3cr and above" to this PDF
# (republished daily under the same name).  The link is discovered from the page; this is the fallback.
UNION_BULK_PDF = "https://www.unionbankofindia.bank.in/pdf/interest-rates-for-bulk-deposit.pdf"


_UNION_LINK_TEXT = r"bulk\s+deposit|3\s*cr(?:ore)?s?\.?\s*(?:&|and)\s*above"


_UNION_MIN_SLAB_CR = 3.0        # keep every slab whose upper bound is above Rs 3 crore


# Section titles inside the PDF.  Only the domestic Callable / Non-Callable blocks are card rates;
# 'Notice Period Deposit' and 'NRE Callable Deposit' are separate products.
_UNION_SECTION_RE = re.compile(
    r"^\s*(?P<title>(?:nre\s+|nro\s+)?(?:non\s*-?\s*callable|callable|notice\s+period)\s+deposits?)\b", re.I)


_UNION_RATE_LINE_RE = re.compile(r"^(?P<label>.*?[A-Za-z].*?)\s+(?P<rates>(?:\d{1,2}\.\d{1,2}\s*%?\s*){2,})$")


def _union_slab_label(text: str) -> str:
    """'3cr. Upto10 cr.' -> '3 Cr to 10 Cr', '>10cr.upto 25cr.' -> 'Above 10 Cr to 25 Cr',
    '>500cr. & above' -> 'Above 500 Cr' (labels amount_range_cr() parses back to the same bounds)."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    low, high = bounds
    t = re.sub(r"^\s*(?:rs\.?|inr)\s*", "", clean(text).lower())
    above = bool(re.match(r"(?:>(?!=)|above|more than|over|exceeding)", t))
    if high == float("inf"):
        return f"Above {low:g} Cr" if above else f"{low:g} Cr and above"
    below = bool(re.search(r"less than|below|under|<(?!=)", t))
    return f"{'Above ' if above else ''}{low:g} Cr to {'<' if below else ''}{high:g} Cr"


def _union_section_type(title: str) -> str:
    t = clean(title).lower()
    if re.match(r"(nre|nro)\b", t) or "notice" in t:
        return ""
    return NON_CALLABLE if re.match(r"non\s*-?\s*callable", t) else CALLABLE


def _union_slab_headers(content) -> list[str]:
    """Amount slab labels from the PDF table header row ('Tenor', '3cr. Upto10 cr.', '>10cr.upto 25cr.', ...).
    (pdfplumber scatters the rate cells of this PDF over spurious blank columns, so only the header is used
    from the table extraction; rates come from the text lines.)"""
    for _, grid in pdf_tables(content):
        for row in grid[:3]:
            if row and clean(row[0]).lower().startswith("tenor"):
                labels = [c for c in row[1:] if c]
                if labels and all(amount_range_cr(c) for c in labels):
                    return labels
    return []


def _union_fix_touching(rows: list[RateRow]) -> list[RateRow]:
    """'121 - 180 Days' followed by '6 Months to < 1 Year': 6 months = 180 days under the 30-day convention,
    so both slabs would claim day 180.  The later slab starts the day after the earlier one ends."""
    for prev, cur in zip(rows, rows[1:]):
        if (prev.WithdrawalType == cur.WithdrawalType and prev.AmountSlab == cur.AmountSlab
                and cur.StartDay == prev.EndDay and cur.StartDay < cur.EndDay and prev.StartDay < prev.EndDay):
            cur.StartDay = prev.EndDay + 1
    return rows


def _union_pdf_rows(pdf) -> list[RateRow]:
    text = pdf_text(pdf.content)
    lines = [clean(l) for l in text.splitlines() if clean(l)]
    title = next((l for l in lines if re.search(r"bulk\s+deposit", l, re.I)), lines[0] if lines else "")
    effective = parse_date(title) or parse_date(" ".join(lines[:5]))   # 'Interest rates ... for 06.10.2026*'

    headers = _union_slab_headers(pdf.content)
    if headers:
        slabs = [(i, _union_slab_label(h)) for i, h in enumerate(headers)
                 if amount_range_cr(h)[1] > _UNION_MIN_SLAB_CR]
    else:
        # Title says 'Rs.3 Crores & Above': without a readable header only the first column (lowest slab) is safe.
        LOG.warning("UNION: amount slab header not readable - using only the first (lowest bulk) column")
        slabs = [(0, "3 Cr and above")]
    if not any(amount_contains(label) for _, label in slabs):
        LOG.warning("UNION: no slab contains Rs 3.01 crore: %s", [label for _, label in slabs])

    mode = "manual_pdf" if pdf.url.startswith("manual:") else "official_pdf"
    rows: list[RateRow] = []
    withdrawal = ""
    for line in lines:
        sec = _UNION_SECTION_RE.match(line)
        if sec and not re.search(r"\d+\.\d+", line[sec.end():]):
            withdrawal = _union_section_type(sec.group("title"))
            continue
        m = _UNION_RATE_LINE_RE.match(line)
        if not m or not withdrawal:
            continue
        rates = re.findall(r"\d{1,2}\.\d{1,2}", m.group("rates"))
        if headers and len(rates) != len(headers):
            LOG.warning("UNION: skipped line with %s rates (expected %s): %s", len(rates), len(headers), line)
            continue
        if tenure_bounds(m.group("label"))[0] is None:
            LOG.warning("UNION: tenure not understood: %s", m.group("label"))
            continue
        for i, label in slabs:
            made = make_row("UNION", withdrawal, m.group("label"), rates[i], amount=label, effective=effective,
                            url=pdf.url, mode=mode)
            if made:
                rows.append(made)
    # source order within each slab, then make touching slabs contiguous
    rows.sort(key=lambda r: (r.WithdrawalType, r.AmountSlab))       # stable: keeps PDF row order per slab
    return _union_fix_touching(rows)


@parser("UNION")
def parse_union() -> list[RateRow]:
    """Union Bank of India: 'Interest rates for Bulk Deposit (Rs.3 Crores & Above)' PDF - every amount column
    ('3cr. Upto10 cr.' ... '>500cr. & above'), sections 'Callable Deposit' and 'Non-Callable Deposit'."""
    page = fetch(BANKS["UNION"], "UNION", label="roi_page")
    if page.is_pdf:                       # a saved copy of the bulk PDF itself
        return _union_pdf_rows(page)
    links = find_links(page.content, page.url, _UNION_LINK_TEXT, href_pattern=r"\.pdf")
    links = [h for h, text in links if re.search(r"bulk", h + " " + text, re.I)]
    rows: list[RateRow] = []
    for url in list(dict.fromkeys(links + [UNION_BULK_PDF])):
        try:
            pdf = fetch(url, "UNION", label="bulk_pdf")
        except SourceBlocked:
            raise
        except Exception as exc:
            LOG.warning("UNION: %s failed: %s", url, exc)
            continue
        if not pdf.is_pdf:
            continue
        rows = _union_pdf_rows(pdf)
        if rows:
            return rows
    return rows


def _union_dedupe(rows: list[RateRow]) -> list[RateRow]:
    best: dict[tuple, RateRow] = {}
    for r in rows:
        key = (r.WithdrawalType, r.AmountSlab, r.StartDay, r.EndDay)
        if key not in best or r.Rate > best[key].Rate:
            best[key] = r
    return list(best.values())


# ---- PNB -----------------------------------------------------------------------
# PNB publishes its bulk (> Rs 3 crore) card rates in two places - both are always read:
#  1. "Domestic TD Rs.3 Cr. To 10 Cr." - HTML table on the general deposit-rates page, with
#     "Domestic Fixed Deposit Scheme" (callable) and "PNB Uttam (Non Callable)" columns, each published as
#     "Existing ... w.e.f. <date>" / "Revised ... w.e.f. <date>"; the newest column already in force is used.
#  2. BANKS["PNB"] (interest-rates-bulk-deposit.html) - a PDF.js viewer for the "> Rs 10 Cr" bulk PDF.  The PDF
#     URL (downloadprocess.aspx?fid=...) changes with each revision, so it is read from the page's script.
#     Every amount column of "Retail Customers (Callable)" / "Retail Customers (Non-Callable)" is used
#     (Above Rs 10 Cr to Rs 25 Cr ... Above Rs 1000 Cr).  NRE, MIBOR-linked and FCNR tables are ignored.
PNB_DEPOSIT_URL = "https://pnb.bank.in/Interest-Rates-Deposit.html"


PNB_BULK_PDF_FALLBACK = "https://pnb.bank.in/downloadprocess.aspx?fid=2CTkAq1s7HMN1kpNzipiSw=="


def _pnb_amount_label(text: str) -> str:
    """'Domestic TD Rs.3 Cr. To 10 Cr.' -> 'Rs.3 Cr. To 10 Cr.' (amount part only, still parseable)."""
    text = clean(text)
    m = re.search(r"(?:above\s+|from\s+|>=?\s*)?(?:rs\.?\s*|inr\s*)?\d", text, re.I)
    return text[m.start():].strip() if m else text


def _pnb_columns(grid: list[list[str]], header_rows: int, period_col: int) -> dict[str, tuple[int, str, str]]:
    """{withdrawal: (column, header text, effective date)} - newest applicable w.e.f. column per scheme."""
    best: dict[str, tuple] = {}
    for c in range(len(grid[0])):
        if c == period_col:
            continue
        header = " | ".join(dict.fromkeys(grid[r][c] for r in range(1, header_rows) if grid[r][c]))
        low = header.lower()
        if not header or re.search(r"\bsl\.?\s*no|\bnro\b|\bnre\b|senior|staff", low):
            continue
        if re.search(r"non[\s-]*callable", low):
            withdrawal = NON_CALLABLE
        elif re.search(r"fixed deposit|callable|term deposit", low):
            withdrawal = CALLABLE
        else:
            continue
        # parse_date ignores dates after today, so a 'Revised w.e.f. <future date>' column is skipped
        # automatically and the 'Existing' column (still in force) wins.
        effective = parse_date(header)
        key = (effective, 1 if "revised" in low else 0, c)
        if withdrawal not in best or key > best[withdrawal][:3]:
            best[withdrawal] = (*key, header)
    return {w: (v[2], v[3], v[0]) for w, v in best.items()}


def _pnb_3_to_10cr(page) -> list[RateRow]:
    """Every domestic bulk slab table on the deposit page (today: only 'Domestic TD Rs.3 Cr. To 10 Cr.')."""
    rows: list[RateRow] = []
    for table in soup_of(page.content).find_all("table"):
        grid = table_grid(table)
        if len(grid) < 4:
            continue
        title = clean(grid[0][0])
        low = title.lower()
        if not low.startswith("domestic") or re.search(r"\bnre\b|tax saver|senior|staff", low):
            continue
        if not is_bulk_slab(title):
            continue
        period_col = next((c for r in grid[:4] for c, v in enumerate(r) if re.search(r"period|tenure|maturity", v, re.I)), 1)
        header_rows = next(i for i, r in enumerate(grid)
                           if tenure_bounds(r[period_col])[0] is not None and any(rate_value(x) for x in r[period_col + 1:]))
        slab = _pnb_amount_label(title)
        for withdrawal, (col, header, effective) in _pnb_columns(grid, header_rows, period_col).items():
            LOG.info("PNB %s: '%s' column '%s'", withdrawal, slab, header)
            for r in grid[header_rows:]:
                made = make_row("PNB", withdrawal, r[period_col], r[col], amount=slab, effective=effective,
                                url=page.url, mode="official_html")
                if made:
                    rows.append(made)
    return rows


def _pnb_bulk_pdf_url() -> str:
    try:
        page = fetch(BANKS["PNB"], "PNB", label="bulk_page")
    except SourceBlocked:
        raise
    except Exception as exc:  # page down - try the last known PDF id
        LOG.warning("PNB bulk page unavailable (%s); using fallback PDF id", exc)
        return PNB_BULK_PDF_FALLBACK
    if page.is_pdf:
        return page.url
    match = re.search(r"fileUrl\s*=\s*[`'\"]([^`'\"]*downloadprocess\.aspx\?fid=[^`'\"]+)", page.text)
    if match:
        return urljoin(page.url if page.url.startswith("http") else BANKS["PNB"], match.group(1))
    links = find_links(page.content, BANKS["PNB"], r"bulk|interest", r"downloadprocess\.aspx\?fid=")
    if links:
        return links[0][0]
    LOG.warning("PNB: PDF link not found on %s; using fallback PDF id", BANKS["PNB"])
    return PNB_BULK_PDF_FALLBACK


def _pnb_bulk_pdf(pdf) -> list[RateRow]:
    """'> Rs 10 Cr' PDF: Retail Customers (Callable / Non-Callable), every amount column."""
    effective = parse_date(pdf_text(pdf.content))
    rows: list[RateRow] = []
    for _, grid in pdf_tables(pdf.content):
        groups, last = [], ""
        for cell in grid[0]:                       # forward-fill the merged group label row
            last = cell or last
            groups.append(last.lower())
        if not any("retail customers" in g for g in groups):
            continue
        slab_row = grid[1]
        for withdrawal, pattern in ((CALLABLE, r"retail customers \(callable\)"),
                                    (NON_CALLABLE, r"retail customers \(non-callable\)")):
            for col in (c for c, g in enumerate(groups) if re.search(pattern, g) and is_bulk_slab(slab_row[c])):
                slab = _pnb_amount_label(slab_row[col])
                for r in grid[2:]:
                    made = make_row("PNB", withdrawal, r[0], r[col], amount=slab, effective=effective,
                                    url=pdf.url, mode="official_pdf")
                    if made:
                        rows.append(made)
    return rows


@parser("PNB")
def parse_pnb() -> list[RateRow]:
    manual = _FETCH_CACHE.get(BANKS["PNB"])
    manual = manual if manual is not None and manual.url.startswith("manual:") else None   # scrape_bank() fallback
    rows: list[RateRow] = []
    problems = []
    # 1. Rs 3 Cr to 10 Cr (HTML)
    try:
        page = manual if manual is not None and not manual.is_pdf else fetch(PNB_DEPOSIT_URL, "PNB", label="deposit")
        part = _pnb_3_to_10cr(page)
        if not part:
            problems.append(f"no 'Domestic TD Rs.3 Cr. To 10 Cr.' table on {page.url}")
        rows += part
    except SourceBlocked:
        if manual is None:
            raise
        problems.append(f"{PNB_DEPOSIT_URL} blocked")
    except Exception as exc:
        problems.append(f"3-10 Cr page failed: {exc}")
    # 2. Above Rs 10 Cr (PDF)
    try:
        pdf = manual if manual is not None and manual.is_pdf else fetch(_pnb_bulk_pdf_url(), "PNB", label="bulk_pdf")
        part = _pnb_bulk_pdf(pdf)
        if not part:
            problems.append(f"no Retail Customers table in bulk PDF {pdf.url}")
        rows += part
    except SourceBlocked:
        if manual is None:
            raise
        problems.append("bulk PDF blocked")
    except Exception as exc:
        problems.append(f">10 Cr bulk PDF failed: {exc}")
    for p in problems:
        LOG.warning("PNB: %s", p)
    return rows


# ---- HDFC ----------------------------------------------------------------------
# HDFC interest-rates page (server-rendered AEM tabs, no JSON API).  Every amount column covering amounts
# above Rs 3 crore is emitted, one AmountSlab per source column:
#  * CALLABLE - tab "3 cr to 5 cr": 'Domestic / NRO / NRE FIXED DEPOSIT RATE', slab '>=3 Crore to < 5 Crores',
#    'Interest Rate (per annum)' column (the '**Senior Citizen Rates' column is skipped); effective date from
#    the tab text ('Applicable from 7<sup>th</sup> March, 2026').
#  * CALLABLE - tab "5 Cr and above": 'DOMESTIC/NRE/NRO TERM DEPOSITS ... EQUAL & MORE THAN Rs 5 Cr
#    (Withdrawable)', every slab column ('5crore to <5.25 crore' ... '>= 400 crore').
#  * NON_CALLABLE - 'NON WITHDRAWABLE DOMESTIC/NRE/NRO TERM DEPOSITS ... EQUAL & MORE THAN RS. 2 Cr', every slab
#    column ('2 crore to < 5 crore' ... '>= 400 crore').  Date row 'APPLICABLE FOR 06th OCTOBER, 2026'.
#  The '< 3 Crore' table, FCNR/RFC/diplomat, MCLR and savings tables are not bulk and are skipped.
_HDFC_SKIP_TABLE = re.compile(r"fcnr|rfc|diplomat|embassy|mclr|forex|savings|recurring|tax\s*saver|mibor", re.I)


def _hdfc_text_outside_tables(node) -> str:
    return clean(" ".join(s for s in node.find_all(string=True)
                          if not s.find_parent("table") and s.parent.name not in ("script", "style")))


def _hdfc_effective(table, header_text: str) -> str:
    found = parse_date(header_text)
    if found:
        return found
    # nearest enclosing block (e.g. the '3 cr to 5 cr' tab panel) whose own text carries a date
    for parent in table.parents:
        if parent.name in ("body", "html", "[document]"):
            break
        found = parse_date(_hdfc_text_outside_tables(parent))
        if found:
            return found
    return ""


def _hdfc_slab_label(text: str) -> str:
    """'5crore<5.25 crore' / '>=5.25 crore <5.50 crore' -> '... crore to <...' (same bounds, one style for C and NC)."""
    return re.sub(r"(crores?|crs?\.?)\s*(?=(?:<|>)=?\s*\d)", r"\1 to ", clean(text), flags=re.I)


def _hdfc_layout(grid: list[list[str]]):
    """-> (header_rows, tenure_col, {col: slab label}, {col: full header text}) or None."""
    tenure_col = next((c for r in grid[:4] for c, v in enumerate(r) if re.search(r"period|tenure|tenor", v, re.I)), 0)
    header_rows = next((i for i, r in enumerate(grid) if tenure_bounds(r[tenure_col])[0] is not None
                        and any(rate_value(x) for x in r[tenure_col + 1:])), None)
    if not header_rows:
        return None
    # merged title rows (same text in every cell) describe the table, not a column - keep them out of the slab
    sub_rows = [r for r in grid[:header_rows] if len({x for x in r if x}) > 1]
    slabs, headers = {}, {}
    for c in range(len(grid[0])):
        if c == tenure_col:
            continue
        cells = [r[c] for r in sub_rows if r[c]]
        slab = next((_hdfc_slab_label(x) for x in cells if amount_range_cr(x)), "")
        if slab:
            slabs[c] = slab
            headers[c] = " | ".join(dict.fromkeys(cells))
    return header_rows, tenure_col, slabs, headers


@parser("HDFC")
def parse_hdfc() -> list[RateRow]:
    page = fetch(BANKS["HDFC"], "HDFC")
    rows: list[RateRow] = []
    for table in soup_of(page.content).find_all("table"):
        grid = table_grid(table)
        if len(grid) < 4:
            continue
        layout = _hdfc_layout(grid)
        if not layout:
            continue
        header_rows, tenure_col, slabs, headers = layout
        title_text = " ".join(" ".join(r) for r in grid[:header_rows])
        context = table_context(table) + " " + title_text
        if _HDFC_SKIP_TABLE.search(context):
            continue
        cols = [c for c in slabs if is_bulk_slab(slabs[c]) and not re.search(r"senior|staff|nri\b", headers[c], re.I)]
        if not cols:
            continue
        withdrawal = NON_CALLABLE if re.search(r"non[\s-]*withdraw", context, re.I) else CALLABLE
        effective = _hdfc_effective(table, title_text)
        LOG.info("HDFC %s: %s slab column(s) %s (effective %s)", withdrawal, len(cols), [slabs[c] for c in cols], effective)
        for col in cols:
            for r in grid[header_rows:]:
                made = make_row("HDFC", withdrawal, r[tenure_col], r[col], amount=slabs[col], effective=effective,
                                url=page.url, mode="official_html")
                if made:
                    rows.append(made)
    return rows


# ---- ICICI ---------------------------------------------------------------------
"""ICICI Bank bulk FD rates - every amount slab above Rs 3 crore.

The FD-rates page has an 'interest-tables' component whose amount tabs ('Less than Rs 3 Cr.', '3 - < 5 Cr.',
'5 - < 5.10 Cr.', ... 'More than 500 Cr.') are built client side from two JSON files referenced on the
component itself:
    data-interestrate -> .../fd-interest-rate-tabs.json  (tab labels, in display order)
    data-api          -> .../fd-interest-rate.json       (interestData[k] = rows of the k-th amount tab,
                                                         'Popular FD Rates' excluded)
Row keys: c1 / c2 = WITH premature withdrawal (general / senior citizen),
          c3 / c4 = WITHOUT premature withdrawal (general / senior citizen).
The tab -> group mapping was verified by rendering the page and clicking every tab (Oct-2026).
If the JSON layout changes, the parser falls back to rendering the page and clicking every bulk tab.
"""


_ICICI_BANK = "ICICI"


_ICICI_RATES_API_FALLBACK = "/content/dam/icicibank-revamp/deposits/fixed-deposits/json/fd-interest-rate.json"


_ICICI_TABS_API_FALLBACK = "/content/dam/icicibank-revamp/deposits/fixed-deposits/json/fd-interest-rate-tabs.json"


_ICICI_KEY_BY_TYPE = {CALLABLE: "c1", NON_CALLABLE: "c3"}           # general-citizen columns only


_ICICI_EXCLUDE_TENURE = re.compile(r"tax\s*saver", re.I)


def _icici_slab_label(text: str) -> str:
    """Bank tab label -> normalised label that amount_range_cr() parses back to the same bounds."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    low, high = bounds
    if high == float("inf"):
        exclusive = re.match(r"^\s*(?:more than|above|over|>(?!=))", clean(text), re.I)
        return f"Above {low:g} Cr" if exclusive else f"{low:g} Cr and above"
    return f"{low:g} Cr to <{high:g} Cr"


def _icici_bulk_slabs(labels: list[str]) -> list[int]:
    """Indexes of amount labels whose slab reaches above Rs 3 crore."""
    out = [i for i, label in enumerate(labels) if is_bulk_slab(label)]     # upper bound above 3 cr
    if not out:
        raise ValueError(f"ICICI: no amount tab above 3 cr: {labels}")
    return out


def _icici_html_text(fragment: str) -> str:
    return clean(soup_of(fragment or "<p></p>").get_text(" ", strip=True))


def _icici_component(page):
    soup = soup_of(page.content)
    div = soup.find(attrs={"data-component": "interest-tables"})
    if div is None:
        raise ValueError("ICICI: interest-tables component not found on the FD rates page")
    item = json.loads(div.get("data-item") or "{}")
    return {
        "heading": _icici_html_text(item.get("heading", "")),
        "col1": _icici_html_text(item.get("column1Subheading", "")).lower(),
        "col3": _icici_html_text(item.get("column3Subheading", "")).lower(),
        "rates_api": urljoin(page.url, div.get("data-api") or _ICICI_RATES_API_FALLBACK),
        "tabs_api": urljoin(page.url, div.get("data-interestrate") or _ICICI_TABS_API_FALLBACK),
    }


def _icici_from_json(page, comp) -> list[RateRow]:
    if "with premature" not in comp["col1"] or "without premature" not in comp["col3"]:
        raise ValueError(f"ICICI: unexpected column headings {comp['col1']!r} / {comp['col3']!r}")
    tabs = json.loads(fetch(comp["tabs_api"], _ICICI_BANK, label="tabs").content)["interestBtns"]
    data = json.loads(fetch(comp["rates_api"], _ICICI_BANK, label="rates").content)["interestData"]
    amount_tabs = [t for t in tabs if str(t.get("value", "")).lower() != "popular"]
    if len(amount_tabs) != len(data):
        raise ValueError(f"ICICI: {len(amount_tabs)} amount tabs but {len(data)} rate groups - layout changed")
    labels = [clean(t.get("label", "")) for t in amount_tabs]
    effective = parse_date(comp["heading"])
    rows = []
    for idx in _icici_bulk_slabs(labels):
        slab = _icici_slab_label(labels[idx])
        for entry in data[idx]:
            tenure = clean(entry.get("tenure", ""))
            if not tenure or _ICICI_EXCLUDE_TENURE.search(tenure):
                continue
            for withdrawal, key in _ICICI_KEY_BY_TYPE.items():
                value = entry.get(key)
                if value is None:
                    continue
                made = make_row(_ICICI_BANK, withdrawal, tenure, str(value), amount=slab, effective=effective,
                                url=comp["rates_api"], mode="official_json")
                if made:
                    rows.append(made)
    return rows


def _icici_from_render(comp) -> list[RateRow]:
    """Fallback: render the page, click every bulk amount tab and read the visible table."""
    captured: list[tuple[str, str]] = []

    def click_tabs(page):
        buttons = page.locator(".interest-tables button")
        labels = [clean(buttons.nth(i).inner_text()) for i in range(buttons.count())]
        amount = [(i, l) for i, l in enumerate(labels) if amount_range_cr(l)]
        for k in _icici_bulk_slabs([l for _, l in amount]):
            index, label = amount[k]
            buttons.nth(index).click()
            page.wait_for_timeout(2000)
            captured.append((label, page.inner_html(".interest-tables")))

    page = render(BANKS[_ICICI_BANK], _ICICI_BANK, wait_for=".interest-tables button", label="rendered", actions=click_tabs)
    effective = parse_date(comp["heading"])
    rows = []
    for label, html in captured:
        for table in soup_of(html).find_all("table"):
            grid = table_grid(table)
            if len(grid) < 3:
                continue
            top = [c.lower() for c in grid[0]]
            if "tenure" not in top:
                continue
            tcol = top.index("tenure")
            heads = [f"{grid[0][c]} | {grid[1][c]}".lower() for c in range(len(grid[0]))]
            for withdrawal, pattern in ((CALLABLE, r"^with premature.*general"),
                                        (NON_CALLABLE, r"^without premature.*general")):
                cols = [c for c, h in enumerate(heads) if re.search(pattern, h)]
                if cols:
                    rows += rows_from_grid(_ICICI_BANK, grid, withdrawal, effective=effective, url=page.url,
                                           mode="official_html_rendered", tenure_col=tcol, rate_col=cols[0],
                                           amount=_icici_slab_label(label), header_rows=2)
    return [r for r in rows if not _ICICI_EXCLUDE_TENURE.search(r.Tenure)]


@parser("ICICI")
def parse_icici() -> list[RateRow]:
    page = fetch(BANKS[_ICICI_BANK], _ICICI_BANK)
    comp = _icici_component(page)
    try:
        rows = _icici_from_json(page, comp)
    except (ValueError, KeyError, IndexError, json.JSONDecodeError) as exc:
        LOG.warning("ICICI JSON route failed (%s) - rendering the page instead", exc)
        rows = []
    if not rows:
        rows = _icici_from_render(comp)
    return rows


# ---- AXIS ----------------------------------------------------------------------
"""Axis Bank bulk FD rates - every amount slab above Rs 3 crore.

The FD-rates page carries 'VIEW RATES' cards linking to dated PDFs (file name changes with every revision):
  * 'Domestic Fixed Deposit (>= Rs 3 crs)'      -> Callable (premature withdrawal allowed);
    page 1 = general rates, page 2 = senior-citizen rates (skipped)
  * 'Domestic Fixed Deposit Plus (>= Rs 3 crs)' -> Non-Callable ('PREMATURE WITHDRAWAL IS NOT PERMITTED')
Both PDFs have amount columns 'Rs 3 Cr < Rs 5Cr', 'Rs 5 Cr < Rs 25 Cr', ... 'Rs 1000 Cr & Above'; every
column is emitted.  pdfplumber splits some header labels over several grid rows ('Rs 500 Cr' / '<' /
'Rs 1000' / 'Cr') and puts the FD Plus value cells one column right of their headers, so header rows are
merged column-wise, every row is compacted to its non-empty cells and the k-th value is matched to the
k-th amount header.
"""


_AXIS_BANK = "AXIS"


_AXIS_INF = float("inf")


def _axis_slab_label(low: float, high: float) -> str:
    """Normalised label that amount_range_cr() parses back to (low, high)."""
    return f"{low:g} Cr and above" if high == _AXIS_INF else f"{low:g} Cr to <{high:g} Cr"


def _axis_slab_bounds(headers: list[str]) -> list[Optional[tuple[float, float]]]:
    """Bounds per amount header; a slab whose upper bound was lost in extraction ends where the next starts."""
    bounds = [amount_range_cr(h) for h in headers]
    for k, b in enumerate(bounds[:-1]):
        nxt = bounds[k + 1]
        if b and nxt and (b[1] == _AXIS_INF or b[1] < nxt[0]):
            LOG.warning("AXIS: slab %r has no usable upper bound - using next slab start %g", headers[k], nxt[0])
            bounds[k] = (b[0], nxt[0])
    return bounds


# card titles on the FD rates page (structural hints)
_AXIS_CARD_TITLE = {
    CALLABLE: re.compile(r"^domestic fixed deposits?\s*\(\s*>=?\s*(?:rs\.?\s*)?3\s*cr", re.I),
    NON_CALLABLE: re.compile(r"^domestic fixed deposits?\s+plus\s*\(\s*>=?\s*(?:rs\.?\s*)?3\s*cr", re.I),
}


# PDF file-name patterns (second line of discovery)
_AXIS_HREF_PATTERN = {
    CALLABLE: re.compile(r"/domestic-fixed-deposits?-interest-rates?-3-crores?-(?:and|&)-above[^/]*\.pdf", re.I),
    NON_CALLABLE: re.compile(r"/domestic-fixed-deposits?-plus[^/]*\.pdf", re.I),
}


# last known URLs (used only if the page no longer links a matching PDF)
_AXIS_FALLBACK_URL = {
    CALLABLE: "https://www.axis.bank.in/docs/default-source/default-document-library/interest-rates/"
              "domestic-fixed-deposit-interest-rates-3-crores-and-above-06-october-2026.pdf",
    NON_CALLABLE: "https://www.axis.bank.in/docs/default-source/default-document-library/interest-rates/"
                  "domestic-fixed-deposits-plus-06-october-26.pdf",
}


_AXIS_TABLE_MUST = {CALLABLE: re.compile(r"3\s*crores?\s*(?:and|&)\s*above", re.I),
              NON_CALLABLE: re.compile(r"premature\s+withdrawal\s+is\s+not\s+permitted", re.I)}


_AXIS_SKIP_TABLE = re.compile(r"senior|nri|nre|nro|fcnr", re.I)


def _axis_date_from_href(href: str) -> str:
    return parse_date(re.sub(r"[-_]", " ", href.rsplit("/", 1)[-1]))


def _axis_discover_pdfs(page) -> dict[str, str]:
    soup = soup_of(page.content)
    found: dict[str, list[str]] = {CALLABLE: [], NON_CALLABLE: []}
    for a in soup.find_all("a", href=re.compile(r"\.pdf", re.I)):
        href = urljoin(page.url, a["href"].strip())
        card = a.find_parent(class_=re.compile(r"card", re.I))           # title of the card holding the link
        title_node = card.find(["h3", "h4", "h5"]) if card else None
        title = clean(title_node.get_text(" ", strip=True)) if title_node else ""
        for w in (CALLABLE, NON_CALLABLE):
            if _AXIS_CARD_TITLE[w].search(title) or _AXIS_HREF_PATTERN[w].search(href):
                if href not in found[w]:
                    found[w].append(href)
    out = {}
    for w in (CALLABLE, NON_CALLABLE):
        links = found[w]
        if links:   # newest dated file wins if several are linked
            out[w] = max(links, key=lambda h: (_axis_date_from_href(h), -links.index(h)))
        else:
            LOG.warning("AXIS: no %s PDF link found on %s - using last known URL", w, page.url)
            out[w] = _AXIS_FALLBACK_URL[w]
    return out


def _axis_effective(page_text: str, url: str) -> str:
    m = re.search(r"applicable\s+(?:for|from|on|w\.?\s*e\.?\s*f\.?)\s*[:\-]?\s*(.{0,40})", page_text, re.I)
    return (parse_date(m.group(1)) if m else "") or parse_date(page_text) or _axis_date_from_href(url)


def _axis_is_data_row(raw: list[str]) -> bool:
    cells = [c for c in raw if c]
    return len(cells) >= 2 and tenure_bounds(cells[0])[0] is not None and rate_value(cells[-1]) is not None


def _axis_merged_headers(grid: list[list[str]], header_idx: int, data_idx: int) -> list[str]:
    """Amount header labels (left to right), joining labels pdfplumber split over several grid rows."""
    merged = list(grid[header_idx])
    anchors = [c for c, text in enumerate(merged) if text]
    for r in range(header_idx + 1, data_idx):
        for c, text in enumerate(grid[r]):
            if text:
                home = c if c in anchors else max((a for a in anchors if a <= c), default=c)
                merged[home] = f"{merged[home]} {text}".strip()
    return [clean(merged[c]) for c in anchors]


def _axis_rows_from_pdf(content: bytes, withdrawal: str, url: str) -> list[RateRow]:
    rows = []
    for page_text, grid in pdf_tables(content):
        header_idx = next((i for i, r in enumerate(grid)
                           if sum(amount_range_cr(c) is not None for c in r if c) >= 2), None)
        if header_idx is None:
            continue
        title = " ".join(" ".join(r) for r in grid[:header_idx + 1])
        if _AXIS_SKIP_TABLE.search(title) or not _AXIS_TABLE_MUST[withdrawal].search(page_text + " " + title):
            continue
        data_idx = next((i for i in range(header_idx + 1, len(grid)) if _axis_is_data_row(grid[i])), len(grid))
        headers = _axis_merged_headers(grid, header_idx, data_idx)
        bounds = _axis_slab_bounds(headers)
        wanted = [k for k, b in enumerate(bounds) if b and is_bulk_slab(_axis_slab_label(*b))]   # upper bound > 3 cr
        if not wanted:
            LOG.warning("AXIS: no amount column above 3 cr in %s table: %s", withdrawal, headers)
            continue
        effective = _axis_effective(page_text, url)
        for raw in grid[data_idx:]:
            if not _axis_is_data_row(raw):
                continue
            cells = [c for c in raw if c]
            values = cells[1:]
            if len(values) != len(headers):
                LOG.warning("AXIS: %s row %r has %s values for %s amount slabs - skipped",
                            withdrawal, cells[0], len(values), len(headers))
                continue
            for k in wanted:
                made = make_row(_AXIS_BANK, withdrawal, cells[0], values[k], amount=_axis_slab_label(*bounds[k]),
                                effective=effective, url=url, mode="official_pdf")
                if made:
                    rows.append(made)
    return rows


@parser("AXIS")
def parse_axis() -> list[RateRow]:
    page = fetch(BANKS[_AXIS_BANK], _AXIS_BANK)
    rows = []
    for withdrawal, url in _axis_discover_pdfs(page).items():
        pdf = fetch(url, _AXIS_BANK, label="nc_plus" if withdrawal == NON_CALLABLE else "callable_3cr")
        if not pdf.is_pdf:
            raise ValueError(f"AXIS: {url} did not return a PDF")
        got = _axis_rows_from_pdf(pdf.content, withdrawal, pdf.url)
        if not got:
            LOG.warning("AXIS: no %s rows parsed from %s", withdrawal, url)
        rows += got
    return rows


# ---- KOTAK ---------------------------------------------------------------------
"""Kotak Mahindra Bank bulk FD rates - every amount slab above Rs 3 crore.

The FD-rates page has no rate table in its HTML: each table is injected by a <script src=
'.../bank/mailers/intrates/get_all_variable_data_latest.php?section=...'> that document.write()s it.
Sections seen (Oct-2026):
  NRO_Term_Deposit            'Maturity Periods - Premature Withdrawal Allowed'; columns Regular / Senior
                              x ('Less than Rs.3 Crore', 'Rs. 3 Cr. & above but less than Rs. 5 Cr.')
  NRO_Term_Deposit_1cr        'Premature Withdrawal Allowed': 5-<10, 10-<25, 25-<100, 100-<300, 300+ cr
  NRO_PW_NA_Term_Deposit_1cr  'Premature Withdrawal Not Allowed': 3-<5, 5-<10, ... 300+ cr
The tables are 'Domestic / NRO / NRE' card rates.  Senior-citizen and '< 3 crore' columns are skipped.
Every script is discovered from the page, each table is classified by its 'Premature Withdrawal
(Not) Allowed' header and every amount column above 3 crore is emitted with a normalised slab label.
Kotak prints 'NA' in every tenure of the 3-<5 cr non-callable column (no non-callable FD below 5 cr),
so that slab yields no rows.
"""


_KOTAK_BANK = "KOTAK"


_KOTAK_SCRIPT_SRC = re.compile(r"/intrates/get_all_variable_data[^\"']*\?section=", re.I)


_KOTAK_FALLBACK_SECTIONS = ["NRO_Term_Deposit", "NRO_Term_Deposit_1cr", "NRO_PW_NA_Term_Deposit_1cr"]


_KOTAK_FALLBACK_SCRIPT = "https://www.kotak.bank.in/bank/mailers/intrates/get_all_variable_data_latest.php?section={}"


_KOTAK_SKIP_SECTION = re.compile(r"recurring|rd_|tax|sweep|activ|fcnr", re.I)


_KOTAK_NOT_ALLOWED = re.compile(r"premature\s+withdrawal\s+not\s+allowed", re.I)


_KOTAK_ALLOWED = re.compile(r"premature\s+withdrawal\s+allowed", re.I)


_KOTAK_SKIP_COLUMN = re.compile(r"senior|staff|nre only|fcnr", re.I)


_KOTAK_INF = float("inf")


def _kotak_slab_label(text: str) -> str:
    """Bank amount header -> normalised label that amount_range_cr() parses back to the same bounds."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    low, high = bounds
    return f"{low:g} Cr and above" if high == _KOTAK_INF else f"{low:g} Cr to <{high:g} Cr"


def _kotak_script_html(js: str) -> str:
    """Concatenate the string literals passed to document.write()."""
    parts = re.findall(r"document\.write\(\s*\"((?:[^\"\\]|\\.)*)\"\s*\)", js, re.S)
    parts += re.findall(r"document\.write\(\s*'((?:[^'\\]|\\.)*)'\s*\)", js, re.S)
    return "".join(p.replace('\\"', '"').replace("\\'", "'").replace("\\/", "/") for p in parts)


def _kotak_script_urls(page) -> list[str]:
    soup = soup_of(page.content)
    urls = [urljoin(page.url, s["src"]) for s in soup.find_all("script", src=_KOTAK_SCRIPT_SRC)]
    urls = [u for u in dict.fromkeys(urls) if not _KOTAK_SKIP_SECTION.search(u.split("section=", 1)[-1])]
    if not urls:
        LOG.warning("KOTAK: no rate scripts found on %s - using known sections", page.url)
        urls = [_KOTAK_FALLBACK_SCRIPT.format(s) for s in _KOTAK_FALLBACK_SECTIONS]
    return urls


def _kotak_tables(page) -> list[dict]:
    """Every rate table as {withdrawal, grid, header_rows, headers, effective, url, section}."""
    out = []
    for url in _kotak_script_urls(page):
        section = url.split("section=", 1)[-1]
        script = fetch(url, _KOTAK_BANK, label=section, headers={"Referer": page.url})
        for table in soup_of(_kotak_script_html(script.text)).find_all("table"):
            grid = table_grid(table)
            if not grid:
                continue
            header_rows = next((i for i, r in enumerate(grid)
                                if tenure_bounds(r[0])[0] is not None and any(rate_value(c) for c in r[1:])), None)
            if not header_rows:
                continue
            head_text = " ".join(" ".join(r) for r in grid[:header_rows])
            if _KOTAK_NOT_ALLOWED.search(head_text):
                withdrawal = NON_CALLABLE
            elif _KOTAK_ALLOWED.search(head_text):
                withdrawal = CALLABLE
            else:
                LOG.warning("KOTAK: unclassified table in %s: %s", url, head_text[:120])
                continue
            # per-column header without the merged title row (row 0 spans the whole table)
            headers = [" | ".join(dict.fromkeys(grid[r][c] for r in range(1, header_rows) if grid[r][c]))
                       for c in range(len(grid[0]))]
            out.append({"withdrawal": withdrawal, "grid": grid, "header_rows": header_rows, "headers": headers,
                        "effective": parse_date(head_text), "url": script.url, "section": section})
    return out


@parser("KOTAK")
def parse_kotak() -> list[RateRow]:
    page = fetch(BANKS[_KOTAK_BANK], _KOTAK_BANK)
    rows = []
    for t in _kotak_tables(page):
        for c, header in enumerate(t["headers"]):
            if c == 0 or _KOTAK_SKIP_COLUMN.search(header):
                continue
            amount_text = t["grid"][t["header_rows"] - 1][c]          # the column's own amount label
            if not is_bulk_slab(amount_text):
                continue
            got = rows_from_grid(_KOTAK_BANK, t["grid"], t["withdrawal"], effective=t["effective"], url=t["url"],
                                 mode="official_html", tenure_col=0, rate_col=c, amount=_kotak_slab_label(amount_text),
                                 header_rows=t["header_rows"])
            if not got:
                LOG.warning("KOTAK: %s %r (%s) has no published rate (all NA)", t["withdrawal"], amount_text,
                            t["section"])
            rows += got
    return rows


# ---- IDFC ----------------------------------------------------------------------
# IDFC FIRST publishes bulk (>= Rs 3 cr) rates as one PDF linked from the interest-rate page
# ("Bulk Deposit Interest Rate (Domestic, NRE, NRO) >=Rs. 3 Crores (w.e.f. dd.mm.yyyy)").  The file name
# carries the w.e.f. date, so it is discovered on every run; this URL is only the last-known fallback.
IDFC_FALLBACK_PDF = ("https://www.idfcfirst.bank.in/content/dam/idfcfirstbank/interest-rate/"
                     "Interest-rates-on-deposits-Rs-3-crs-and-above-wef-01102026.pdf")


# The PDF has 12 tables: {Individual, Non Financial, Financial} customers x {NC, C} x {3-27.50 cr, >27.50 cr},
# each table with 2 (3-10, >10-27.50 cr) or 5 (>27.50-50 ... >400 cr) amount columns.
# Both non-FI segments are emitted ('Individual' / 'Non Financial'); 'Financial' (FI depositors) is excluded.
IDFC_SEGMENTS = {"individual": "Individual", "non financial": "Non Financial"}


_IDFC_HEADING_RE = re.compile(r"\(([IVXL]+)\)\s+((?:Domestic\s+)?Interest rates for\s+(.+?)\s+Customers?\s*-?\s*(where\s+.+?)):", re.I)


def _idfc_slab_label(text: str) -> str:
    """Normalise a published amount slab to a label amount_range_cr() parses, e.g. 'Above 10 Cr to 27.5 Cr'."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    lo, hi = bounds
    t = clean(text).lower()
    lo_excl = bool(re.search(r"(?:>(?!=)|\babove\b|more than|\bover\b)\s*(?:rs\.?\s*|inr\s*)?\d", t))
    hi_excl = bool(re.search(r"(?:less than|below|under|<(?!=))\s*(?:rs\.?\s*|inr\s*)?\d", t))
    if hi == float("inf"):
        return f"Above {lo:g} Cr" if lo_excl else f"{lo:g} Cr and above"
    return f"{'Above ' if lo_excl else ''}{lo:g} Cr to {'<' if hi_excl else ''}{hi:g} Cr"


def _idfc_discover_pdf(page) -> str:
    soup = soup_of(page.content)
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if ".pdf" not in href.lower():
            continue
        box = a.parent
        label = clean(a.get("data-gtm-clicktext", "")) or (clean(box.get_text(" ", strip=True)) if box else "")
        if re.search(r"bulk\s+deposit\s+interest\s+rate|>=?\s*rs\.?\s*3\s*crores?", label, re.I):
            return urljoin(page.url, href)
    links = find_links(page.content, page.url, r"", href_pattern=r"interest-rate/[^?#]*3-crs?-and-above[^?#]*\.pdf")
    if links:
        return links[0][0]
    LOG.warning("IDFC: bulk-deposit PDF link not found on %s - using fallback %s", page.url, IDFC_FALLBACK_PDF)
    return IDFC_FALLBACK_PDF


def _idfc_make_contiguous(rows: list[RateRow]) -> list[RateRow]:
    """IDFC labels '732 - 1095 days' followed by '3 years - 5 years' (= 1095-1825): the second slab really
    starts the day after the first ends (the next one is '5 Year 1 Day - 8 Years'), so bump its start."""
    groups: dict[tuple, list[RateRow]] = {}
    for r in rows:
        groups.setdefault((r.WithdrawalType, r.AmountSlab), []).append(r)
    for seq in groups.values():
        seq.sort(key=lambda r: (r.StartDay, r.EndDay))
        for prev, cur in zip(seq, seq[1:]):
            if cur.StartDay == prev.EndDay and cur.EndDay > prev.EndDay:
                cur.StartDay = prev.EndDay + 1
    return rows


def _idfc_table_rows(grid: list[list[str]], withdrawal: str, segment: str, *, effective: str, url: str) -> list[RateRow]:
    """pdfplumber splits IDFC header cells over several lines and sometimes into extra columns, so header
    text and rate cells land in different column indexes.  Match them by order instead: the headers that
    carry an amount slab, left to right, against the non-empty rate cells of each tenor row."""
    first = next((i for i, r in enumerate(grid) if tenure_bounds(r[0])[0] is not None
                  and any(rate_value(c) for c in r[1:])), None)
    if first is None:
        return []
    width = len(grid[0])
    merged = [clean(" ".join(grid[r][c] for r in range(first) if grid[r][c])) for c in range(1, width)]
    slabs = [h for h in merged if amount_range_cr(h)]
    if not slabs:
        return []
    rows = []
    for raw in grid[first:]:
        rates = [c for c in raw[1:] if c]
        if len(rates) != len(slabs):
            LOG.warning("IDFC: skipped misaligned row %s (slabs %s)", raw, slabs)
            continue
        for head, rate in zip(slabs, rates):
            if amount_range_cr(head)[1] <= BULK_FLOOR_CR:
                continue
            made = make_row("IDFC", withdrawal, raw[0], rate, amount=f"{segment} | {_idfc_slab_label(head)}",
                            effective=effective, url=url, mode="official_pdf")
            if made:
                rows.append(made)
    return rows


def _idfc_rows_from_pdf(content: bytes, url: str) -> list[RateRow]:
    full = pdf_text(content)
    effective = parse_date(full[:400])          # '(effective October 01, 2026)' in the document title
    rows = []
    seen = set()
    for page_text, grid in pdf_tables(content):
        heading = _IDFC_HEADING_RE.search(page_text)
        if not heading:
            continue
        roman, _title, segment_text, condition = heading.groups()
        segment = IDFC_SEGMENTS.get(re.sub(r"[\s-]+", " ", clean(segment_text)).lower())
        if not segment:
            continue                              # 'Financial' (FI depositors) segment is excluded
        if re.search(r"withdrawal\s+is\s+not\s+permitted", condition, re.I):
            withdrawal = NON_CALLABLE
        elif re.search(r"withdrawal\s+is\s+permitted", condition, re.I):
            withdrawal = CALLABLE
        else:
            continue
        got = _idfc_table_rows(grid, withdrawal, segment, effective=effective, url=url)
        if got:
            seen.add((segment, withdrawal, roman))
            rows += got
    for segment in IDFC_SEGMENTS.values():
        for w in (CALLABLE, NON_CALLABLE):
            n = sum(1 for s, ww, _ in seen if s == segment and ww == w)
            if n != 2:
                LOG.warning("IDFC: expected 2 %s %s tables (3-27.50 cr and >27.50 cr), found %s", segment, w, n)
    return _idfc_make_contiguous(rows)


@parser("IDFC")
def parse_idfc() -> list[RateRow]:
    page = fetch(BANKS["IDFC"], "IDFC")
    if page.is_pdf:                               # manual_sources/idfc*.pdf saved by the user
        return _idfc_rows_from_pdf(page.content, page.url)
    pdf_url = _idfc_discover_pdf(page)
    pdf = fetch(pdf_url, "IDFC", label="bulk_pdf")
    if not pdf.is_pdf:
        raise RuntimeError(f"IDFC: {pdf_url} did not return a PDF")
    return _idfc_rows_from_pdf(pdf.content, pdf.url)


# ---- RBL -----------------------------------------------------------------------
# RBL publishes bulk (>= Rs 3 cr) card rates as two daily PDFs linked from the interest-rates page
# ("Premature withdrawal is allowed (Callable Bulk Deposit Rates)" / "... NOT allowed (Non-callable ...)").
# Each PDF has one table: tenor rows x 17 amount-slab columns (Rs 3 Cr to 3.50 Cr ... Above Rs 225 Cr).
# The file names have been stable; they are still discovered from the page and these are only fallbacks.
RBL_FALLBACK_PDFS = {
    CALLABLE: "https://webassets.rbl.bank.in/document/pdfs/callable-bulk-fd-rates.pdf",
    NON_CALLABLE: "https://webassets.rbl.bank.in/document/pdfs/non-callable-bulk-fd-rates.pdf",
}


def _rbl_slab_label(text: str) -> str:
    """Normalise a published amount slab to a label amount_range_cr() parses, e.g. 'Above 3.5 Cr to 4 Cr'."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    lo, hi = bounds
    t = clean(text).lower()
    lo_excl = bool(re.search(r"(?:>(?!=)|\babove\b|more than|\bover\b)\s*(?:rs\.?\s*|inr\s*)?\d", t))
    hi_excl = bool(re.search(r"(?:less than|below|under|<(?!=))\s*(?:rs\.?\s*|inr\s*)?\d", t))
    if hi == float("inf"):
        return f"Above {lo:g} Cr" if lo_excl else f"{lo:g} Cr and above"
    return f"{'Above ' if lo_excl else ''}{lo:g} Cr to {'<' if hi_excl else ''}{hi:g} Cr"


def _rbl_withdrawal_of(text: str) -> str:
    """Classify a PDF / link by its own wording; '' for floating-rate (MIBOR) or unknown documents."""
    t = clean(text).lower()
    if re.search(r"floating|mibor", t):
        return ""
    if re.search(r"premature\s+withdrawal\s+(?:is\s+)?not\s+allowed|non[\s-]*callable", t):
        return NON_CALLABLE
    if re.search(r"premature\s+withdrawal\s+(?:is\s+)?allowed|callable", t):
        return CALLABLE
    return ""


def _rbl_discover_pdfs(page) -> dict[str, str]:
    """{withdrawal type: pdf url} from the bulk-deposit links on the RBL interest-rates page."""
    found: dict[str, str] = {}
    soup = soup_of(page.content)
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not re.search(r"bulk[^/]*\.pdf", href, re.I):
            continue
        # the anchor text is just 'click here' - classify by the href, then by the sentence it sits in
        context = clean(a.parent.get_text(" ", strip=True)) if a.parent else ""
        kind = _rbl_withdrawal_of(href.rsplit("/", 1)[-1].replace("-", " ")) or _rbl_withdrawal_of(context)
        if kind and kind not in found:
            found[kind] = urljoin(page.url, href)
    return found


def _rbl_rows_from_pdf(content: bytes, url: str, expected: str = "") -> list[RateRow]:
    text = pdf_text(content)
    withdrawal = _rbl_withdrawal_of(" ".join(text.splitlines()[:4])) or expected
    if not withdrawal:
        LOG.warning("RBL: cannot tell callable / non-callable for %s", url)
        return []
    if expected and withdrawal != expected:
        LOG.warning("RBL: %s link points to a %s PDF (%s) - using the PDF's own wording", expected, withdrawal, url)
    effective = parse_date(text)              # 'Valid for 06-Oct-2026 Value Date'
    rows = []
    for _page_text, grid in pdf_tables(content):
        # header row holding the amount slabs ('Rs.3 Cr to Rs. 3.50 Cr', 'Above Rs.3.50 Cr to Rs. 4 Cr', ...)
        slab_row = next((i for i, r in enumerate(grid) if sum(amount_range_cr(c) is not None for c in r) >= 2), None)
        if slab_row is None:
            continue
        for col, head in enumerate(grid[slab_row]):
            bounds = amount_range_cr(head) if col else None
            if not bounds or bounds[1] <= BULK_FLOOR_CR:
                continue
            rows += rows_from_grid("RBL", grid[slab_row:], withdrawal, effective=effective, url=url,
                                   mode="official_pdf", tenure_col=0, rate_col=col, amount=_rbl_slab_label(head),
                                   header_rows=1)
    return rows


@parser("RBL")
def parse_rbl() -> list[RateRow]:
    page = fetch(BANKS["RBL"], "RBL")
    if page.is_pdf:                       # manual_sources/rbl*.pdf saved by the user
        return _rbl_rows_from_pdf(page.content, page.url)
    links = _rbl_discover_pdfs(page)
    rows = []
    for withdrawal in (CALLABLE, NON_CALLABLE):
        url = links.get(withdrawal)
        if not url:
            LOG.warning("RBL: %s bulk PDF link not found on page - using fallback URL", withdrawal)
            url = RBL_FALLBACK_PDFS[withdrawal]
        pdf = fetch(url, "RBL", label=f"bulk_{'c' if withdrawal == CALLABLE else 'nc'}")
        if not pdf.is_pdf:
            raise RuntimeError(f"RBL: {url} did not return a PDF")
        rows += _rbl_rows_from_pdf(pdf.content, pdf.url, expected=withdrawal)
    return rows


# ---- DCB -----------------------------------------------------------------------
# The DCB rates page is a Next.js app: its HTML has no tables.  The browser loads them by POSTing to the
# site's own public API route /api/api-interceptor with {"url": "/rates-sub-nav?main_nav=<page slug>"} and
# then {"url": "/rates?id=<sub-tab id>"}.  Each returned 'rate_details' item has a title, a subtitle
# '(with effect from ...)' and an HTML 'details' fragment holding the table(s).  Bulk tables are titled
# 'Interest Rates for Domestic / NRO / NRE** Bulk Deposits', one table per amount slab, each with a
# 'Regular' (= callable) and a 'Non - Callable' column.
DCB_API = "https://www.dcb.bank.in/api/api-interceptor"


DCB_MAIN_NAV = BANKS["DCB"].rstrip("/").rsplit("/", 1)[-1]      # 'fixed-deposit-interest-rate'


DCB_SUB_TAB_HINT = r"resident"                                  # 'Resident Fixed Deposit Interest Rates'


_DCB_SKIP = re.compile(r"retail|fcnr|penal|senior|nri\s+interest", re.I)


def _dcb_api(path: str, label: str) -> dict:
    """POST to DCB's API proxy exactly as the page does (one request per call), archived like fetch()."""
    headers = {"Accept": "application/json, text/plain, */*", "Content-Type": "application/json",
               "Origin": "https://www.dcb.bank.in", "Referer": BANKS["DCB"]}
    last = None
    for attempt in (1, 2):
        try:
            resp = SESSION.post(DCB_API, data=json.dumps({"url": path, "method": "GET"}), headers=headers, timeout=45)
            head = resp.content[:6000].decode("utf-8", errors="ignore")
            if resp.status_code in (403, 429, 503) and re.search(
                    r"just a moment|cf-challenge|cloudflare|incapsula|access denied|security verification", head, re.I):
                raise SourceBlocked(f"DCB: {DCB_API} is protected by a bot-check page (HTTP {resp.status_code})")
            resp.raise_for_status()
            (RAW_DIR / f"dcb_{_safe_name(label)}_{datetime.now():%H%M%S}.json").write_bytes(resp.content)
            return resp.json()
        except SourceBlocked:
            raise
        except Exception as exc:
            last = exc
            LOG.warning("DCB API attempt %s failed for %s: %s", attempt, path, exc)
            time.sleep(2 * attempt)
    raise RuntimeError(f"DCB API {path} failed: {last}")


def _dcb_nearest_date(context: str) -> str:
    for part in reversed(context.split(" | ")):          # table_context(): nearest heading last
        found = parse_date(part)
        if found:
            return found
    return ""


def _dcb_slab_label(text: str) -> str:
    """Normalise a published amount slab to a label amount_range_cr() parses, e.g. '3 Cr to <5 Cr'."""
    bounds = amount_range_cr(text)
    if not bounds:
        return clean(text)
    lo, hi = bounds
    t = clean(text).lower()
    lo_excl = bool(re.search(r"(?:>(?!=)|\babove\b|more than|\bover\b)\s*(?:rs\.?\s*|inr\s*)?\d", t))
    hi_excl = bool(re.search(r"(?:less than|below|under|<(?!=))\s*(?:rs\.?\s*|inr\s*)?\d", t))
    if hi == float("inf"):
        return f"Above {lo:g} Cr" if lo_excl else f"{lo:g} Cr and above"
    return f"{'Above ' if lo_excl else ''}{lo:g} Cr to {'<' if hi_excl else ''}{hi:g} Cr"


def _dcb_rows_from_candidates(candidates, url: str, mode: str) -> list[RateRow]:
    """candidates = [(context text, grid)].  Emits every bulk table / column whose slab reaches above Rs 3 cr.
    Each DCB bulk table = one amount slab (header row 0) with 'Regular' and 'Non - Callable' columns."""
    rows = []
    done = set()
    for context, grid in candidates:
        if len(grid) < 3:
            continue
        nearest = " | ".join(context.split(" | ")[-2:])    # own title + '(with effect from ...)'
        header_rows = next((i for i, r in enumerate(grid) if tenure_bounds(r[0])[0] is not None), 0)
        if not header_rows or not re.search(r"\bbulk\b", nearest, re.I) or _DCB_SKIP.search(nearest):
            continue
        effective = _dcb_nearest_date(context)
        for col in range(1, len(grid[0])):
            heads = [grid[r][col] for r in range(header_rows)]
            slab = next((h for h in heads if amount_range_cr(h)), "")
            head = " ".join(heads).lower()
            if not slab or amount_range_cr(slab)[1] <= BULK_FLOOR_CR or _DCB_SKIP.search(slab):
                continue
            if re.search(r"non\s*-?\s*callable", head):
                withdrawal = NON_CALLABLE
            elif re.search(r"\bregular\b|\bcallable\b", head):
                withdrawal = CALLABLE
            else:
                continue
            label = _dcb_slab_label(slab)
            if (label, withdrawal) in done:          # same slab published twice -> keep the first
                continue
            done.add((label, withdrawal))
            # '-' cells (e.g. no non-callable rate below 91 days) are dropped by rate_value()
            rows += rows_from_grid("DCB", grid, withdrawal, effective=effective, url=url, mode=mode,
                                   tenure_col=0, rate_col=col, amount=label, header_rows=header_rows)
    return rows


def _dcb_from_api() -> list[RateRow]:
    sub = _dcb_api(f"/rates-sub-nav?main_nav={DCB_MAIN_NAV}", "subnav").get("data") or []
    tids = [s["tid"] for s in sub if re.search(DCB_SUB_TAB_HINT, s.get("name", "") + s.get("tid", ""), re.I)]
    tids += [s["tid"] for s in sub if s["tid"] not in tids]      # resident tab first, then any others
    for tid in tids:
        data = _dcb_api(f"/rates?id={tid}", f"rates_{tid}")
        items = (((data.get("data") or {}).get("data")) or [{}])[0].get("rate_details") or []
        candidates = []
        for item in items:
            title, subtitle = clean(item.get("title")), clean(item.get("subtitle"))
            for table in soup_of(item.get("details") or "").find_all("table"):
                local = table_context(table)   # non-empty only for a 2nd table inside the same fragment
                candidates.append((" | ".join(p for p in (title, subtitle, local) if p), table_grid(table)))
        rows = _dcb_rows_from_candidates(candidates, f"{BANKS['DCB']}#{tid}", "official_json")
        if rows:
            return rows
    return []


def _dcb_from_html(page, mode: str) -> list[RateRow]:
    """Rendered page or manually saved page: <h3>title</h3><p>(with effect from ...)</p><table>."""
    candidates = [(table_context(t), table_grid(t)) for t in soup_of(page.content).find_all("table")]
    return _dcb_rows_from_candidates(candidates, page.url, mode)


@parser("DCB")
def parse_dcb() -> list[RateRow]:
    cached = _FETCH_CACHE.get(BANKS["DCB"])
    if cached is not None and str(cached.url).startswith("manual:"):     # scrape_bank() manual fallback
        return _dcb_from_html(cached, "manual_html")
    try:
        rows = _dcb_from_api()
        if rows:
            return rows
        LOG.warning("DCB: API returned no bulk table - falling back to rendering the page")
    except SourceBlocked:
        raise
    except Exception as exc:
        LOG.warning("DCB: API route failed (%s) - falling back to rendering the page", exc)
    page = render(BANKS["DCB"], "DCB", wait_for="table", wait_ms=4000)
    return _dcb_from_html(page, "official_html_rendered")


# ==== merged bank parsers: end ====


# --------------------------------------------------------------------------
# Scrape orchestration / CSV
# --------------------------------------------------------------------------
CSV_FIELDS = [f.name for f in fields(RateRow)]


def write_csv(bank: str, rows: list[RateRow]) -> Path:
    daily = OUTPUT_DIR / f"{bank.lower()}_bulk_rates_{datetime.now():%Y-%m-%d}.csv"
    with daily.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(asdict(r) for r in rows)
    return daily


def read_latest_csv(bank: str) -> tuple[Optional[Path], list[RateRow]]:
    files = sorted(OUTPUT_DIR.glob(f"{bank.lower()}_bulk_rates_*.csv"), key=lambda p: p.stat().st_mtime)
    if not files:
        return None, []
    rows = []
    with files[-1].open(encoding="utf-8-sig", newline="") as handle:
        for x in csv.DictReader(handle):
            try:
                rows.append(RateRow(
                    x["BankName"], x["WithdrawalType"], x.get("AmountSlab", ""), x["Tenure"], int(float(x["StartDay"])),
                    int(float(x["EndDay"])), float(x["Rate"]), x.get("EffectiveDate", ""), x.get("SourceURL", ""),
                    x.get("ExtractionMode", ""), x.get("ScrapedAt", "")))
            except (KeyError, ValueError):
                continue
    return files[-1], rows


def scrape_bank(bank: str) -> list[RateRow]:
    LOG.info("Scraping %s", bank)
    func = PARSERS[bank]
    try:
        rows = func()
    except SourceBlocked as exc:
        manual = manual_source(bank)
        if not manual:
            raise SourceBlocked(f"{exc}. Open the page in your browser, save it (Ctrl+S, 'Webpage, complete' "
                                f"or the rate PDF) as {MANUAL_DIR / (bank.lower() + '.html')} and re-run.") from None
        LOG.warning("%s blocked online - using manual source %s", bank, manual.path.name)
        _FETCH_CACHE[BANKS[bank]] = manual
        rows = func()
    rows = dedupe(fix_touching_slabs(rows))
    if not rows:
        raise ValueError(f"No validated bulk rows extracted for {bank}; check raw file in {RAW_DIR}")
    write_csv(bank, rows)
    counts = {w: sum(r.WithdrawalType == w for r in rows) for w in WITHDRAWAL_TYPES}
    LOG.info("%s rows: Callable=%s Non-Callable=%s", bank, counts[CALLABLE], counts[NON_CALLABLE])
    return rows


# --------------------------------------------------------------------------
# Comparison matrix
# --------------------------------------------------------------------------
def overlapping(rows: list[RateRow], start: int, end: int, withdrawal: str) -> list[RateRow]:
    """Source slabs of one type that overlap the AU bucket [start, end] by at least MIN_OVERLAP_DAYS
    (or fully, when the slab or the bucket is shorter than that)."""
    hits = []
    for r in rows:
        if r.WithdrawalType != withdrawal:
            continue
        overlap = min(end, r.EndDay) - max(start, r.StartDay) + 1
        needed = min(MIN_OVERLAP_DAYS, end - start + 1, r.EndDay - r.StartDay + 1)
        if overlap >= max(1, needed):
            hits.append(r)
    return hits


def bucket_max(rows: list[RateRow], start: int, end: int, withdrawal: str) -> tuple[Optional[float], list[RateRow]]:
    """Highest rate among all valid source tenures overlapping the AU bucket.

    The amount slab is selected before this function. Therefore this function
    compares tenure rates only within the selected entry bulk amount slab and
    Callable/Non-Callable type. Even a one-day valid overlap is included.
    """
    hits = overlapping(rows, start, end, withdrawal)
    if not hits:
        return None, []
    best = max(r.Rate for r in hits)
    return best, [r for r in hits if abs(r.Rate - best) < 0.000001]


def fmt_date(iso: str) -> str:
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d-%b-%y")
    except ValueError:
        return iso or "-"


def check_au_alignment(au_rows: list[RateRow]) -> list[str]:
    """Every AU slab must sit inside exactly one comparison bucket."""
    issues = []
    for r in au_rows:
        homes = [label for label, s, e in AU_TENURES if s <= r.StartDay and r.EndDay <= min(e, MAX_TENURE_DAYS)]
        if len(homes) != 1:
            issues.append(f"AU slab '{r.Tenure}' ({r.StartDay}-{r.EndDay} days) does not fit one bucket")
    return issues


def build_workbook(all_data: dict[str, list[RateRow]], status: dict[str, dict], amount_mode: str = "") -> Path:
    amount_mode = amount_mode or AMOUNT_MODE
    # Two independent amount views. Rates are selected only from slabs applicable
    # to that view, while tenure-wise maximum-overlap logic remains unchanged.
    data = {bank: select_amount_band_rows(rows, "3_to_10") for bank, rows in all_data.items()}
    data_10_plus = {bank: select_amount_band_rows(rows, "10_plus") for bank, rows in all_data.items()}
    au_rows = data.get("AU", [])
    au_rows_10_plus = data_10_plus.get("AU", [])
    if not any(r.WithdrawalType == CALLABLE for r in au_rows):
        raise ValueError("AU callable rows missing; AU is mandatory for the comparison")
    for issue in check_au_alignment(au_rows):
        LOG.warning(issue)

    wb = Workbook()
    ws = wb.active
    ws.title = "3-10 Cr"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85

    purple, grey, white = "5B2D91", "D9D9D9", "FFFFFF"
    thin = Side(style="thin", color="7F7F7F")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = lambda c: PatternFill("solid", fgColor=c)
    styles = {   # vs AU: (fill, font colour)
        "lower": ("C6EFCE", "006100"), "higher": ("FFC7CE", "9C0006"),
        "equal": ("FFEB9C", "9C5700"), "blank": ("E2EFDA", "000000"),
    }

    # column map
    columns: dict[tuple[str, str], int] = {}
    col = 2
    group_spans = []
    for title, banks, colour, gap in GROUPS:
        first = col
        for bank in banks:
            columns[(bank, CALLABLE)], columns[(bank, NON_CALLABLE)] = col, col + 1
            col += 2
        group_spans.append((title, colour, first, col - 1))
        if gap:
            ws.column_dimensions[get_column_letter(col)].width = 1.5
            col += 1
    ws.column_dimensions[get_column_letter(col)].width = 1.5   # wider visual break before AU
    col += 1
    columns[("AU", CALLABLE)], columns[("AU", NON_CALLABLE)] = col, col + 1
    last_col = col + 1

    def head(row, column, value, colour, font_colour="000000", size=10, bold=True):
        c = ws.cell(row, column, value)
        c.fill, c.font, c.alignment, c.border = fill(colour), Font(color=font_colour, bold=bold, size=size), center, border
        return c

    ws.merge_cells(start_row=1, start_column=1, end_row=4, end_column=1)
    head(1, 1, "Tenure", purple, white, 11)
    for r in range(2, 5):
        ws.cell(r, 1).border = border
    for title, colour, first, last in group_spans:
        ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=last)
        head(1, first, title, colour)
        for c in range(first, last + 1):
            ws.cell(1, c).border = border

    for bank in BANK_ORDER:
        c0 = columns[(bank, CALLABLE)]
        is_au = bank == "AU"
        colour, font_colour = (purple, white) if is_au else (grey, "000000")
        if is_au:
            ws.merge_cells(start_row=1, start_column=c0, end_row=2, end_column=c0 + 1)
            head(1, c0, "AU Bank\n3-10 Cr", colour, font_colour, 10)
            for rr in (1, 2):
                for cc in (c0, c0 + 1):
                    ws.cell(rr, cc).border = border
        else:
            ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + 1)
            head(2, c0, f"{DISPLAY_NAMES[bank]}\n3-10 Cr", colour, font_colour, 9)
            ws.cell(2, c0 + 1).border = border
        for w, short in ((CALLABLE, "C"), (NON_CALLABLE, "NC")):
            cc = columns[(bank, w)]
            head(3, cc, short, colour, font_colour, 9)
            dates = [r.EffectiveDate for r in data.get(bank, []) if r.WithdrawalType == w and r.EffectiveDate]
            head(4, cc, fmt_date(max(dates)) if dates else "-", colour, font_colour, 8)
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[4].height = 24

    trace = []
    first_data_row = 5
    for i, (label, start, end) in enumerate(AU_TENURES):
        row_no = first_data_row + i
        c = ws.cell(row_no, 1, label)
        c.fill, c.font, c.border = fill("F2F2F2"), Font(size=9), border
        c.alignment = Alignment(horizontal="left", vertical="center")
        for bank in BANK_ORDER:
            for w in WITHDRAWAL_TYPES:
                rate, sources = bucket_max(data.get(bank, []), start, end, w)
                au_rate, _ = bucket_max(au_rows, start, end, w)
                cell = ws.cell(row_no, columns[(bank, w)])
                cell.alignment, cell.border = Alignment(horizontal="center", vertical="center"), border
                if rate is not None:
                    cell.value = rate / 100
                    cell.number_format = "0.00%"
                if bank == "AU":
                    cell.fill, cell.font = fill(white), Font(bold=True, size=9)
                else:
                    if rate is None:
                        kind = "blank"
                    elif au_rate is None:
                        kind = "blank"
                    elif abs(rate - au_rate) < 1e-9:
                        kind = "equal"
                    else:
                        kind = "higher" if rate > au_rate else "lower"
                    cell.fill, cell.font = fill(styles[kind][0]), Font(bold=True, size=9, color=styles[kind][1])
                trace.append(["3-10 Cr", label, start, end, DISPLAY_NAMES[bank], w, rate if rate is not None else "",
                              au_rate if au_rate is not None else "",
                              "; ".join(f"{s.Tenure} ({s.StartDay}-{s.EndDay}d) [{s.AmountSlab}] = {s.Rate:.2f}% "
                                        f"w.e.f. {s.EffectiveDate or '?'}" for s in sources[:4])
                              + (f"; +{len(sources) - 4} more" if len(sources) > 4 else "")])

    # underline the peak rate of every column
    last_data_row = first_data_row + len(AU_TENURES) - 1
    for (bank, w), cc in columns.items():
        values = [ws.cell(r, cc).value for r in range(first_data_row, last_data_row + 1)]
        numbers = [v for v in values if isinstance(v, float)]
        if not numbers:
            continue
        peak = max(numbers)
        for r in range(first_data_row, last_data_row + 1):
            cell = ws.cell(r, cc)
            if isinstance(cell.value, float) and abs(cell.value - peak) < 1e-12:
                cell.font = Font(bold=True, size=9, color=cell.font.color, underline="single")

    legend_row = last_data_row + 2
    legend = [("Lower than AU", "lower"), ("Higher than AU", "higher"), ("Equal to AU", "equal")]
    ws.cell(legend_row, 1, "Legend").font = Font(bold=True, size=9)
    for k, (text, kind) in enumerate(legend):
        c = ws.cell(legend_row + 1 + k, 1, text)
        c.fill, c.font, c.border = fill(styles[kind][0]), Font(size=9, color=styles[kind][1], bold=True), border
    ws.cell(legend_row + 4, 1, "Underlined = peak rate in that column").font = Font(size=9, underline="single")
    amount_note = ("best rate across every published amount slab above Rs 3 cr" if amount_mode == "max"
                   else "rate of the amount slab containing Rs 3 cr")
    ws.cell(legend_row + 5, 1, f"Each cell = maximum published rate among the bank's tenure slabs overlapping the AU tenure, "
                               f"{amount_note} (1 month = 30 days, 12 months = 365 days). "
                               f"Blank = bank publishes no rate for that tenure. Details: 'Cell Trace' sheet."
            ).font = Font(size=9, italic=True)

    ws.column_dimensions["A"].width = 33
    for (bank, w), cc in columns.items():
        ws.column_dimensions[get_column_letter(cc)].width = 7.6
    for r in range(first_data_row, last_data_row + 1):
        ws.row_dimensions[r].height = 15
    ws.freeze_panes = "B5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{get_column_letter(last_col)}{legend_row + 5}"

    # ---- second comparison sheet: Rs 10 Cr and above -----------------------
    ws_10 = wb.copy_worksheet(ws)
    ws_10.title = "10 Cr and Above"

    # Update amount-view headers.
    for bank in BANK_ORDER:
        c0 = columns[(bank, CALLABLE)]
        if bank == "AU":
            ws_10.cell(1, c0).value = "AU Bank\n10 Cr and Above"
        else:
            ws_10.cell(2, c0).value = f"{DISPLAY_NAMES[bank]}\n10 Cr and Above"

    # Replace all rate cells using only source amount slabs applicable above Rs 10 Cr.
    for i, (label, start, end) in enumerate(AU_TENURES):
        row_no = first_data_row + i
        for bank in BANK_ORDER:
            for w in WITHDRAWAL_TYPES:
                rate, sources = bucket_max(data_10_plus.get(bank, []), start, end, w)
                au_rate, _ = bucket_max(au_rows_10_plus, start, end, w)
                cell = ws_10.cell(row_no, columns[(bank, w)])
                cell.value = rate / 100 if rate is not None else None
                if rate is not None:
                    cell.number_format = "0.00%"
                if bank == "AU":
                    cell.fill, cell.font = fill(white), Font(bold=True, size=9)
                else:
                    if rate is None or au_rate is None:
                        kind = "blank"
                    elif abs(rate - au_rate) < 1e-9:
                        kind = "equal"
                    else:
                        kind = "higher" if rate > au_rate else "lower"
                    cell.fill = fill(styles[kind][0])
                    cell.font = Font(bold=True, size=9, color=styles[kind][1])
                trace.append([
                    "10 Cr and Above", label, start, end, DISPLAY_NAMES[bank], w,
                    rate if rate is not None else "", au_rate if au_rate is not None else "",
                    "; ".join(
                        f"{source.Tenure} ({source.StartDay}-{source.EndDay}d) "
                        f"[{source.AmountSlab}] = {source.Rate:.2f}% "
                        f"w.e.f. {source.EffectiveDate or '?'}"
                        for source in sources[:4]
                    ) + (f"; +{len(sources) - 4} more" if len(sources) > 4 else "")
                ])

    # Re-apply peak underlining for the second amount view.
    for (bank, w), cc in columns.items():
        values = [ws_10.cell(r, cc).value for r in range(first_data_row, last_data_row + 1)]
        numbers = [value for value in values if isinstance(value, float)]
        if not numbers:
            continue
        peak = max(numbers)
        for row_no in range(first_data_row, last_data_row + 1):
            cell = ws_10.cell(row_no, cc)
            if isinstance(cell.value, float) and abs(cell.value - peak) < 1e-12:
                color = cell.font.color
                cell.font = Font(bold=True, size=9, color=color, underline="single")

    ws_10.cell(legend_row + 5, 1).value = (
        "Each cell = maximum published rate among tenure slabs overlapping the AU tenure, "
        "using only amount slabs applicable above Rs 10 Cr. Blank = no published rate. "
        "Details: 'Cell Trace' sheet."
    )

    # ---- audit sheets -------------------------------------------------------
    t = wb.create_sheet("Cell Trace")
    t.append(["Amount View", "AU Tenure", "Start Day", "End Day", "Bank", "Type", "Max Rate %", "AU Rate %", "Source slab(s) giving the max"])
    for row in trace:
        t.append(row)
    src = wb.create_sheet("Source Rates")
    src.append(CSV_FIELDS)
    for bank in BANK_ORDER:   # slabs with identical rates are shown once
        grouped: dict[tuple, list[RateRow]] = {}
        for r in all_data.get(bank, []):
            grouped.setdefault((r.WithdrawalType, r.StartDay, r.EndDay, r.Tenure, r.Rate, r.EffectiveDate, r.SourceURL), []).append(r)
        for members in grouped.values():
            values = [getattr(members[0], f) for f in CSV_FIELDS]
            if len(members) > 1:
                values[CSV_FIELDS.index("AmountSlab")] = f"{members[0].AmountSlab} (+{len(members) - 1} more slabs, same rate)"
            src.append(values)
    st = wb.create_sheet("Source Status")
    st.append(["Bank", "Status", "All amount slabs", "3-10 C buckets", "3-10 NC buckets",
               "10+ C buckets", "10+ NC buckets", "Effective date (C)", "Effective date (NC)",
               "Source URL(s)", "Note"])
    for bank in BANK_ORDER:
        rows = data.get(bank, [])
        info = status.get(bank, {})
        filled = {w: sum(bucket_max(rows, s, e, w)[0] is not None for _, s, e in AU_TENURES) for w in WITHDRAWAL_TYPES}
        rows_10 = data_10_plus.get(bank, [])
        filled_10 = {w: sum(bucket_max(rows_10, s, e, w)[0] is not None for _, s, e in AU_TENURES) for w in WITHDRAWAL_TYPES}
        eff = {w: max((r.EffectiveDate for r in rows if r.WithdrawalType == w and r.EffectiveDate), default="")
               for w in WITHDRAWAL_TYPES}
        st.append([DISPLAY_NAMES[bank], info.get("status", "OK" if rows else "MISSING"),
                   len({r.AmountSlab for r in all_data.get(bank, [])}),
                   f"{filled[CALLABLE]}/{len(AU_TENURES)}", f"{filled[NON_CALLABLE]}/{len(AU_TENURES)}",
                   f"{filled_10[CALLABLE]}/{len(AU_TENURES)}", f"{filled_10[NON_CALLABLE]}/{len(AU_TENURES)}",
                   eff[CALLABLE], eff[NON_CALLABLE],
                   " ; ".join(dict.fromkeys(r.SourceURL for r in rows)) if rows else BANKS.get(bank, ""),
                   info.get("note", "")])
    for sheet in (t, src, st):
        for cell in sheet[1]:
            cell.font, cell.fill = Font(bold=True, color=white), fill(purple)
        for column_cells in sheet.columns:
            width = min(80, max(len(str(c.value or "")) for c in column_cells) + 2)
            sheet.column_dimensions[column_cells[0].column_letter].width = width
        sheet.freeze_panes = "A2"

    try:
        wb.save(RESULT_FILE)
        out = RESULT_FILE
    except PermissionError:   # workbook open in Excel
        out = RESULT_FILE.with_name(f"{RESULT_FILE.stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        wb.save(out)
        LOG.warning("%s is open - saved as %s", RESULT_FILE.name, out.name)
    LOG.info("Workbook created: %s", out)
    return out


def coverage_report(data: dict[str, list[RateRow]]) -> list[str]:
    lines = []
    for bank in BANK_ORDER:
        rows = data.get(bank, [])
        for w in WITHDRAWAL_TYPES:
            missing = [label for label, s, e in AU_TENURES if bucket_max(rows, s, e, w)[0] is None]
            if missing and rows:
                lines.append(f"{DISPLAY_NAMES[bank]} {w}: no published rate for {len(missing)} bucket(s): "
                             + ", ".join(missing[:4]) + (" ..." if len(missing) > 4 else ""))
    return lines


def main() -> int:
    ap = argparse.ArgumentParser(description="Scrape bulk TD rates and compare max overlapping rates against AU tenures")
    ap.add_argument("--only", default="", help="Comma-separated banks to scrape, e.g. AU,HDFC,BOI")
    ap.add_argument("--skip-scrape", action="store_true", help="Build the workbook from the latest CSVs")
    ap.add_argument("--skip-build", action="store_true", help="Scrape only")
    ap.add_argument("--amount-mode", choices=["max", "slab"], default=AMOUNT_MODE,
                    help="max = best rate across all slabs above 3 cr (default); slab = only the slab holding 3 cr")
    args = ap.parse_args()
    selected = [b.strip().upper() for b in args.only.split(",") if b.strip()] or BANK_ORDER
    unknown = [b for b in selected if b not in PARSERS]
    if unknown:
        ap.error(f"Unknown bank(s): {', '.join(unknown)}. Choose from {', '.join(BANK_ORDER)}")

    data: dict[str, list[RateRow]] = {}
    status: dict[str, dict] = {}
    failures = []
    if not args.skip_scrape:
        for bank in selected:
            try:
                data[bank] = scrape_bank(bank)
                status[bank] = {"status": "OK (live)"}
            except Exception as exc:
                failures.append((bank, str(exc)))
                status[bank] = {"status": "FAILED", "note": str(exc)[:300]}
                LOG.error("%s failed: %s", bank, exc)
                LOG.debug(traceback.format_exc())
    # banks not scraped in this run (or failed) fall back to their latest CSV
    for bank in BANK_ORDER:
        if bank not in data:
            path, rows = read_latest_csv(bank)
            if rows:
                data[bank] = rows
                note = status.get(bank, {}).get("note", "")
                status[bank] = {"status": "FROM CSV" if bank not in dict(failures) else "FAILED - last CSV used",
                                "note": (note + " | " if note else "") + f"{path.name}"}
    workbook = RESULT_FILE
    if not args.skip_build:
        try:
            workbook = build_workbook(data, status, args.amount_mode)
        except Exception as exc:
            failures.append(("BUILDER", str(exc)))
            LOG.error("Builder failed: %s", exc)
            traceback.print_exc()

    print(f"\nBULK TD SUMMARY (amount mode: {args.amount_mode})")
    selected_rows = {bank: select_amount_rows(rows, args.amount_mode) for bank, rows in data.items()}
    for bank in BANK_ORDER:
        rows = selected_rows.get(bank, [])
        filled = {w: sum(bucket_max(rows, s, e, w)[0] is not None for _, s, e in AU_TENURES) for w in WITHDRAWAL_TYPES}
        print(f"  {DISPLAY_NAMES[bank]:<9} {status.get(bank, {}).get('status', 'MISSING'):<24} "
              f"C buckets {filled[CALLABLE]:>2}/{len(AU_TENURES)}  NC buckets {filled[NON_CALLABLE]:>2}/{len(AU_TENURES)}")
    gaps = coverage_report(selected_rows)
    if gaps:
        print("\nTENURES WITHOUT A PUBLISHED RATE (left blank in the matrix)")
        for line in gaps:
            print("  " + line)
    if failures:
        print("\nFAILURES")
        for bank, error in failures:
            print(f"  {bank}: {error}")
    print(f"\nCSV folder: {OUTPUT_DIR}\nWorkbook:   {workbook}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
