"""Pure cleaning functions (no file I/O) used by scripts/clean_sales_data.py.

Design rules
------------
* Explicit mappings only: every accepted label spelling is listed below. An
  unknown spelling is never guessed; it is reported.
* Line_ID (e.g. LAT-HOM01-202403-004) is the record key. It also carries the
  region, product and month, so it is used to recover blank labels/dates and
  to detect labels or dates that contradict it.
* A missing or wrong number is only repaired when an arithmetic identity
  gives the answer EXACTLY:
      Sales        = Qty x Unit_Price
      Gross_Profit = Sales - Qty x Unit_Cost
  Anything that cannot be solved exactly goes to review - nothing is estimated.
"""
import re
from datetime import datetime

import pandas as pd

from . import config
from .finance import money_text_to_cents

# ------------------------------------------------------------ text --
_ZERO_WIDTH = dict.fromkeys(map(ord, "​‌‍﻿"), None)


def clean_text(value):
    """Strip ordinary and hidden whitespace (non-breaking, zero-width). '' -> None."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:          # NaN
        return None
    s = str(value).translate(_ZERO_WIDTH).replace(" ", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s or None


NULL_TOKENS = {"null", "n/a", "na", "nan", "none", "-", "--", "#n/a", "missing"}
# For Region, "NA" is a real value (North America), so it is NOT a null token there.
NULL_TOKENS_REGION = NULL_TOKENS - {"na"}


def is_null(value, column=""):
    """True if the cell is empty or a text spelling of 'missing' for this column."""
    s = clean_text(value)
    if s is None:
        return True
    tokens = NULL_TOKENS_REGION if column == "Region" else NULL_TOKENS
    return s.lower() in tokens


def norm_key(value):
    """Lower-case letters and digits only: 'N. America ' -> 'namerica'."""
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


# ------------------------------------------------------- explicit maps --
REGION_ALIASES = {
    "northamerica": "North America", "namerica": "North America", "na": "North America", "nam": "North America",
    "europe": "Europe", "eu": "Europe", "europa": "Europe", "eur": "Europe",
    "asiapacific": "Asia Pacific", "apac": "Asia Pacific", "asiapac": "Asia Pacific", "ap": "Asia Pacific",
    "latinamerica": "Latin America", "latam": "Latin America",
    "middleeast": "Middle East", "me": "Middle East", "mideast": "Middle East",
}
PRODUCT_ALIASES = {norm_key(p): p for p in config.PRODUCTS}
CATEGORY_ALIASES = {norm_key(c): c for c in config.CATEGORIES}
NOTE_ALIASES = {
    "promotion": "Promotion", "promo": "Promotion",
    "stockout": "Stock-out",
    "lateshipment": "Late shipment",
    "returns": "Returns", "return": "Returns",
    "priceincrease": "Price increase",
    "costincrease": "Cost increase",
}


def map_label(value, aliases):
    """Return (standard value or None, cleaned original). None = blank or unknown."""
    s = clean_text(value)
    if s is None:
        return None, None
    return aliases.get(norm_key(s)), s


# ----------------------------------------------------------- Line_ID --
LINE_ID_RE = re.compile(r"^(NAM|EUR|APA|LAT|MEA)-([A-Z]{3}\d{2})-(\d{4})(\d{2})-(\d{3})$")


def parse_line_id(value):
    """'LAT-HOM01-202403-004' -> dict(region, product, period) or None if malformed.

    Line_ID is kept as TEXT everywhere so leading zeros ('004') are preserved.
    """
    s = clean_text(value)
    if s is None:
        return None
    m = LINE_ID_RE.match(s)
    if not m or m.group(2) not in config.PRODUCT_BY_CODE:
        return None
    month = int(m.group(4))
    if not 1 <= month <= 12:
        return None
    return {"line_id": s, "region": config.REGIONS[m.group(1)], "product": config.PRODUCT_BY_CODE[m.group(2)],
            "period": pd.Period(year=int(m.group(3)), month=month, freq="M")}


# ------------------------------------------------------------ numbers --
def parse_qty(value):
    """Parse a quantity. Returns (int or None, status) where status is
    'ok', 'text' (number stored as text), 'missing' or 'invalid'."""
    if is_null(value):
        return None, "missing"
    if isinstance(value, bool):
        return None, "invalid"
    if isinstance(value, int):
        return int(value), "ok"
    if isinstance(value, float):
        return (int(value), "ok") if value.is_integer() else (None, "invalid")
    s = clean_text(value).replace(",", "").replace(" ", "")
    if re.fullmatch(r"-?\d+", s):
        return int(s), "text"
    if re.fullmatch(r"-?\d+\.0+", s):
        return int(float(s)), "text"
    return None, "invalid"


def parse_money(value):
    """Parse an amount into integer cents. Returns (cents or None, status) like parse_qty.
    More than 2 decimals is 'invalid' (never silently rounded)."""
    if is_null(value):
        return None, "missing"
    if isinstance(value, bool):
        return None, "invalid"
    if isinstance(value, (int, float)):
        c = money_text_to_cents(repr(float(value)) if isinstance(value, float) else str(value))
        if c is None and isinstance(value, float):                 # e.g. 0.1+0.2 artefacts
            c = round(value * 100)
            if abs(value * 100 - c) > 1e-6:
                return None, "invalid"
        return (c, "ok") if c is not None else (None, "invalid")
    c = money_text_to_cents(clean_text(value))
    return (c, "text") if c is not None else (None, "invalid")


# --------------------------------------------------------------- dates --
_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def date_candidates(value):
    """All month-start dates a raw cell could mean. Returns (set of Periods, form)."""
    if is_null(value):
        return set(), "missing"
    if isinstance(value, (datetime, pd.Timestamp)):
        ts = pd.Timestamp(value)
        return ({ts.to_period("M")} if ts.day == 1 else set()), "date"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 < float(value) < 80000 and float(value).is_integer():
            ts = _EXCEL_EPOCH + pd.Timedelta(days=int(value))
            return ({ts.to_period("M")} if ts.day == 1 else set()), "excel serial"
        return set(), "invalid"
    s = clean_text(value)
    out = set()
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%b-%Y", "%B %Y", "%b %Y", "%Y/%m/%d"):
        try:
            ts = datetime.strptime(s, fmt)
            if ts.day == 1:
                out.add(pd.Period(year=ts.year, month=ts.month, freq="M"))
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2})[/.](\d{1,2})[/.](\d{4})", s)
    if m:                                                   # day-first and month-first readings
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        for day, month in ((a, b), (b, a)):
            if day == 1 and 1 <= month <= 12:
                out.add(pd.Period(year=y, month=month, freq="M"))
    return out, ("text" if out else "invalid")


# ------------------------------------------------ identities & repair --
SIDES = {
    "Target": dict(q="Target_Qty", p="Target_Unit_Price_USD", u="Target_Unit_Cost_USD",
                   s="Target_Sales_USD", g="Target_Gross_Profit_USD"),
    "Actual": dict(q="Actual_Qty", p="Actual_Unit_Price_USD", u="Actual_Unit_Cost_USD",
                   s="Actual_Sales_USD", g="Actual_Gross_Profit_USD"),
}


def identities_hold(v):
    """v: dict with q, p, u, s, g (ints). Sales = q*p and Gross profit = Sales - q*u."""
    return v["s"] == v["q"] * v["p"] and v["g"] == v["s"] - v["q"] * v["u"]


def _exact_div(num, den):
    if den is None or num is None or den <= 0 or num % den:
        return None
    return num // den


def fill_missing(v):
    """Fill missing values of one side using only exact identity solutions.

    v: dict q, p, u, s, g with ints or None. Returns (new dict, list of (key, rule)).
    Repeats until nothing more can be solved.
    """
    v = dict(v)
    filled = []
    changed = True
    while changed:
        changed = False
        q, p, u, s, g = v["q"], v["p"], v["u"], v["s"], v["g"]
        cands = []
        if s is None and q is not None and p is not None:
            cands.append(("s", q * p, "Sales = Qty x Unit_Price"))
        elif s is None and g is not None and q is not None and u is not None:
            cands.append(("s", g + q * u, "Sales = Gross_Profit + Qty x Unit_Cost"))
        if q is None and s is not None:
            x = _exact_div(s, p)
            if x is not None:
                cands.append(("q", x, "Qty = Sales / Unit_Price"))
            elif g is not None and u is not None:
                x = _exact_div(s - g, u)
                if x is not None:
                    cands.append(("q", x, "Qty = (Sales - Gross_Profit) / Unit_Cost"))
        if p is None and s is not None and q is not None:
            x = _exact_div(s, q)
            if x is not None:
                cands.append(("p", x, "Unit_Price = Sales / Qty"))
        if u is None and s is not None and g is not None and q is not None:
            x = _exact_div(s - g, q)
            if x is not None:
                cands.append(("u", x, "Unit_Cost = (Sales - Gross_Profit) / Qty"))
        if g is None and s is not None and q is not None and u is not None:
            cands.append(("g", s - q * u, "Gross_Profit = Sales - Qty x Unit_Cost"))
        for key, val, rule in cands[:1]:                     # one at a time, then re-evaluate
            v[key] = val
            filled.append((key, rule))
            changed = True
    return v, filled


def correct_single_error(v):
    """If the identities fail, test the hypothesis 'exactly one field is wrong'
    for each field. Returns (key, corrected value, rule) when exactly ONE field
    explains the failure, otherwise None (the row must be reviewed)."""
    q, p, u, s, g = v["q"], v["p"], v["u"], v["s"], v["g"]
    hyps = []
    for key in ("q", "p", "s", "g", "u"):
        t = dict(v)
        if key == "q":
            x = _exact_div(s, p)
            if x is None and u:
                x = _exact_div(s - g, u)
            rule = "Qty = Sales / Unit_Price"
        elif key == "p":
            x, rule = _exact_div(s, q), "Unit_Price = Sales / Qty"
        elif key == "s":
            x, rule = q * p, "Sales = Qty x Unit_Price"
        elif key == "g":
            x, rule = s - q * u, "Gross_Profit = Sales - Qty x Unit_Cost"
        else:
            x, rule = _exact_div(s - g, q), "Unit_Cost = (Sales - Gross_Profit) / Qty"
        if x is None or x == v[key]:
            continue
        if key == "q" and x < 0:
            continue
        if key in ("p", "u") and x <= 0:
            continue
        t[key] = x
        if identities_hold(t):
            hyps.append((key, x, rule))
    return hyps[0] if len(hyps) == 1 else None
