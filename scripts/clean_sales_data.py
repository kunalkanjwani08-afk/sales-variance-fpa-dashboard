# -*- coding: utf-8 -*-
# =============================================================================
#  SALES DATA CLEANING PIPELINE  -  sales_raw_messy_csv.xlsx
#  Run in VS Code:  open this file -> press the "Run Python File" (play) button,
#  or in a terminal:   python clean_sales_data.py   [optional: path\to\file.xlsx]
#
#  Repo layout: run from anywhere; the raw file is read from  data/raw/  and the
#  outputs are written to  data/processed/  (when this script sits in scripts/).
#  Otherwise the outputs are saved next to this script.
#
#  Output:
#       Cleaned_Sales_Data.xlsx   multi-sheet, formatted, with a full audit trail
#       Cleaned_Sales_Data.csv    clean data only (Power BI / SQL / Python / AI tools)
# =============================================================================

# ----------------------------------------------------------------------------
# 0. IMPORTS & SETUP
# ----------------------------------------------------------------------------
import os
import re
import sys
import glob
import time
import difflib
import datetime as _dt
import warnings
import subprocess
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)
try:                                    # keep the Windows console from crashing on odd characters
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

T0 = time.time()


def banner(title):
    print("\n" + "=" * 78)
    print(title)
    print("=" * 78)


def ensure_package(import_name, pip_name=None):
    """Import a package; if it is missing, pip-install it into the current Python."""
    try:
        __import__(import_name)
        return True
    except ImportError:
        try:
            print(f"   Installing missing package '{pip_name or import_name}' ...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pip_name or import_name])
            __import__(import_name)
            return True
        except Exception as exc:                                  # pragma: no cover
            print(f"   Could not install {pip_name or import_name}: {exc}")
            return False


ensure_package("openpyxl")
HAS_XLSXWRITER = ensure_package("xlsxwriter", "XlsxWriter")

# ----------------------------------------------------------------------------
# 1. CONFIGURATION  (the only part you might ever want to edit)
# ----------------------------------------------------------------------------
INPUT_FILE_NAME = "sales_raw_messy_csv.xlsx"
try:
    SCRIPT_DIR = Path(__file__).resolve().parent
except NameError:                                                  # e.g. VS Code interactive window
    SCRIPT_DIR = Path.cwd()
# Repo layout (scripts/ + data/raw + data/processed). Path wiring only - no cleaning logic.
IN_REPO_SCRIPTS = SCRIPT_DIR.name.lower() == "scripts"
REPO_ROOT = SCRIPT_DIR.parent if IN_REPO_SCRIPTS else SCRIPT_DIR
RAW_DIR = REPO_ROOT / "data" / "raw"
OUTPUT_DIR = (REPO_ROOT / "data" / "processed") if IN_REPO_SCRIPTS else SCRIPT_DIR
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_XLSX = OUTPUT_DIR / "Cleaned_Sales_Data.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "Cleaned_Sales_Data.csv"

DROP_NOTES = False               # False = keep the (mostly empty) Notes column as an optional field
INCLUDE_SOURCE_ROW = True        # keep "Source_Excel_Row" so every clean row can be traced to the raw file
REMOVE_LOGICAL_DUPLICATES = True # remove copies that were "noised" differently (see STEP 5)
RECOVER_REGION_BY_BLOCK = True   # rebuild blank Regions from the proven block layout of the file
OPEN_FOLDER_WHEN_DONE = True     # open the output folder when the script finishes
VERIFY_OUTPUT_FILE = True        # re-open the saved Excel file and double-check it
DAYFIRST_DEFAULT = True          # only used for text dates like 03/04/2024 that cannot be resolved otherwise

MONEY_TOL = 0.02                 # rounding tolerance for money identities (2 cents)
PCT_TOL = 0.011                  # rounding tolerance for percentage identities
QTY_INT_TOL = 0.001              # Actual_Sales / Unit_Price must be a whole number within this tolerance
PRICE_TOL = 0.0051               # Target_Sales / Target_Qty must equal Unit_Price within this tolerance

EXPECTED_COLUMNS = [
    "Date", "Year", "Month", "Month_Name", "Region", "Category", "Product", "Target_Qty",
    "Actual_Qty", "Unit_Price_USD", "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD",
    "Gross_Margin_USD", "Gross_Margin_Pct", "Notes",
]
TEXT_COLS = ["Month_Name", "Region", "Category", "Product", "Notes"]
NUMERIC_COLS = ["Year", "Month", "Target_Qty", "Actual_Qty", "Unit_Price_USD", "Unit_Cost_USD",
                "Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_USD", "Gross_Margin_Pct"]
CANON_REGIONS = ["North America", "Europe", "Asia Pacific", "Latin America", "Middle East"]

# Region spellings -> standard region. Keys are lower-case with all spaces/dots/hyphens removed,
# so "N. America", "North-America", "Northamerica" and "North America " all become "northamerica".
REGION_ALIASES = {
    "northamerica": "North America", "namerica": "North America", "na": "North America",
    "nam": "North America",
    "europe": "Europe", "eu": "Europe", "europa": "Europe", "eur": "Europe",
    "asiapacific": "Asia Pacific", "apac": "Asia Pacific", "asiapac": "Asia Pacific", "ap": "Asia Pacific",
    "latinamerica": "Latin America", "latam": "Latin America", "lam": "Latin America",
    "middleeast": "Middle East", "me": "Middle East", "mideast": "Middle East",
}
# Text that really means "empty". NOTE: "NA" is deliberately NOT in the Region list, because for
# Region "NA" means North America (pandas silently turns "NA" into a null - this script does not).
NULL_TOKENS_COMMON = {"", "nan", "none", "null", "n/a", "#n/a", "nil", "-", "--", "?", "missing",
                      "unknown", "#null!", "<na>", "nat"}
NULL_TOKENS_DEFAULT = NULL_TOKENS_COMMON | {"na"}
NULL_TOKENS_REGION = set(NULL_TOKENS_COMMON)

# A row's identity: these six columns were never damaged in the source (verified), and together they
# are unique for every genuine row. Copies of a row share them even when other cells were corrupted.
DEDUP_KEY = ["Date", "Product", "Target_Qty", "Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_Pct"]
CONFLICT_COLS = ["Region", "Category", "Unit_Price_USD", "Unit_Cost_USD", "Gross_Margin_USD"]

# ----------------------------------------------------------------------------
# 2. LOGGING HELPERS
# ----------------------------------------------------------------------------
LOG = []              # every cleaning action
CORRECTIONS = []      # cell-level audit trail (old value -> new value)
REVIEW_PARTS = []     # rows that could not be repaired with certainty (never silently dropped)
MAPPING_ROWS = []     # label standardisation report


def log(step, column, issue, rows, action, kind="Fix"):
    LOG.append({"Step": step, "Column": column, "Issue_Found": issue,
                "Rows_Affected": int(rows), "Action_Taken": action, "Type": kind})
    if rows or kind in ("Observation", "Warning", "Decision"):
        print(f"   [{kind:<11}] {column:<18} {issue} -> {int(rows):,} rows | {action}")


def send_to_review(frame, mask, reason):
    mask = pd.Series(mask, index=frame.index).fillna(False).astype(bool)
    if mask.any():
        part = frame.loc[mask].copy()
        part["Review_Reason"] = reason
        REVIEW_PARTS.append(part)
    return frame.loc[~mask].copy()


def audit(frame, mask, field, old, new, method):
    """Record cell-level corrections (for the Corrections_Audit sheet)."""
    idx = frame.index[mask]
    if len(idx) == 0:
        return
    CORRECTIONS.append(pd.DataFrame({
        "Source_Excel_Row": frame.loc[idx, "_src_row"].astype(int).values,
        "Date": frame.loc[idx, "Date"].values,
        "Region": frame.loc[idx, "Region"].values,
        "Product": frame.loc[idx, "Product"].values,
        "Field": field,
        "Old_Value": pd.Series(old).loc[idx].values,
        "New_Value": pd.Series(new).loc[idx].values,
        "Method": method,
    }))


# ----------------------------------------------------------------------------
# 3. LOCATE & LOAD THE FILE  (safe read: the text "NA" is NOT turned into a null)
# ----------------------------------------------------------------------------
banner("STEP 1/12  -  LOCATING AND LOADING THE RAW FILE")


def find_input_file():
    if len(sys.argv) > 1:
        cand = Path(sys.argv[1].strip('"'))
        if cand.is_file() and cand.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".csv"):
            return cand
    folders = [RAW_DIR, SCRIPT_DIR, Path.cwd(), Path.home() / "Downloads", Path.home() / "Desktop",
               Path.home() / "Documents"]
    for folder in folders:
        p = folder / INPUT_FILE_NAME
        if p.is_file():
            return p
    for folder in folders:
        hits = sorted(glob.glob(str(folder / "sales_raw_messy*.*")))
        hits = [h for h in hits if h.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv"))]
        if hits:
            return Path(hits[0])
    try:                                                        # file-picker window
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        chosen = filedialog.askopenfilename(title="Select sales_raw_messy_csv.xlsx",
                                            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xls *.csv")])
        root.destroy()
        if chosen:
            return Path(chosen)
    except Exception:
        pass
    try:
        typed = input("Could not find the file automatically. Paste the full path to it: ").strip().strip('"')
    except (EOFError, KeyboardInterrupt):
        typed = ""
    if typed and Path(typed).is_file():
        return Path(typed)
    raise FileNotFoundError(
        f"Could not find '{INPUT_FILE_NAME}'. Put it in the same folder as this script "
        f"({SCRIPT_DIR}) and run again, or pass the path:  python clean_sales_data.py \"C:\\path\\file.xlsx\"")


def read_raw(path):
    if path.suffix.lower() == ".csv":
        last = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(path, dtype=object, keep_default_na=False, na_values=[""], encoding=enc)
            except UnicodeDecodeError as exc:
                last = exc
        raise last
    return pd.read_excel(path, sheet_name=0, dtype=object, keep_default_na=False, na_values=[""],
                         engine="openpyxl")


INPUT_PATH = find_input_file()
print(f"Reading: {INPUT_PATH}")
raw = read_raw(INPUT_PATH)
raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]
print(f"Loaded {raw.shape[0]:,} rows x {raw.shape[1]} columns")
missing_cols = [c for c in EXPECTED_COLUMNS if c not in raw.columns]
if missing_cols:
    raise KeyError(f"These expected columns are missing from the file: {missing_cols}. Found: {list(raw.columns)}")
extra_cols = [c for c in raw.columns if c not in EXPECTED_COLUMNS]
if extra_cols:
    print(f"   (ignoring unexpected extra columns: {extra_cols})")
raw = raw[EXPECTED_COLUMNS].copy()
raw["_src_row"] = np.arange(len(raw)) + 2                       # row number in the Excel sheet
RAW_ROWS = len(raw)
raw = raw.dropna(how="all", subset=EXPECTED_COLUMNS).copy() if raw[EXPECTED_COLUMNS].isna().all(axis=1).any() else raw
BLANK_ROWS_DROPPED = RAW_ROWS - len(raw)
if BLANK_ROWS_DROPPED:
    log("1. Load", "ALL", "Completely empty rows", BLANK_ROWS_DROPPED, "Removed", "Fix")

# ----------------------------------------------------------------------------
# 4. "BEFORE" PROFILE
# ----------------------------------------------------------------------------
def dominant_type(series):
    t = series.dropna().map(lambda v: type(v).__name__)
    return t.value_counts().index[0] if len(t) else "empty"


def make_profile(frame, cols, tag):
    rows = []
    n = len(frame)
    for c in cols:
        s = frame[c]
        miss = int(s.isna().sum())
        rows.append({"Column": c, f"Type_{tag}": dominant_type(s) if tag == "Before" else str(s.dtype),
                     f"Missing_{tag}": miss, f"Missing_%_{tag}": round(100.0 * miss / n, 2) if n else 0.0,
                     f"Unique_{tag}": int(s.nunique(dropna=True))})
    return pd.DataFrame(rows)


banner("STEP 2/12  -  PROFILING THE RAW DATA (BEFORE)")
profile_before = make_profile(raw, EXPECTED_COLUMNS, "Before")
print(profile_before.to_string(index=False))
na_text_region = int(raw["Region"].map(lambda v: isinstance(v, str) and v.strip().upper() == "NA").sum())
if na_text_region:
    log("2. Profile", "Region", "Literal text 'NA' (means North America; pandas would silently read it as a null)",
        na_text_region, "Preserved and mapped to North America", "Observation")

# ----------------------------------------------------------------------------
# 5. EXACT DUPLICATE ROWS
# ----------------------------------------------------------------------------
banner("STEP 3/12  -  REMOVING EXACT DUPLICATE ROWS")
dup_mask = raw.duplicated(subset=EXPECTED_COLUMNS, keep="first")
removed_exact = raw.loc[dup_mask].copy()
removed_exact["Removal_Reason"] = "Exact duplicate of an earlier row (all 16 columns identical)"
removed_exact["Duplicate_Of_Excel_Row"] = np.nan
df = raw.loc[~dup_mask].copy().reset_index(drop=True)
EXACT_DUPS = int(dup_mask.sum())
log("3. Duplicates", "ALL", "Exact duplicate rows (all 16 columns identical)", EXACT_DUPS,
    "Removed - first occurrence kept (see 'Removed_Duplicates' sheet)", "Fix")
snapshot = df[["Region", "Category", "Product", "Notes"]].copy()   # raw text for the mapping report

# ----------------------------------------------------------------------------
# 6. HIDDEN CHARACTERS / WHITESPACE / FAKE NULLS
# ----------------------------------------------------------------------------
banner("STEP 4/12  -  CLEANING HIDDEN SPACES, INVISIBLE CHARACTERS AND FAKE NULLS")
INVISIBLE = re.compile("[\u200b\u200c\u200d\u2060\ufeff\u00ad]")
WHITESPACE = re.compile(r"\s+")


def clean_text(x, null_tokens):
    if x is None:
        return np.nan
    if not isinstance(x, str):
        if pd.isna(x):
            return np.nan
        x = str(x)
    s = WHITESPACE.sub(" ", INVISIBLE.sub("", x)).strip()
    return np.nan if s.lower() in null_tokens else s


for col in TEXT_COLS:
    tokens = NULL_TOKENS_REGION if col == "Region" else NULL_TOKENS_DEFAULT
    before = df[col].copy()
    cleaned = before.map(lambda v: clean_text(v, tokens)).astype(object)
    both = before.notna() & cleaned.notna()
    changed = int((both & (before.astype(str) != cleaned.astype(str))).sum())
    newly_null = int((before.notna() & cleaned.isna()).sum())
    df[col] = cleaned
    if changed:
        log("4. Text hygiene", col, "Leading/trailing/double spaces or hidden characters", changed,
            "Trimmed and collapsed to single spaces", "Fix")
    if newly_null:
        log("4. Text hygiene", col, "Fake-null text ('N/A', 'null', '-', blank ...)", newly_null,
            "Converted to a real missing value", "Fix")

# ----------------------------------------------------------------------------
# 7. DATA TYPES (numbers) AND DATES
# ----------------------------------------------------------------------------
banner("STEP 5/12  -  FIXING DATA TYPES AND DATES")


def to_numeric_safe(series):
    def conv(v):
        if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
            return v
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return np.nan
        t = str(v).strip()
        neg = t.startswith("(") and t.endswith(")")
        t = re.sub(r"[,\s$\u20ac\u00a3\u20b9%()]|USD", "", t, flags=re.I)
        if t.lower() in NULL_TOKENS_DEFAULT:
            return np.nan
        try:
            val = float(t)
        except ValueError:
            return np.nan
        return -val if neg else val
    return pd.to_numeric(series.map(conv), errors="coerce")


N_NUMTEXT = 0
for col in NUMERIC_COLS:
    before_nn = int(df[col].notna().sum())
    was_text = int(df[col].map(lambda v: isinstance(v, str)).sum())
    df[col] = to_numeric_safe(df[col]).astype("float64")
    lost = before_nn - int(df[col].notna().sum())
    N_NUMTEXT += was_text
    if was_text:
        log("5. Data types", col, "Numbers stored as text (e.g. '1,234' or '$12.50')", was_text,
            "Converted to real numbers", "Fix")
    if lost:
        log("5. Data types", col, "Values that are not valid numbers", lost, "Converted to missing", "Fix")
    if not was_text and not lost:
        log("5. Data types", col, "Stored as text / invalid numbers", 0,
            "None found - already numeric", "Check")

_DATE_ISO = re.compile(r"^\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")
_DATE_DMY = re.compile(r"^\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})")


def _make_ts(y, m, d):
    try:
        return pd.Timestamp(int(y), int(m), int(d))
    except (ValueError, TypeError):
        return pd.NaT


def parse_one_date(v, month_hint):
    """Return (Timestamp, how). how: ok | iso | dmy | mdy | ambiguous | serial | failed"""
    if isinstance(v, (_dt.datetime, pd.Timestamp)):
        return pd.Timestamp(v).normalize(), "ok"
    if isinstance(v, _dt.date):
        return pd.Timestamp(v), "ok"
    if isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v):
        if 20000 < float(v) < 80000:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v)), "serial"
        return pd.NaT, "failed"
    if isinstance(v, str):
        m = _DATE_ISO.match(v)
        if m:
            return _make_ts(m.group(1), m.group(2), m.group(3)), "iso"
        m = _DATE_DMY.match(v)
        if m:
            a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if a > 12:
                return _make_ts(y, b, a), "dmy"
            if b > 12:
                return _make_ts(y, a, b), "mdy"
            if a == b:                                              # e.g. 01/01/2024 - same either way
                return _make_ts(y, b, a), "dmy"
            if month_hint in (a, b) and a != b:                     # Month column decides
                return (_make_ts(y, b, a), "dmy") if month_hint == b else (_make_ts(y, a, b), "mdy")
            return (_make_ts(y, b, a) if DAYFIRST_DEFAULT else _make_ts(y, a, b)), "ambiguous"
        ts = pd.to_datetime(v, errors="coerce", dayfirst=DAYFIRST_DEFAULT)
        return (ts.normalize() if pd.notna(ts) else pd.NaT), ("iso" if pd.notna(ts) else "failed")
    return pd.NaT, "failed"


hints = df["Month"].map(lambda m: int(m) if pd.notna(m) else -1).tolist()
parsed = [parse_one_date(v, h) for v, h in zip(df["Date"].tolist(), hints)]
df["Date"] = pd.to_datetime(pd.Series([p[0] for p in parsed], index=df.index), errors="coerce")
how = pd.Series([p[1] for p in parsed], index=df.index)
n_text_dates = int((how != "ok").sum())
n_bad_dates = int(df["Date"].isna().sum())
n_ambig = int((how == "ambiguous").sum())
log("5. Dates", "Date", "Dates stored as text / Excel serial numbers (mixed formats)", n_text_dates,
    "Parsed (YYYY-MM-DD and DD/MM/YYYY understood; Month column used to resolve ambiguity)"
    if n_text_dates else "None - every Date is already a true Excel date", "Fix" if n_text_dates else "Check")
if n_ambig:
    log("5. Dates", "Date", "Ambiguous text dates (both parts <= 12)", n_ambig,
        "Resolved with DAYFIRST_DEFAULT - please verify", "Warning")
log("5. Dates", "Date", "Unparseable / missing dates", n_bad_dates,
    "Moved to Review_Required" if n_bad_dates else "None", "Fix" if n_bad_dates else "Check")

# Day/month swap detector: Excel often reads DD/MM/YYYY as MM/DD/YYYY. The Month column exposes it.
mcol = df["Month"]
swap = (df["Date"].notna() & mcol.notna() & (df["Date"].dt.month != mcol)
        & (df["Date"].dt.day == mcol) & (df["Date"].dt.month <= 12))
if swap.any():
    fixed = [_make_ts(d.year, d.day, d.month) for d in df.loc[swap, "Date"]]
    df.loc[swap, "Date"] = pd.Series(fixed, index=df.index[swap])
log("5. Dates", "Date", "Day and month swapped (Month column disagrees with the date)", int(swap.sum()),
    "Swapped back using the Month column" if swap.any() else "None found", "Fix" if swap.any() else "Check")
df["Date"] = df["Date"].dt.normalize()
df = send_to_review(df, df["Date"].isna(), "Missing or unparseable Date")
DATE_SPAN = (df["Date"].min(), df["Date"].max())
ALL_FIRST_OF_MONTH = bool((df["Date"].dt.day == 1).all())
distinct_dates = int(df["Date"].nunique())
log("5. Dates", "Date", f"Date range {DATE_SPAN[0]:%Y-%m-%d} to {DATE_SPAN[1]:%Y-%m-%d}; distinct dates = {distinct_dates}",
    0, "Informational: all dates fall on the 1st of a month (monthly data), so no DD/MM ambiguity exists"
    if (df["Date"].dt.day == 1).all() else "Informational", "Observation")

# ----------------------------------------------------------------------------
# 8. STANDARDISING REGION / CATEGORY / PRODUCT / NOTES
# ----------------------------------------------------------------------------
banner("STEP 6/12  -  STANDARDISING REGION, CATEGORY, PRODUCT AND NOTES")


def norm_key(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def standardise_region(series):
    keymap = {}
    alias_keys = list(REGION_ALIASES.keys())
    unmapped = {}
    out = []
    for v in series:
        if not isinstance(v, str):
            out.append(np.nan)
            continue
        k = norm_key(v)
        if k not in keymap:
            if k in REGION_ALIASES:
                keymap[k] = REGION_ALIASES[k]
            else:
                m = difflib.get_close_matches(k, [a for a in alias_keys if len(a) > 3], n=1, cutoff=0.86)
                keymap[k] = REGION_ALIASES[m[0]] if m else None
        res = keymap[k]
        if res is None:
            unmapped[v] = unmapped.get(v, 0) + 1
            out.append(np.nan)
        else:
            out.append(res)
    return pd.Series(out, index=series.index, dtype=object), unmapped


def smart_title(s):
    return re.sub(r"[A-Za-z]+(?:'[A-Za-z]+)?", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), s)


def standardise_labels(series, prefer="freq"):
    """Case-insensitive unification. prefer='freq': most frequent spelling; 'title': proper Title Case."""
    s = series.astype(object)
    key = s.map(lambda v: v.lower() if isinstance(v, str) else np.nan)
    tmp = pd.DataFrame({"k": key, "raw": s}).dropna()
    if tmp.empty:
        return s
    sp = tmp.groupby(["k", "raw"]).size().reset_index(name="n")
    sp["is_title"] = sp["raw"].map(lambda v: v == smart_title(v))
    if prefer == "title":
        sp = sp.sort_values(["k", "is_title", "n", "raw"], ascending=[True, False, False, True])
    else:
        sp = sp.sort_values(["k", "n", "is_title", "raw"], ascending=[True, False, False, True])
    best = sp.drop_duplicates("k").set_index("k")["raw"].to_dict()
    if prefer == "title":
        best = {k: (v if v == smart_title(v) else smart_title(v)) for k, v in best.items()}
    return key.map(lambda k: best.get(k, np.nan) if isinstance(k, str) else np.nan).astype(object)


flag_cols = ["region_std", "region_rec", "cat_std", "prod_std", "notes_std", "merged", "price_rec",
             "qty_missing", "qty_sign", "qty_scaled", "qty_other", "cost_rec", "cost_est", "gm_rec", "conflict"]
FLAGS = pd.DataFrame(False, index=df.index, columns=flag_cols)

# --- Region ---
reg_before = df["Region"].copy()
reg_after, reg_unmapped = standardise_region(reg_before)
chg = reg_before.notna() & reg_after.notna() & (snapshot.loc[df.index, "Region"].astype(str) != reg_after.astype(str))
df["Region"] = reg_after
FLAGS.loc[chg[chg].index, "region_std"] = True
N_REG_VARIANTS = int(snapshot.loc[df.index, "Region"].dropna().nunique())
N_REG_FIX = int(chg.sum())
log("6. Region", "Region", f"Dirty spellings ({N_REG_VARIANTS} different spellings found, e.g. NA, N. America, Europa, APAC, LatAm, ME)",
    N_REG_FIX, "Standardised to the 5 official regions", "Fix")
if reg_unmapped:
    log("6. Region", "Region", f"Unrecognised labels {reg_unmapped}", sum(reg_unmapped.values()),
        "Treated as missing (will be recovered or sent to review)", "Warning")
print("      Region variants -> ", sorted(df["Region"].dropna().unique().tolist()))

# --- Category ---
cat_before = df["Category"].copy()
df["Category"] = standardise_labels(cat_before, prefer="freq")
chg = snapshot.loc[df.index, "Category"].notna() & (snapshot.loc[df.index, "Category"].astype(str) != df["Category"].astype(str))
FLAGS.loc[chg[chg].index, "cat_std"] = True
N_CAT_FIX = int(chg.sum())
log("6. Category", "Category", "Leading/trailing spaces or case differences", int(chg.sum()),
    f"Standardised to {df['Category'].nunique()} clean categories", "Fix" if chg.any() else "Check")
print("      Category -> ", sorted(df["Category"].dropna().unique().tolist()))

# --- Product ---
df["Product"] = standardise_labels(df["Product"], prefer="title")
chg = snapshot.loc[df.index, "Product"].notna() & (snapshot.loc[df.index, "Product"].astype(str) != df["Product"].astype(str))
FLAGS.loc[chg[chg].index, "prod_std"] = True
N_PROD_FIX = int(chg.sum())
N_PRODUCTS = int(df["Product"].nunique())
log("6. Product", "Product", "Mixed case (lower / UPPER) and spacing", int(chg.sum()),
    f"Standardised to proper Title Case ({df['Product'].nunique()} products)", "Fix" if chg.any() else "Check")

# --- Notes ---
df["Notes"] = standardise_labels(df["Notes"], prefer="freq")
chg = snapshot.loc[df.index, "Notes"].notna() & (snapshot.loc[df.index, "Notes"].astype(str) != df["Notes"].astype(str))
FLAGS.loc[chg[chg].index, "notes_std"] = True

# mapping report (original -> standard)
for col in ["Region", "Category", "Product"]:
    orig = snapshot.loc[df.index, col]
    diff = orig.notna() & (orig.astype(str) != df[col].astype(str))
    if diff.any():
        mp = (pd.DataFrame({"Original_Value": orig[diff].map(repr), "Standardised_Value": df.loc[diff, col]})
              .groupby(["Original_Value", "Standardised_Value"]).size().reset_index(name="Row_Count"))
        mp.insert(0, "Column", col)
        MAPPING_ROWS.append(mp)

# Product -> Category must be a one-to-one relationship in a clean file
pc = df.dropna(subset=["Category"]).groupby("Product")["Category"].nunique()
multi = pc[pc > 1]
log("6. Product", "Product/Category", "Products that appear under more than one Category", len(multi),
    "None - every product belongs to exactly one category" if multi.empty else f"Review: {list(multi.index)}",
    "Check" if multi.empty else "Warning")
for col in ["Category", "Product", "Month_Name"]:
    n_missing = int(df[col].isna().sum())
    if col != "Month_Name" and n_missing:
        log("6. Missing", col, "Missing", n_missing, "Moved to Review_Required", "Fix")
        df = send_to_review(df, df[col].isna(), f"Missing {col}")
FLAGS = FLAGS.loc[df.index]

# ----------------------------------------------------------------------------
# 9. LOGICAL DUPLICATES  (copies that were damaged differently, so they are no longer "exact")
# ----------------------------------------------------------------------------
banner("STEP 7/12  -  REMOVING LOGICAL (HIDDEN) DUPLICATES")
LOGICAL_DUPS = 0
removed_logical = pd.DataFrame()
if REMOVE_LOGICAL_DUPLICATES:
    key_ok = df[DEDUP_KEY].notna().all(axis=1)
    sub = df.loc[key_ok].copy()
    sub["_g"] = sub.groupby(DEDUP_KEY, sort=False).ngroup()
    is_copy = sub["_g"].duplicated(keep="first")
    multi_groups = sub.loc[sub["_g"].duplicated(keep=False), "_g"].unique()
    if len(multi_groups):
        nun = sub[sub["_g"].isin(multi_groups)].groupby("_g")[CONFLICT_COLS].nunique(dropna=True)
        conflict_groups = set(nun.index[(nun > 1).any(axis=1)])
        notes_conf = int((sub[sub["_g"].isin(multi_groups)].groupby("_g")["Notes"].nunique(dropna=True) > 1).sum())
        conflict_copy = is_copy & sub["_g"].isin(conflict_groups)
        merge_copy = is_copy & ~sub["_g"].isin(conflict_groups)
        # fill blanks of the kept (first) row from its copies, then drop the copies
        fill_cols = ["Region", "Unit_Price_USD", "Unit_Cost_USD", "Gross_Margin_USD", "Notes"]
        firsts = sub[sub["_g"].isin(multi_groups) & ~sub["_g"].isin(conflict_groups)].groupby("_g")[fill_cols].first()
        kept_multi = (~is_copy) & sub["_g"].isin(firsts.index)
        for c in fill_cols:
            m = kept_multi & sub[c].isna()
            if m.any():
                vals = firsts.loc[sub.loc[m, "_g"], c].values
                filled = pd.Series(vals, index=sub.index[m])
                good = filled.notna()
                if good.any():
                    df.loc[good[good].index, c] = filled[good]
                    if c == "Region":
                        log("7. Logical dups", c, "Blank in the kept row but present in its copy", int(good.sum()),
                            "Filled from the copy", "Fix")
        FLAGS.loc[sub.index[kept_multi], "merged"] = True
        first_row = sub.groupby("_g")["_src_row"].transform("first")
        removed_logical = df.loc[sub.index[merge_copy]].copy()
        removed_logical["Removal_Reason"] = "Same row as an earlier one (identical Date, Product, Target and Actual sales); a copy that was damaged differently"
        removed_logical["Duplicate_Of_Excel_Row"] = first_row[merge_copy].values
        LOGICAL_DUPS = int(merge_copy.sum())
        df = send_to_review(df, pd.Series(df.index.isin(sub.index[conflict_copy]), index=df.index),
                            "Same identity key as another row but with conflicting values")
        df = df.loc[~df.index.isin(sub.index[merge_copy])].copy()
        if notes_conf:
            log("7. Logical dups", "Notes", "Copies carried two different notes", notes_conf,
                "Kept the first non-empty note", "Observation")
    FLAGS = FLAGS.loc[df.index]
log("7. Logical dups", "ALL",
    "Rows that repeat an earlier row (same Date, Product, Target_Qty, Target/Actual sales, margin %) but were "
    "re-damaged so they are not exact copies", LOGICAL_DUPS,
    "Removed - the earlier (original) row is kept", "Fix" if LOGICAL_DUPS else "Check")
removed_duplicates = pd.concat([removed_exact, removed_logical], ignore_index=True, sort=False)
if "Duplicate_Of_Excel_Row" not in removed_duplicates.columns:
    removed_duplicates["Duplicate_Of_Excel_Row"] = np.nan
still_dup = int(df.duplicated(subset=DEDUP_KEY).sum()) if len(df) else 0
log("7. Logical dups", "ALL", "Identity-key duplicates remaining after clean-up", still_dup,
    "None" if still_dup == 0 else "Check manually", "Check" if still_dup == 0 else "Warning")
df = df.sort_values("_src_row", kind="mergesort")             # keep the original file order (needed for the region rule)

# ----------------------------------------------------------------------------
# 10. RECOVERING BLANK REGIONS
# ----------------------------------------------------------------------------
banner("STEP 8/12  -  RECOVERING BLANK REGIONS (NO GUESSING)")


def recover_regions(frame):
    """Return (Series of recovered regions, method text). The file is laid out in fixed blocks:
    for every (Date, Product) the rows run through the regions in a fixed order, N equal rows each.
    The rule is learned from the rows that DO have a region and is applied only if it is 100% consistent."""
    miss = frame["Region"].isna()
    out = pd.Series(np.nan, index=frame.index, dtype=object)
    if not miss.any():
        return out, "nothing to recover"
    n_reg = len(CANON_REGIONS)
    grp = frame.groupby(["Date", "Product"], sort=False)
    pos = grp.cumcount()
    size = grp["Date"].transform("size")
    sizes = size.unique()
    if RECOVER_REGION_BY_BLOCK and len(sizes) == 1 and sizes[0] % n_reg == 0:
        per = int(sizes[0] // n_reg)
        slot = pos // per
        known = ~miss
        ct = pd.crosstab(slot[known], frame.loc[known, "Region"])
        slot_map = ct.idxmax(axis=1)
        purity = ct.max(axis=1).sum() / max(ct.values.sum(), 1)
        if purity == 1.0 and slot_map.nunique() == n_reg and len(slot_map) == n_reg and ct.sum(axis=1).min() >= 30:
            out[miss] = slot[miss].map(slot_map)
            return out, (f"block rule: every Date+Product block holds {int(sizes[0])} rows = {per} rows per region "
                         f"in the order {' > '.join(slot_map.sort_index().tolist())}; rule matched "
                         f"{int(ct.values.sum()):,} of {int(ct.values.sum()):,} known rows (100%)")
    # fallback: a blank sandwiched between two identical neighbours inside the same date
    prev, nxt = frame["Region"].shift(1), frame["Region"].shift(-1)
    same_date = (frame["Date"] == frame["Date"].shift(1)) & (frame["Date"] == frame["Date"].shift(-1))
    cand = prev.notna() & (prev == nxt) & same_date
    known = ~miss
    acc = (frame.loc[known & cand, "Region"] == prev[known & cand]).mean() if (known & cand).any() else 0
    if acc >= 0.999:
        out[miss & cand] = prev[miss & cand]
        return out, f"neighbour rule (accuracy on known rows {acc:.4f})"
    return out, "no reliable rule found"


n_blank_region = int(df["Region"].isna().sum())
rec, rec_method = recover_regions(df)
got = rec.notna() & df["Region"].isna()
old_reg = df["Region"].copy()
df.loc[got, "Region"] = rec[got]
FLAGS.loc[got[got].index, "region_rec"] = True
audit(df, got, "Region", old_reg.where(old_reg.notna(), ""), df["Region"], "Recovered from the block layout of the file")
log("8. Region", "Region", "Blank Region", n_blank_region,
    f"Recovered {int(got.sum()):,} with certainty - {rec_method}", "Fix" if n_blank_region else "Check")
still = df["Region"].isna()
log("8. Region", "Region", "Blank Regions that could NOT be recovered with certainty", int(still.sum()),
    "Moved to Review_Required (never labelled with a made-up value)" if still.any() else "None - 100% recovered",
    "Fix" if still.any() else "Check")
df = send_to_review(df, still, "Region blank and not recoverable with certainty")
FLAGS = FLAGS.loc[df.index]

# ----------------------------------------------------------------------------
# 11. NUMERIC RECOVERY: PRICE, QUANTITY, COST, MARGIN
# ----------------------------------------------------------------------------
banner("STEP 9/12  -  REPAIRING PRICE, QUANTITY, COST AND MARGIN FROM THE FILE'S OWN MATHS")
print("   Identities that hold in this file (verified on every row where the cells exist):")
print("     Target_Sales  = Target_Qty x Unit_Price")
print("     Actual_Sales  = Actual_Qty x Unit_Price      (Actual_Sales itself was never damaged)")
print("     Gross_Margin  = Actual_Sales - Actual_Qty x Unit_Cost ;  Margin % = Gross_Margin / Actual_Sales x 100")

# basic validity of the never-damaged anchor columns
bad_anchor = (~(df["Target_Qty"] > 0)) | (~(df["Target_Sales_USD"] > 0)) | (~(df["Actual_Sales_USD"] > 0))
log("9. Anchors", "Target_Qty / Target_Sales / Actual_Sales", "Missing, zero or negative anchor values",
    int(bad_anchor.sum()), "Moved to Review_Required" if bad_anchor.any() else "None - all anchors valid",
    "Fix" if bad_anchor.any() else "Check")
df = send_to_review(df, bad_anchor, "Invalid Target_Qty / Target_Sales_USD / Actual_Sales_USD")
FLAGS = FLAGS.loc[df.index]

TQ, TS, AS = df["Target_Qty"], df["Target_Sales_USD"], df["Actual_Sales_USD"]
AQ0, P0, C0 = df["Actual_Qty"].copy(), df["Unit_Price_USD"].copy(), df["Unit_Cost_USD"].copy()
GM0, GP0 = df["Gross_Margin_USD"].copy(), df["Gross_Margin_Pct"].copy()
GP_ORIGINAL = GP0.copy()                                   # never-damaged reference, used for validation at the end

# ---- 9a. Unit_Price ----
bad_price = P0.notna() & ~(P0 > 0)
log("9a. Price", "Unit_Price_USD", "Zero or negative price", int(bad_price.sum()), "Treated as missing and rebuilt",
    "Fix" if bad_price.any() else "Check")
P0 = P0.where(P0 > 0)
P_t = (TS / TQ).round(2)
conflict_p = P0.notna() & ((P0 - P_t).abs() > PRICE_TOL)
price_missing = P0.isna()
log("9a. Price", "Unit_Price_USD", "Missing price", int(price_missing.sum()),
    "Rebuilt exactly as Target_Sales / Target_Qty (the price is not guessed from a median)", "Fix")
log("9a. Price", "Unit_Price_USD", "Known price that contradicts Target_Sales / Target_Qty", int(conflict_p.sum()),
    "Moved to Review_Required" if conflict_p.any() else "None - price agrees with Target_Sales/Target_Qty on every row",
    "Fix" if conflict_p.any() else "Check")
df.loc[price_missing, "Unit_Price_USD"] = P_t[price_missing]
FLAGS.loc[price_missing[price_missing].index, "price_rec"] = True
audit(df, price_missing, "Unit_Price_USD", P0.where(P0.notna(), ""), df["Unit_Price_USD"],
      "Rebuilt = Target_Sales / Target_Qty")
df = send_to_review(df, conflict_p, "Unit_Price_USD contradicts Target_Sales_USD / Target_Qty")
FLAGS = FLAGS.loc[df.index]
TQ, TS, AS = df["Target_Qty"], df["Target_Sales_USD"], df["Actual_Sales_USD"]
AQ0, C0, GM0, GP0 = AQ0.loc[df.index], C0.loc[df.index], GM0.loc[df.index], GP0.loc[df.index]
P = df["Unit_Price_USD"].astype("float64")

# ---- 9b. Actual_Qty (negatives, x8 outliers, blanks) ----
Qc = AS / P
Qr = Qc.round()
integral = ((Qc - Qr).abs() <= QTY_INT_TOL) & (Qr >= 1)
raw_q = AQ0
ratio = raw_q / Qr
kind = pd.Series("other", index=df.index, dtype=object)
kind[raw_q.isna()] = "missing"
kind[raw_q.notna() & (raw_q == Qr)] = "ok"
kind[raw_q.notna() & (raw_q == -Qr)] = "negative"
scaled = raw_q.notna() & (ratio.abs() >= 2) & ((ratio.abs() - ratio.abs().round()).abs() < 1e-9)
kind[scaled & (ratio < 0)] = "negative_scaled"
kind[scaled & (ratio > 0)] = "scaled"
n_missing_q = int((kind == "missing").sum())
n_neg = int((raw_q < 0).sum())
n_scaled = int(((kind == "scaled") | (kind == "negative_scaled")).sum())
mult = (ratio[(kind == "scaled") | (kind == "negative_scaled")].abs().round().value_counts().to_dict())
n_other = int(((kind == "other") & integral).sum())

_sc = kind.index[kind == "scaled"]
EXAMPLE_SCALED = (f"{int(raw_q[_sc[0]])} = {int(round(ratio[_sc[0]]))} x {int(Qr[_sc[0]])}" if len(_sc) else "none found")
RETURN_NOTE_NEG = int((kind.isin(["negative", "negative_scaled"]) & (df["Notes"].astype(str).str.lower() == "return")).sum())
PRICE_SPREAD = float(P.max() / P.min()) if len(P) else 0.0
# statistical outlier check on the RAW quantity - shown only so the IQR method can be compared with exact recovery
q_non = raw_q.dropna()
q1, q3 = q_non.quantile(0.25), q_non.quantile(0.75)
fence_lo, fence_hi = q1 - 1.5 * (q3 - q1), q3 + 1.5 * (q3 - q1)
iqr_flag = (raw_q < fence_lo) | (raw_q > fence_hi)
iqr_caught = int((iqr_flag & (kind != "ok")).sum())
iqr_false = int((iqr_flag & (kind == "ok")).sum())
iqr_missed = int(((kind.isin(["negative", "scaled", "negative_scaled"])) & ~iqr_flag).sum())

qty_final = pd.Series(np.nan, index=df.index, dtype="float64")
qty_final[integral] = Qr[integral]
unverified = ~integral
if unverified.any():                       # fallback: only a positive, present quantity can be kept (flagged)
    keep_raw = unverified & raw_q.notna() & (raw_q > 0)
    qty_final[keep_raw] = raw_q[keep_raw]
    FLAGS.loc[keep_raw[keep_raw].index, "qty_other"] = True
log("9b. Quantity", "Actual_Qty", "Missing quantity", n_missing_q,
    "Rebuilt exactly as Actual_Sales / Unit_Price (always a whole number in this file)", "Fix")
log("9b. Quantity", "Actual_Qty", "Negative quantity (impossible; sign flipped by data entry)", n_neg,
    "Restored to the true positive value (Actual_Sales confirms it; not just abs() - verified)", "Fix")
log("9b. Quantity", "Actual_Qty", f"Extreme values that are exact multiples of the true quantity ({', '.join(f'{int(k)}x: {v:,}' for k, v in mult.items())})", n_scaled,
    "Restored to the true quantity (capping/removing would have thrown away good rows)", "Fix")
log("9b. Quantity", "Actual_Qty", "Quantities that disagree with Actual_Sales/Unit_Price in some OTHER way", n_other,
    "None - every damaged quantity is explained" if n_other == 0 else "Rebuilt from Actual_Sales / Unit_Price",
    "Check" if n_other == 0 else "Warning")
log("9b. Quantity", "Actual_Qty",
    f"IQR method (fences {fence_lo:.0f} to {fence_hi:.0f}) comparison: flags {int(iqr_flag.sum())} rows - "
    f"{iqr_false} are genuine (false alarms) and it misses {iqr_missed} damaged values",
    0, "IQR NOT used for cleaning because exact recovery is more accurate", "Observation")
chg_q = (kind != "ok") & qty_final.notna()
FLAGS.loc[(kind == "missing")[kind == "missing"].index, "qty_missing"] = True
FLAGS.loc[kind.isin(["negative", "negative_scaled"]), "qty_sign"] = True
FLAGS.loc[kind.isin(["scaled", "negative_scaled"]), "qty_scaled"] = True
audit(df, chg_q, "Actual_Qty", raw_q.where(raw_q.notna(), ""), qty_final,
      "Rebuilt = Actual_Sales / Unit_Price (fixes blanks, negatives and x8 values)")
df["Actual_Qty"] = qty_final
no_qty = df["Actual_Qty"].isna()
df = send_to_review(df, no_qty, "Actual_Qty could not be rebuilt")
FLAGS = FLAGS.loc[df.index]
AQ0, C0, GM0, GP0 = AQ0.loc[df.index], C0.loc[df.index], GM0.loc[df.index], GP0.loc[df.index]
Q = df["Actual_Qty"]
S = df["Actual_Sales_USD"]
P = df["Unit_Price_USD"]

# ---- 9c. Unit_Cost, Gross_Margin_USD, Gross_Margin_Pct ----
C = C0.where(C0 > 0)
log("9c. Cost", "Unit_Cost_USD", "Zero or negative cost", int((C0.notna() & ~(C0 > 0)).sum()),
    "Treated as missing and rebuilt", "Fix" if (C0.notna() & ~(C0 > 0)).any() else "Check")
cost_known = C.notna()
gm_known = GM0.notna()
# known cost vs known GM: do they agree (with the repaired quantity)?
both = cost_known & gm_known
gm_conf = both & ((GM0 - (S - Q * C)).abs() > MONEY_TOL)
conflict_rows = pd.Series(False, index=df.index)
if gm_conf.any():
    gp_cost = ((S - Q * C) / S * 100)
    cost_ok = (gp_cost - GP0).abs() <= PCT_TOL
    trust_gm = gm_conf & ~cost_ok & ((GM0 / S * 100 - GP0).abs() <= PCT_TOL)
    conflict_rows = gm_conf & ~cost_ok & ~trust_gm
    c_from_gm_conf = ((S - GM0) / Q).round(2)
    C = C.where(~trust_gm, c_from_gm_conf)
log("9c. Cost", "Unit_Cost_USD / Gross_Margin_USD", "Known cost and known margin contradict each other",
    int(gm_conf.sum()), "Resolved with the never-damaged Gross_Margin_Pct" if gm_conf.any()
    else "None - cost and margin agree on every row where both exist", "Fix" if gm_conf.any() else "Check")
df = send_to_review(df, conflict_rows, "Unit_Cost and Gross_Margin contradict each other and cannot be resolved")
FLAGS = FLAGS.loc[df.index]
C, GM0, GP0, Q, S, P = [x.loc[df.index] for x in (C, GM0, GP0, Q, S, P)]
cost_known = C.notna()
gm_known = GM0.notna()

c_from_gm = ((S - GM0) / Q).round(2)
use_gm = (~cost_known) & gm_known & (c_from_gm > 0) & (c_from_gm < P) & ((S - Q * c_from_gm - GM0).abs() <= MONEY_TOL)
C = C.where(~use_gm, c_from_gm)
FLAGS.loc[use_gm[use_gm].index, "cost_rec"] = True
if gm_conf.any():
    FLAGS.loc[gm_conf[gm_conf].index, "conflict"] = True
n_cost_missing = int((~cost_known).sum())
log("9c. Cost", "Unit_Cost_USD", "Missing cost", n_cost_missing,
    f"{int(use_gm.sum()):,} rebuilt exactly from Gross_Margin_USD: (Actual_Sales - GM) / Actual_Qty", "Fix")

need_pct = C.isna()
est_rows = []
exact_pct = 0
for idx in df.index[need_pct]:
    s_i, q_i, p_i, gp_i = S.at[idx], Q.at[idx], P.at[idx], GP0.at[idx]
    if not (s_i > 0 and q_i > 0 and p_i > 0 and pd.notna(gp_i)):
        continue
    c0 = p_i * (1 - gp_i / 100.0)
    base = int(round(c0 * 100))
    cands = [k / 100.0 for k in range(base - 3, base + 4)
             if k > 0 and round((s_i - q_i * (k / 100.0)) / s_i * 100, 2) == round(gp_i, 2)]
    if not cands:
        continue
    pick = min(cands, key=lambda z: abs(z - c0))
    C.at[idx] = pick
    FLAGS.at[idx, "cost_rec"] = True
    if len(cands) == 1:
        exact_pct += 1
    else:
        FLAGS.at[idx, "cost_est"] = True
        est_rows.append(idx)
log("9c. Cost", "Unit_Cost_USD", "Missing cost AND missing Gross_Margin_USD (rebuilt from Gross_Margin_Pct)",
    int(need_pct.sum()), f"{exact_pct} rebuilt exactly; {len(est_rows)} are ambiguous between adjacent cents "
    f"(max error 0.01 USD) and are flagged 'Cost_Estimated'", "Fix" if need_pct.any() else "Check")
audit(df, df["Unit_Cost_USD"].isna() | (df["Unit_Cost_USD"] != C), "Unit_Cost_USD",
      df["Unit_Cost_USD"].where(df["Unit_Cost_USD"].notna(), ""), C,
      "Rebuilt from Gross_Margin_USD / Gross_Margin_Pct")
df["Unit_Cost_USD"] = C
no_cost = df["Unit_Cost_USD"].isna()
df = send_to_review(df, no_cost, "Unit_Cost_USD could not be rebuilt")
FLAGS = FLAGS.loc[df.index]
GM0, GP0, Q, S, P = [x.loc[df.index] for x in (GM0, GP0, Q, S, P)]

# Recalculate the derived fields from the repaired inputs (as in your checklist)
gm_new = (S - Q * df["Unit_Cost_USD"]).round(2)
gp_new = (gm_new / S * 100).round(2)
gm_missing = GM0.isna()
FLAGS.loc[gm_missing[gm_missing].index, "gm_rec"] = True
gm_diff = (~gm_missing) & ((gm_new - GM0).abs() > MONEY_TOL)
audit(df, gm_missing, "Gross_Margin_USD", GM0.where(GM0.notna(), ""), gm_new, "Recalculated = Actual_Sales - Qty x Cost")
log("9c. Margin", "Gross_Margin_USD", "Missing", int(gm_missing.sum()),
    "Recalculated = Actual_Sales - Actual_Qty x Unit_Cost", "Fix")
log("9c. Margin", "Gross_Margin_USD", "Existing values that differ from the recalculation (> 2 cents)", int(gm_diff.sum()),
    "None - every existing value is reproduced exactly" if not gm_diff.any() else "Overwritten with the recalculation",
    "Check" if not gm_diff.any() else "Fix")
df["Gross_Margin_USD"] = gm_new
df["Gross_Margin_Pct"] = gp_new
gp_dev = (gp_new - GP_ORIGINAL.loc[df.index]).abs()
log("9c. Margin", "Gross_Margin_Pct", "Recalculated % that differs from the original untouched column (> 0.011)",
    int((gp_dev > PCT_TOL).sum()), "None - recalculated margin % matches the original everywhere"
    if not (gp_dev > PCT_TOL).any() else "Review", "Check" if not (gp_dev > PCT_TOL).any() else "Warning")

# Actual_Sales / Target_Sales identities - recompute only if they disagree
as_gap = ((Q * P).round(2) - S).abs() > MONEY_TOL
ts_gap = ((TQ * P).round(2) - TS).abs() > MONEY_TOL
log("9d. Sales", "Actual_Sales_USD", "Actual_Sales != Actual_Qty x Unit_Price after repair", int(as_gap.sum()),
    "Moved to Review_Required" if as_gap.any() else "Holds on every row - Actual_Sales was never damaged, it was the quantity that was",
    "Fix" if as_gap.any() else "Check")
log("9d. Sales", "Target_Sales_USD", "Target_Sales != Target_Qty x Unit_Price", int(ts_gap.sum()),
    "Moved to Review_Required" if ts_gap.any() else "Holds on every row", "Fix" if ts_gap.any() else "Check")
df = send_to_review(df, as_gap | ts_gap, "Sales do not reconcile with quantity x price after repair")
FLAGS = FLAGS.loc[df.index]
cost_gt = df["Unit_Cost_USD"] >= df["Unit_Price_USD"]
df = send_to_review(df, cost_gt, "Unit_Cost is not below Unit_Price (negative margin per unit)")
FLAGS = FLAGS.loc[df.index]

# ----------------------------------------------------------------------------
# 12. FINAL TYPES, DERIVED COLUMNS, CLEANING NOTES, SORT
# ----------------------------------------------------------------------------
banner("STEP 10/12  -  FINAL TYPES, VARIANCE COLUMNS AND SORT")
year_raw, month_raw = df["Year"], df["Month"]
df["Year"] = df["Date"].dt.year.astype("int64")
df["Month"] = df["Date"].dt.month.astype("int64")
df["Month_Name"] = df["Date"].dt.month_name()
n_year_fix = int((year_raw != df["Year"]).sum())
n_month_fix = int((month_raw != df["Month"]).sum())
log("10. Types", "Year / Month / Month_Name", "Disagree with the Date", n_year_fix + n_month_fix,
    "Re-derived from Date as integers" if (n_year_fix + n_month_fix) else
    "Already consistent - stored as proper integers, Month_Name re-derived from Date", "Fix" if (n_year_fix + n_month_fix) else "Check")
for c in ["Target_Qty", "Actual_Qty"]:
    df[c] = df[c].round().astype("int64")
for c in ["Unit_Price_USD", "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_USD", "Gross_Margin_Pct"]:
    df[c] = df[c].astype("float64").round(2)

df["Quarter"] = "Q" + df["Date"].dt.quarter.astype(str)
df["Year_Month"] = df["Date"].dt.strftime("%Y-%m")
df["Qty_Variance"] = (df["Actual_Qty"] - df["Target_Qty"]).astype("int64")
df["Qty_Attainment_Pct"] = (df["Actual_Qty"] / df["Target_Qty"] * 100).round(2)
df["Sales_Variance_USD"] = (df["Actual_Sales_USD"] - df["Target_Sales_USD"]).round(2)
df["Sales_Variance_Pct"] = (df["Sales_Variance_USD"] / df["Target_Sales_USD"] * 100).round(2)
qa = df["Actual_Qty"]
qq1, qq3 = qa.quantile(0.25), qa.quantile(0.75)
df["Qty_IQR_Outlier"] = np.where((qa < qq1 - 1.5 * (qq3 - qq1)) | (qa > qq3 + 1.5 * (qq3 - qq1)), "Yes", "No")
n_iqr_clean = int((df["Qty_IQR_Outlier"] == "Yes").sum())
log("10. Outliers", "Actual_Qty", "Statistical (IQR) outliers remaining AFTER exact repair", n_iqr_clean,
    "Kept and flagged in Qty_IQR_Outlier: they are real (they reconcile with Actual_Sales / Unit_Price)",
    "Decision" if n_iqr_clean else "Check")

NOTE_TEXT = {
    "region_std": "Region standardised", "region_rec": "Blank Region recovered", "cat_std": "Category trimmed",
    "prod_std": "Product case fixed", "notes_std": "Notes tidied", "merged": "Hidden duplicate merged",
    "price_rec": "Price rebuilt", "qty_missing": "Missing quantity rebuilt", "qty_sign": "Negative quantity fixed",
    "qty_scaled": "Extreme quantity restored", "qty_other": "Quantity unverified", "cost_rec": "Cost rebuilt",
    "cost_est": "Cost estimated (+/-0.01)", "gm_rec": "Gross margin recalculated", "conflict": "Conflict resolved",
}
FLAGS = FLAGS.loc[df.index]
notes_txt = pd.Series("", index=df.index, dtype=object)
for fl, txt in NOTE_TEXT.items():
    m = FLAGS[fl]
    notes_txt = notes_txt.where(~m, notes_txt + txt + "; ")
df["Cleaning_Notes"] = notes_txt.str.rstrip("; ").replace("", "No changes needed")
df["Source_Excel_Row"] = df["_src_row"].astype("int64")

df = df.sort_values(["Date", "_src_row"], kind="mergesort").reset_index(drop=True)
final_cols = [c for c in EXPECTED_COLUMNS if c != "Notes"]
derived = ["Quarter", "Year_Month", "Qty_Variance", "Qty_Attainment_Pct", "Sales_Variance_USD",
           "Sales_Variance_Pct", "Qty_IQR_Outlier"]
final_cols = final_cols + ([] if DROP_NOTES else ["Notes"]) + derived + ["Cleaning_Notes"]
if INCLUDE_SOURCE_ROW:
    final_cols.append("Source_Excel_Row")
clean = df[final_cols].copy()
log("10. Sort", "Date", "File was not in date order (appended duplicates sat at the bottom)",
    0, "Sorted by Date (original order kept inside each date)", "Fix")
if DROP_NOTES:
    log("10. Columns", "Notes", "97% empty column", 0, "Dropped (DROP_NOTES = True)", "Decision")
else:
    log("10. Columns", "Notes", f"Mostly empty column ({int(clean['Notes'].notna().sum()):,} filled cells: Late shipment, Stock issue, ...)",
        0, "Kept as an optional, cleaned column (blank = no note); set DROP_NOTES = True to remove it", "Decision")

# ----------------------------------------------------------------------------
# 13. FINAL VALIDATION  (brutal self-audit - includes every business rule on your checklist)
# ----------------------------------------------------------------------------
banner("STEP 11/12  -  FINAL VALIDATION OF THE CLEANED DATA")
checks = []


def check(name, ok, detail=""):
    ok = bool(ok)
    checks.append({"Check": name, "Result": "PASS" if ok else "FAIL", "Detail": detail})
    print(f"   [{'PASS' if ok else 'FAIL'}] {name}  {detail}")


n_review = sum(len(x) for x in REVIEW_PARTS)
core = ["Date", "Year", "Month", "Month_Name", "Region", "Category", "Product", "Target_Qty", "Actual_Qty",
        "Unit_Price_USD", "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_USD",
        "Gross_Margin_Pct"]
check("No exact duplicate rows", not clean[[c for c in clean.columns if c not in ("Source_Excel_Row", "Cleaning_Notes")]].duplicated().any())
key_dups = int(clean.duplicated(subset=DEDUP_KEY).sum())
check("No logical duplicates (Date + Product + targets + actual sales is unique)", key_dups == 0, f"({key_dups} found)")
check("Region has ONLY the 5 standard values", set(clean["Region"].unique()) == set(CANON_REGIONS)
      or set(clean["Region"].unique()) <= set(CANON_REGIONS), f"{sorted(clean['Region'].unique())}")
for c in core:
    check(f"No missing values in {c}", clean[c].notna().all())
check("Date is a real datetime column", str(clean["Date"].dtype).startswith("datetime64"))
check("Date is sorted ascending", clean["Date"].is_monotonic_increasing)
check("Year and Month are integers", str(clean["Year"].dtype).startswith("int") and str(clean["Month"].dtype).startswith("int"))
check("Year / Month / Month_Name agree with Date",
      ((clean["Year"] == clean["Date"].dt.year) & (clean["Month"] == clean["Date"].dt.month)
       & (clean["Month_Name"] == clean["Date"].dt.month_name())).all())
check("Quantities are integers", str(clean["Target_Qty"].dtype).startswith("int") and str(clean["Actual_Qty"].dtype).startswith("int"))
check("Money columns are floats", all(str(clean[c].dtype).startswith("float") for c in
      ["Unit_Price_USD", "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_USD", "Gross_Margin_Pct"]))
check("No negative or zero quantities", ((clean["Target_Qty"] > 0) & (clean["Actual_Qty"] > 0)).all(),
      f"(Actual_Qty range {int(clean['Actual_Qty'].min())} to {int(clean['Actual_Qty'].max())})")
check("No negative or zero prices, costs or sales",
      (clean[["Unit_Price_USD", "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD"]] > 0).all().all())
check("Unit_Cost is below Unit_Price on every row", (clean["Unit_Cost_USD"] < clean["Unit_Price_USD"]).all())
check("Actual_Sales_USD = Actual_Qty x Unit_Price_USD on EVERY row",
      ((clean["Actual_Qty"] * clean["Unit_Price_USD"] - clean["Actual_Sales_USD"]).abs() <= MONEY_TOL).all())
check("Target_Sales_USD = Target_Qty x Unit_Price_USD on EVERY row",
      ((clean["Target_Qty"] * clean["Unit_Price_USD"] - clean["Target_Sales_USD"]).abs() <= MONEY_TOL).all())
check("Gross_Margin_USD = Actual_Sales - Actual_Qty x Unit_Cost on EVERY row",
      ((clean["Actual_Sales_USD"] - clean["Actual_Qty"] * clean["Unit_Cost_USD"] - clean["Gross_Margin_USD"]).abs() <= MONEY_TOL).all())
check("Gross_Margin_Pct = Gross_Margin_USD / Actual_Sales x 100 on EVERY row",
      ((clean["Gross_Margin_USD"] / clean["Actual_Sales_USD"] * 100 - clean["Gross_Margin_Pct"]).abs() <= PCT_TOL).all())
# independent cross-check: recalculated margin % vs the original (never damaged) margin % column
orig_by_src = to_numeric_safe(raw.set_index("_src_row")["Gross_Margin_Pct"])
cmp_orig = orig_by_src.reindex(df["Source_Excel_Row"].values).values
max_dev = float(np.nanmax(np.abs(clean["Gross_Margin_Pct"].values - cmp_orig)))
check("Recalculated Gross_Margin_Pct reproduces the ORIGINAL untouched % column", max_dev <= PCT_TOL,
      f"(max deviation {max_dev:.4f} percentage points)")
orig_sales = to_numeric_safe(raw.set_index("_src_row")["Actual_Sales_USD"]).reindex(df["Source_Excel_Row"].values).values
orig_tsales = to_numeric_safe(raw.set_index("_src_row")["Target_Sales_USD"]).reindex(df["Source_Excel_Row"].values).values
check("Actual_Sales_USD and Target_Sales_USD are identical to the source (never altered)",
      np.allclose(clean["Actual_Sales_USD"].values, orig_sales) and np.allclose(clean["Target_Sales_USD"].values, orig_tsales))
for c in ["Region", "Category", "Product", "Month_Name"] + ([] if DROP_NOTES else ["Notes"]):
    s = clean[c].dropna().astype(str)
    check(f"No leading/trailing/double spaces in {c}", ((s != s.str.strip()) | s.str.contains(r"\s{2,}")).sum() == 0)
check("Product names are in proper Title Case", clean["Product"].map(lambda v: v == smart_title(v)).all(),
      f"({clean['Product'].nunique()} products)")
pcn = clean.groupby("Product")["Category"].nunique()
check("Every product belongs to exactly one category", (pcn == 1).all())
grid = clean.groupby(["Date", "Product", "Region"]).size()
deficit = int((grid.max() - grid).sum()) if len(grid) else 0
check("Layout is balanced: every Date x Product x Region holds the same number of rows",
      grid.nunique() == 1 or (n_review > 0 and deficit <= n_review),
      f"({int(grid.max()) if len(grid) else 0} rows each, {len(grid):,} combinations)" if grid.nunique() == 1
      else f"(short by {deficit} rows, fully explained by the {n_review} row(s) in Review_Required)")
reg_counts = clean["Region"].value_counts()
check("Regions are equally represented", reg_counts.nunique() == 1 or (n_review > 0 and int(reg_counts.max() - reg_counts.min()) <= n_review),
      f"{reg_counts.to_dict()}")
check("No infinite values anywhere", not np.isinf(clean.select_dtypes("number")).any().any())
check("Row reconciliation: raw = clean + exact dups + logical dups + review + empty rows",
      RAW_ROWS == len(clean) + EXACT_DUPS + LOGICAL_DUPS + n_review + BLANK_ROWS_DROPPED,
      f"({RAW_ROWS:,} = {len(clean):,} + {EXACT_DUPS:,} + {LOGICAL_DUPS:,} + {n_review:,} + {BLANK_ROWS_DROPPED:,})")
n_fail = sum(1 for c in checks if c["Result"] == "FAIL")
print(f"\n   >>> {len(checks) - n_fail}/{len(checks)} checks passed" + ("  - ALL CLEAR" if n_fail == 0 else "  - REVIEW THE FAILS ABOVE"))

# ----------------------------------------------------------------------------
# 14. REPORT TABLES
# ----------------------------------------------------------------------------
profile_after = make_profile(clean, EXPECTED_COLUMNS if not DROP_NOTES else [c for c in EXPECTED_COLUMNS if c != "Notes"], "After")
quality = profile_before.merge(profile_after, on="Column", how="left")
log_df = pd.DataFrame(LOG)
corr_df = (pd.concat(CORRECTIONS, ignore_index=True) if CORRECTIONS else
           pd.DataFrame(columns=["Source_Excel_Row", "Date", "Region", "Product", "Field", "Old_Value", "New_Value", "Method"]))
corr_df["Old_Value"] = corr_df["Old_Value"].astype(object)
corr_df = corr_df.sort_values(["Field", "Source_Excel_Row"]).reset_index(drop=True)
mapping_df = (pd.concat(MAPPING_ROWS, ignore_index=True) if MAPPING_ROWS
              else pd.DataFrame(columns=["Column", "Original_Value", "Standardised_Value", "Row_Count"]))
mapping_df = mapping_df.sort_values(["Column", "Standardised_Value", "Row_Count"], ascending=[True, True, False])
checks_df = pd.DataFrame(checks)
dup_cols = ["Removal_Reason", "Duplicate_Of_Excel_Row", "_src_row"] + EXPECTED_COLUMNS
removed_out = removed_duplicates.reindex(columns=dup_cols).rename(columns={"_src_row": "Source_Excel_Row"})
if REVIEW_PARTS:
    review_df = pd.concat(REVIEW_PARTS, sort=False).rename(columns={"_src_row": "Source_Excel_Row"})
    review_df = review_df.reindex(columns=["Review_Reason", "Source_Excel_Row"] + EXPECTED_COLUMNS)
else:
    review_df = pd.DataFrame(columns=["Review_Reason", "Source_Excel_Row"] + EXPECTED_COLUMNS)

# ---- checklist review: each point you raised, verified against the real data ----
vb = profile_before.set_index("Column")["Missing_Before"].to_dict()
blank_region_raw = int(raw["Region"].isna().sum())
conf = lambda n: "Confirmed" if n > 0 else "NOT confirmed"
note_vals = sorted(clean["Notes"].dropna().unique().tolist()) if "Notes" in clean.columns else []
notes_pct = vb.get("Notes", 0) / max(len(raw), 1) * 100
n_destroyed = int(kind.isin(["negative", "scaled", "negative_scaled"]).sum())
checklist = pd.DataFrame([
    ["1. Missing values", "Region", f"{conf(blank_region_raw + na_text_region)}. {blank_region_raw:,} truly blank cells PLUS {na_text_region:,} cells holding the text 'NA' "
     f"that pandas silently converts to null (they mean North America).", f"Blanks recovered exactly ({int(got.sum()):,} after de-duplication) from the block layout of the file; 'NA' mapped to North America. Nothing labelled 'Unknown'."],
    ["1. Missing values", "Actual_Qty", f"{conf(vb.get('Actual_Qty', 0))}: {vb.get('Actual_Qty', 0):,} blanks.", "Rebuilt exactly = Actual_Sales / Unit_Price (always a whole number)."],
    ["1. Missing values", "Unit_Price_USD", f"{conf(vb.get('Unit_Price_USD', 0))}: {vb.get('Unit_Price_USD', 0):,} blanks.", f"Rebuilt exactly = Target_Sales / Target_Qty. A median would have been wrong: prices vary about {PRICE_SPREAD:.0f}x across rows."],
    ["1. Missing values", "Unit_Cost_USD", f"{conf(vb.get('Unit_Cost_USD', 0))}: {vb.get('Unit_Cost_USD', 0):,} blanks.", f"Rebuilt exactly from Gross_Margin_USD; {len(est_rows)} rows are ambiguous by 1 cent and flagged Cost_Estimated."],
    ["1. Missing values", "Gross_Margin_USD", f"{conf(vb.get('Gross_Margin_USD', 0))}: {vb.get('Gross_Margin_USD', 0):,} blanks.", "Recalculated = Actual_Sales - Qty x Cost."],
    ["1. Missing values", "Gross_Margin_Pct", f"{conf(vb.get('Gross_Margin_Pct', 0))}: {vb.get('Gross_Margin_Pct', 0):,} blanks" + (" (this column is complete)." if vb.get('Gross_Margin_Pct', 0) == 0 else "."), "Recalculated for every row; it matches the original untouched column, which independently proves the repairs."],
    ["1. Missing values", "Other columns", f"Notes is {notes_pct:.1f}% empty; all other columns: " + ("complete." if all(v == 0 for k, v in vb.items() if k not in ("Region", "Actual_Qty", "Unit_Price_USD", "Unit_Cost_USD", "Gross_Margin_USD", "Notes")) else "see Data_Quality_Report."), "Notes kept as an optional column."],
    ["2. Dirty categories", "Region", f"{conf(N_REG_VARIANTS - len(CANON_REGIONS))}: {N_REG_VARIANTS} different spellings for 5 regions (incl. NA, N. America, Northamerica, North-America, EU, Europa, europe, APAC, AsiaPac, Asia-Pacific, LATAM, LatAm, LatinAmerica, ME, MiddleEast, trailing spaces).", "Standardised to exactly 5 regions."],
    ["2. Dirty categories", "Category", f"{conf(N_CAT_FIX)}: {N_CAT_FIX:,} values carry extra leading/trailing spaces.", f"Trimmed; {clean['Category'].nunique()} clean categories."],
    ["2. Dirty categories", "Product", f"{conf(N_PROD_FIX)}: {N_PROD_FIX:,} values in lower / UPPER case (every product exists in several spellings).", f"{N_PRODUCTS} products in proper Title Case."],
    ["3. Data types", "Date", (f"NOT confirmed: all dates are real Excel dates and there are only {distinct_dates} distinct dates" + (" (the 1st of each month), so no DD/MM/YYYY ambiguity or day/month swap is possible." if ALL_FIRST_OF_MONTH else ".")) if n_text_dates == 0 else f"Confirmed: {n_text_dates:,} dates were text/serial values.", "Parsed defensively anyway (text, serial numbers, swapped day/month) and sorted."],
    ["3. Data types", "Numeric columns", (f"{conf(N_NUMTEXT)}: " + ("no numbers are stored as text; they were float/int mixed with blanks." if N_NUMTEXT == 0 else f"{N_NUMTEXT:,} numbers were stored as text.")), "Converted to proper int64 (quantities) and float64 (money)."],
    ["3. Data types", "Year / Month", "Integers and consistent with Date." if (n_year_fix + n_month_fix) == 0 else f"{n_year_fix + n_month_fix:,} values disagreed with Date.", "Re-derived from Date and kept as integers."],
    ["4. Negative values", "Actual_Qty", f"{conf(n_neg)}: {n_neg:,} negative quantities. " + (f"They are not genuine returns (only {RETURN_NOTE_NEG} carry a 'Return' note) and" if RETURN_NOTE_NEG else "None of them carries a 'Return' note, so they are not genuine returns, and") + " Actual_Sales_USD is positive.", "Restored to the true positive value, verified against Actual_Sales / Unit_Price."],
    ["5. Extreme outliers", "Actual_Qty", f"{conf(n_scaled)}: {n_scaled:,} values are exact multiples ({', '.join(f'{int(k)}x: {v:,}' for k, v in mult.items())}) of the true quantity (e.g. {EXAMPLE_SCALED}).", f"Restored exactly. An IQR rule flags {int(iqr_flag.sum())} rows, of which {iqr_false} are genuine, and misses {iqr_missed} damaged values - so it was not used to clean."],
    ["6. Duplicates", "All", f"{conf(EXACT_DUPS + LOGICAL_DUPS)}: {EXACT_DUPS:,} exact duplicates AND {LOGICAL_DUPS:,} hidden duplicates (copies that were damaged differently). Total {EXACT_DUPS + LOGICAL_DUPS:,} = {(EXACT_DUPS + LOGICAL_DUPS) / max(len(clean), 1) * 100:.1f}% of the {len(clean):,} true rows.", "Both removed; each removed row is listed with the Excel row it duplicates."],
    ["7. Calculated fields", "Margins / Sales", f"Only partly: Actual_Sales_USD, Target_Sales_USD and Gross_Margin_Pct were never wrong - only the quantity was. Re-computing Actual_Sales from the damaged quantity would have corrupted about {n_destroyed:,} correct sales values.", "Quantity repaired first, then Gross_Margin_USD / Pct recalculated from the repaired inputs."],
    ["8. Notes column", "Notes", f"{notes_pct:.1f}% empty; values: {', '.join(note_vals) if note_vals else 'none'}.", "Kept as an optional cleaned column (DROP_NOTES = True removes it)."],
    ["9. Business rules", "All rules", "Every rule is tested on every row (see Validation_Checks).", f"{len(checks) - n_fail}/{len(checks)} checks pass."],
], columns=["Your_Checklist_Item", "Column", "What_The_Data_Actually_Shows", "What_Was_Done"])


def summarise(frame, by):
    g = frame.groupby(by, dropna=False).agg(
        Rows=("Region", "size"), Target_Qty=("Target_Qty", "sum"), Actual_Qty=("Actual_Qty", "sum"),
        Target_Sales_USD=("Target_Sales_USD", "sum"), Actual_Sales_USD=("Actual_Sales_USD", "sum"),
        Gross_Margin_USD=("Gross_Margin_USD", "sum")).reset_index()
    g["Qty_Attainment_%"] = (g["Actual_Qty"] / g["Target_Qty"] * 100).round(2)
    g["Sales_Variance_USD"] = (g["Actual_Sales_USD"] - g["Target_Sales_USD"]).round(2)
    g["Sales_Variance_%"] = (g["Sales_Variance_USD"] / g["Target_Sales_USD"] * 100).round(2)
    g["Gross_Margin_%"] = (g["Gross_Margin_USD"] / g["Actual_Sales_USD"] * 100).round(2)
    for c in ["Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_USD"]:
        g[c] = g[c].round(2)
    return g


SUMMARIES = {
    "By_Month": summarise(clean, ["Year_Month"]),
    "By_Region": summarise(clean, ["Region"]).sort_values("Actual_Sales_USD", ascending=False),
    "By_Category": summarise(clean, ["Category"]).sort_values("Actual_Sales_USD", ascending=False),
    "By_Product": summarise(clean, ["Category", "Product"]).sort_values("Actual_Sales_USD", ascending=False),
    "By_Region_Category": summarise(clean, ["Region", "Category"]),
}

readme = pd.DataFrame({
    "Item": ["Source file", "Raw rows", "Exact duplicates removed", "Hidden (logical) duplicates removed",
             "Rows sent to Review_Required", "Clean rows delivered", "Blank Regions recovered",
             "Quantities repaired (blank / negative / x8)", "Prices rebuilt", "Costs rebuilt",
             "Validation checks passed", "Date range", "Total Actual Sales (USD)", "Total Target Sales (USD)",
             "Total Gross Margin (USD)", "", "SHEETS", "Clean_Data", "Checklist_Review", "Cleaning_Log",
             "Data_Quality_Report", "Validation_Checks", "Corrections_Audit", "Value_Mappings",
             "Removed_Duplicates", "Review_Required", "By_Month ... By_Region_Category"],
    "Value": [INPUT_PATH.name, f"{RAW_ROWS:,}", f"{EXACT_DUPS:,}", f"{LOGICAL_DUPS:,}", f"{n_review:,}", f"{len(clean):,}",
              f"{int(got.sum()):,}", f"{int(chg_q.sum()):,}", f"{int(price_missing.sum()):,}",
              f"{int(FLAGS['cost_rec'].sum()):,}", f"{len(checks) - n_fail}/{len(checks)}",
              f"{clean['Date'].min():%Y-%m-%d} to {clean['Date'].max():%Y-%m-%d}",
              f"{clean['Actual_Sales_USD'].sum():,.2f}", f"{clean['Target_Sales_USD'].sum():,.2f}",
              f"{clean['Gross_Margin_USD'].sum():,.2f}", "", "",
              "The final cleaned dataset (use this one) + variance columns",
              "Each point of your checklist verified against the data, and what was done",
              "Every issue found, rows affected and the action taken",
              "Before-vs-after profile of every original column",
              "PASS/FAIL results of the final self-audit (all 9 business rules + more)",
              "Cell-level old -> new value for every repaired cell",
              "Every spelling that was standardised and how many rows it affected",
              "Every duplicate that was removed and which row it duplicates",
              "Rows needing a human decision (empty = nothing was uncertain)",
              "Ready-made variance summary tables"],
})

# ----------------------------------------------------------------------------
# 15. WRITE THE EXCEL FILE (formatted) + CSV
# ----------------------------------------------------------------------------
banner("STEP 12/12  -  SAVING FILES")
SHEETS = [("Read_Me", readme), ("Clean_Data", clean), ("Checklist_Review", checklist), ("Cleaning_Log", log_df),
          ("Data_Quality_Report", quality), ("Validation_Checks", checks_df), ("Corrections_Audit", corr_df),
          ("Value_Mappings", mapping_df), ("Removed_Duplicates", removed_out), ("Review_Required", review_df)]
SHEETS += list(SUMMARIES.items())

COL_FORMATS = {
    "Date": "yyyy-mm-dd", "Year": "0", "Month": "0", "Target_Qty": "#,##0", "Actual_Qty": "#,##0",
    "Unit_Price_USD": "#,##0.00", "Unit_Cost_USD": "#,##0.00", "Target_Sales_USD": "#,##0.00",
    "Actual_Sales_USD": "#,##0.00", "Gross_Margin_USD": "#,##0.00", "Gross_Margin_Pct": "0.00",
    "Qty_Variance": "#,##0;[Red]-#,##0", "Qty_Attainment_Pct": "0.00", "Sales_Variance_USD": "#,##0.00;[Red]-#,##0.00",
    "Sales_Variance_Pct": "0.00;[Red]-0.00", "Rows": "#,##0", "Rows_Affected": "#,##0", "Row_Count": "#,##0",
    "Qty_Attainment_%": "0.00", "Sales_Variance_%": "0.00;[Red]-0.00", "Gross_Margin_%": "0.00",
    "Source_Excel_Row": "0", "Duplicate_Of_Excel_Row": "0",
}
WRAP_SHEETS = {"Cleaning_Log": ("Issue_Found", "Action_Taken"), "Read_Me": ("Value",),
               "Checklist_Review": ("What_The_Data_Actually_Shows", "What_Was_Done")}


def write_workbook(path):
    if HAS_XLSXWRITER:
        with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
            wb = writer.book
            head_fmt = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1,
                                      "align": "center", "valign": "vcenter", "text_wrap": True})
            wrap_fmt = wb.add_format({"text_wrap": True, "valign": "top"})
            cache = {}
            for name, frame in SHEETS:
                frame.to_excel(writer, sheet_name=name, index=False)
                ws = writer.sheets[name]
                for j, col in enumerate(frame.columns):
                    ws.write(0, j, col, head_fmt)
                    lens = frame[col].head(2000).map(lambda v: len(str(v)) if pd.notna(v) else 0)
                    longest = int(lens.max()) if len(lens) else 0
                    width = min(max(len(str(col)) + 2, longest + 2), 45)
                    if name in WRAP_SHEETS and col in WRAP_SHEETS[name]:
                        ws.set_column(j, j, 75, wrap_fmt)
                        continue
                    nf = COL_FORMATS.get(col)
                    if nf:
                        if nf not in cache:
                            cache[nf] = wb.add_format({"num_format": nf})
                        ws.set_column(j, j, max(width, 12), cache[nf])
                    else:
                        ws.set_column(j, j, width)
                ws.freeze_panes(1, 0)
                ws.set_row(0, 30)
                if len(frame) > 0 and name != "Read_Me":
                    ws.autofilter(0, 0, len(frame), len(frame.columns) - 1)
            writer.sheets["Clean_Data"].freeze_panes(1, 4)
            writer.sheets["Clean_Data"].activate()
    else:                                                          # plain fallback without styling
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, frame in SHEETS:
                frame.to_excel(writer, sheet_name=name, index=False)


def save_with_fallback(path):
    try:
        write_workbook(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{_dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        print(f"   '{path.name}' is open in Excel - saving as '{alt.name}' instead.")
        write_workbook(alt)
        return alt


FINAL_XLSX = save_with_fallback(OUTPUT_XLSX)
try:
    clean.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    FINAL_CSV = OUTPUT_CSV
except PermissionError:
    FINAL_CSV = OUTPUT_CSV.with_name(f"{OUTPUT_CSV.stem}_{_dt.datetime.now():%Y%m%d_%H%M%S}.csv")
    clean.to_csv(FINAL_CSV, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
print(f"   Saved: {FINAL_XLSX}   ({FINAL_XLSX.stat().st_size / 1e6:.1f} MB)")
print(f"   Saved: {FINAL_CSV}   ({FINAL_CSV.stat().st_size / 1e6:.1f} MB)")

if VERIFY_OUTPUT_FILE:
    print("   Re-opening the saved Excel file to verify it ...")
    back = pd.read_excel(FINAL_XLSX, sheet_name="Clean_Data", engine="openpyxl", keep_default_na=False, na_values=[""])
    ok_shape = back.shape == clean.shape
    ok_sales = abs(back["Actual_Sales_USD"].sum() - clean["Actual_Sales_USD"].sum()) < 0.5
    ok_dates = pd.api.types.is_datetime64_any_dtype(back["Date"])
    ok_nulls = int(back[["Region", "Category", "Product", "Actual_Qty", "Unit_Price_USD", "Unit_Cost_USD",
                         "Gross_Margin_USD"]].isna().sum().sum()) == 0
    ok_region = set(back["Region"].unique()) <= set(CANON_REGIONS)
    print(f"   [{'PASS' if ok_shape else 'FAIL'}] Saved file has the same shape as the clean data: {back.shape}")
    print(f"   [{'PASS' if ok_sales else 'FAIL'}] Actual sales total matches after save/reload")
    print(f"   [{'PASS' if ok_dates else 'FAIL'}] Date is stored as real Excel dates")
    print(f"   [{'PASS' if ok_nulls else 'FAIL'}] No blanks in the repaired columns of the saved file")
    print(f"   [{'PASS' if ok_region else 'FAIL'}] Only the 5 standard regions in the saved file")

# ----------------------------------------------------------------------------
# 16. FINAL REPORT
# ----------------------------------------------------------------------------
banner("DONE - CLEANING REPORT")
print(f"   Raw rows ............................. {RAW_ROWS:>8,}")
print(f"   - Exact duplicates ................... {EXACT_DUPS:>8,}")
print(f"   - Hidden (re-damaged) duplicates ..... {LOGICAL_DUPS:>8,}")
print(f"   - Sent to Review_Required ............ {n_review:>8,}")
print(f"   = CLEAN ROWS DELIVERED ............... {len(clean):>8,}")
print(f"   Blank Regions recovered .............. {int(got.sum()):>8,}   (+ {na_text_region:,} 'NA' cells kept as North America)")
print(f"   Quantities repaired .................. {int(chg_q.sum()):>8,}   ({n_missing_q:,} blank, {n_neg:,} negative, {n_scaled:,} extreme; a few rows are both)")
print(f"   Prices rebuilt ....................... {int(price_missing.sum()):>8,}")
print(f"   Costs rebuilt ........................ {int(FLAGS['cost_rec'].sum()):>8,}   ({len(est_rows)} flagged as +/-0.01 estimates)")
print(f"   Gross margins recalculated ........... {int(gm_missing.sum()):>8,}")
print(f"   Validation checks passed ............. {len(checks) - n_fail:>5}/{len(checks)}")
print(f"   Total runtime ........................ {time.time() - T0:>7.0f}s")
print(f"\n   Open  {FINAL_XLSX.name}  in Excel - start with the 'Read_Me' and 'Checklist_Review' sheets.")

if OPEN_FOLDER_WHEN_DONE:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(FINAL_XLSX.parent))                     # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(FINAL_XLSX.parent)])
        else:
            subprocess.Popen(["xdg-open", str(FINAL_XLSX.parent)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass
