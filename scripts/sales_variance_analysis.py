# -*- coding: utf-8 -*-
"""
==============================================================================
 SALES VARIANCE & PERFORMANCE ANALYSIS      (VS Code / terminal, no input needed)
==============================================================================
 Input  : Cleaned_Sales_Data.csv   (output of clean_sales_data.py)
 Output : Sales_Variance_Analysis.xlsx          multi-sheet workbook
          Sales_Variance_Enriched_Data.csv      cleaned data + new columns
          (saved next to the input file; when run from scripts/ in the repo layout,
           the CSV goes to data/processed/ and the workbook to outputs/)

 Run    : python sales_variance_analysis.py
          python sales_variance_analysis.py "C:\\path\\to\\Cleaned_Sales_Data.csv"

 Libraries: pandas + numpy do all the analysis. Writing a formatted .xlsx needs
 an Excel "engine" that pandas calls (xlsxwriter, else openpyxl); the script
 pip-installs xlsxwriter automatically if neither is present.

 Design rules (the things that usually go wrong in variance work):
  * Money is handled as integer CENTS, so a row sitting exactly on +3 % / -3 %
    / 10 % / $75,000 is classified exactly, with no floating-point surprises.
  * Group percentages are always  sum(variance) / sum(target)  - never an
    average of row percentages (that is mathematically wrong).
  * Group status / materiality is re-derived from the group totals.
  * Month-on-month / year-on-year are matched on a calendar month number,
    never on row position.
  * Original column names are never renamed. The four variance columns that the
    cleaning script already wrote are RE-COMPUTED here and compared with the
    input; any difference is reported.
==============================================================================
"""
import sys
import subprocess
import datetime as _dt
from pathlib import Path

import numpy as np
import pandas as pd

# ------------------------------------------------------------------ CONFIG --
INPUT_FILE_NAME = "Cleaned_Sales_Data.csv"
OUT_XLSX_NAME = "Sales_Variance_Analysis.xlsx"
OUT_CSV_NAME = "Sales_Variance_Enriched_Data.csv"

FAVORABLE_THRESHOLD_PCT = 3.0       # Favorable   if Sales_Variance_Pct  >  +3
UNFAVORABLE_THRESHOLD_PCT = -3.0    # Unfavorable if Sales_Variance_Pct  <  -3
MATERIAL_PCT_THRESHOLD = 10.0       # Material if |Sales_Variance_Pct|  >= 10
MATERIAL_USD_THRESHOLD = 75000.0    # ... or |Sales_Variance_USD|       >= 75,000
TOP_N = 10
PERSISTENCE_SHARE = 0.75            # "chronic" = below target in >= 75 % of months
TRACKING_SIGNAL_LIMIT = 4.0         # |cumulative variance / MAD| above this = biased
PV_ROW_TOLERANCE_USD = 0.01         # allowed rounding gap per row in price/volume split
YOY_BASE_YEAR = None                # None = auto (last two years in the data)
YOY_CURRENT_YEAR = None

REQUIRED_COLUMNS = ["Date", "Region", "Category", "Product", "Target_Qty",
                    "Actual_Qty", "Unit_Price_USD", "Target_Sales_USD",
                    "Actual_Sales_USD"]

# ------------------------------------------------------------ CHECK LOGGER --
CHECKS = []          # (status, check, detail)


def _p(msg=""):
    print(msg, flush=True)


def record(status, check, detail=""):
    CHECKS.append((status, check, detail))
    tag = {"PASS": "[ PASS ]", "WARN": "[ WARN ]", "FAIL": "[ FAIL ]", "INFO": "[ INFO ]"}[status]
    _p(f"   {tag} {check}" + (f"  -  {detail}" if detail else ""))


def check(cond, name, detail_ok="", detail_bad="", severity="FAIL"):
    record("PASS" if cond else severity, name, detail_ok if cond else detail_bad)
    return bool(cond)


def section(title):
    _p("\n" + "=" * 78 + f"\n {title}\n" + "=" * 78)


def money(x):
    return f"${x:,.2f}"


# ------------------------------------------------------------------ HELPERS --
def safe_div(num, den):
    """Element-wise num/den, NaN where den is 0 / NaN (never raises, never inf)."""
    num = np.asarray(num, dtype="float64")
    den = np.asarray(den, dtype="float64")
    ok = np.isfinite(den) & (den != 0)
    out = np.full(np.broadcast(num, den).shape, np.nan)
    np.divide(num, den, out=out, where=ok)
    return out


def to_cents(s):
    return np.rint(np.asarray(s, dtype="float64") * 100.0).astype("int64")


def thousandths(x, name):
    v = float(x) * 1000.0
    if abs(v - round(v)) > 1e-9:
        raise ValueError(f"{name}={x} must be a multiple of 0.001 for exact classification")
    return int(round(v))


FAV_T = thousandths(FAVORABLE_THRESHOLD_PCT, "FAVORABLE_THRESHOLD_PCT")
UNF_T = thousandths(UNFAVORABLE_THRESHOLD_PCT, "UNFAVORABLE_THRESHOLD_PCT")
MAT_T = thousandths(MATERIAL_PCT_THRESHOLD, "MATERIAL_PCT_THRESHOLD")
MAT_USD_C = int(round(MATERIAL_USD_THRESHOLD * 100))


def classify(var_c, tc):
    """Exact integer classification. var_c, tc are int64 cents.
    pct > T%  <=>  var*100/tc > T/1 <=>  var*100000 > (T*1000) * tc  (tc > 0)."""
    var_c = np.asarray(var_c, dtype="int64")
    tc = np.asarray(tc, dtype="int64")
    valid = tc > 0
    lhs = var_c * 100000
    fav = valid & (lhs > FAV_T * tc)
    unf = valid & (lhs < UNF_T * tc)
    m_pct = valid & (np.abs(var_c) * 100000 >= MAT_T * tc)
    m_usd = np.abs(var_c) >= MAT_USD_C
    status = np.select([~valid, fav, unf], ["No Target", "Favorable", "Unfavorable"],
                       default="On Target")
    mat = m_pct | m_usd
    trig = np.select([m_pct & m_usd, m_pct, m_usd], ["Pct and USD", "Pct", "USD"], default="")
    return status, mat, trig, fav, unf, m_pct, m_usd


def pct2(num, den):
    """num/den x 100 rounded to 2 dp, half AWAY from zero (like Excel), exact in integers.
    NaN where den <= 0.  (np.round would round exact ties like 21.875 to even.)"""
    num = np.asarray(num, dtype="int64")
    den = np.asarray(den, dtype="int64")
    ok = den > 0
    d = np.where(ok, den, 1)
    q = (2 * np.abs(num) * 10000 + d) // (2 * d)
    return np.where(ok, np.sign(num) * q / 100.0, np.nan)


def growth_pct(cur, prev):
    prev = np.asarray(prev, dtype="float64")
    return np.where(prev > 0, safe_div(np.asarray(cur, dtype="float64"), prev) * 100 - 100, np.nan)


# ===========================================================================
# 1. LOAD & VALIDATE INPUT
# ===========================================================================
section("1. LOAD & VALIDATE INPUT")


def find_input():
    here = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    if len(sys.argv) > 1:
        p = Path(sys.argv[1].strip('"'))
        if p.is_file():
            return p
        sys.exit(f"ERROR: file not found: {p}")
    repo = here.parent if here.name.lower() == "scripts" else here
    for folder in (here, Path.cwd(), repo / "data" / "processed"):
        p = folder / INPUT_FILE_NAME
        if p.is_file():
            return p
    sys.exit(f"ERROR: '{INPUT_FILE_NAME}' not found next to this script ({here}) or in the current "
             f"folder ({Path.cwd()}). Put it there, or run:\n"
             f"   python sales_variance_analysis.py \"full\\path\\to\\{INPUT_FILE_NAME}\"")


INPUT_PATH = find_input()
OUT_DIR = INPUT_PATH.parent
# Repo layout: the Excel workbook goes to outputs/ when the input sits in data/processed/ (path wiring only).
_here = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
_repo = _here.parent if _here.name.lower() == "scripts" else _here
XLSX_DIR = OUT_DIR
if _here.name.lower() == "scripts" and OUT_DIR.resolve() == (_repo / "data" / "processed").resolve():
    XLSX_DIR = _repo / "outputs"
    XLSX_DIR.mkdir(parents=True, exist_ok=True)
_p(f"   Reading: {INPUT_PATH}")
# keep_default_na=False: the literal text "NA" (e.g. a region/notes value) must not become NaN
raw = pd.read_csv(INPUT_PATH, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
raw.columns = [str(c).strip() for c in raw.columns]
missing_cols = [c for c in REQUIRED_COLUMNS if c not in raw.columns]
if missing_cols:
    sys.exit(f"ERROR: required column(s) missing: {missing_cols}\n   Found: {list(raw.columns)}")
ORIGINAL_COLUMNS = list(raw.columns)
N = len(raw)
if N == 0:
    sys.exit("ERROR: the input file has no rows.")
record("INFO", "Rows / columns loaded", f"{N:,} rows x {raw.shape[1]} columns")

df = raw.copy()
for c in ("Target_Qty", "Actual_Qty", "Unit_Price_USD", "Target_Sales_USD", "Actual_Sales_USD"):
    df[c] = pd.to_numeric(df[c], errors="coerce")
if "Gross_Margin_USD" in df.columns:
    df["Gross_Margin_USD"] = pd.to_numeric(df["Gross_Margin_USD"], errors="coerce")
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
for c in ("Region", "Category", "Product"):
    df[c] = df[c].astype("string").str.strip()

null_req = {c: int(df[c].isna().sum()) for c in REQUIRED_COLUMNS}
bad = {k: v for k, v in null_req.items() if v}
check(not bad, "No nulls / unparseable values in required columns",
      "all 9 required columns fully populated", f"nulls found: {bad}")
if bad:
    sys.exit("ERROR: fix the nulls above (run clean_sales_data.py first) and re-run.")
for c in ("Target_Qty", "Actual_Qty"):
    if (df[c] % 1 != 0).any():
        record("WARN", f"{c} has non-integer values", "quantities are expected to be whole numbers")
check((df["Actual_Qty"] >= 0).all() and (df["Target_Qty"] >= 0).all(), "No negative quantities")
check((df["Actual_Sales_USD"] >= 0).all() and (df["Target_Sales_USD"] >= 0).all(),
      "No negative sales values")
check((df["Unit_Price_USD"] >= 0).all(), "No negative unit prices")
n_zero_tq = int((df["Target_Qty"] == 0).sum())
n_zero_ts = int((df["Target_Sales_USD"] == 0).sum())
check(n_zero_tq == 0 and n_zero_ts == 0, "No zero targets (divisor safety)",
      "every row has a positive target",
      f"{n_zero_tq} rows Target_Qty=0, {n_zero_ts} rows Target_Sales=0 - handled as 'No Target', never divided",
      severity="WARN")
n_full_dup = int(raw.duplicated().sum())
check(n_full_dup == 0, "No fully identical duplicate rows", "", f"{n_full_dup} identical rows present", "WARN")
# internal period keys come from Date (single source of truth)
df["Month_Ord"] = df["Date"].dt.year * 12 + df["Date"].dt.month - 1
df["Year_Chk"] = df["Date"].dt.year
if "Year" in df.columns:
    check((pd.to_numeric(df["Year"], errors="coerce") == df["Year_Chk"]).all(), "Year column agrees with Date")
else:
    df["Year"] = df["Year_Chk"]
if "Month" in df.columns:
    check((pd.to_numeric(df["Month"], errors="coerce") == df["Date"].dt.month).all(), "Month column agrees with Date")
else:
    df["Month"] = df["Date"].dt.month
df["Year"] = df["Year"].astype("int64")
df["Month"] = df["Month"].astype("int64")
df["Year_Month"] = df["Date"].dt.strftime("%Y-%m")
df["Quarter"] = "Q" + ((df["Date"].dt.month - 1) // 3 + 1).astype(str)
df = df.drop(columns=["Year_Chk"])
months_present = np.sort(df["Month_Ord"].unique())
gaps = int(np.sum(np.diff(months_present) > 1))
check(gaps == 0, "Calendar months are continuous",
      f"{len(months_present)} consecutive months: {df['Year_Month'].min()} to {df['Year_Month'].max()}",
      f"{gaps} gap(s) in the calendar - growth / rolling metrics are blank across gaps", "WARN")
record("INFO", "Regions", ", ".join(sorted(df["Region"].unique())))
record("INFO", "Categories / products", f"{df['Category'].nunique()} categories, {df['Product'].nunique()} products")
prod_cat = df.groupby("Product")["Category"].nunique()
check((prod_cat == 1).all(), "Each product belongs to exactly one category", "",
      f"{int((prod_cat > 1).sum())} product(s) appear in several categories", "WARN")

# ===========================================================================
# 2. ROW-LEVEL METRICS
# ===========================================================================
section("2. ROW-LEVEL METRICS (exact integer-cent arithmetic)")
df["Tq"] = df["Target_Qty"].astype("int64")
df["Aq"] = df["Actual_Qty"].astype("int64")
df["Tc"] = to_cents(df["Target_Sales_USD"])
df["Ac"] = to_cents(df["Actual_Sales_USD"])
df["Vc"] = df["Ac"] - df["Tc"]
big = max(int(df["Tc"].abs().sum()), int(df["Ac"].abs().sum()))
check(big * 100000 < 9.0e18, "Integer arithmetic cannot overflow", f"largest total {big/100:,.0f} USD")

Qty_Variance = df["Aq"] - df["Tq"]
Qty_Att = pct2(df["Aq"].values, df["Tq"].values)
Sales_Var = df["Vc"] / 100.0
Sales_Var_Pct = pct2(df["Vc"].values, df["Tc"].values)

# compare with what the cleaning script already wrote (if present)
compare = {"Qty_Variance": Qty_Variance, "Qty_Attainment_Pct": Qty_Att,
           "Sales_Variance_USD": Sales_Var, "Sales_Variance_Pct": Sales_Var_Pct}
for col, new in compare.items():
    if col in df.columns:
        old = pd.to_numeric(df[col], errors="coerce")
        diff = (~np.isclose(old.fillna(-9e9), pd.Series(new).fillna(-9e9), atol=0.0051, rtol=0)).sum()
        check(diff == 0, f"{col}: re-computed value equals input column", f"all {N:,} rows identical",
              f"{int(diff)} rows differ by 0.01 (exact .xx5 ties: this script rounds half away from zero like Excel; the input was rounded inconsistently); the re-computed value is used", "WARN")
    df[col] = new

status, mat, trig, fav, unf, m_pct, m_usd = classify(df["Vc"].values, df["Tc"].values)
df["is_fav"], df["is_unf"] = fav, unf
df["is_nt"] = (df["Tc"] <= 0).values
df["mat"] = mat
df["Sales_Variance_Status"] = status
df["Material_Variance_Flag"] = np.where(mat, "Yes", "No")
df["Material_Trigger"] = trig
n_f, n_u, n_o, n_n = int(fav.sum()), int(unf.sum()), int((status == "On Target").sum()), int((status == "No Target").sum())
check(n_f + n_u + n_o + n_n == N, "Every row has exactly one status", f"{n_f:,} Favorable + {n_u:,} Unfavorable + {n_o:,} On Target + {n_n:,} No Target = {N:,}")
# cross-check the exact test against an independent float test away from the boundary
pct_f = safe_div(df["Vc"].values, df["Tc"].values) * 100
away = np.isfinite(pct_f) & (np.abs(np.abs(pct_f) - 3) > 1e-6)
chk_status = np.where(~np.isfinite(pct_f), "No Target",
                      np.where(pct_f > FAVORABLE_THRESHOLD_PCT, "Favorable",
                               np.where(pct_f < UNFAVORABLE_THRESHOLD_PCT, "Unfavorable", "On Target")))
n_flip = int(((chk_status != status) & away).sum())
n_edge = int(((chk_status != status) & ~away).sum())
check(n_flip == 0, "Status agrees with a plain float test (rows away from the +/-3 % edge)",
      f"0 disagreements; {n_edge} row(s) within 1e-6 of an edge are decided exactly by cents",
      f"{n_flip} disagreements")
on_edge = int(((df["Vc"] * 100000) == FAV_T * df["Tc"]).sum() + ((df["Vc"] * 100000) == UNF_T * df["Tc"]).sum())
record("INFO", "Rows exactly on a +/-3 % boundary (classified On Target)", f"{on_edge:,}")
check(not (df.loc[df["is_fav"], "Vc"] <= 0).any() and not (df.loc[df["is_unf"], "Vc"] >= 0).any(),
      "Flags are logical: Favorable => variance > 0, Unfavorable => variance < 0")
check((df.loc[df["mat"], "Material_Trigger"] != "").all() and (df.loc[~df["mat"], "Material_Trigger"] == "").all(),
      "Every Material = Yes row has a trigger, every No row has none")
share_mat = df["mat"].mean() * 100
if share_mat > 40:
    record("WARN", "Row-level materiality is not selective",
           f"{share_mat:.1f} % of rows are 'Material' (|pct| >= {MATERIAL_PCT_THRESHOLD:g} % or |USD| >= {MATERIAL_USD_THRESHOLD:,.0f}). "
           f"By USD alone: {int(m_usd.sum())} rows. Use the group-level sheets (Exceptions_RegCat_Month) for triage.")
else:
    record("INFO", "Row-level material share", f"{share_mat:.1f} % of rows; USD rule alone: {int(m_usd.sum())} rows")

section("3. PRICE / VOLUME DECOMPOSITION")
tq_ok = df["Tq"] > 0
tp = np.where(tq_ok, safe_div(df["Target_Sales_USD"], df["Tq"]), np.nan)
df["Target_Price_USD"] = np.round(tp, 6)
vol = np.where(tq_ok, (df["Aq"] - df["Tq"]) * tp, Sales_Var)       # no target price => all volume
prc = np.where(tq_ok, (df["Unit_Price_USD"] - tp) * df["Aq"], 0.0)
df["Volume_Effect_USD"] = np.round(vol, 6)
df["Price_Effect_USD"] = np.round(prc, 6)
df["PV_Residual_USD"] = np.round(Sales_Var - df["Volume_Effect_USD"] - df["Price_Effect_USD"], 6)
res_max = float(df["PV_Residual_USD"].abs().max())
check(res_max <= PV_ROW_TOLERANCE_USD, "Volume + Price effect reconciles to Sales_Variance_USD on every row",
      f"largest row residual {res_max:.6f} USD (cent rounding of Actual_Sales_USD)",
      f"largest row residual {res_max:.4f} USD exceeds {PV_ROW_TOLERANCE_USD} - Actual_Sales_USD is not Actual_Qty x Unit_Price_USD on some rows", "WARN")
tot_var = df["Vc"].sum() / 100
tot_vp = df["Volume_Effect_USD"].sum() + df["Price_Effect_USD"].sum()
check(abs(tot_var - tot_vp) < max(1.0, N * 0.005), "Totals: Volume + Price effect = Total variance",
      f"{money(df['Volume_Effect_USD'].sum())} + {money(df['Price_Effect_USD'].sum())} = {money(tot_vp)} vs {money(tot_var)} "
      f"(gap {money(tot_var - tot_vp)} = pure cent-rounding residual)")
if (df["Price_Effect_USD"].abs() < 1e-6).all():
    record("WARN", "Price effect is zero on every row",
           "Unit_Price_USD equals Target_Sales/Target_Qty everywhere: the target was set at the same price as actuals, "
           "so ALL variance is volume. A price/mix story cannot be told from this dataset.")
elif (df["Price_Effect_USD"].abs() > 0.01).mean() < 0.01:
    record("WARN", "Price effect is almost always zero", "price/volume split is of limited analytical value")

# ===========================================================================
# 4. AGGREGATION ENGINE
# ===========================================================================
TOT_TC = int(df["Tc"].sum())
TOT_AC = int(df["Ac"].sum())
TOT_VC = TOT_AC - TOT_TC
TOT_TQ, TOT_AQ = int(df["Tq"].sum()), int(df["Aq"].sum())
df["Gross_Margin_USD_n"] = df["Gross_Margin_USD"] if "Gross_Margin_USD" in df.columns else np.nan
df["GM"] = df["Gross_Margin_USD_n"].fillna(0.0)
df["Vol"], df["Prc"] = df["Volume_Effect_USD"], df["Price_Effect_USD"]
df["Res"] = df["PV_Residual_USD"]
HAS_GM = "Gross_Margin_USD" in df.columns and df["Gross_Margin_USD"].notna().any()


def summarise(frame, keys):
    g = frame.groupby(keys, sort=True, observed=True)
    s = g.agg(Rows=("Tc", "size"), Target_Qty=("Tq", "sum"), Actual_Qty=("Aq", "sum"),
              _Tc=("Tc", "sum"), _Ac=("Ac", "sum"), Volume_Effect_USD=("Vol", "sum"),
              Price_Effect_USD=("Prc", "sum"), PV_Residual_USD=("Res", "sum"), _GM=("GM", "sum"),
              Rows_Favorable=("is_fav", "sum"), Rows_Unfavorable=("is_unf", "sum"),
              Rows_No_Target=("is_nt", "sum"), Rows_Material=("mat", "sum")).reset_index()
    for c in ("Rows_Favorable", "Rows_Unfavorable", "Rows_No_Target", "Rows_Material"):
        s[c] = s[c].astype("int64")
    s["_Vc"] = s["_Ac"] - s["_Tc"]
    s["Qty_Variance"] = s["Actual_Qty"] - s["Target_Qty"]
    s["Qty_Attainment_Pct"] = pct2(s["Actual_Qty"].values, s["Target_Qty"].values)
    s["Target_Sales_USD"] = s["_Tc"] / 100.0
    s["Actual_Sales_USD"] = s["_Ac"] / 100.0
    s["Sales_Variance_USD"] = s["_Vc"] / 100.0
    s["Sales_Variance_Pct"] = pct2(s["_Vc"].values, s["_Tc"].values)
    st, mt, tg, *_ = classify(s["_Vc"].values, s["_Tc"].values)
    s["Sales_Variance_Status"], s["Material_Variance_Flag"] = st, np.where(mt, "Yes", "No")
    s["Material_Trigger"] = tg
    s["Rows_On_Target"] = s["Rows"] - s["Rows_Favorable"] - s["Rows_Unfavorable"] - s["Rows_No_Target"]
    s["Pct_Rows_Unfavorable"] = np.round(safe_div(s["Rows_Unfavorable"], s["Rows"]) * 100, 2)
    # volume split: pure volume (at the group's average target price) and mix (the remainder)
    pure = (s["Actual_Qty"] - s["Target_Qty"]) * safe_div(s["_Tc"] / 100.0, s["Target_Qty"])
    pure = np.where(s["Target_Qty"] > 0, pure, s["Volume_Effect_USD"])
    s["Volume_Pure_USD"] = np.round(pure, 2)
    s["Mix_Effect_USD"] = np.round(s["Volume_Effect_USD"] - pure, 2)
    for c in ("Volume_Effect_USD", "Price_Effect_USD", "PV_Residual_USD"):
        s[c] = np.round(s[c], 2)
    s["Share_of_Target_Pct"] = np.round(safe_div(s["_Tc"], TOT_TC) * 100, 3)
    s["Contribution_to_Total_Var_pp"] = np.round(safe_div(s["_Vc"], TOT_TC) * 100, 4)
    s["Avg_Realised_Price_USD"] = np.round(safe_div(s["_Ac"] / 100.0, s["Actual_Qty"]), 2)
    if HAS_GM:
        s["Gross_Margin_USD"] = np.round(s["_GM"], 2)
        s["Gross_Margin_Pct"] = np.round(safe_div(s["_GM"], s["_Ac"] / 100.0) * 100, 2)
    return s


def tie_out(tbl, name):
    """Every summary table must add back to the grand totals."""
    ok = (int(tbl["_Tc"].sum()) == TOT_TC and int(tbl["_Ac"].sum()) == TOT_AC and
          int(tbl["Rows"].sum()) == N and int(tbl["Actual_Qty"].sum()) == TOT_AQ and
          int(tbl["Target_Qty"].sum()) == TOT_TQ)
    check(ok, f"{name}: rows, quantities, target and actual sales tie to grand totals",
          f"{len(tbl):,} groups", "totals do not tie to the row-level data")
    return ok


def lag_cols(tbl, ord_col, lag, tag):
    prev = tbl[[ord_col, "_Tc", "_Ac", "_Vc"]].copy()
    prev[ord_col] = prev[ord_col] + lag
    prev = prev.rename(columns={"_Tc": f"_Tc_{tag}", "_Ac": f"_Ac_{tag}", "_Vc": f"_Vc_{tag}"})
    return tbl.merge(prev, on=ord_col, how="left")


def add_growth(tbl, ord_col, lag, tag, label):
    t = lag_cols(tbl, ord_col, lag, tag)
    t[f"{label}_Actual_Growth_Pct"] = np.round(growth_pct(t["_Ac"], t[f"_Ac_{tag}"]), 2)
    t[f"{label}_Target_Growth_Pct"] = np.round(growth_pct(t["_Tc"], t[f"_Tc_{tag}"]), 2)
    t[f"{label}_Variance_Change_USD"] = np.round((t["_Vc"] - t[f"_Vc_{tag}"]) / 100.0, 2)
    att_cur = safe_div(t["_Ac"], t["_Tc"]) * 100
    att_prev = safe_div(t[f"_Ac_{tag}"], t[f"_Tc_{tag}"]) * 100
    t[f"{label}_Attainment_Change_pp"] = np.round(att_cur - att_prev, 2)
    return t.drop(columns=[f"_Tc_{tag}", f"_Ac_{tag}", f"_Vc_{tag}"])


section("4. SUMMARY TABLES")
TABLES = {}

# ---- monthly -----------------------------------------------------------------
m = summarise(df, ["Month_Ord"]).sort_values("Month_Ord").reset_index(drop=True)
m["Year"] = m["Month_Ord"] // 12
m["Month"] = m["Month_Ord"] % 12 + 1
m["Year_Month"] = m["Year"].astype(str) + "-" + m["Month"].astype(str).str.zfill(2)
m["Quarter"] = "Q" + ((m["Month"] - 1) // 3 + 1).astype(str)
tie_out(m, "By_Year_Month")
m = add_growth(m, "Month_Ord", 1, "m1", "MoM")
m = add_growth(m, "Month_Ord", 12, "y1", "YoY")
m = m.sort_values("Month_Ord").reset_index(drop=True)
m["YTD_Target_USD"] = m.groupby("Year")["_Tc"].cumsum() / 100.0
m["YTD_Actual_USD"] = m.groupby("Year")["_Ac"].cumsum() / 100.0
m["YTD_Variance_USD"] = np.round(m["YTD_Actual_USD"] - m["YTD_Target_USD"], 2)
m["YTD_Variance_Pct"] = np.round(safe_div(m["YTD_Variance_USD"], m["YTD_Target_USD"]) * 100, 2)
full = pd.RangeIndex(m["Month_Ord"].min(), m["Month_Ord"].max() + 1)
r3 = (m.set_index("Month_Ord")[["_Tc", "_Ac"]].reindex(full).rolling(3, min_periods=3).sum())
m["Rolling3M_Variance_USD"] = np.round(((r3["_Ac"] - r3["_Tc"]) / 100.0).reindex(m["Month_Ord"]).values, 2)
m["Rolling3M_Variance_Pct"] = np.round(safe_div((r3["_Ac"] - r3["_Tc"]).reindex(m["Month_Ord"]).values,
                                                r3["_Tc"].reindex(m["Month_Ord"]).values) * 100, 2)
m["Cumulative_Variance_USD"] = m["_Vc"].cumsum() / 100.0
mad = m["_Vc"].abs().expanding().mean() / 100.0
m["Tracking_Signal"] = np.round(safe_div(m["Cumulative_Variance_USD"], mad), 2)
m["Forecast_Bias"] = np.where(m["Tracking_Signal"].abs() > TRACKING_SIGNAL_LIMIT,
                              np.where(m["Tracking_Signal"] < 0, "Persistent under-delivery", "Persistent over-delivery"), "Within limit")
TABLES["By_Year_Month"] = m
mom_ok = m["MoM_Actual_Growth_Pct"].notna().sum() == len(m) - 1 - gaps
check(mom_ok, "MoM growth exists for every month except the first", f"{int(m['MoM_Actual_Growth_Pct'].notna().sum())} values")
n_yoy_exp = int(m["Month_Ord"].isin(m["Month_Ord"] + 12).sum())
check(int(m["YoY_Actual_Growth_Pct"].notna().sum()) == n_yoy_exp,
      "YoY growth exists for every month that has a same month one year earlier",
      f"{int(m['YoY_Actual_Growth_Pct'].notna().sum())} values")
last_ytd = m.groupby("Year").tail(1)
check(abs(last_ytd["YTD_Variance_USD"].sum() - TOT_VC / 100) < 0.01,
      "Final YTD figure of each year adds to the grand total variance")

# ---- quarter / year ---------------------------------------------------------
q = summarise(df, ["Year", "Quarter"]).sort_values(["Year", "Quarter"]).reset_index(drop=True)
q["Q_Ord"] = q["Year"] * 4 + q["Quarter"].str[1].astype(int) - 1
tie_out(q, "By_Quarter")
q = add_growth(q, "Q_Ord", 1, "q1", "QoQ")
q = add_growth(q, "Q_Ord", 4, "q4", "YoY").drop(columns=["Q_Ord"])
TABLES["By_Quarter"] = q
y = summarise(df, ["Year"]).sort_values("Year").reset_index(drop=True)
tie_out(y, "By_Year")
y = add_growth(y, "Year", 1, "y1", "YoY")
TABLES["By_Year"] = y

# ---- dimension tables with YoY ------------------------------------------------
years = sorted(df["Year"].unique())
if YOY_CURRENT_YEAR is not None and YOY_BASE_YEAR is not None:
    Y0, Y1 = int(YOY_BASE_YEAR), int(YOY_CURRENT_YEAR)
elif len(years) >= 2:
    Y0, Y1 = int(years[-2]), int(years[-1])
else:
    Y0 = Y1 = None
record("INFO", "Year-on-year comparison", f"{Y1} vs {Y0}" if Y0 is not None and Y0 != Y1 else "only one year in the data - YoY columns skipped")


def add_yoy(tbl, keys):
    if Y0 is None or Y0 == Y1:
        return tbl
    yy = df.groupby(keys + ["Year"], observed=True).agg(_a=("Ac", "sum"), _t=("Tc", "sum")).reset_index()
    a = yy[yy["Year"] == Y0].drop(columns="Year").rename(columns={"_a": "_a0", "_t": "_t0"})
    b = yy[yy["Year"] == Y1].drop(columns="Year").rename(columns={"_a": "_a1", "_t": "_t1"})
    w = a.merge(b, on=keys, how="outer").fillna({"_a0": 0, "_t0": 0, "_a1": 0, "_t1": 0})
    out = pd.DataFrame({k: w[k] for k in keys})
    out[f"Actual_Sales_{Y0}"] = w["_a0"] / 100.0
    out[f"Actual_Sales_{Y1}"] = w["_a1"] / 100.0
    out["YoY_Actual_Growth_Pct"] = np.round(growth_pct(w["_a1"], w["_a0"]), 2)
    out[f"Variance_USD_{Y0}"] = (w["_a0"] - w["_t0"]) / 100.0
    out[f"Variance_USD_{Y1}"] = (w["_a1"] - w["_t1"]) / 100.0
    out[f"Attainment_Pct_{Y0}"] = np.round(safe_div(w["_a0"], w["_t0"]) * 100, 2)
    out[f"Attainment_Pct_{Y1}"] = np.round(safe_div(w["_a1"], w["_t1"]) * 100, 2)
    out["YoY_Attainment_Change_pp"] = np.round(out[f"Attainment_Pct_{Y1}"] - out[f"Attainment_Pct_{Y0}"], 2)
    return tbl.merge(out, on=keys, how="left")


for name, keys in [("By_Region", ["Region"]), ("By_Category", ["Category"]),
                   ("By_Region_Category", ["Region", "Category"]), ("By_Product", ["Category", "Product"]),
                   ("_RegProd", ["Region", "Product"])]:
    t = summarise(df, keys)
    tie_out(t, name.lstrip("_"))
    t = add_yoy(t, keys)
    t = t.sort_values("_Vc").reset_index(drop=True) if name in ("By_Product", "_RegProd") else t
    TABLES[name] = t
check(len(TABLES["By_Region"]) == 5, "Exactly 5 regions in the summary", ", ".join(TABLES["By_Region"]["Region"]),
      f"{len(TABLES['By_Region'])} regions found", "WARN")

# ---- top / bottom ---------------------------------------------------------------
bp = TABLES["By_Product"].copy()
bp["Above_Target"] = np.where(bp["_Vc"] > 0, "Yes", "No")
top = bp.sort_values(["_Vc", "Product"], ascending=[False, True]).head(TOP_N).reset_index(drop=True)
bot = bp.sort_values(["_Vc", "Product"], ascending=[True, True]).head(TOP_N).reset_index(drop=True)
for t_ in (top, bot):
    t_.insert(0, "Rank", range(1, len(t_) + 1))
TABLES["Top_10_Products"], TABLES["Bottom_10_Products"] = top, bot
n_above = int((bp["_Vc"] > 0).sum())
check(top["Sales_Variance_USD"].is_monotonic_decreasing and bot["Sales_Variance_USD"].is_monotonic_increasing,
      "Top / Bottom lists are correctly ordered")
if n_above < len(bp) / 2:
    record("WARN", "'Top 10' is a ranking, not a list of winners",
           f"only {n_above} of {len(bp)} products beat target; {len(bp) - n_above} are below. "
           f"Read Top-10 as 'least bad' where Above_Target = No.")

# ---- PVM bridge ----------------------------------------------------------------
pv_parts = []
tot_row = summarise(df.assign(_All="Total"), ["_All"])
tot_row.insert(0, "Level", "Total")
tot_row.insert(1, "Name", "All sales")
pv_parts.append(tot_row.drop(columns=["_All"]))
for lvl, key in (("Region", "Region"), ("Category", "Category")):
    t = TABLES[f"By_{lvl}"].copy()
    t.insert(0, "Level", lvl)
    t.insert(1, "Name", t[key])
    pv_parts.append(t)
pvm = pd.concat(pv_parts, ignore_index=True)
pvm["Bridge_Check_USD"] = np.round(pvm["Target_Sales_USD"] + pvm["Volume_Pure_USD"] + pvm["Mix_Effect_USD"] +
                                   pvm["Price_Effect_USD"] + pvm["PV_Residual_USD"] - pvm["Actual_Sales_USD"], 2)
TABLES["PVM_Bridge"] = pvm
check(pvm["Bridge_Check_USD"].abs().max() <= 0.05, "PVM bridge: Target + Volume + Mix + Price + Residual = Actual (all levels)",
      f"largest gap {pvm['Bridge_Check_USD'].abs().max():.2f} USD (2-dp rounding of components)",
      f"largest gap {pvm['Bridge_Check_USD'].abs().max():.2f} USD", "WARN")

# ---- monthly region x category (exceptions / persistence / heatmaps) -----------------
rcm = summarise(df, ["Region", "Category", "Month_Ord"])
rcm["Year_Month"] = (rcm["Month_Ord"] // 12).astype(str) + "-" + (rcm["Month_Ord"] % 12 + 1).astype(str).str.zfill(2)
tie_out(rcm, "Region x Category x Month")
exc = rcm[rcm["Material_Variance_Flag"] == "Yes"].copy()
exc["_abs"] = exc["_Vc"].abs()
exc = exc.sort_values(["_abs", "Region", "Category", "Month_Ord"], ascending=[False, True, True, True]).drop(columns="_abs")
TABLES["Exceptions_RegCat_Month"] = exc
record("INFO", "Region x Category x Month cells flagged Material", f"{len(exc)} of {len(rcm)}")

rm = summarise(df, ["Region", "Month_Ord"])
rm["Year_Month"] = (rm["Month_Ord"] // 12).astype(str) + "-" + (rm["Month_Ord"] % 12 + 1).astype(str).str.zfill(2)
tie_out(rm, "Region x Month")
heat = rm.pivot(index="Region", columns="Year_Month", values="Sales_Variance_Pct")
heat = heat.reindex(sorted(heat.columns), axis=1)
heat["Full_Period"] = TABLES["By_Region"].set_index("Region")["Sales_Variance_Pct"]
TABLES["Heatmap_Region_Month_Pct"] = heat.reset_index()

gm_ = rcm.groupby(["Region", "Category"]).agg(
    Months=("Month_Ord", "nunique"), Months_Below_Target=("_Vc", lambda s: int((s < 0).sum())),
    Months_Unfavorable=("Sales_Variance_Status", lambda s: int((s == "Unfavorable").sum())),
    Avg_Monthly_Variance_USD=("Sales_Variance_USD", "mean"), Worst_Month_Variance_USD=("Sales_Variance_USD", "min")
).reset_index()
gm_["Share_Months_Below_Target_Pct"] = np.round(gm_["Months_Below_Target"] / gm_["Months"] * 100, 1)
gm_["Chronic_Shortfall"] = np.where(gm_["Months_Below_Target"] / gm_["Months"] >= PERSISTENCE_SHARE, "Yes", "No")
gm_["Avg_Monthly_Variance_USD"] = gm_["Avg_Monthly_Variance_USD"].round(2)
worst_idx = rcm.loc[rcm.groupby(["Region", "Category"])["_Vc"].idxmin(), ["Region", "Category", "Year_Month"]]
gm_ = gm_.merge(worst_idx.rename(columns={"Year_Month": "Worst_Month"}), on=["Region", "Category"], how="left")
TABLES["Persistence_Region_Category"] = gm_.sort_values("Share_Months_Below_Target_Pct", ascending=False).reset_index(drop=True)

short = TABLES["By_Product"][TABLES["By_Product"]["_Vc"] < 0].sort_values("_Vc").copy()
if len(short):
    cs = short["_Vc"].cumsum()
    short["Shortfall_USD"] = -short["Sales_Variance_USD"]
    short["Share_of_Shortfall_Pct"] = np.round(short["_Vc"] / short["_Vc"].sum() * 100, 2)
    short["Cumulative_Share_Pct"] = np.round(cs / short["_Vc"].sum() * 100, 2)
    short["Pareto_Class"] = np.where((cs - short["_Vc"]) / short["_Vc"].sum() < 0.8, "A (first 80%)", "B/C (tail)")
    TABLES["Pareto_Shortfall_Product"] = short[["Category", "Product", "Shortfall_USD", "Sales_Variance_Pct",
                                                "Share_of_Shortfall_Pct", "Cumulative_Share_Pct", "Pareto_Class", "_Tc", "_Ac", "_Vc"]]
n_chronic = int((gm_["Chronic_Shortfall"] == "Yes").sum())
record("INFO", "Region x Category pairs below target in >= 75 % of months", f"{n_chronic} of {len(gm_)}")

# ---- headline checks on the monthly table --------------------------------------------
n_below = int((m["_Vc"] < 0).sum())
ts_last = float(m["Tracking_Signal"].iloc[-1])
if n_below == len(m):
    record("WARN", "EVERY month is below target", f"{n_below} of {len(m)} months; tracking signal {ts_last:.1f} "
           f"(limit +/-{TRACKING_SIGNAL_LIMIT:g}). Targets look systematically too high - a bias, not random noise.")
elif abs(ts_last) > TRACKING_SIGNAL_LIMIT:
    record("WARN", "Tracking signal breaches limit", f"{ts_last:.1f}")

# ===========================================================================
# 5. ENRICHED DATA
# ===========================================================================
section("5. ENRICHED DATASET")
NEW_COLUMNS = ["Quarter", "Year_Month", "Target_Price_USD", "Volume_Effect_USD", "Price_Effect_USD",
               "PV_Residual_USD", "Sales_Variance_Status", "Material_Variance_Flag", "Material_Trigger"]
ordered = [c for c in ORIGINAL_COLUMNS if c in df.columns] + [c for c in NEW_COLUMNS if c not in ORIGINAL_COLUMNS]
enriched = df[ordered].copy()
enriched["Date"] = enriched["Date"].dt.date
check(list(enriched.columns[:len(ORIGINAL_COLUMNS)]) == ORIGINAL_COLUMNS, "Original columns keep their names and order")
check(len(enriched) == N, "No rows added or dropped", f"{N:,} rows")
check(abs(enriched["Sales_Variance_USD"].sum() - TOT_VC / 100) < 0.01, "Enriched variance sums to grand total")
new_num = ["Target_Price_USD", "Volume_Effect_USD", "Price_Effect_USD", "PV_Residual_USD"]
check(not enriched[new_num].isna().any().any() or n_zero_tq > 0, "No unexpected nulls in new numeric columns",
      "Target_Price_USD, Volume, Price, Residual fully populated")
check(not np.isinf(enriched.select_dtypes("number").to_numpy(dtype="float64")).any(), "No infinite values anywhere")
n_nan_pct = int(enriched["Sales_Variance_Pct"].isna().sum())
check(n_nan_pct == n_zero_ts, "Blank percentages only where target is zero", f"{n_nan_pct} blank")

# ===========================================================================
# 6. INSIGHTS
# ===========================================================================
section("6. INSIGHTS")
reg, cat, prod = TABLES["By_Region"], TABLES["By_Category"], TABLES["By_Product"]
tot_att = TOT_AC / TOT_TC * 100 if TOT_TC else float("nan")
tot_pct = TOT_VC / TOT_TC * 100 if TOT_TC else float("nan")
worst_reg, best_reg = reg.loc[reg["_Vc"].idxmin()], reg.loc[reg["_Vc"].idxmax()]
worst_cat, best_cat = cat.loc[cat["_Vc"].idxmin()], cat.loc[cat["_Vc"].idxmax()]
worst_prod = prod.iloc[0]
rc = TABLES["By_Region_Category"]
worst_rc = rc.loc[rc["_Vc"].idxmin()]
top3_short = prod.head(3)
share_top3 = top3_short["_Vc"].sum() / prod[prod["_Vc"] < 0]["_Vc"].sum() * 100 if (prod["_Vc"] < 0).any() else float("nan")
ins = []
ins.append(f"Overall: actual sales {money(TOT_AC/100)} vs target {money(TOT_TC/100)} = variance {money(TOT_VC/100)} ({tot_pct:+.2f} %). "
           f"Quantity attainment {TOT_AQ/TOT_TQ*100:.2f} %.")
ins.append(f"Worst region: {worst_reg['Region']} ({money(worst_reg['Sales_Variance_USD'])}, {worst_reg['Sales_Variance_Pct']:+.2f} %). "
           f"Best region: {best_reg['Region']} ({money(best_reg['Sales_Variance_USD'])}, {best_reg['Sales_Variance_Pct']:+.2f} %).")
ins.append(f"Worst category: {worst_cat['Category']} ({money(worst_cat['Sales_Variance_USD'])}); best: {best_cat['Category']} ({money(best_cat['Sales_Variance_USD'])}).")
ins.append(f"Largest Region x Category gap: {worst_rc['Region']} / {worst_rc['Category']} ({money(worst_rc['Sales_Variance_USD'])}, {worst_rc['Sales_Variance_Pct']:+.2f} %).")
ins.append(f"Product concentration: {worst_prod['Product']} is the largest shortfall ({money(worst_prod['Sales_Variance_USD'])}); the three worst products "
           f"explain {share_top3:.1f} % of all product-level shortfall.")
ins.append(f"Breadth: {n_above} of {len(prod)} products are above target; {n_below} of {len(m)} months are below target.")
if Y0 is not None and Y0 != Y1:
    yy_ = TABLES["By_Year"].set_index("Year")
    ins.append(f"{Y1} vs {Y0}: actual sales {growth_pct(yy_.loc[Y1, '_Ac'], yy_.loc[Y0, '_Ac']):+.2f} %, target {growth_pct(yy_.loc[Y1, '_Tc'], yy_.loc[Y0, '_Tc']):+.2f} %, "
               f"attainment {yy_.loc[Y0, 'Sales_Variance_Pct']:+.2f} % -> {yy_.loc[Y1, 'Sales_Variance_Pct']:+.2f} % (variance %).")
ins.append(f"Tracking signal at the last month: {ts_last:.1f} (limit +/-{TRACKING_SIGNAL_LIMIT:g}) -> {m['Forecast_Bias'].iloc[-1]}."
           + (" A signal this far out means the target is biased, not unlucky." if abs(ts_last) > TRACKING_SIGNAL_LIMIT else ""))
pv_vol, pv_prc = df["Volume_Effect_USD"].sum(), df["Price_Effect_USD"].sum()
ins.append(f"Driver split: volume {money(pv_vol)}, price {money(pv_prc)}, rounding residual {money(df['PV_Residual_USD'].sum())}."
           + (" Price explains nothing - the gap is entirely about units sold." if abs(pv_prc) < 1 else ""))
ins.append(f"Row-level materiality flags {share_mat:.1f} % of lines; at Region x Category x Month level {len(exc)} of {len(rcm)} cells are material.")
for s_ in ins:
    _p("   * " + s_)
ACTIONS = [
    f"Re-base or challenge the target for {worst_reg['Region']} and {worst_cat['Category']}: they carry the biggest absolute gap.",
    f"Run a root-cause review on the worst products (start with {', '.join(top3_short['Product'])}).",
    f"Review the {n_chronic} Region x Category pairs that miss target in >= {PERSISTENCE_SHARE*100:.0f} % of months (sheet Persistence_Region_Category) - these are structural, not one-offs.",
    "Treat Top/Bottom product lists as rankings; check Above_Target before calling anything a 'winner'.",
]

# ===========================================================================
# 7. WORKBOOK
# ===========================================================================
section("7. WRITE EXCEL + CSV")
try:
    import xlsxwriter  # noqa: F401
    ENGINE = "xlsxwriter"
except ImportError:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "xlsxwriter"])
        import xlsxwriter  # noqa: F401
        ENGINE = "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa: F401
            ENGINE = "openpyxl"
            record("WARN", "xlsxwriter unavailable", "using openpyxl: workbook is written without formatting")
        except ImportError:
            sys.exit("ERROR: no Excel writer found. Run:  pip install xlsxwriter")

DICT = [
    ("Qty_Variance", "Actual_Qty - Target_Qty", "Units over (+) or under (-) plan."),
    ("Qty_Attainment_Pct", "Actual_Qty / Target_Qty x 100", "100 = exactly on plan."),
    ("Sales_Variance_USD", "Actual_Sales_USD - Target_Sales_USD", "Positive = favourable for revenue."),
    ("Sales_Variance_Pct", "Sales_Variance_USD / Target_Sales_USD x 100", "Blank when target is zero. Groups use sum/sum, never an average of row %."),
    ("Sales_Variance_Status", f"Favorable if pct > {FAVORABLE_THRESHOLD_PCT:g}; Unfavorable if pct < {UNFAVORABLE_THRESHOLD_PCT:g}; else On Target; 'No Target' if target = 0",
     "Decided in integer cents, so a value exactly on the threshold is On Target."),
    ("Material_Variance_Flag", f"Yes if |pct| >= {MATERIAL_PCT_THRESHOLD:g} OR |USD| >= {MATERIAL_USD_THRESHOLD:,.0f}", "Material_Trigger says which rule fired."),
    ("Target_Price_USD", "Target_Sales_USD / Target_Qty", "Price the plan assumed."),
    ("Volume_Effect_USD", "(Actual_Qty - Target_Qty) x Target_Price", "Variance caused by selling more / fewer units at plan price."),
    ("Price_Effect_USD", "(Unit_Price_USD - Target_Price) x Actual_Qty", "Variance caused by selling at a different price than planned."),
    ("PV_Residual_USD", "Sales_Variance_USD - Volume - Price", "Cent rounding only; should be ~0."),
    ("Volume_Pure_USD / Mix_Effect_USD", "Pure = group qty variance x group avg target price; Mix = Volume - Pure", "Mix = shift between cheap and dear products."),
    ("Contribution_to_Total_Var_pp", "Group variance / TOTAL target x 100", "Percentage points of the company-wide variance; adds to the total."),
    ("MoM / YoY growth", "current / prior - 1, matched on calendar month", "Blank when no comparable period exists."),
    ("YTD_*", "Running total within the calendar year", ""),
    ("Rolling3M_*", "Sum of the last 3 consecutive months", "Blank until 3 months exist."),
    ("Tracking_Signal", "Cumulative variance / mean absolute monthly variance", f"|value| > {TRACKING_SIGNAL_LIMIT:g} = persistent bias."),
    ("Chronic_Shortfall", f"Below target in >= {PERSISTENCE_SHARE*100:.0f} % of months", ""),
    ("Pareto_Class", "Products ordered by shortfall; 'A' = first 80 % of total shortfall", ""),
]


def public(t):
    return t[[c for c in t.columns if not str(c).startswith("_")]]


def build_sheets():
    sh = {}
    es = [("Rows analysed", f"{N:,}", ""),
          ("Period", f"{df['Year_Month'].min()} to {df['Year_Month'].max()}", f"{len(m)} months"),
          ("Target sales (USD)", TOT_TC / 100, ""), ("Actual sales (USD)", TOT_AC / 100, ""),
          ("Total variance (USD)", TOT_VC / 100, f"{tot_pct:+.2f} %"),
          ("Quantity attainment %", round(TOT_AQ / TOT_TQ * 100, 2), ""),
          ("Rows Favorable / Unfavorable / On Target", f"{n_f:,} / {n_u:,} / {n_o:,}", f"No Target: {n_n}"),
          ("Rows flagged Material", int(mat.sum()), f"{share_mat:.1f} % of rows"),
          ("Volume effect (USD)", round(pv_vol, 2), ""), ("Price effect (USD)", round(pv_prc, 2), ""),
          ("Tracking signal (last month)", ts_last, m["Forecast_Bias"].iloc[-1]), ("", "", "")]
    es += [("Insight", s_, "") for s_ in ins] + [("Suggested action", a, "") for a in ACTIONS]
    sh["Executive_Summary"] = pd.DataFrame(es, columns=["Item", "Value", "Comment"])
    sh["Validation_Summary"] = pd.DataFrame(CHECKS, columns=["Status", "Check", "Detail"])
    sh["Metric_Dictionary"] = pd.DataFrame(DICT, columns=["Metric", "Definition", "Notes"])
    sh["Enriched_Data"] = enriched
    for k in ("By_Year_Month", "By_Quarter", "By_Year", "By_Region", "By_Category", "By_Region_Category",
              "By_Product", "Top_10_Products", "Bottom_10_Products", "PVM_Bridge"):
        sh[k] = public(TABLES[k])
    sh["Material_Variances"] = enriched[enriched["Material_Variance_Flag"] == "Yes"].assign(
        _a=lambda d: d["Sales_Variance_USD"].abs()).sort_values("_a", ascending=False).drop(columns="_a")
    sh["Exceptions_RegCat_Month"] = public(TABLES["Exceptions_RegCat_Month"])
    sh["Heatmap_Region_Month_Pct"] = TABLES["Heatmap_Region_Month_Pct"]
    sh["Persistence_Region_Category"] = TABLES["Persistence_Region_Category"]
    if "Pareto_Shortfall_Product" in TABLES:
        sh["Pareto_Shortfall_Product"] = public(TABLES["Pareto_Shortfall_Product"])
    return sh


def fmt_for(col, dtype):
    c = str(col)
    if c in ("Year", "Month", "Source_Excel_Row", "Rank") or c.endswith("_Year"):
        return "0"
    if pd.api.types.is_integer_dtype(dtype):
        return "#,##0"
    if pd.api.types.is_float_dtype(dtype):
        if "USD" in c:
            return "#,##0.00;[Red]-#,##0.00"
        return "0.00"
    return None


def write_workbook(path, sheets):
    with pd.ExcelWriter(path, engine=ENGINE, datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
        for name, d in sheets.items():
            d = d.copy()
            d.to_excel(xw, sheet_name=name[:31], index=False)
            if ENGINE != "xlsxwriter":
                continue
            wb, ws = xw.book, xw.sheets[name[:31]]
            head = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1,
                                  "text_wrap": True, "valign": "vcenter", "align": "center"})
            wrap = wb.add_format({"text_wrap": True, "valign": "top"})
            for j, col in enumerate(d.columns):
                ws.write(0, j, col, head)
                f = fmt_for(col, d[col].dtype)
                sample = d[col].astype(str).fillna("").head(300)
                mx = sample.str.len().max() if len(sample) else 8
                mx = 8 if pd.isna(mx) else int(mx)
                width = min(max(len(str(col)) + 2, mx + 2, 9), 45)
                if name in ("Executive_Summary", "Validation_Summary", "Metric_Dictionary"):
                    width = {"Item": 38, "Value": 95, "Comment": 40, "Status": 9, "Check": 70, "Detail": 90,
                             "Metric": 34, "Definition": 70, "Notes": 70}.get(str(col), width)
                    ws.set_column(j, j, width, wrap)
                else:
                    ws.set_column(j, j, width, wb.add_format({"num_format": f}) if f else None)
            ws.set_row(0, 32)
            ws.freeze_panes(1, 0)
            if len(d) and len(d.columns):
                ws.autofilter(0, 0, len(d), len(d.columns) - 1)
            ws.set_landscape()
            ws.repeat_rows(0)
            if len(d.columns) <= 14:
                ws.fit_to_pages(1, 0)
            g_ = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
            r_ = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            a_ = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C5700"})
            n = len(d)
            for j, col in enumerate(d.columns):
                rng = (1, j, max(n, 1), j)
                if col == "Sales_Variance_Status" or col == "Forecast_Bias":
                    for txt, fm in (("Favorable", g_), ("Unfavorable", r_), ("Persistent under-delivery", r_),
                                    ("Persistent over-delivery", g_)):
                        ws.conditional_format(*rng, {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": fm})
                elif col in ("Material_Variance_Flag", "Chronic_Shortfall"):
                    ws.conditional_format(*rng, {"type": "cell", "criteria": "==", "value": '"Yes"', "format": a_})
                elif name == "Validation_Summary" and col == "Status":
                    for txt, fm in (("PASS", g_), ("FAIL", r_), ("WARN", a_)):
                        ws.conditional_format(*rng, {"type": "cell", "criteria": "==", "value": f'"{txt}"', "format": fm})
            if name.startswith("Heatmap") and n:
                ws.conditional_format(1, 1, n, len(d.columns) - 1, {"type": "3_color_scale", "min_color": "#F8696B",
                                      "mid_color": "#FFFFFF", "max_color": "#63BE7B", "mid_type": "num", "mid_value": 0})


def save_with_fallback(path, writer):
    try:
        writer(path)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{_dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        record("WARN", f"{path.name} is open in another program", f"saved as {alt.name} instead")
        writer(alt)
        return alt


# the validation sheet is written last-known: build once now, rebuild after the final checks
SHEETS = build_sheets()
FINAL_XLSX = save_with_fallback(XLSX_DIR / OUT_XLSX_NAME, lambda p: write_workbook(p, SHEETS))
FINAL_CSV = save_with_fallback(OUT_DIR / OUT_CSV_NAME,
                               lambda p: enriched.to_csv(p, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d"))
_p(f"   Saved: {FINAL_XLSX}")
_p(f"   Saved: {FINAL_CSV}")

# ---- verify what was saved -------------------------------------------------------
xl = pd.ExcelFile(FINAL_XLSX)
check(set(xl.sheet_names) == {k[:31] for k in SHEETS}, "Workbook contains every expected sheet", f"{len(xl.sheet_names)} sheets")
back = pd.read_excel(xl, sheet_name="Enriched_Data", keep_default_na=False)
check(len(back) == N and abs(pd.to_numeric(back["Actual_Sales_USD"]).sum() - TOT_AC / 100) < 0.01,
      "Saved Enriched_Data sheet re-opens with the same rows and sales total")
csv_back = pd.read_csv(FINAL_CSV, keep_default_na=False, encoding="utf-8-sig")
check(len(csv_back) == N and abs(csv_back["Sales_Variance_USD"].sum() - TOT_VC / 100) < 0.01,
      "Saved CSV re-opens with the same rows and variance total")
mv = pd.read_excel(xl, sheet_name="Material_Variances", keep_default_na=False)
check(len(mv) == int(mat.sum()), "Material_Variances sheet row count equals the material flag count", f"{len(mv):,}")

# ---- final summary ----------------------------------------------------------------------
n_pass = sum(1 for c in CHECKS if c[0] == "PASS")
n_warn = sum(1 for c in CHECKS if c[0] == "WARN")
n_fail = sum(1 for c in CHECKS if c[0] == "FAIL")
section("VALIDATION SUMMARY")
_p(f"   Total rows analysed ............ {N:,}")
_p(f"   Total target sales ............. {money(TOT_TC/100)}")
_p(f"   Total actual sales ............. {money(TOT_AC/100)}")
_p(f"   Total sales variance ........... {money(TOT_VC/100)}  ({tot_pct:+.2f} %)")
_p(f"   Favorable / Unfavorable / On Target / No Target ... {n_f:,} / {n_u:,} / {n_o:,} / {n_n:,}")
_p(f"   Material variances (rows) ...... {int(mat.sum()):,}  ({share_mat:.1f} % of rows)")
_p(f"   Price + Volume reconcile ....... {'YES' if abs(tot_var - tot_vp) < max(1.0, N * 0.005) else 'NO'}  "
   f"(volume {money(pv_vol)} + price {money(pv_prc)} + residual {money(df['PV_Residual_USD'].sum())} = {money(tot_var)})")
_p(f"   Checks: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL")
if n_warn:
    _p("\n   WARNINGS (read these - they are findings, not errors):")
    for s_, c_, d_ in CHECKS:
        if s_ == "WARN":
            _p(f"     - {c_}: {d_}")
if n_fail:
    _p("\n   FAILURES:")
    for s_, c_, d_ in CHECKS:
        if s_ == "FAIL":
            _p(f"     - {c_}: {d_}")
_p("\n   RESULT: " + ("ALL CHECKS PASSED" if n_fail == 0 else "FAILURES PRESENT - do not use the output until fixed"))
# refresh the in-workbook validation sheet so it includes the final checks
SHEETS["Validation_Summary"] = pd.DataFrame(CHECKS, columns=["Status", "Check", "Detail"])
save_with_fallback(FINAL_XLSX, lambda p: write_workbook(p, SHEETS))
_p(f"\n   Done. Open {FINAL_XLSX.name} - start with Executive_Summary and Validation_Summary.")
sys.exit(1 if n_fail else 0)
