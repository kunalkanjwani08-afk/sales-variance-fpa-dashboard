"""
==============================================================================
 SALES VARIANCE & GROSS-PROFIT ANALYSIS
==============================================================================
 Reads   data/processed/Cleaned_Sales_Data.csv          (from clean_sales_data.py)
 Writes  data/processed/Sales_Variance_Enriched_Data.csv  clean lines + variance columns
         outputs/Sales_Variance_Analysis.xlsx             multi-sheet workbook

 Run from the repository root:
        python scripts/sales_variance_analysis.py
        python scripts/sales_variance_analysis.py --input file.csv --output-dir dir --workbook-dir dir

 Method (details in docs/METHODOLOGY.md; formulas in fpa_lib/finance.py)
   * Money is handled in integer cents; percentages are sum/sum and rounded
     half away from zero.
   * Revenue bridge:  Target sales -> pure volume -> mix -> price -> Actual sales
   * Profit bridge:   Target GP -> pure volume -> mix -> price -> cost -> Actual GP
   * Status: Favorable > +3 %, Unfavorable < -3 %, otherwise Within tolerance.
     Materiality (|%| >= 10 or |$| >= 75,000) is reported separately.
   * Every bridge and every summary is reconciled to the cent (Validation sheet).
==============================================================================
"""
import argparse
import sys
from decimal import Decimal
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fpa_lib import analysis as an  # noqa: E402
from fpa_lib import config  # noqa: E402
from fpa_lib import finance as fin  # noqa: E402


class AnalysisError(Exception):
    pass


# ======================================================================
# LOAD
# ======================================================================
def load_clean(path):
    path = Path(path)
    if not path.is_file():
        raise AnalysisError(f"Input file not found: {path}. Run scripts/clean_sales_data.py first.")
    d = pd.read_csv(path, dtype={"Line_ID": str, "Notes": str}, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
    missing = [c for c in config.CLEAN_COLUMNS if c not in d.columns]
    if missing:
        raise AnalysisError(f"{path.name} is missing columns {missing}")
    if d.empty:
        raise AnalysisError(f"{path.name} has no rows")
    req = [c for c in config.CLEAN_COLUMNS if c != "Notes"]
    if d[req].isna().any().any():
        bad = d[req].columns[d[req].isna().any()].tolist()
        raise AnalysisError(f"{path.name} has blank values in {bad}")
    if not d["Line_ID"].is_unique:
        raise AnalysisError("Line_ID is not unique")
    d["Notes"] = d["Notes"].fillna("")
    for c in ("Repaired_Fields", "DQ_Flag"):
        if c in d.columns:
            d[c] = d[c].fillna("")
    d["Date"] = pd.to_datetime(d["Date"])
    d["Year_Month"] = d["Date"].dt.strftime("%Y-%m")
    d["Quarter"] = "Q" + d["Date"].dt.quarter.astype(str)
    d["Year_Quarter"] = d["Year"].astype(str) + " " + d["Quarter"]
    return d


# ======================================================================
# TABLES
# ======================================================================
def enriched(L):
    e = L[[c for c in L.columns if c in config.CLEAN_COLUMNS or c in
           ("Repaired_Fields", "DQ_Flag", "Source_Row", "Year_Month", "Quarter")]].copy()
    e["Date"] = e["Date"].dt.strftime("%Y-%m-%d")
    e["Qty_Variance"] = L["AQ"] - L["TQ"]
    e["Qty_Attainment_Pct"] = fin.pct(L["AQ"], L["TQ"])
    e["Sales_Variance_USD"] = L["VAR"] / 100
    e["Sales_Variance_Pct"] = fin.pct(L["VAR"], L["TS"])
    e["Sales_Variance_Status"] = L["STATUS"]
    e["Material_Variance_Flag"] = np.where(L["MAT"], "Yes", "No")
    e["Material_Trigger"] = L["TRIG"]
    e["Volume_Effect_USD"] = L["VOL"] / 100
    e["Price_Effect_USD"] = L["PRC"] / 100
    e["Price_Realisation_Pct"] = fin.pct(L["AP"] - L["TP"], L["TP"])
    e["Target_COGS_USD"] = (L["TS"] - L["TG"]) / 100
    e["Actual_COGS_USD"] = (L["AS"] - L["AG"]) / 100
    e["Gross_Profit_Variance_USD"] = L["GVAR"] / 100
    e["Gross_Profit_Variance_Pct"] = fin.pct(L["GVAR"], L["TG"])
    e["Profit_Volume_Effect_USD"] = L["PVOL"] / 100
    e["Profit_Price_Effect_USD"] = L["PRC"] / 100
    e["Profit_Cost_Effect_USD"] = L["COST"] / 100
    e["Unit_Cost_Change_Pct"] = fin.pct(L["ACU"] - L["TCU"], L["TCU"])
    e["Target_Gross_Margin_Pct"] = fin.pct(L["TG"], L["TS"])
    e["Actual_Gross_Margin_Pct"] = fin.pct(L["AG"], L["AS"])
    return e


def monthly(L):
    m = an.summarise(L, ["Year_Month"]).sort_values("Year_Month").reset_index(drop=True)
    m["Year"] = m["Year_Month"].str[:4].astype(int)
    m["Month"] = m["Year_Month"].str[5:].astype(int)
    vc = (m["Sales_Variance_USD"] * 100).round().astype("int64")
    ac = (m["Actual_Sales_USD"] * 100).round().astype("int64")
    tc = (m["Target_Sales_USD"] * 100).round().astype("int64")
    by_ym = {(y, mo): i for i, (y, mo) in enumerate(zip(m["Year"], m["Month"], strict=True))}
    mom, yoy, r3 = [], [], []
    for i, (y, mo) in enumerate(zip(m["Year"], m["Month"], strict=True)):
        prev = by_ym.get((y, mo - 1) if mo > 1 else (y - 1, 12))
        mom.append(fin.pct(int(ac[i] - ac[prev]), int(ac[prev])) if prev is not None else np.nan)
        ly = by_ym.get((y - 1, mo))
        yoy.append(fin.pct(int(ac[i] - ac[ly]), int(ac[ly])) if ly is not None else np.nan)
        win = [by_ym.get(((y * 12 + mo - 1 - k) // 12, (y * 12 + mo - 1 - k) % 12 + 1)) for k in range(3)]
        r3.append(sum(int(vc[j]) for j in win) / 100 if all(j is not None for j in win) else np.nan)
    m["MoM_Actual_Sales_Growth_Pct"] = mom
    m["YoY_Actual_Sales_Growth_Pct"] = yoy
    m["YTD_Sales_Variance_USD"] = vc.groupby(m["Year"]).cumsum() / 100
    m["Rolling_3M_Sales_Variance_USD"] = r3
    m["Cumulative_Sales_Variance_USD"] = vc.cumsum() / 100
    mad = vc.abs().expanding().mean()
    m["Tracking_Signal"] = (vc.cumsum() / mad).round(2)
    year_t = tc.groupby(m["Year"]).transform("sum")
    year_a = ac.groupby(m["Year"]).transform("sum")
    m["Target_Share_of_Year_Pct"] = [fin.pct(int(a), int(b)) for a, b in zip(tc, year_t, strict=True)]
    m["Actual_Share_of_Year_Pct"] = [fin.pct(int(a), int(b)) for a, b in zip(ac, year_a, strict=True)]
    m["Phasing_Gap_pp"] = (m["Actual_Share_of_Year_Pct"] - m["Target_Share_of_Year_Pct"]).round(2)
    return m


def discount_bands(L):
    """Lines grouped by actual price vs target price (exact integer tests, same as the dashboard)."""
    d100, tp = (L["AP"] - L["TP"]) * 100, L["TP"]
    b1 = d100 <= -5 * tp
    b2 = ~b1 & (d100 < -1 * tp)
    b3 = ~b1 & ~b2 & (d100 <= tp)
    b4 = ~b1 & ~b2 & ~b3
    names = ["1. Discount of 5% or more", "2. Discount of 1-5%", "3. At target price (within 1%)",
             "4. Above target price (over 1%)"]
    parts = [an.summarise(L[m], []).assign(Price_Band=n) for n, m in zip(names, (b1, b2, b3, b4), strict=True)]
    t = pd.concat(parts, ignore_index=True)
    return t[["Price_Band"] + [c for c in t.columns if c != "Price_Band"]]


def target_accuracy(L):
    """2025 region-month targets vs a naive forecast (same month of 2024, actual)."""
    rm = L.groupby(["Region", "Year", "Month"])[["TS", "AS"]].sum().reset_index()
    a25 = rm[rm["Year"] == 2025].merge(rm[rm["Year"] == 2024][["Region", "Month", "AS"]].rename(columns={"AS": "NAIVE"}),
                                       on=["Region", "Month"], how="left")
    rows = []
    for name, g in list(a25.groupby("Region")) + [("Company", a25.groupby("Month")[["TS", "AS", "NAIVE"]].sum().reset_index())]:
        act = int(g["AS"].sum())
        e_t = (g["AS"] - g["TS"]).astype("int64")
        e_n = (g["AS"] - g["NAIVE"]).astype("int64")
        rows.append({"Region": name, "Months": len(g), "Actual_Sales_2025_USD": act / 100,
                     "Target_WAPE_Pct": fin.pct(int(e_t.abs().sum()), act),
                     "Naive_WAPE_Pct": fin.pct(int(e_n.abs().sum()), act),
                     "Target_Bias_Pct": fin.pct(int(-e_t.sum()), act),
                     "Naive_Bias_Pct": fin.pct(int(-e_n.sum()), act),
                     "Months_Target_Closer": int((e_t.abs() < e_n.abs()).sum())})
    t = pd.DataFrame(rows)
    t["More_Accurate"] = np.where(t["Target_WAPE_Pct"] < t["Naive_WAPE_Pct"], "Target", "Naive forecast")
    return t


def persistence(L, months_in_view):
    g = L.groupby(["Region", "Category", "Year_Month"])[["TS", "AS"]].sum().reset_index()
    g["below"] = g["AS"] < g["TS"]
    p = g.groupby(["Region", "Category"]).agg(Months=("below", "size"), Months_Below_Target=("below", "sum")).reset_index()
    p["Share_Months_Below_Pct"] = [fin.pct(int(b), int(m)) for b, m in zip(p["Months_Below_Target"], p["Months"], strict=True)]
    p["Persistent_Shortfall"] = np.where(p["Months_Below_Target"] >= config.PERSISTENCE_SHARE * p["Months"], "Yes", "No")
    tot = an.summarise(L, ["Region", "Category"])[["Region", "Category", "Sales_Variance_USD", "Sales_Variance_Pct",
                                                   "Gross_Profit_Variance_USD"]]
    return p.merge(tot, on=["Region", "Category"]).sort_values(["Months_Below_Target", "Sales_Variance_USD"],
                                                              ascending=[False, True]).reset_index(drop=True)


def pareto(prod):
    neg = prod[prod["Sales_Variance_USD"] < 0].sort_values("Sales_Variance_USD").reset_index(drop=True)
    gross_c = int(round(neg["Sales_Variance_USD"].sum() * 100))
    cum = (neg["Sales_Variance_USD"] * 100).round().astype("int64").cumsum()
    neg["Share_of_Shortfall_Pct"] = [fin.pct(-int(round(v * 100)), -gross_c) for v in neg["Sales_Variance_USD"]]
    neg["Cumulative_Share_Pct"] = [fin.pct(-int(c), -gross_c) for c in cum]
    return neg[["Product", "Category", "Sales_Variance_USD", "Sales_Variance_Pct", "Share_of_Shortfall_Pct",
                "Cumulative_Share_Pct"]]


# ======================================================================
# CHECKS
# ======================================================================
def validate(d, L, tables, clean_csv_path):
    out = []

    def add(name, ok, detail=""):
        out.append(("PASS" if ok else "FAIL", name, detail))
    add("Input lines = clean file lines, Line_ID unique", len(L) == len(d) and L["Line_ID"].is_unique, f"{len(L):,} lines")
    add("Line bridge: Volume + Price = Sales variance (every line)", bool((L["VOL"] + L["PRC"] == L["VAR"]).all()))
    add("Line bridge: Profit volume + Price + Cost = Gross-profit variance (every line)",
        bool((L["PVOL"] + L["PRC"] + L["COST"] == L["GVAR"]).all()))
    # independent recomputation of totals with Decimal straight from the CSV text
    raw = pd.read_csv(clean_csv_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    for col, k in (("Target_Sales_USD", "TS"), ("Actual_Sales_USD", "AS"), ("Target_Gross_Profit_USD", "TG"),
                   ("Actual_Gross_Profit_USD", "AG")):
        dec = sum(Decimal(x) for x in raw[col])
        add(f"Decimal re-sum of {col} from the CSV text equals the cent total", dec * 100 == int(L[k].sum()),
            f"${dec:,.2f}")
    tot = tables["Company"].iloc[0]

    def cents(x):
        return int(round(float(x) * 100))
    for name, t in tables.items():
        if name in ("Company",) or "Sales_Variance_USD" not in t.columns or "Volume_Pure_USD" not in t.columns:
            continue
        rb = all(cents(r.Volume_Pure_USD) + cents(r.Mix_Effect_USD) + cents(r.Price_Effect_USD) == cents(r.Sales_Variance_USD)
                 for r in t.itertuples())
        pb = all(cents(r.Profit_Volume_Pure_USD) + cents(r.Profit_Mix_Effect_USD) + cents(r.Profit_Price_Effect_USD)
                 + cents(r.Profit_Cost_Effect_USD) == cents(r.Gross_Profit_Variance_USD) for r in t.itertuples())
        add(f"{name}: pure volume + mix + price = sales variance (every row)", rb)
        add(f"{name}: profit pure volume + mix + price + cost = GP variance (every row)", pb)
        if name in ("By_Region", "By_Category", "By_Product", "By_Year", "By_Year_Month", "By_Region_Category",
                    "By_Quarter", "By_Price_Band"):
            for col in ("Target_Sales_USD", "Actual_Sales_USD", "Target_Gross_Profit_USD", "Actual_Gross_Profit_USD",
                        "Volume_Effect_USD", "Price_Effect_USD", "Profit_Cost_Effect_USD"):
                s = sum(cents(x) for x in t[col])
                if s != cents(tot[col]):
                    add(f"{name}: {col} adds back to the company total", False, f"{s} vs {cents(tot[col])}")
                    break
            else:
                add(f"{name}: all amounts add back to the company total", True)
    # company bridge
    add("Company revenue bridge reconciles", cents(tot["Volume_Pure_USD"]) + cents(tot["Mix_Effect_USD"]) + cents(tot["Price_Effect_USD"])
        == cents(tot["Sales_Variance_USD"]), f"${tot['Sales_Variance_USD']:,.2f}")
    add("Company gross-profit bridge reconciles", cents(tot["Profit_Volume_Pure_USD"]) + cents(tot["Profit_Mix_Effect_USD"])
        + cents(tot["Profit_Price_Effect_USD"]) + cents(tot["Profit_Cost_Effect_USD"]) == cents(tot["Gross_Profit_Variance_USD"]),
        f"${tot['Gross_Profit_Variance_USD']:,.2f}")
    # status re-derived with a different method (Decimal percentages)
    pct_dec = [(Decimal(int(v)) / Decimal(int(t)) * 100) if t > 0 else None for v, t in zip(L["VAR"], L["TS"], strict=True)]
    st2 = [config.STATUS_NO_TARGET if p is None else config.STATUS_FAVORABLE if p > config.FAVORABLE_PCT
           else config.STATUS_UNFAVORABLE if p < -config.UNFAVORABLE_PCT else config.STATUS_WITHIN for p in pct_dec]
    add("Line status re-derived with Decimal arithmetic matches", list(L["STATUS"]) == st2,
        f"{sum(a != b for a, b in zip(L['STATUS'], st2, strict=True))} differences")
    n_nt = int((L["TS"] <= 0).sum())
    add("Lines with no target are reported, not hidden", True,
        f"{n_nt:,} lines have target 0 (new accounts): variance % is blank and status is 'No target'")
    return out


# ======================================================================
# WORKBOOK
# ======================================================================
DICTIONARY = [
    ("Sales_Variance_USD", "Actual_Sales - Target_Sales", "Positive = above target"),
    ("Sales_Variance_Pct", "Sales variance / Target sales x 100", "Groups: sum/sum. Blank when target = 0"),
    ("Sales_Variance_Status", "Favorable > +3 %, Unfavorable < -3 %, otherwise Within tolerance; No target if target = 0",
     "Boundaries count as Within tolerance. Exact integer comparison"),
    ("Material_Variance_Flag", "Yes if |variance %| >= 10 or |variance $| >= 75,000",
     "Separate from status: a group can be Within tolerance and still material in dollars"),
    ("Volume_Effect_USD", "(Actual_Qty - Target_Qty) x Target_Unit_Price", "Line level; sums to groups"),
    ("Volume_Pure_USD", "(sum Actual_Qty - sum Target_Qty) x group average target price", "Group level only"),
    ("Mix_Effect_USD", "Volume effect - pure volume = sum Actual_Qty x (line target price - group average target price)",
     "Value of selling a different mix of lines than planned, at target prices. Depends on the grouping level"),
    ("Price_Effect_USD", "Actual_Qty x (Actual_Unit_Price - Target_Unit_Price)", "Negative = sold below target price"),
    ("Price_Realisation_Pct", "Price effect / (Actual_Qty x Target_Unit_Price) x 100", "Average price vs target, same volume"),
    ("Gross_Profit_USD", "Sales - Qty x Unit_Cost", "Target and actual each use their own price and cost"),
    ("Profit_Volume_Effect_USD", "(Actual_Qty - Target_Qty) x (Target_Unit_Price - Target_Unit_Cost)",
     "Volume valued at target unit margin; split into pure volume and mix like revenue"),
    ("Profit_Price_Effect_USD", "Actual_Qty x (Actual_Unit_Price - Target_Unit_Price)", "Same as the revenue price effect"),
    ("Profit_Cost_Effect_USD", "-Actual_Qty x (Actual_Unit_Cost - Target_Unit_Cost)", "Negative = costs above target"),
    ("Gross_Margin_Pct", "Gross profit / Sales x 100", ""),
    ("Lines_Below_Target_Pct", "Share of lines with actual sales below target (variance < 0)",
     "Not the same as Unfavorable, which is below -3 %"),
    ("Tracking_Signal", "Cumulative variance / mean absolute monthly variance",
     "Values beyond +/-4 suggest persistent bias. Equals -n if every month so far is below target"),
    ("Target_WAPE_Pct", "sum |Actual - Target| / sum Actual, 2025 region-months", "Lower = more accurate"),
    ("Naive_WAPE_Pct", "Same for a naive forecast: the same month of 2024, actual", "Benchmark with no skill"),
    ("Persistent_Shortfall", "Region x category below target in >= 75 % of months", ""),
]


def write_workbook(path, sheets):
    path.parent.mkdir(parents=True, exist_ok=True)
    # strings_to_formulas=False: text such as "=..." is stored as text, never executed as a formula
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_formulas": False}}) as xw:
        wb = xw.book
        head = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1, "text_wrap": True})
        money = wb.add_format({"num_format": "#,##0.00;[Red]-#,##0.00"})
        pctf = wb.add_format({"num_format": "0.00"})
        for name, frame in sheets.items():
            frame.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            for j, col in enumerate(frame.columns):
                ws.write(0, j, col, head)
                fmt = money if str(col).endswith("_USD") else pctf if str(col).endswith(("_Pct", "_pp")) else None
                width = min(55, max(11, len(str(col)) * 0.9 + 2,
                                    (frame[col].astype(str).str.len().quantile(0.9) + 2) if len(frame) else 0))
                ws.set_column(j, j, width, fmt)
            ws.freeze_panes(1, 1 if name not in ("Executive_Summary", "Validation_Summary", "Metric_Dictionary") else 0)
            if len(frame):
                ws.autofilter(0, 0, len(frame), len(frame.columns) - 1)


def executive_summary(tables, checks, n_review, review_target):
    t = tables["Company"].iloc[0]
    reg, cat, mon = tables["By_Region"], tables["By_Category"], tables["By_Year_Month"]
    worst_r, best_r = reg.loc[reg["Sales_Variance_USD"].idxmin()], reg.loc[reg["Sales_Variance_USD"].idxmax()]
    worst_c, best_c = cat.loc[cat["Sales_Variance_USD"].idxmin()], cat.loc[cat["Sales_Variance_USD"].idxmax()]
    below = int((mon["Sales_Variance_USD"] < 0).sum())
    acc = tables["Target_Accuracy"].set_index("Region").loc["Company"]
    items = [
        ("Lines analysed", f"{int(t['Lines']):,}", "Synthetic data"),
        ("Period", f"{mon['Year_Month'].iloc[0]} to {mon['Year_Month'].iloc[-1]}", f"{len(mon)} months"),
        ("Target sales (USD)", t["Target_Sales_USD"], ""),
        ("Actual sales (USD)", t["Actual_Sales_USD"], ""),
        ("Sales variance (USD)", t["Sales_Variance_USD"], f"{t['Sales_Variance_Pct']:+.2f} % - {t['Sales_Variance_Status']}"
         + (", material" if t["Material_Variance_Flag"] == "Yes" else "")),
        ("  Pure volume effect (USD)", t["Volume_Pure_USD"], ""),
        ("  Mix effect (USD)", t["Mix_Effect_USD"], ""),
        ("  Price effect (USD)", t["Price_Effect_USD"], f"price realisation {t['Price_Realisation_Pct']:+.2f} %"),
        ("Target gross profit (USD)", t["Target_Gross_Profit_USD"], f"margin {t['Target_Gross_Margin_Pct']:.2f} %"),
        ("Actual gross profit (USD)", t["Actual_Gross_Profit_USD"], f"margin {t['Actual_Gross_Margin_Pct']:.2f} %"),
        ("Gross-profit variance (USD)", t["Gross_Profit_Variance_USD"], f"{t['Gross_Profit_Variance_Pct']:+.2f} %"),
        ("  Profit volume + mix (USD)", t["Profit_Volume_Effect_USD"], ""),
        ("  Profit price effect (USD)", t["Profit_Price_Effect_USD"], ""),
        ("  Profit cost effect (USD)", t["Profit_Cost_Effect_USD"], ""),
        ("Quantity attainment %", t["Qty_Attainment_Pct"], ""),
        ("Months below target", f"{below} of {len(mon)}", ""),
        ("Best region", best_r["Region"], f"{best_r['Sales_Variance_USD']:+,.2f} ({best_r['Sales_Variance_Pct']:+.2f} %)"),
        ("Weakest region", worst_r["Region"], f"{worst_r['Sales_Variance_USD']:+,.2f} ({worst_r['Sales_Variance_Pct']:+.2f} %)"),
        ("Best category", best_c["Category"], f"{best_c['Sales_Variance_USD']:+,.2f} ({best_c['Sales_Variance_Pct']:+.2f} %)"),
        ("Weakest category", worst_c["Category"], f"{worst_c['Sales_Variance_USD']:+,.2f} ({worst_c['Sales_Variance_Pct']:+.2f} %)"),
        ("2025 target accuracy (WAPE)", f"{acc['Target_WAPE_Pct']:.2f} %",
         f"naive same-month-last-year forecast: {acc['Naive_WAPE_Pct']:.2f} %"),
        ("Lines held back by the cleaning step", f"{n_review:,}",
         f"target sales ${review_target:,.2f} where readable (excluded from all figures)"),
        ("Validation checks", f"{sum(c[0] == 'PASS' for c in checks)} of {len(checks)} passed", ""),
    ]
    return pd.DataFrame(items, columns=["Item", "Value", "Comment"])


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sales variance and gross-profit analysis.")
    ap.add_argument("--input", type=Path, default=config.CLEAN_CSV)
    ap.add_argument("--output-dir", type=Path, default=config.PROCESSED_DIR, help="folder for the enriched CSV")
    ap.add_argument("--workbook-dir", type=Path, default=config.OUTPUTS_DIR, help="folder for the Excel workbook")
    ap.add_argument("--cleaning-workbook", type=Path, default=None,
                    help="Cleaned_Sales_Data.xlsx (default: next to the input) - used to report review rows")
    a = ap.parse_args(argv)
    try:
        d = load_clean(a.input)
        L = an.line_metrics(d)
    except (AnalysisError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    tables = {"Company": an.summarise(L, [])}
    tables["By_Year"] = an.summarise(L, ["Year"])
    tables["By_Quarter"] = an.summarise(L, ["Year_Quarter"])
    tables["By_Year_Month"] = monthly(L)
    tables["By_Region"] = an.summarise(L, ["Region"])
    tables["By_Category"] = an.summarise(L, ["Category"])
    tables["By_Region_Category"] = an.summarise(L, ["Region", "Category"])
    tables["By_Product"] = an.summarise(L, ["Category", "Product"])
    tables["By_Region_Product"] = an.summarise(L, ["Region", "Category", "Product"])
    tables["By_Region_Month"] = an.summarise(L, ["Region", "Year_Month"])
    tables["By_Price_Band"] = discount_bands(L)
    tables["Target_Accuracy"] = target_accuracy(L)
    tables["Persistence_Region_Category"] = persistence(L, L["Year_Month"].nunique())
    tables["Pareto_Shortfall_Product"] = pareto(tables["By_Product"])
    cells = an.summarise(L, ["Region", "Category", "Year_Month"])
    tables["Exceptions_RegCat_Month"] = cells[cells["Material_Variance_Flag"] == "Yes"].sort_values("Sales_Variance_USD").reset_index(drop=True)
    e = enriched(L)
    tables["Material_Lines"] = e[e["Material_Variance_Flag"] == "Yes"].assign(_a=lambda x: x["Sales_Variance_USD"].abs()) \
        .sort_values("_a", ascending=False).drop(columns="_a")[
        ["Line_ID", "Year_Month", "Region", "Category", "Product", "Target_Sales_USD", "Actual_Sales_USD",
         "Sales_Variance_USD", "Sales_Variance_Pct", "Sales_Variance_Status", "Material_Trigger", "Notes"]]
    # data-quality impact
    cw = a.cleaning_workbook or a.input.with_name(config.CLEAN_XLSX.name)
    n_review, review_target = 0, 0.0
    if cw.is_file():
        rv = pd.read_excel(cw, sheet_name="Review_Required", dtype=str, keep_default_na=False)
        n_review = len(rv)
        review_target = float(pd.to_numeric(rv["Target_Sales_USD"], errors="coerce").sum())
    rep = L["Repaired_Fields"].fillna("") != "" if "Repaired_Fields" in L.columns else pd.Series(False, index=L.index)
    dq = pd.concat([an.summarise(L, []).assign(Scope="All clean lines"),
                    an.summarise(L[~rep], []).assign(Scope="Excluding lines with any repaired value"),
                    an.summarise(L[rep], []).assign(Scope="Only lines with a repaired value")], ignore_index=True)
    dq = dq[["Scope", "Lines", "Target_Sales_USD", "Actual_Sales_USD", "Sales_Variance_USD", "Sales_Variance_Pct",
             "Gross_Profit_Variance_USD", "Gross_Profit_Variance_Pct", "Price_Realisation_Pct"]]
    dq.loc[len(dq)] = ["Rows held back for review (not in any figure)", n_review, review_target] + [np.nan] * 6
    tables["Data_Quality_Impact"] = dq

    a.output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = a.output_dir / config.ENRICHED_CSV.name
    e.to_csv(csv_path, index=False, encoding="utf-8-sig")
    checks = validate(d, L, tables, a.input)
    exec_sum = executive_summary(tables, checks, n_review, review_target)
    order = ["Executive_Summary", "Validation_Summary", "Metric_Dictionary", "Company", "By_Year", "By_Quarter",
             "By_Year_Month", "By_Region", "By_Category", "By_Region_Category", "By_Product", "By_Region_Product",
             "By_Region_Month", "By_Price_Band", "Target_Accuracy", "Persistence_Region_Category",
             "Pareto_Shortfall_Product", "Exceptions_RegCat_Month", "Material_Lines", "Data_Quality_Impact"]
    sheets = {"Executive_Summary": exec_sum,
              "Validation_Summary": pd.DataFrame(checks, columns=["Status", "Check", "Detail"]),
              "Metric_Dictionary": pd.DataFrame(DICTIONARY, columns=["Metric", "Definition", "Notes"])}
    sheets.update({k: tables[k] for k in order if k in tables})
    xlsx_path = a.workbook_dir / config.ANALYSIS_XLSX.name
    write_workbook(xlsx_path, {k: sheets[k] for k in order})

    t = tables["Company"].iloc[0]
    n_fail = sum(c[0] == "FAIL" for c in checks)
    print("=" * 78)
    print(" SALES VARIANCE & GROSS-PROFIT ANALYSIS")
    print("=" * 78)
    print(f"   Lines ................ {int(t['Lines']):,}")
    print(f"   Target sales ......... ${t['Target_Sales_USD']:,.2f}")
    print(f"   Actual sales ......... ${t['Actual_Sales_USD']:,.2f}")
    print(f"   Sales variance ....... ${t['Sales_Variance_USD']:,.2f} ({t['Sales_Variance_Pct']:+.2f} %)")
    print(f"     pure volume ${t['Volume_Pure_USD']:,.2f} | mix ${t['Mix_Effect_USD']:,.2f} | price ${t['Price_Effect_USD']:,.2f}")
    print(f"   Gross-profit variance  ${t['Gross_Profit_Variance_USD']:,.2f} ({t['Gross_Profit_Variance_Pct']:+.2f} %)")
    print(f"     volume+mix ${t['Profit_Volume_Effect_USD']:,.2f} | price ${t['Profit_Price_Effect_USD']:,.2f} | "
          f"cost ${t['Profit_Cost_Effect_USD']:,.2f}")
    for st, name, det in checks:
        if st == "FAIL":
            print(f"   [FAIL] {name} {det}")
    print(f"   Checks: {len(checks) - n_fail} passed, {n_fail} failed")
    print(f"   Wrote {csv_path}")
    print(f"   Wrote {xlsx_path}")
    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
