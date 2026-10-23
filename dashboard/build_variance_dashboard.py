"""
==============================================================================
 SALES VARIANCE DASHBOARD BUILDER
==============================================================================
 Bundles the dashboard source (dashboard/template/*.html, *.js) and the
 project data into ONE self-contained HTML file that opens offline in any
 browser - no server, no internet, no external libraries.

 Reads (all produced by the earlier pipeline steps)
   data/processed/Sales_Variance_Enriched_Data.csv   every sales line (required)
   outputs/Sales_Variance_Analysis.xlsx              used to reconcile the dashboard
   data/processed/Cleaned_Sales_Data.xlsx            cleaning summary and checks
   docs/Insights_and_Recommendations.md              shown on the Insights tab
   docs/key_metrics.json                             headline statements to re-check
 Writes
   dashboard/Sales_Variance_Dashboard.html

 Run from the repository root:
        python dashboard/build_variance_dashboard.py

 All numbers are computed in the browser from the line data in integer cents,
 with the same rules as fpa_lib/finance.py, and reconciled against the Excel
 workbook when the page opens (Data & Validation tab).
==============================================================================
"""
import argparse
import datetime as _dt
import html
import json
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fpa_lib import config  # noqa: E402
from fpa_lib import finance as fin  # noqa: E402

TEMPLATE_DIR = Path(__file__).resolve().parent / "template"
JS_ORDER = ["core.js", "charts.js", "ui.js", "views_overview.js", "views_drivers_trends.js",
            "views_breakdowns.js", "views_insights_data.js", "boot.js"]
TITLE = "Sales Variance Dashboard"
STATUS_CODE = {config.STATUS_FAVORABLE: 0, config.STATUS_WITHIN: 1, config.STATUS_UNFAVORABLE: 2, config.STATUS_NO_TARGET: 3}
TRIGGER_CODE = {"": 0, "Percent": 1, "Amount": 2, "Percent and amount": 3}
REQUIRED = config.CLEAN_COLUMNS + ["Sales_Variance_Status", "Material_Variance_Flag", "Material_Trigger", "Repaired_Fields", "DQ_Flag"]


class BuildError(Exception):
    pass


# ------------------------------------------------------------- markdown --
def _inline(t):
    t = html.escape(t, quote=False)
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    t = re.sub(r"(?<![\*\w])\*(?!\s)(.+?)(?<!\s)\*(?![\*\w])", r"<em>\1</em>", t)
    t = re.sub(r"`(.+?)`", r"<code>\1</code>", t)
    return t


def md_to_html(md):
    """Small Markdown subset -> HTML (headings, paragraphs, bullet and numbered
    lists, tables, bold, italic, code). Every piece of text is HTML-escaped."""
    out, i, lines = [], 0, md.replace("\r\n", "\n").split("\n")
    lst = None

    def close():
        nonlocal lst
        if lst:
            out.append(f"</{lst}>")
            lst = None
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s\-:|]+\|$", lines[i + 1].strip()):
            close()
            head = [c.strip() for c in ln.strip("|").split("|")]
            out.append("<div class='md-table'><table><thead><tr>" + "".join(f"<th>{_inline(c)}</th>" for c in head) + "</tr></thead><tbody>")
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                out.append("<tr>" + "".join(f"<td>{_inline(c)}</td>" for c in cells) + "</tr>")
                i += 1
            out.append("</tbody></table></div>")
            continue
        m = re.match(r"^(#{1,4})\s+(.*)$", ln)
        if m:
            close()
            lvl = len(m.group(1))
            out.append(f"<h{lvl + 1}>{_inline(m.group(2))}</h{lvl + 1}>")
        elif re.match(r"^\s*[-*]\s+", ln):
            if lst != "ul":
                close()
                out.append("<ul>")
                lst = "ul"
            out.append("<li>" + _inline(re.sub(r"^\s*[-*]\s+", "", ln)) + "</li>")
        elif re.match(r"^\s*\d+\.\s+", ln):
            if lst != "ol":
                close()
                out.append("<ol>")
                lst = "ol"
            out.append("<li>" + _inline(re.sub(r"^\s*\d+\.\s+", "", ln)) + "</li>")
        elif ln.strip() == "":
            close()
        elif ln.startswith(">"):
            close()
            out.append("<blockquote>" + _inline(ln.lstrip("> ")) + "</blockquote>")
        else:
            close()
            out.append("<p>" + _inline(ln) + "</p>")
        i += 1
    close()
    return "\n".join(out)


# ------------------------------------------------------------- payload --
def to_c(series):
    return fin.to_cents(series).astype("int64")


def build_payload(e):
    missing = [c for c in REQUIRED if c not in e.columns]
    if missing:
        raise BuildError(f"Enriched CSV is missing {missing}. Run scripts/sales_variance_analysis.py first.")
    regions = sorted(e["Region"].unique())
    cats = sorted(e["Category"].unique())
    prods = sorted(e["Product"].unique(), key=lambda p: (config.PRODUCTS[p][1], p))
    months = sorted(e["Year_Month"].unique())
    notes = [""] + sorted(x for x in e["Notes"].unique() if x)
    ri, ci = {r: i for i, r in enumerate(regions)}, {c: i for i, c in enumerate(cats)}
    pi, mi, ni = {p: i for i, p in enumerate(prods)}, {m: i for i, m in enumerate(months)}, {n: i for i, n in enumerate(notes)}
    unknown_status = set(e["Sales_Variance_Status"]) - set(STATUS_CODE)
    if unknown_status:
        raise BuildError(f"Unknown status values {unknown_status}")
    rows = {
        "R": e["Region"].map(ri).tolist(), "C": e["Category"].map(ci).tolist(), "P": e["Product"].map(pi).tolist(),
        "M": e["Year_Month"].map(mi).tolist(),
        "TQ": e["Target_Qty"].astype("int64").tolist(), "AQ": e["Actual_Qty"].astype("int64").tolist(),
        "TP": to_c(e["Target_Unit_Price_USD"]).tolist(), "AP": to_c(e["Actual_Unit_Price_USD"]).tolist(),
        "TCU": to_c(e["Target_Unit_Cost_USD"]).tolist(), "ACU": to_c(e["Actual_Unit_Cost_USD"]).tolist(),
        "ST": e["Sales_Variance_Status"].map(STATUS_CODE).tolist(),
        "MT": (e["Material_Variance_Flag"] == "Yes").astype(int).tolist(),
        "TRG": e["Material_Trigger"].map(TRIGGER_CODE).tolist(),
        "NT": e["Notes"].map(ni).tolist(),
        "RP": (e["Repaired_Fields"] != "").astype(int).tolist(),
        "DQ": (e["DQ_Flag"] != "").astype(int).tolist(),
    }
    dims = {"regions": regions, "cats": cats, "prods": prods, "months": months, "notes": notes,
            "prodCat": [ci[config.PRODUCTS[p][1]] for p in prods]}
    return rows, dims, e["Line_ID"].astype(str).tolist()


def check_lines(e, rows):
    """Re-derive totals and every line's status/materiality from cents (independent of the CSV labels)."""
    tq, aq = np.array(rows["TQ"], dtype="int64"), np.array(rows["AQ"], dtype="int64")
    tp, ap = np.array(rows["TP"], dtype="int64"), np.array(rows["AP"], dtype="int64")
    ts, as_ = tq * tp, aq * ap
    if not np.array_equal(ts, to_c(e["Target_Sales_USD"])) or not np.array_equal(as_, to_c(e["Actual_Sales_USD"])):
        raise BuildError("Sales do not equal quantity x unit price on every line")
    st = pd.Series(fin.status(as_ - ts, ts)).map(STATUS_CODE).to_numpy()
    _, trig = fin.material(as_ - ts, ts)
    trg = pd.Series(trig).map(TRIGGER_CODE).to_numpy()
    bad = int((st != np.array(rows["ST"])).sum() + (trg != np.array(rows["TRG"])).sum())
    if bad:
        raise BuildError(f"{bad} line status/materiality values differ from a re-derivation in cents")
    return int(ts.sum()), int(as_.sum())


def read_workbook(path):
    if not path.is_file():
        return None
    X = pd.ExcelFile(path)
    P = lambda s: X.parse(s, keep_default_na=False, na_values=[""])  # noqa: E731
    pick = lambda d, key: {str(r[key]): {"v": float(r["Sales_Variance_USD"]), "gv": float(r["Gross_Profit_Variance_USD"]),  # noqa: E731
                                         "prc": float(r["Price_Effect_USD"]), "mix": float(r["Mix_Effect_USD"])}
                           for _, r in d.iterrows()}
    month = {}
    for _, r in P("By_Year_Month").iterrows():
        month[str(r["Year_Month"])] = {"var": float(r["Sales_Variance_USD"]), "pct": _f(r["Sales_Variance_Pct"]),
                                       "gv": float(r["Gross_Profit_Variance_USD"]), "ts": _f(r["Tracking_Signal"]),
                                       "ytd": _f(r["YTD_Sales_Variance_USD"]), "r3": _f(r["Rolling_3M_Sales_Variance_USD"]),
                                       "mom": _f(r["MoM_Actual_Sales_Growth_Pct"]), "yoy": _f(r["YoY_Actual_Sales_Growth_Pct"])}
    year = {str(int(r["Year"])): {"v": float(r["Sales_Variance_USD"]), "gv": float(r["Gross_Profit_Variance_USD"])}
            for _, r in P("By_Year").iterrows()}
    comp = P("Company").iloc[0]
    company = {"pure": float(comp["Volume_Pure_USD"]), "mix": float(comp["Mix_Effect_USD"]), "prc": float(comp["Price_Effect_USD"]),
               "cost": float(comp["Profit_Cost_Effect_USD"]), "ppure": float(comp["Profit_Volume_Pure_USD"]),
               "pmix": float(comp["Profit_Mix_Effect_USD"])}
    pers_df = P("Persistence_Region_Category")
    pers = {f"{r['Region']}|{r['Category']}": int(r["Months_Below_Target"]) for _, r in pers_df.iterrows()}
    bands_df = P("By_Price_Band").sort_values("Price_Band")
    bands = [{"n": int(r["Lines"]), "gv": float(r["Gross_Profit_Variance_USD"])} for _, r in bands_df.iterrows()]
    return {
        "region": pick(P("By_Region"), "Region"), "category": pick(P("By_Category"), "Category"),
        "product": pick(P("By_Product"), "Product"), "month": month, "year": year, "company": company,
        "pers": pers, "persistent": int((pers_df["Persistent_Shortfall"] == "Yes").sum()),
        "exc_cells": int(len(P("Exceptions_RegCat_Month"))), "mat_lines": int(len(P("Material_Lines"))),
        "bands": bands,
        "validation": P("Validation_Summary").fillna("").astype(str).values.tolist(),
        "dictionary": P("Metric_Dictionary").fillna("").astype(str).values.tolist(),
    }


def _f(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    return None if np.isnan(v) else v


def read_cleaning(path):
    if not path.is_file():
        return None
    X = pd.ExcelFile(path)
    rec = X.parse("Row_Reconciliation")
    vals = dict(zip(rec["Step"], rec["Rows"], strict=True))
    checks = X.parse("Validation_Checks", keep_default_na=False).astype(str).values.tolist()
    audit = X.parse("Corrections_Audit", usecols=[0])
    return {"raw_rows": int(vals["Raw data rows"]), "blank": -int(vals["- entirely blank rows"]),
            "exact": -int(vals["- exact duplicate rows"]), "merged": -int(vals["- duplicate copies merged"]),
            "review": -int(vals["- rows held for review"]), "clean": int(vals["= clean rows"]),
            "repaired": int(len(audit)), "checks": [[r[1], r[0], r[2]] for r in checks]}


def safe_json(obj):
    """JSON that cannot break out of a <script> element: <, > and & are escaped."""
    return (json.dumps(obj, separators=(",", ":"), ensure_ascii=True)
            .replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026"))


def bundle(payload):
    page = (TEMPLATE_DIR / "page.html").read_text(encoding="utf-8")
    parts = []
    for name in JS_ORDER:
        src = (TEMPLATE_DIR / name).read_text(encoding="utf-8")
        if "</script" in src.lower():
            raise BuildError(f"{name} contains '</script' which would end the script early")
        parts.append(f"/* ---- {name} ---- */\n{src}")
    page = page.replace("__SCRIPT__", "\n".join(parts))
    page = page.replace("__TITLE__", html.escape(TITLE))
    return page.replace("__DATA_JSON__", safe_json(payload))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build the offline HTML dashboard.")
    ap.add_argument("--enriched", type=Path, default=config.ENRICHED_CSV)
    ap.add_argument("--workbook", type=Path, default=config.ANALYSIS_XLSX)
    ap.add_argument("--cleaning-workbook", type=Path, default=config.CLEAN_XLSX)
    ap.add_argument("--insights", type=Path, default=config.INSIGHTS_MD)
    ap.add_argument("--metrics", type=Path, default=config.KEY_METRICS_JSON)
    ap.add_argument("--output", type=Path, default=config.DASHBOARD_HTML)
    a = ap.parse_args(argv)
    print("=" * 78)
    print(" SALES VARIANCE DASHBOARD BUILDER")
    print("=" * 78)
    try:
        if not a.enriched.is_file():
            raise BuildError(f"{a.enriched} not found. Run scripts/sales_variance_analysis.py first.")
        e = pd.read_csv(a.enriched, dtype={"Line_ID": str, "Notes": str, "Material_Trigger": str, "Repaired_Fields": str,
                                           "DQ_Flag": str}, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
        for c in ("Notes", "Material_Trigger", "Repaired_Fields", "DQ_Flag"):
            if c in e.columns:
                e[c] = e[c].fillna("")
        rows, dims, ids = build_payload(e)
        ts, as_ = check_lines(e, rows)
        print(f"   [PASS] {len(e):,} lines loaded; sales = qty x price and every status/materiality re-derived from cents")
        print(f"   [INFO] Target ${ts / 100:,.2f} | Actual ${as_ / 100:,.2f} | Variance ${(as_ - ts) / 100:,.2f}")
        xl = read_workbook(a.workbook)
        print(f"   [{'PASS' if xl else 'WARN'}] Analysis workbook {'read' if xl else 'not found - reconciliation disabled'}")
        cleaning = read_cleaning(a.cleaning_workbook)
        print(f"   [{'PASS' if cleaning else 'WARN'}] Cleaning workbook {'read' if cleaning else 'not found'}")
        md = md_to_html(a.insights.read_text(encoding="utf-8")) if a.insights.is_file() else None
        print(f"   [{'PASS' if md else 'WARN'}] Insights document {'read' if md else 'not found'}")
        metrics = json.loads(a.metrics.read_text(encoding="utf-8")) if a.metrics.is_file() else None
        print(f"   [{'PASS' if metrics else 'WARN'}] Key metrics {'read (' + str(len(metrics['claims'])) + ' headline statements)' if metrics else 'not found'}")
    except BuildError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    payload = {
        "title": TITLE, "note": "Synthetic data", "source": a.enriched.name,
        "built": _dt.datetime.now().strftime("%d %b %Y %H:%M"),
        "cfg": {"favPct": config.FAVORABLE_PCT, "unfPct": config.UNFAVORABLE_PCT, "favBp": int(round(config.FAVORABLE_PCT * 100)),
                "unfBp": int(round(config.UNFAVORABLE_PCT * 100)), "matPctV": config.MATERIAL_PCT,
                "matBp": int(round(config.MATERIAL_PCT * 100)), "matUsd": config.MATERIAL_USD,
                "matUsdC": int(round(config.MATERIAL_USD * 100)), "persist": config.PERSISTENCE_SHARE},
        "dims": dims, "rows": rows, "ids": ids, "xl": xl, "cleaning": cleaning, "md": md, "metrics": metrics,
    }
    try:
        page = bundle(payload)
    except BuildError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    a.output.parent.mkdir(parents=True, exist_ok=True)
    a.output.write_text(page, encoding="utf-8")
    ext = re.findall(r"""(?:src|href)\s*=\s*["']https?://""", page)
    print(f"   [{'PASS' if not ext else 'FAIL'}] No external scripts or stylesheets ({len(ext)} found)")
    print(f"   Wrote {a.output}  ({a.output.stat().st_size / 1024 / 1024:.1f} MB)")
    return 0 if not ext else 1


if __name__ == "__main__":
    sys.exit(main())
