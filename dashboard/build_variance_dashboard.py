# -*- coding: utf-8 -*-
"""
==============================================================================
 SALES VARIANCE DASHBOARD BUILDER   (VS Code / terminal - no input needed)
==============================================================================
 Reads the three files produced by the variance project and writes ONE
 self-contained, offline HTML dashboard (no internet, no libraries needed):

   Sales_Variance_Enriched_Data.csv      (required - every sales line)
   Sales_Variance_Analysis.xlsx          (optional - validation log, dictionary,
                                          reference tables used for reconciliation)
   Insights_and_Recommendations.md       (optional - shown on the Insights tab)

 Repo layout (this script sits in dashboard/): the files are found automatically in
 data/processed/, outputs/ and docs/, and the dashboard is written next to this script.
   Otherwise put this script in the same folder as the three files.

 Run:   python build_variance_dashboard.py
   or   python build_variance_dashboard.py "C:\\path\\to\\folder"

 Output: Sales_Variance_Dashboard.html   (double-click to open in any browser)

 Python needs: pandas + numpy (+ openpyxl only if you want the xlsx tabs).
 Every number in the dashboard is computed in the browser from the line-level
 data with exact integer-cent arithmetic - the same rules as the analysis script:
   Favorable   : variance % >  +3        Unfavorable : variance % < -3
   Material    : |variance %| >= 10  OR  |variance $| >= 75,000
 and is reconciled to the Excel workbook when the page loads (see "Data &
 Validation" tab).
==============================================================================
"""
import sys, re, json, html, datetime as _dt
from pathlib import Path
import numpy as np
import pandas as pd

# ------------------------------------------------------------------ CONFIG --
CSV_NAME = "Sales_Variance_Enriched_Data.csv"
XLSX_NAME = "Sales_Variance_Analysis.xlsx"
MD_NAME = "Insights_and_Recommendations.md"
OUT_NAME = "Sales_Variance_Dashboard.html"
DASHBOARD_TITLE = "Sales Variance Dashboard"
FAV_PCT, UNF_PCT = 3.0, -3.0          # must match the analysis script
MAT_PCT, MAT_USD = 10.0, 75000.0
PRIVACY_NOTE = "Synthetic portfolio dataset"

REQUIRED = ["Region", "Category", "Product", "Year_Month", "Target_Qty", "Actual_Qty", "Unit_Price_USD",
            "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD", "Gross_Margin_USD", "Notes",
            "Source_Excel_Row", "Volume_Effect_USD", "Price_Effect_USD", "PV_Residual_USD",
            "Sales_Variance_Status", "Material_Variance_Flag", "Material_Trigger"]
STATUSES = ["Favorable", "On Target", "Unfavorable", "No Target"]
TRIGGERS = {"": 0, "Pct": 1, "USD": 2, "Pct and USD": 3}

BUILD_LOG = []


def log(kind, msg):
    BUILD_LOG.append((kind, msg))
    print(f"   [{kind}] {msg}", flush=True)


def die(msg):
    sys.exit("ERROR: " + msg)


def find_dir():
    if len(sys.argv) > 1:
        p = Path(sys.argv[1].strip('"'))
        if p.is_file():
            return p.parent
        if p.is_dir():
            return p
        die(f"folder not found: {p}")
    here = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    repo = here.parent if here.name.lower() in ("dashboard", "scripts") else here
    for folder in (here, Path.cwd(), repo / "data" / "processed"):
        if (folder / CSV_NAME).is_file():
            return folder
    die(f"'{CSV_NAME}' not found next to this script ({here}) or in the current folder ({Path.cwd()}).\n"
        f"   Put the three project files in one folder with this script, or run:\n"
        f"   python build_variance_dashboard.py \"full\\path\\to\\folder\"")


def locate(name, folder):
    """Find an optional input: next to the CSV first, then the repo's outputs/ and docs/ folders."""
    here = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    repo = here.parent if here.name.lower() in ("dashboard", "scripts") else here
    for d in (folder, repo / "outputs", repo / "docs", here, Path.cwd()):
        if (d / name).is_file():
            return d / name
    return folder / name


def out_dir_for(folder):
    """Repo layout: write the dashboard next to this script (dashboard/). Otherwise next to the data."""
    here = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    if len(sys.argv) == 1 and here.name.lower() == "dashboard":
        return here
    return folder


def to_cents(s):
    return np.rint(np.asarray(s, dtype="float64") * 100.0).astype("int64")


def classify(var_c, tc):
    """Same exact integer rule as the analysis script."""
    var_c = np.asarray(var_c, dtype="int64"); tc = np.asarray(tc, dtype="int64")
    valid = tc > 0
    lhs = var_c * 100000
    fav = valid & (lhs > int(round(FAV_PCT * 1000)) * tc)
    unf = valid & (lhs < int(round(UNF_PCT * 1000)) * tc)
    m_pct = valid & (np.abs(var_c) * 100000 >= int(round(MAT_PCT * 1000)) * tc)
    m_usd = np.abs(var_c) >= int(round(MAT_USD * 100))
    status = np.select([~valid, fav, unf], [3, 0, 2], default=1)
    return status, (m_pct | m_usd)


# ------------------------------------------------------ tiny Markdown -> HTML --
def _inline(t):
    t = html.escape(t, quote=False)
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    t = re.sub(r"(?<![\*\w])\*(?!\s)(.+?)(?<!\s)\*(?![\*\w])", r"<em>\1</em>", t)
    t = re.sub(r"`(.+?)`", r"<code>\1</code>", t)
    return t


def md_to_html(md):
    out, i, lines = [], 0, md.replace("\r\n", "\n").split("\n")
    in_list = False
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s\-:|]+\|$", lines[i + 1].strip()):
            if in_list: out.append("</ul>"); in_list = False
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
            if in_list: out.append("</ul>"); in_list = False
            lvl = len(m.group(1)); out.append(f"<h{lvl + 1}>{_inline(m.group(2))}</h{lvl + 1}>")
        elif re.match(r"^\s*[-*]\s+", ln):
            if not in_list: out.append("<ul>"); in_list = True
            out.append("<li>" + _inline(re.sub(r"^\s*[-*]\s+", "", ln)) + "</li>")
        elif ln.strip() == "":
            if in_list: out.append("</ul>"); in_list = False
        else:
            if in_list: out.append("</ul>"); in_list = False
            out.append("<p>" + _inline(ln) + "</p>")
        i += 1
    if in_list: out.append("</ul>")
    return "\n".join(out)


TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
:root{
  color-scheme: light;
  --page:#f4f4f1; --surface:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e; --muted:#898781;
  --grid:#e1e0d9; --axis:#c3c2b7; --border:rgba(11,11,11,.10); --wash:rgba(11,11,11,.045);
  --s1:#2a78d6; --s2:#eb6834; --s3:#1baf7a; --s4:#eda100; --s5:#e87ba4; --s6:#008300; --s7:#6250d6; --s8:#e34948;
  --fav:#2a78d6; --unf:#e34948; --mid:#f0efec; --target:#898781;
  --good:#0ca30c; --warn:#fab219; --serious:#ec835a; --crit:#d03b3b; --goodtext:#006300; --warntext:#8a5b00;
  --shadow:0 1px 2px rgba(11,11,11,.06),0 4px 14px rgba(11,11,11,.05);
}
:root[data-theme="dark"]{
  color-scheme: dark;
  --page:#0d0d0d; --surface:#1a1a19; --ink:#f0efec; --ink2:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --axis:#383835; --border:rgba(255,255,255,.10); --wash:rgba(255,255,255,.06);
  --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500; --s5:#d55181; --s6:#008300; --s7:#9085e9; --s8:#e66767;
  --fav:#3987e5; --unf:#e66767; --mid:#383835; --target:#898781; --goodtext:#0ca30c; --warntext:#fab219;
  --shadow:0 1px 2px rgba(0,0,0,.4),0 4px 14px rgba(0,0,0,.35);
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    color-scheme: dark;
    --page:#0d0d0d; --surface:#1a1a19; --ink:#f0efec; --ink2:#c3c2b7; --muted:#898781;
    --grid:#2c2c2a; --axis:#383835; --border:rgba(255,255,255,.10); --wash:rgba(255,255,255,.06);
    --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500; --s5:#d55181; --s6:#008300; --s7:#9085e9; --s8:#e66767;
    --fav:#3987e5; --unf:#e66767; --mid:#383835; --target:#898781; --goodtext:#0ca30c; --warntext:#fab219;
    --shadow:0 1px 2px rgba(0,0,0,.4),0 4px 14px rgba(0,0,0,.35);
  }
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--page);color:var(--ink);font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased}
button,select,input{font:inherit;color:inherit}
.app{max-width:1560px;margin:0 auto;padding:0 20px 60px}
/* header */
.top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;padding:22px 0 14px;flex-wrap:wrap}
.brand h1{margin:0;font-size:22px;font-weight:650;letter-spacing:-.01em}
.brand p{margin:3px 0 0;color:var(--ink2);font-size:13px}
.top-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.btn{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:6px 12px;cursor:pointer;font-size:13px;display:inline-flex;gap:6px;align-items:center;white-space:nowrap}
.btn:hover{background:var(--wash)}
.btn:focus-visible,.chip:focus-visible,.tab:focus-visible,select:focus-visible,input:focus-visible,.seg button:focus-visible{outline:2px solid var(--s1);outline-offset:1px}
.btn.primary{background:var(--ink);color:var(--surface);border-color:var(--ink)}
.badge{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:3px 10px;font-size:12px;font-weight:600;border:1px solid var(--border);background:var(--surface);white-space:nowrap}
.badge.ok{color:var(--goodtext)} .badge.bad{color:var(--crit)} .badge.note{color:var(--ink2);font-weight:500}
/* filters */
.filters{position:sticky;top:0;z-index:30;background:color-mix(in srgb,var(--page) 92%,transparent);backdrop-filter:blur(8px);border-bottom:1px solid var(--border);margin:0 -20px;padding:10px 20px}
.frow{display:flex;gap:18px 22px;flex-wrap:wrap;align-items:flex-end}
.fg{display:flex;flex-direction:column;gap:4px;min-width:0}
.fg>label,.fg>.lbl{font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:var(--muted);font-weight:600}
.chips{display:flex;gap:5px;flex-wrap:wrap}
.chip{border:1px solid var(--border);background:var(--surface);border-radius:999px;padding:3px 10px;font-size:12.5px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}
.chip:hover{background:var(--wash)}
.chip[aria-pressed="true"]{background:var(--ink);color:var(--surface);border-color:var(--ink)}
.chip .dot{width:8px;height:8px;border-radius:50%;display:inline-block}
.sel{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:4px 8px;font-size:12.5px;max-width:200px}
.active{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-top:8px;min-height:22px;font-size:12px;color:var(--ink2)}
.tag{background:var(--wash);border-radius:6px;padding:2px 4px 2px 8px;display:inline-flex;gap:4px;align-items:center}
.tag button{border:0;background:transparent;cursor:pointer;color:var(--ink2);font-size:14px;line-height:1;padding:0 4px;border-radius:4px}
.tag button:hover{background:var(--grid)}
/* tabs */
.tabs{display:flex;gap:2px;flex-wrap:wrap;margin:16px 0 14px;border-bottom:1px solid var(--border)}
.tab{background:transparent;border:0;border-bottom:2px solid transparent;padding:9px 14px;cursor:pointer;color:var(--ink2);font-weight:550;font-size:13.5px;margin-bottom:-1px}
.tab:hover{color:var(--ink)}
.tab[aria-selected="true"]{color:var(--ink);border-bottom-color:var(--ink)}
/* grid & cards */
.grid{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:14px}
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:14px 16px 12px;box-shadow:var(--shadow);min-width:0;position:relative}
.card h3{margin:0;font-size:14px;font-weight:650}
.card .sub{color:var(--ink2);font-size:12px;margin-top:2px}
.card-h{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;margin-bottom:8px}
.card-h .tools{display:flex;gap:4px;flex:none}
.tool{border:1px solid var(--border);background:transparent;border-radius:6px;font-size:11.5px;padding:2px 8px;cursor:pointer;color:var(--ink2)}
.tool:hover{background:var(--wash)} .tool[aria-pressed="true"]{background:var(--ink);color:var(--surface)}
.s3{grid-column:span 3}.s4{grid-column:span 4}.s5{grid-column:span 5}.s6{grid-column:span 6}.s7{grid-column:span 7}.s8{grid-column:span 8}.s12{grid-column:span 12}
@media (max-width:1180px){.s3,.s4,.s5,.s6,.s7,.s8{grid-column:span 12}.s3.keep,.s4.keep{grid-column:span 6}}
@media (max-width:640px){.s3.keep,.s4.keep{grid-column:span 12}.app{padding:0 12px 40px}.filters{margin:0 -12px;padding:8px 12px}}
/* KPI */
.hero{display:flex;flex-direction:column;gap:6px;justify-content:space-between}
.hero .big{font-size:52px;font-weight:650;letter-spacing:-.02em;line-height:1.05;margin:2px 0}
.hero .row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.kpi .lab{font-size:12px;color:var(--ink2)}
.kpi .val{font-size:26px;font-weight:650;letter-spacing:-.01em;margin-top:2px;line-height:1.15}
.kpi .dl{font-size:12px;color:var(--ink2);margin-top:3px}
.pos{color:var(--fav)} .neg{color:var(--unf)}
.st{display:inline-flex;align-items:center;gap:5px;font-weight:600;font-size:12px;padding:2px 9px;border-radius:999px;border:1px solid currentColor}
.st.fav{color:var(--fav)} .st.unf{color:var(--unf)} .st.on{color:var(--ink2)} .st.nt{color:var(--muted)}
.mat{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:650;color:var(--warntext);background:color-mix(in srgb,var(--warntext) 13%,transparent);border:1px solid color-mix(in srgb,var(--warntext) 38%,transparent);border-radius:999px;padding:0 8px;line-height:1.55;white-space:nowrap}
.chr{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:700;color:var(--unf);background:color-mix(in srgb,var(--unf) 13%,transparent);border:1px solid color-mix(in srgb,var(--unf) 42%,transparent);border-radius:999px;padding:0 8px;line-height:1.55;white-space:nowrap}
/* key takeaway banner */
.takeaway{display:flex;gap:18px;align-items:stretch;justify-content:space-between;flex-wrap:wrap;margin:0 0 14px;padding:16px 20px 16px 18px;border-radius:12px;border:1px solid var(--border);border-left:5px solid var(--unf);background:var(--surface);box-shadow:var(--shadow)}
.takeaway.pos{border-left-color:var(--fav)}
.takeaway .tk-l{flex:1 1 520px;min-width:0}
.takeaway .tk-k{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--ink2);margin-bottom:6px}
.takeaway p{margin:0;font-size:16px;line-height:1.5;color:var(--ink)}
.takeaway p+p{margin-top:4px}
.takeaway strong{font-weight:700}
.takeaway .neg{color:var(--unf)} .takeaway .pos{color:var(--fav)}
.takeaway .tk-r{display:flex;gap:10px;align-items:stretch;flex-wrap:wrap}
.takeaway .stat{border:1px solid var(--border);border-radius:10px;padding:8px 14px;min-width:112px;background:var(--wash);display:flex;flex-direction:column;justify-content:center}
.takeaway .stat b{font-size:22px;font-weight:700;letter-spacing:-.01em;line-height:1.15}
.takeaway .stat span{font-size:11.5px;color:var(--ink2);margin-top:2px}
/* KPI tiers */
.kpi-wrap{display:flex;flex-direction:column;gap:14px}
.kpi-pri{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}
.kpi-sec{display:grid;grid-template-columns:repeat(auto-fit,minmax(138px,1fr));gap:12px}
.kpi.pri{padding:16px 16px 14px;border-top:3px solid var(--axis)}
.kpi.pri.neg-t{border-top-color:var(--unf)} .kpi.pri.pos-t{border-top-color:var(--fav)}
.kpi.pri .lab{font-size:12.5px;font-weight:600;color:var(--ink2)}
.kpi.pri .val{font-size:34px;margin-top:4px}
.kpi:not(.pri){padding:10px 12px 9px;box-shadow:none}
.kpi:not(.pri) .val{font-size:19px;font-weight:650}
.kpi:not(.pri) .lab{font-size:11.5px}
.kpi:not(.pri) .dl{font-size:11.5px}
@media (max-width:1180px){.kpi-pri{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media (max-width:640px){.kpi-pri{grid-template-columns:minmax(0,1fr)}}
.hero{border-left:5px solid var(--unf)} .hero.pos-t{border-left-color:var(--fav)}
.hero .big{font-size:60px}
/* filter bar: open = inline panel, closed = compact sticky strip */
.filters.open{position:static;background:transparent;backdrop-filter:none}
.filters.open .active{border-top:1px dashed var(--border);padding-top:8px}
.filters:not(.open){background:var(--page);box-shadow:0 2px 10px rgba(0,0,0,.06)}
.ftoggle{border:1px solid var(--border);background:var(--surface);border-radius:999px;padding:2px 11px;font-size:12px;font-weight:600;cursor:pointer;color:var(--ink)}
.ftoggle:hover{background:var(--wash)}
/* insights document */
.md{max-width:82ch;font-size:14px;line-height:1.58}
.md h3{padding-top:16px;margin-top:22px;border-top:1px solid var(--grid);scroll-margin-top:70px}
.md h2{font-size:22px}
.md-jump{display:flex;gap:6px;flex-wrap:wrap;margin:0 0 10px}

/* charts */
.chart{width:100%;position:relative}
.chart svg{display:block;width:100%;overflow:visible}
.chart text{font-family:inherit;fill:var(--ink2)}
.axis-t{font-size:11px;fill:var(--muted)}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--ink2);margin:0 0 6px}
.legend span{display:inline-flex;align-items:center;gap:6px}
.legend i{display:inline-block;width:14px;height:0;border-top:2px solid;border-radius:1px}
.legend b{display:inline-block;width:10px;height:10px;border-radius:2px}
.clk{cursor:pointer}
#tip{position:fixed;z-index:100;pointer-events:none;background:var(--surface);color:var(--ink);border:1px solid var(--border);box-shadow:0 6px 24px rgba(0,0,0,.18);border-radius:8px;padding:8px 10px;font-size:12.5px;max-width:320px;opacity:0;transition:opacity .08s}
#tip .t{font-weight:650;margin-bottom:4px} #tip .r{display:flex;gap:8px;align-items:center;justify-content:space-between;white-space:nowrap}
#tip .r .k{display:flex;gap:6px;align-items:center;color:var(--ink2)} #tip .r .v{font-weight:650;font-variant-numeric:tabular-nums}
#tip .sw{display:inline-block;width:12px;height:0;border-top:2px solid}
#tip .f{color:var(--muted);margin-top:4px;font-size:11.5px;white-space:normal}
/* tables */
.tbl-wrap{overflow:auto;max-height:560px;border:1px solid var(--border);border-radius:8px}
table.t{border-collapse:collapse;width:100%;font-size:12.5px}
table.t th{position:sticky;top:0;background:var(--surface);text-align:left;font-weight:650;padding:7px 10px;border-bottom:1px solid var(--axis);white-space:nowrap;cursor:pointer;user-select:none;z-index:1}
table.t th.r,table.t td.r{text-align:right}
table.t td{padding:6px 10px;border-bottom:1px solid var(--grid);white-space:nowrap;font-variant-numeric:tabular-nums}
table.t tr:hover td{background:var(--wash)}
table.t tr.clk td{cursor:pointer}
.tbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:0 0 8px}
.tbar input[type=search]{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:5px 9px;min-width:200px}
.pager{display:flex;gap:8px;align-items:center;justify-content:flex-end;margin-top:8px;font-size:12px;color:var(--ink2)}
.seg{display:inline-flex;border:1px solid var(--border);border-radius:8px;overflow:hidden;background:var(--surface)}
.seg button{border:0;background:transparent;padding:4px 10px;font-size:12px;cursor:pointer;border-right:1px solid var(--border)}
.seg button:last-child{border-right:0}
.seg button[aria-pressed="true"]{background:var(--ink);color:var(--surface)}
.ins{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:9px}
.ins li{display:flex;gap:9px;font-size:13px;line-height:1.45}
.ins li::before{content:"";flex:none;width:6px;height:6px;border-radius:50%;background:var(--s1);margin-top:7px}
.ins li.warn::before{background:var(--warn)} .ins li.bad::before{background:var(--unf)}
.note{font-size:12px;color:var(--ink2);background:var(--wash);border-radius:8px;padding:8px 10px;margin-top:8px}
.md h2{font-size:20px;margin:0 0 6px} .md h3{font-size:16px;margin:20px 0 6px} .md h4{font-size:14px;margin:14px 0 4px}
.md p{margin:6px 0} .md ul{margin:6px 0 6px 18px;padding:0} .md li{margin:3px 0}
.md code{background:var(--wash);padding:1px 5px;border-radius:4px}
.md-table{overflow:auto;margin:8px 0}.md-table table{border-collapse:collapse;font-size:12.5px}
.md-table th,.md-table td{border:1px solid var(--grid);padding:5px 9px;text-align:left}.md-table th{background:var(--wash)}
.pass{color:var(--goodtext);font-weight:650}.fail{color:var(--crit);font-weight:650}.warnc{color:var(--warntext);font-weight:650}
.slider{display:flex;gap:10px;align-items:center}.slider input[type=range]{flex:1;accent-color:var(--s1)}
.sr{position:absolute;left:-9999px}
.empty{padding:30px 10px;text-align:center;color:var(--ink2)}
footer{margin-top:28px;color:var(--muted);font-size:12px;text-align:center}
@media print{.filters,.tabs,.top-actions,.tools{display:none!important}.card{break-inside:avoid;box-shadow:none}body{background:#fff}}
@media (forced-colors:active){.card,.btn,.chip{border:1px solid CanvasText}}
</style>
</head>
<body>
<div class="app" id="app">
  <header class="top">
    <div class="brand">
      <h1 id="ttl">__TITLE__</h1>
      <p id="subttl"></p>
    </div>
    <div class="top-actions">
      <span id="recon" class="badge note">Reconciling…</span>
      <button class="btn" id="btnExport" title="Download the lines that match the current filters as CSV">⬇ Export filtered lines</button>
      <button class="btn" id="btnPrint" title="Print this view">⎙ Print</button>
      <button class="btn" id="btnTheme" title="Switch light / dark theme">◐ Theme</button>
    </div>
  </header>
  <section class="filters" id="filters" aria-label="Filters"></section>
  <nav class="tabs" id="tabs" role="tablist"></nav>
  <main id="view"></main>
  <footer id="foot"></footer>
</div>
<div id="tip" role="tooltip"></div>
<script id="payload" type="application/json">__DATA_JSON__</script>
<script>
'use strict';
/* ===========================================================================
   CORE: data decode, exact integer maths, filters, aggregation
   =========================================================================== */
const D = JSON.parse(document.getElementById('payload').textContent);
const DIM = D.dims, CFG = D.cfg, N = D.rows.R.length;
const REG = DIM.regions, CAT = DIM.cats, PRD = DIM.prods, MON = DIM.months, NOTES = DIM.notes, STN = DIM.statuses;
const NR = REG.length, NC = CAT.length, NP = PRD.length, NM = MON.length;
const A = {}; for (const k in D.rows) A[k] = Float64Array.from(D.rows[k]);
const MNAME = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const MYEAR = MON.map(m => +m.slice(0, 4)), MNUM = MON.map(m => +m.slice(5, 7));
const MQ = MNUM.map(n => Math.floor((n - 1) / 3) + 1);
const MORD = MON.map((m, i) => MYEAR[i] * 12 + MNUM[i] - 1);
const MLAB = MON.map((m, i) => MNAME[MNUM[i] - 1] + ' ' + String(MYEAR[i]).slice(2));
const MLONG = MON.map((m, i) => MNAME[MNUM[i] - 1] + ' ' + MYEAR[i]);
const YEARS = [...new Set(MYEAR)];
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const SVGNS = 'http://www.w3.org/2000/svg';

/* ---------- DOM helpers (text always via textContent) ---------- */
function h(tag, props, ...kids) {
  const el = document.createElement(tag);
  if (props) for (const k in props) {
    const v = props[k];
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else if (k === 'text') el.textContent = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(String(c)));
  return el;
}
function sv(tag, attrs, ...kids) {
  const el = document.createElementNS(SVGNS, tag);
  if (attrs) for (const k in attrs) {
    const v = attrs[k]; if (v == null) continue;
    if (typeof v === 'string' && v.startsWith('var(') && (k === 'fill' || k === 'stroke')) el.style.setProperty(k, v);
    else if (k === 'text') el.textContent = v;
    else el.setAttribute(k, v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(String(c)));
  return el;
}

/* ---------- exact integer helpers (money = cents, same rules as the Python analysis) ---------- */
function pct2(num, den) {            // num/den*100, 2 dp, half away from zero, exact on integers
  if (!(den > 0)) return NaN;
  const q = Math.floor((2 * Math.abs(num) * 10000 + den) / (2 * den));
  return (num < 0 ? -q : q) / 100;
}
function statusOf(vc, tc) {          // 0 Favorable, 1 On Target, 2 Unfavorable, 3 No Target
  if (!(tc > 0)) return 3;
  const l = vc * 100000;
  if (l > CFG.fav * tc) return 0;
  if (l < CFG.unf * tc) return 2;
  return 1;
}
function isMat(vc, tc) {
  const byUsd = Math.abs(vc) >= CFG.matUsdC;
  const byPct = tc > 0 && Math.abs(vc) * 100000 >= CFG.matPct * tc;
  return byUsd || byPct;
}
function trigOf(vc, tc) {
  const byUsd = Math.abs(vc) >= CFG.matUsdC, byPct = tc > 0 && Math.abs(vc) * 100000 >= CFG.matPct * tc;
  return byUsd && byPct ? 'Pct and USD' : byPct ? 'Pct' : byUsd ? 'USD' : '';
}

/* ---------- formatting ---------- */
const MINUS = '–';
function sgnNum(v, body) { return (v < 0 ? MINUS : '') + body; }
function usd(v, dp = 0) { if (!isFinite(v)) return '—'; return sgnNum(v, '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp })); }
function usdC(v, dp = 2) {
  if (!isFinite(v)) return '—'; const a = Math.abs(v);
  if (a >= 1e9) return sgnNum(v, '$' + (a / 1e9).toFixed(dp) + 'B');
  if (a >= 1e6) return sgnNum(v, '$' + (a / 1e6).toFixed(dp) + 'M');
  if (a >= 1e3) return sgnNum(v, '$' + (a / 1e3).toFixed(a >= 1e5 ? 0 : 1) + 'K');
  return sgnNum(v, '$' + a.toFixed(0));
}
function usdS(v, dp = 2) { return (v > 0 ? '+' : '') + usdC(v, dp); }          // signed compact
function pctF(v, dp = 2, plus = false) { if (!isFinite(v)) return '—'; return (v < 0 ? MINUS : (plus && v > 0 ? '+' : '')) + Math.abs(v).toFixed(dp) + '%'; }
function numF(v, dp = 0) { if (!isFinite(v)) return '—'; return sgnNum(v, Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp })); }
function ppF(v, dp = 2) { if (!isFinite(v)) return '—'; return (v > 0 ? '+' : v < 0 ? MINUS : '') + Math.abs(v).toFixed(dp) + ' pp'; }
const ST_NAME = ['Favorable', 'On Target', 'Unfavorable', 'No Target'];
const ST_GLYPH = ['▲', '●', '▼', '○'];
const ST_CLS = ['fav', 'on', 'unf', 'nt'];
function stBadge(s) { return h('span', { class: 'st ' + ST_CLS[s] }, ST_GLYPH[s] + ' ' + ST_NAME[s]); }
function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
const regColor = i => `var(--s${(i % 8) + 1})`, catColor = i => `var(--s${(i % 8) + 1})`;
const esc = s => String(s);

/* ---------- state & filters ---------- */
const S = { years: new Set(), qtrs: new Set(), regs: new Set(), cats: new Set(), prods: new Set(), stat: new Set(), mat: 'all', dq: 'all', m0: 0, m1: NM - 1 };
let VER = 0; const maskCache = new Map();
function bump() { VER++; maskCache.clear(); }
function monthOk() {
  const ok = new Uint8Array(NM);
  for (let m = 0; m < NM; m++) ok[m] = (m >= S.m0 && m <= S.m1 && (!S.years.size || S.years.has(MYEAR[m])) && (!S.qtrs.size || S.qtrs.has(MQ[m]))) ? 1 : 0;
  return ok;
}
function getMask(excl = '') {
  const key = excl; if (maskCache.has(key)) return maskCache.get(key);
  const ex = new Set(excl.split(',').filter(Boolean));
  const mOk = ex.has('time') ? null : monthOk();
  const useR = !ex.has('reg') && S.regs.size, useC = !ex.has('cat') && S.cats.size, useP = !ex.has('prod') && S.prods.size;
  const useS = !ex.has('stat') && S.stat.size, useM = !ex.has('mat') && S.mat !== 'all', useQ = !ex.has('dq') && S.dq !== 'all';
  const m = new Uint8Array(N), R = A.R, C = A.C, P = A.P, M = A.M, ST = A.ST, MT = A.MT, QR = A.QR;
  for (let i = 0; i < N; i++) {
    if (mOk && !mOk[M[i]]) continue;
    if (useR && !S.regs.has(R[i])) continue;
    if (useC && !S.cats.has(C[i])) continue;
    if (useP && !S.prods.has(P[i])) continue;
    if (useS && !S.stat.has(ST[i])) continue;
    if (useM && (S.mat === 'yes') !== (MT[i] === 1)) continue;
    if (useQ && (S.dq === 'repaired') !== (QR[i] === 1)) continue;
    m[i] = 1;
  }
  maskCache.set(key, m); return m;
}
function maskCount(m) { let c = 0; for (let i = 0; i < N; i++) c += m[i]; return c; }

/* ---------- aggregation ---------- */
function newG(K) {
  const f = () => new Float64Array(K);
  return { K, n: f(), tq: f(), aq: f(), tc: f(), ac: f(), vol: f(), prc: f(), res: f(), gm: f(), mls: f(), nf: f(), nu: f(), nn: f(), nm: f(), no: f() };
}
function agg(key, K, mask) {
  const g = newG(K);
  const TQ = A.TQ, AQ = A.AQ, TC = A.TC, AC = A.AC, VOL = A.VOL, PRC = A.PRC, RES = A.RES, GM = A.GM, MLS = A.MLS, ST = A.ST, MT = A.MT;
  for (let i = 0; i < N; i++) {
    if (!mask[i]) continue;
    const k = key[i];
    g.n[k]++; g.tq[k] += TQ[i]; g.aq[k] += AQ[i]; g.tc[k] += TC[i]; g.ac[k] += AC[i];
    g.vol[k] += VOL[i]; g.prc[k] += PRC[i]; g.res[k] += RES[i]; g.gm[k] += GM[i]; g.mls[k] += MLS[i];
    const s = ST[i]; if (s === 0) g.nf[k]++; else if (s === 2) g.nu[k]++; else if (s === 3) g.nn[k]++; else g.no[k]++;
    if (MT[i] === 1) g.nm[k]++;
  }
  return g;
}
function dv(g, k) {              // derived metrics of group k (all dollars in USD)
  const tc = g.tc[k], ac = g.ac[k], vc = ac - tc, tq = g.tq[k], aq = g.aq[k];
  const volT = g.vol[k] / 1e4, prcT = g.prc[k] / 1e4, resT = g.res[k] / 1e4;
  const pure = tq > 0 ? (aq - tq) * (tc / 100) / tq : volT;
  const st = statusOf(vc, tc);
  return {
    n: g.n[k], tq, aq, tc, ac, vc, target: tc / 100, actual: ac / 100, v: vc / 100,
    pct: pct2(vc, tc), att: pct2(aq, tq), qv: aq - tq, st, mat: isMat(vc, tc), trig: trigOf(vc, tc),
    vol: volT, prc: prcT, res: vc / 100 - volT - prcT, pure, mix: volT - pure, gm: g.gm[k] / 100, gmp: g.ac[k] > 0 ? g.gm[k] / g.ac[k] * 100 : NaN,
    mls: g.mls[k] / 1e4, nf: g.nf[k], nu: g.nu[k], nn: g.nn[k], no: g.no[k], nm: g.nm[k],
    pu: g.n[k] ? g.nu[k] / g.n[k] * 100 : NaN, price: aq > 0 ? ac / 100 / aq : NaN,
  };
}
const ONE = new Float64Array(N);                 // all zeros: key for grand totals
function totals(mask) { return dv(agg(ONE, 1, mask), 0); }
const KEY_RC = Float64Array.from({ length: N }, (_, i) => A.R[i] * NC + A.C[i]);
const KEY_RM = Float64Array.from({ length: N }, (_, i) => A.R[i] * NM + A.M[i]);
const KEY_CM = Float64Array.from({ length: N }, (_, i) => A.C[i] * NM + A.M[i]);
const KEY_RCM = Float64Array.from({ length: N }, (_, i) => (A.R[i] * NC + A.C[i]) * NM + A.M[i]);
const KEY_RP = Float64Array.from({ length: N }, (_, i) => A.R[i] * NP + A.P[i]);
const KEY_RY = Float64Array.from({ length: N }, (_, i) => A.R[i] * 2 + (MYEAR[A.M[i]] === YEARS[YEARS.length - 1] ? 1 : 0));

/* ---------- monthly series with growth metrics (exact, matched on calendar month) ---------- */
function monthlyFrom(g, off, mask2) {
  // g: group aggregation by (group*NM + m); off = group*NM
  const out = [];
  for (let m = 0; m < NM; m++) out.push(dv(g, off + m));
  const lagIdx = (m, lag) => { const t = MORD[m] - lag; return MORD.indexOf(t); };
  let cum = 0, absSum = 0, cnt = 0;
  const ytd = {};
  for (let m = 0; m < NM; m++) {
    const d = out[m];
    const p1 = lagIdx(m, 1), p12 = lagIdx(m, 12);
    d.mom = p1 >= 0 && out[p1].ac > 0 ? (d.ac / out[p1].ac - 1) * 100 : NaN;
    d.yoy = p12 >= 0 && out[p12].ac > 0 ? (d.ac / out[p12].ac - 1) * 100 : NaN;
    const y = MYEAR[m]; ytd[y] = ytd[y] || { t: 0, a: 0 }; ytd[y].t += d.tc; ytd[y].a += d.ac;
    d.ytdV = (ytd[y].a - ytd[y].t) / 100; d.ytdP = pct2(ytd[y].a - ytd[y].t, ytd[y].t);
    const p2 = lagIdx(m, 2);
    if (p1 >= 0 && p2 >= 0) { const t3 = d.tc + out[p1].tc + out[p2].tc, a3 = d.ac + out[p1].ac + out[p2].ac; d.r3 = (a3 - t3) / 100; d.r3p = pct2(a3 - t3, t3); } else { d.r3 = NaN; d.r3p = NaN; }
    cum += d.vc; absSum += Math.abs(d.vc); cnt++;
    d.cum = cum / 100; d.ts = absSum > 0 ? cum / (absSum / cnt) : NaN;
  }
  return out;
}
function monthlySeries(mask, groupKey, K, gi) {          // gi = group index or -1 for total
  if (gi < 0) return monthlyFrom(agg(A.M, NM, mask), 0);
  return monthlyFrom(agg(groupKey, K * NM, mask), gi * NM);
}

/* ---------- tooltip ---------- */
const tipEl = document.getElementById('tip');
function showTip(ev, title, rows, foot) {
  tipEl.replaceChildren();
  if (title) tipEl.append(h('div', { class: 't' }, title));
  for (const r of rows || []) {
    const k = h('span', { class: 'k' }, r.color ? h('span', { class: 'sw', style: { borderTopColor: r.color } }) : null, r.label);
    tipEl.append(h('div', { class: 'r' }, k, h('span', { class: 'v' }, r.value)));
  }
  if (foot) tipEl.append(h('div', { class: 'f' }, foot));
  tipEl.style.opacity = 1;
  const pad = 14, w = tipEl.offsetWidth, hh = tipEl.offsetHeight;
  let x = ev.clientX + pad, y = ev.clientY + pad;
  if (x + w > innerWidth - 8) x = ev.clientX - w - pad;
  if (y + hh > innerHeight - 8) y = innerHeight - hh - 8;
  tipEl.style.left = Math.max(8, x) + 'px'; tipEl.style.top = Math.max(8, y) + 'px';
}
function hideTip() { tipEl.style.opacity = 0; }
/* ===========================================================================
   CHARTS (hand-built SVG - no libraries, works offline)
   Specs: thin marks (<=24px), 4px rounded data-end, hairline recessive grid,
   one axis per chart, tooltips on every mark, legend whenever >=2 series.
   =========================================================================== */
function niceTicks(lo, hi, n = 5) {
  if (!isFinite(lo) || !isFinite(hi)) return [0, 1];
  if (lo === hi) { const p = Math.abs(lo) || 1; lo -= p * .5; hi += p * .5; }
  const span = hi - lo, raw = span / n, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag, step = (norm >= 5 ? 10 : norm >= 2 ? 5 : norm >= 1 ? 2 : 1) * mag;   // 1,2,5,10
  const t0 = Math.floor(lo / step + 1e-9) * step, out = [];
  let v = t0; for (; v <= hi + step * 1e-9; v += step) out.push(+v.toPrecision(12));
  if (out[out.length - 1] < hi - step * 1e-9) out.push(+v.toPrecision(12));
  return out;
}
function lin(d0, d1, r0, r1) { const k = (r1 - r0) / ((d1 - d0) || 1); return v => r0 + (v - d0) * k; }
function txt(x, y, s, o = {}) { return sv('text', Object.assign({ x, y, 'font-size': 11, class: 'axis-t' }, o), s); }
function chartHost(el) { el.classList.add('chart'); return Math.max(260, el.clientWidth || 600); }
function mixColor(a, b, t) {        // a,b = '#rrggbb' or 'rgb(...)'
  const pa = parseRGB(a), pb = parseRGB(b); const c = pa.map((v, i) => Math.round(v + (pb[i] - v) * t));
  return `rgb(${c[0]},${c[1]},${c[2]})`;
}
function parseRGB(s) {
  s = s.trim();
  if (s[0] === '#') { if (s.length === 4) s = '#' + [...s.slice(1)].map(c => c + c).join(''); return [1, 3, 5].map(i => parseInt(s.slice(i, i + 2), 16)); }
  const m = s.match(/[\d.]+/g); return m ? m.slice(0, 3).map(Number) : [128, 128, 128];
}
function lum(rgb) { const f = c => { c /= 255; return c <= .03928 ? c / 12.92 : Math.pow((c + .055) / 1.055, 2.4); }; return .2126 * f(rgb[0]) + .7152 * f(rgb[1]) + .0722 * f(rgb[2]); }
function inkOn(bg) { return lum(parseRGB(bg)) > .42 ? '#0b0b0b' : '#ffffff'; }
function divColor(v, maxAbs) {          // blue (positive) <-> neutral <-> red (negative)
  if (!isFinite(v)) return cssVar('--wash') || '#eee';
  const t = Math.min(1, Math.abs(v) / (maxAbs || 1)), mid = cssVar('--mid'), pole = cssVar(v >= 0 ? '--fav' : '--unf');
  return mixColor(mid, pole, Math.pow(t, .8));
}
function legendEl(items) {
  return h('div', { class: 'legend' }, items.map(it => h('span', null, it.type === 'bar' ? h('b', { style: { background: it.color } }) : h('i', { style: { borderTopColor: it.color } }), it.label)));
}
function roundedBarPath(a, b, c, w, r, horizontal) {      // data end rounded (4px), baseline end square
  if (horizontal) {      // a = baseline x, b = top y of the bar, c = tip x, w = thickness
    const dir = c >= a ? 1 : -1, len = Math.abs(c - a), rr = Math.min(r, len, w / 2);
    if (len < .5) return '';
    const yt = b, yb = b + w;
    return `M${a},${yt}H${c - dir * rr}Q${c},${yt} ${c},${yt + rr}V${yb - rr}Q${c},${yb} ${c - dir * rr},${yb}H${a}Z`;
  }
  const len = Math.abs(c - b), rr = Math.min(r, len, w / 2);   // vertical: a = left x, b = baseline y, c = tip y
  if (len < .5) return '';
  const e = c < b ? rr : -rr, xl = a, xr = a + w;
  return `M${xl},${b}V${c + e}Q${xl},${c} ${xl + rr},${c}H${xr - rr}Q${xr},${c} ${xr},${c + e}V${b}Z`;
}

/* ---------- 1. Overview trend: two aligned panels, one axis each ---------- */
function trendChart(host, o) {
  // o: {tc[],ac[],vc[],pct[],att[],inSel[], onMonth(m), selected}
  const W = chartHost(host), ml = 58, mr = 14, hA = 190, hB = 110, gap = 34, top = 6, xAx = 26, H = top + hA + gap + hB + xAx;
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Monthly target versus actual sales and variance' });
  const n = o.ac.length, pw = W - ml - mr, band = pw / n, X = i => ml + band * (i + .5);
  const allV = o.ac.concat(o.tc).map(v => v / 1e8);   // $M scale: cents/1e8 = millions
  let lo = Math.min(...allV), hi = Math.max(...allV); const pad = (hi - lo) * .08 || 1;
  const tA = niceTicks(lo - pad, hi + pad, 4), yA = lin(tA[0], tA[tA.length - 1], top + hA, top);
  tA.forEach(t => { svg.append(sv('line', { x1: ml, x2: W - mr, y1: yA(t), y2: yA(t), stroke: 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, yA(t) + 4, '$' + t.toFixed(t % 1 ? 1 : 0) + 'M', { 'text-anchor': 'end' })); });
  // gap wash between lines
  for (let i = 0; i < n - 1; i++) {
    const s = (o.vc[i] + o.vc[i + 1]) / 2, c = s >= 0 ? 'var(--fav)' : 'var(--unf)';
    const pts = [[X(i), yA(o.ac[i] / 1e8)], [X(i + 1), yA(o.ac[i + 1] / 1e8)], [X(i + 1), yA(o.tc[i + 1] / 1e8)], [X(i), yA(o.tc[i] / 1e8)]];
    svg.append(sv('polygon', { points: pts.map(p => p.join(',')).join(' '), fill: c, 'fill-opacity': .16 }));
  }
  const path = a => a.map((v, i) => (i ? 'L' : 'M') + X(i) + ',' + yA(v / 1e8)).join('');
  svg.append(sv('path', { d: path(o.tc), fill: 'none', stroke: 'var(--target)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
  svg.append(sv('path', { d: path(o.ac), fill: 'none', stroke: 'var(--s1)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
  [['tc', 'var(--target)'], ['ac', 'var(--s1)']].forEach(([k, c]) => svg.append(sv('circle', { cx: X(n - 1), cy: yA(o[k][n - 1] / 1e8), r: 4, fill: c, stroke: 'var(--surface)', 'stroke-width': 2 })));
  // panel B: variance columns from a zero baseline
  const yb0 = top + hA + gap; const vv = o.vc.map(v => v / 1e5);   // $K
  const vlo = Math.min(0, ...vv), vhi = Math.max(0, ...vv), tB = niceTicks(vlo * 1.12, vhi * 1.12 || 1, 3);
  const yB = lin(tB[0], tB[tB.length - 1], yb0 + hB, yb0);
  svg.append(txt(ml, yb0 - 10, 'Variance (actual – target), $K', { 'font-weight': 600 }));
  tB.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: yB(t), y2: yB(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, yB(t) + 4, numF(t), { 'text-anchor': 'end' })));
  const bw = Math.min(24, band * .62), y0 = yB(0);
  const bars = [];
  for (let i = 0; i < n; i++) {
    const y1 = yB(vv[i]), c = vv[i] >= 0 ? 'var(--fav)' : 'var(--unf)';
    const p = sv('path', { d: roundedBarPath(X(i) - bw / 2, y0, y1, bw, 4, false), fill: c, opacity: o.inSel[i] ? 1 : .28 }); bars.push(p); svg.append(p);
  }
  // selective labels: extremes
  const iMin = vv.indexOf(Math.min(...vv)), iMax = vv.indexOf(Math.max(...vv));
  [iMin, iMax].filter((v, k, a) => a.indexOf(v) === k).forEach(i => svg.append(txt(X(i), vv[i] < 0 ? yB(vv[i]) + 13 : yB(vv[i]) - 5, numF(vv[i]), { 'text-anchor': 'middle', 'font-size': 10.5, style: 'fill:var(--ink2)' })));
  // x axis
  const every = band < 52 ? 2 : 1;
  for (let i = 0; i < n; i++) if (i % every === 0) svg.append(txt(X(i), yb0 + hB + 17, MLAB[i].replace(' ', ' '), { 'text-anchor': 'middle', 'font-size': 10.5 }));
  // hover
  const cross = sv('line', { y1: top, y2: yb0 + hB, stroke: 'var(--ink2)', 'stroke-width': 1, opacity: 0 });
  const rect = sv('rect', { x: ml, y: 0, width: pw, height: H, fill: 'transparent', style: 'cursor:pointer' });
  svg.append(cross, rect);
  const idxAt = ev => { const r = svg.getBoundingClientRect(); const x = (ev.clientX - r.left) * (W / r.width); return Math.max(0, Math.min(n - 1, Math.floor((x - ml) / band))); };
  rect.addEventListener('pointermove', ev => {
    const i = idxAt(ev); cross.setAttribute('x1', X(i)); cross.setAttribute('x2', X(i)); cross.setAttribute('opacity', .5);
    showTip(ev, MLONG[o.m[i]], [
      { color: 'var(--target)', label: 'Target', value: usd(o.tc[i] / 100) }, { color: 'var(--s1)', label: 'Actual', value: usd(o.ac[i] / 100) },
      { label: 'Variance', value: usd(o.vc[i] / 100) + '  (' + pctF(o.pct[i], 2, true) + ')' }, { label: 'Qty attainment', value: pctF(o.att[i]) },
      { label: 'Status (month total)', value: ST_GLYPH[o.st[i]] + ' ' + ST_NAME[o.st[i]] + (o.mat[i] ? '  ◆ material' : '') }], 'Click to filter to this month');
  });
  rect.addEventListener('pointerleave', () => { cross.setAttribute('opacity', 0); hideTip(); });
  rect.addEventListener('click', ev => o.onMonth && o.onMonth(o.m[idxAt(ev)]));
  host.replaceChildren(legendEl([{ label: 'Target', color: 'var(--target)' }, { label: 'Actual', color: 'var(--s1)' }, { label: 'Gap: actual above target', color: 'var(--fav)', type: 'bar' }, { label: 'Gap: actual below target', color: 'var(--unf)', type: 'bar' }]), svg);
}

/* ---------- 2. Horizontal diverging bars (clickable) ---------- */
function hbars(host, items, o = {}) {
  const W = chartHost(host), rowH = o.rowH || 30, bh = Math.min(o.barH || 18, rowH - 8), mt = 4, H = items.length * rowH + mt + (o.axis === false ? 4 : 22);
  const labW = Math.min(o.labW || 150, Math.max(70, Math.max(0, ...items.map(i => i.label.length)) * 6.4 + 6));
  const vals = items.map(i => i.value), hasNeg = vals.some(v => v < 0), hasPos = vals.some(v => v > 0);
  const lo = Math.min(0, ...vals), hi = Math.max(0, ...vals);
  const xL = labW + 10 + (hasNeg ? 96 : 0), xR = W - (hasPos ? 100 : 16), x = lin(lo, hi || 1, xL, xR), x0 = x(0);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Bar chart' });
  niceTicks(lo, hi, 3).forEach(t => { if (t < lo - 1e-9 || t > hi + 1e-9) return; svg.append(sv('line', { x1: x(t), x2: x(t), y1: mt, y2: H - 18, stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 })); if (o.axis !== false) svg.append(txt(x(t), H - 4, (o.tickFmt || (v => usdC(v, 1)))(t), { 'text-anchor': 'middle', 'font-size': 10.5 })); });
  const anySel = items.some(i => i.selected);
  items.forEach((it, k) => {
    const y = mt + k * rowH + (rowH - bh) / 2, c = it.color || (it.value >= 0 ? 'var(--fav)' : 'var(--unf)');
    const g = sv('g', { class: o.onClick ? 'clk' : '', opacity: anySel && !it.selected ? .38 : 1, tabindex: o.onClick ? 0 : null, role: o.onClick ? 'button' : null, 'aria-label': it.label + ': ' + it.txt });
    g.append(sv('rect', { x: 0, y: mt + k * rowH, width: W, height: rowH, fill: 'transparent' }));
    g.append(txt(labW, y + bh / 2 + 4, it.label, { 'text-anchor': 'end', 'font-size': 12, style: 'fill:var(--ink)' }));
    const bar = sv('path', { d: roundedBarPath(x0, y, x(it.value), bh, 4, true), fill: c }); g.append(bar);
    const tipX = x(it.value), pos = it.value >= 0;
    g.append(txt(pos ? tipX + 6 : tipX - 6, y + bh / 2 + 4, it.txt, { 'text-anchor': pos ? 'start' : 'end', 'font-size': 11.5, style: 'fill:var(--ink)', 'font-weight': 600 }));
    g.addEventListener('pointermove', ev => { bar.style.filter = 'brightness(1.12)'; const t = o.tip ? o.tip(it) : { title: it.label, rows: [{ label: 'Value', value: it.txt }] }; showTip(ev, t.title, t.rows, o.onClick ? (it.selected ? 'Click to remove from filter' : 'Click to filter') : t.foot); });
    g.addEventListener('pointerleave', () => { bar.style.filter = ''; hideTip(); });
    if (o.onClick) { g.addEventListener('click', ev => o.onClick(it, ev)); g.addEventListener('keydown', ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); o.onClick(it, ev); } }); }
    svg.append(g);
  });
  host.replaceChildren(svg);
}

/* ---------- 3. Waterfall (floating bars; truncated axis is labelled) ---------- */
function waterfall(host, steps, o = {}) {
  // steps: [{label, value, kind:'total'|'delta', color}]
  const W = chartHost(host), H = o.height || 300, ml = 60, mr = 12, mt = 26, mb = 52;
  let run = 0; const bars = steps.map(s => { if (s.kind === 'total') { run = s.value; return { s, a: 0, b: s.value, end: run }; } const a = run; run += s.value; return { s, a, b: run, end: run }; });
  const all = bars.flatMap(b => b.s.kind === 'total' ? [b.b] : [b.a, b.b, b.end]); let lo = Math.min(...all), hi = Math.max(...all);
  const totals = steps.filter(s => s.kind === 'total').map(s => s.value);
  let trunc = false, base = 0;
  if (totals.length && Math.min(...totals) > 0 && (hi - lo) < 0.5 * hi) { const dmin = Math.min(...bars.filter(b => b.s.kind === 'delta').flatMap(b => [b.a, b.b])); base = Math.min(dmin, ...totals); const sp = hi - base; base = base - sp * .6; trunc = base > 0; lo = trunc ? base : 0; }
  else lo = Math.min(0, lo);
  const ticks = niceTicks(lo, hi * 1.0, 4), y = lin(ticks[0], ticks[ticks.length - 1], H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Waterfall chart' });
  ticks.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, usdC(t, ticks[1] - ticks[0] >= 1e6 ? 0 : 1), { 'text-anchor': 'end' })));
  const n = steps.length, band = (W - ml - mr) / n, bw = Math.min(46, band * .6), X = i => ml + band * (i + .5);
  bars.forEach((b, i) => {
    const top = Math.max(b.a, b.b), bot = Math.min(b.a, b.b), yT = y(Math.max(top, ticks[0])), yB = y(Math.max(bot, ticks[0]));
    const c = b.s.color || (b.s.kind === 'total' ? 'var(--s1)' : (b.s.value >= 0 ? 'var(--fav)' : 'var(--unf)'));
    const r = sv('rect', { x: X(i) - bw / 2, y: yT, width: bw, height: Math.max(1.5, yB - yT), rx: 3, fill: c, class: 'clk' });
    r.addEventListener('pointermove', ev => showTip(ev, b.s.label, [{ label: b.s.kind === 'total' ? 'Total' : 'Change', value: usd(b.s.value) }, { label: 'Running total', value: usd(b.end) }].concat(b.s.extra || []), b.s.foot));
    r.addEventListener('pointerleave', hideTip); svg.append(r);
    if (i < n - 1) svg.append(sv('line', { x1: X(i) + bw / 2, x2: X(i + 1) - bw / 2, y1: y(b.end), y2: y(b.end), stroke: 'var(--axis)', 'stroke-width': 1 }));
    const lab = b.s.kind === 'total' ? usdC(b.s.value) : usdS(b.s.value);
    const up = b.s.kind === 'total' || b.s.value >= 0;
    svg.append(txt(X(i), up ? yT - 6 : yB + 14, lab, { 'text-anchor': 'middle', 'font-size': 11.5, 'font-weight': 600, style: 'fill:var(--ink)' }));
    const words = b.s.label.split(' '); const l1 = words.slice(0, Math.ceil(words.length / 2)).join(' '), l2 = words.slice(Math.ceil(words.length / 2)).join(' ');
    svg.append(txt(X(i), H - mb + 16, l1, { 'text-anchor': 'middle', 'font-size': 11 }), txt(X(i), H - mb + 29, l2, { 'text-anchor': 'middle', 'font-size': 11 }));
  });
  if (trunc) svg.append(txt(ml, H - 4, 'Axis starts at ' + usdC(ticks[0], 0) + ' so the changes are visible – bar heights are not proportional to totals.', { 'font-size': 10.5 }));
  host.replaceChildren(svg);
}

/* ---------- 4. Heatmap (diverging, symmetric around zero) ---------- */
function heatmap(host, o) {
  // o: {rows[], cols[], colSub[], vals[r][c], fmt(v), cellTip(r,c)->{title,rows}, onCell(r,c), selR:Set, selC:Set, rowMarg[], colMarg[], diverging, legendFmt}
  const W = chartHost(host), nr = o.rows.length, nc = o.cols.length, labW = o.labW || 118, mt = o.colSub ? 36 : 22, cw = (W - labW - 8) / nc;
  const ch = o.cellH || (nc > 12 ? 30 : 44), H = mt + nr * ch + 30;
  const all = o.vals.flat().filter(isFinite); const mx = o.maxAbs || Math.max(1e-9, ...all.map(Math.abs)), mn = Math.min(...all), mxv = Math.max(...all);
  const seqLo = cssVar('--mid'), seqHi = cssVar('--s1');
  const colorOf = v => o.diverging === false ? mixColor(seqLo, seqHi, Math.pow(Math.max(0, (v - mn) / ((mxv - mn) || 1)), .9)) : divColor(v, mx);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Heatmap' });
  o.cols.forEach((c, j) => { svg.append(txt(labW + cw * (j + .5), 14, c, { 'text-anchor': 'middle', 'font-size': cw < 40 ? 10 : 11.5, style: 'fill:var(--ink2)' })); if (o.colSub && o.colSub[j]) svg.append(txt(labW + cw * (j + .5), 27, o.colSub[j], { 'text-anchor': 'middle', 'font-size': 10 })); });
  const anyR = o.selR && o.selR.size, anyC = o.selC && o.selC.size;
  o.rows.forEach((r, i) => {
    svg.append(txt(labW - 8, mt + i * ch + ch / 2 + 4, r, { 'text-anchor': 'end', 'font-size': 12, style: 'fill:var(--ink)' }));
    o.cols.forEach((c, j) => {
      const v = o.vals[i][j], col = colorOf(v), dim = (anyR && !o.selR.has(i)) || (anyC && !o.selC.has(j));
      const g = sv('g', { class: o.onCell ? 'clk' : '', opacity: dim ? .35 : 1 });
      const rc = sv('rect', { x: labW + cw * j + 1, y: mt + i * ch + 1, width: cw - 2, height: ch - 2, rx: 3, fill: col });
      g.append(rc);
      if (cw >= 44 || (cw >= 30 && ch >= 28 && o.fmtShort)) { const s = (cw >= 44 ? o.fmt : o.fmtShort)(v); if (s) g.append(txt(labW + cw * (j + .5), mt + i * ch + ch / 2 + 4, s, { 'text-anchor': 'middle', 'font-size': cw >= 44 ? 11.5 : 9.5, style: `fill:${inkOn(col)}`, 'font-weight': 600 })); }
      g.addEventListener('pointermove', ev => { rc.style.stroke = 'var(--ink)'; rc.style.strokeWidth = 2; const t = o.cellTip(i, j); showTip(ev, t.title, t.rows, t.foot || (o.onCell ? 'Click to filter' : '')); });
      g.addEventListener('pointerleave', () => { rc.style.stroke = ''; rc.style.strokeWidth = ''; hideTip(); });
      if (o.onCell) g.addEventListener('click', () => o.onCell(i, j));
      svg.append(g);
    });
  });
  // scale legend
  const ly = mt + nr * ch + 12, lw = Math.min(260, W - labW), lx = labW; const steps = 40;
  for (let k = 0; k < steps; k++) { const v = o.diverging === false ? mn + (mxv - mn) * k / (steps - 1) : -mx + 2 * mx * k / (steps - 1); svg.append(sv('rect', { x: lx + lw * k / steps, y: ly, width: lw / steps + .5, height: 8, fill: colorOf(v) })); }
  const lf = o.legendFmt || (v => v.toFixed(1));
  svg.append(txt(lx, ly + 20, o.diverging === false ? lf(mn) : lf(-mx), { 'font-size': 10.5 }), txt(lx + lw, ly + 20, o.diverging === false ? lf(mxv) : lf(mx), { 'text-anchor': 'end', 'font-size': 10.5 }));
  if (o.diverging !== false) svg.append(txt(lx + lw / 2, ly + 20, lf(0), { 'text-anchor': 'middle', 'font-size': 10.5 }));
  svg.append(txt(lx + lw + 12, ly + 8, o.legendNote || '', { 'font-size': 10.5 }));
  host.replaceChildren(svg);
}

/* ---------- 5. Multi-series line / single-series column chart with crosshair ---------- */
function lineChart(host, o) {
  // o: {x[], xFull[], series:[{id,name,color,values[]}], fmt(v), tickFmt, zero:boolean, mode:'line'|'bar', onX(i), selX:Uint8Array, height, refs:[{y,label}]}
  const W = chartHost(host), H = o.height || 260, ml = o.ml || 58, mr = 16, mt = 10, mb = 28, n = o.x.length, pw = W - ml - mr, band = pw / n, X = i => ml + band * (i + .5);
  const vals = o.series.flatMap(s => s.values).filter(isFinite); let lo = Math.min(...vals, ...(o.refs || []).map(r => r.y)), hi = Math.max(...vals, ...(o.refs || []).map(r => r.y));
  if (o.zero || o.mode === 'bar') { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  const pad = (hi - lo) * .08 || 1; const ticks = niceTicks(lo - (o.mode === 'bar' || o.zero ? 0 : pad), hi + pad, 5), y = lin(ticks[0], ticks[ticks.length - 1], H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Line chart' });
  ticks.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, (o.tickFmt || o.fmt)(t), { 'text-anchor': 'end' })));
  (o.refs || []).forEach(r => { svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(r.y), y2: y(r.y), stroke: 'var(--ink2)', 'stroke-width': 1, opacity: .55 }), txt(W - mr - 4, y(r.y) - 4, r.label, { 'text-anchor': 'end', 'font-size': 10.5 })); });
  const every = band < 52 ? 2 : 1; for (let i = 0; i < n; i++) if (i % every === 0) svg.append(txt(X(i), H - 8, o.x[i].replace(' ', ' '), { 'text-anchor': 'middle', 'font-size': 10.5 }));
  if (o.mode === 'bar') {
    const s = o.series[0], bw = Math.min(24, band * .62), y0 = y(0);
    s.values.forEach((v, i) => { if (!isFinite(v)) return; svg.append(sv('path', { d: roundedBarPath(X(i) - bw / 2, y0, y(v), bw, 4, false), fill: v >= 0 ? 'var(--fav)' : 'var(--unf)', opacity: !o.selX || o.selX[i] ? 1 : .28 })); });
  } else {
    o.series.forEach(s => {
      let d = '', pen = false; s.values.forEach((v, i) => { if (!isFinite(v)) { pen = false; return; } d += (pen ? 'L' : 'M') + X(i) + ',' + y(v); pen = true; });
      svg.append(sv('path', { d, fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round', opacity: s.dim ? .35 : 1 }));
      let li = -1; s.values.forEach((v, i) => { if (isFinite(v)) li = i; });
      if (li >= 0) svg.append(sv('circle', { cx: X(li), cy: y(s.values[li]), r: 4, fill: s.color, stroke: 'var(--surface)', 'stroke-width': 2 }));
    });
  }
  const cross = sv('line', { y1: mt, y2: H - mb, stroke: 'var(--ink2)', 'stroke-width': 1, opacity: 0 }), dots = o.series.map(s => sv('circle', { r: 4, fill: s.color, stroke: 'var(--surface)', 'stroke-width': 2, opacity: 0 }));
  const rect = sv('rect', { x: ml, y: 0, width: pw, height: H, fill: 'transparent', style: o.onX ? 'cursor:pointer' : '' });
  svg.append(cross, ...dots, rect);
  const idxAt = ev => { const r = svg.getBoundingClientRect(); const px = (ev.clientX - r.left) * (W / r.width); return Math.max(0, Math.min(n - 1, Math.floor((px - ml) / band))); };
  rect.addEventListener('pointermove', ev => {
    const i = idxAt(ev); cross.setAttribute('x1', X(i)); cross.setAttribute('x2', X(i)); cross.setAttribute('opacity', .5);
    dots.forEach((d, k) => { const v = o.series[k].values[i]; if (isFinite(v) && o.mode !== 'bar') { d.setAttribute('cx', X(i)); d.setAttribute('cy', y(v)); d.setAttribute('opacity', 1); } else d.setAttribute('opacity', 0); });
    const rows = o.series.map(s => ({ color: s.color, label: s.name, value: o.fmt(s.values[i]) })).sort((a, b) => 0);
    showTip(ev, (o.xFull || o.x)[i], rows.concat(o.extraTip ? o.extraTip(i) : []), o.onX ? 'Click to filter to this month' : '');
  });
  rect.addEventListener('pointerleave', () => { cross.setAttribute('opacity', 0); dots.forEach(d => d.setAttribute('opacity', 0)); hideTip(); });
  if (o.onX) rect.addEventListener('click', ev => o.onX(idxAt(ev)));
  const kids = []; if (o.series.length >= 2) kids.push(legendEl(o.series.map(s => ({ label: s.name, color: s.color })))); kids.push(svg);
  host.replaceChildren(...kids);
}

/* ---------- 6. Bubble / scatter with nearest-point hit layer ---------- */
function bubble(host, o) {
  // o: {pts:[{id,label,x,y,r,color,tipRows,selected}], xFmt,yFmt, xTitle,yTitle, bands:[{y0,y1,label}], labelIds:Set, onClick(p), height}
  const W = chartHost(host), H = o.height || 340, ml = 58, mr = 16, mt = 12, mb = 40;
  const xs = o.pts.map(p => p.x), ys = o.pts.map(p => p.y);
  const tx = niceTicks(Math.min(...xs) - (Math.max(...xs) - Math.min(...xs)) * .06, Math.max(...xs) + (Math.max(...xs) - Math.min(...xs)) * .06, 5);
  const yl = Math.min(...ys, 0), yh = Math.max(...ys, 0), ty = niceTicks(yl - (yh - yl) * .08, yh + (yh - yl) * .08, 5);
  const x = lin(tx[0], tx[tx.length - 1], ml, W - mr), y = lin(ty[0], ty[ty.length - 1], H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Scatter chart' });
  (o.bands || []).forEach(b => { const a = y(Math.min(ty[ty.length - 1], b.y1)), c = y(Math.max(ty[0], b.y0)); svg.append(sv('rect', { x: ml, y: a, width: W - ml - mr, height: Math.max(0, c - a), fill: 'var(--wash)' }), txt(W - mr - 4, a + 12, b.label, { 'text-anchor': 'end', 'font-size': 10.5 })); });
  ty.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, o.yFmt(t), { 'text-anchor': 'end' })));
  tx.forEach(t => svg.append(txt(x(t), H - mb + 16, o.xFmt(t), { 'text-anchor': 'middle' })));
  svg.append(txt((ml + W - mr) / 2, H - 4, o.xTitle || '', { 'text-anchor': 'middle', 'font-weight': 600 }));
  svg.append(txt(ml, mt - 1, o.yTitle || '', { 'font-weight': 600 }));
  const anySel = o.pts.some(p => p.selected);
  const dots = o.pts.map(p => { const c = sv('circle', { cx: x(p.x), cy: y(p.y), r: p.r, fill: p.color, 'fill-opacity': .82, stroke: 'var(--surface)', 'stroke-width': 2, opacity: anySel && !p.selected ? .35 : 1 }); svg.append(c); return c; });
  o.pts.forEach(p => { if (o.labelIds && o.labelIds.has(p.id)) svg.append(txt(x(p.x) + p.r + 4, y(p.y) + 4, p.label, { 'font-size': 10.5, style: 'fill:var(--ink)' })); });
  const rect = sv('rect', { x: ml, y: 0, width: W - ml - mr, height: H, fill: 'transparent', style: o.onClick ? 'cursor:pointer' : '' }); svg.append(rect);
  const near = ev => { const r = svg.getBoundingClientRect(), k = W / r.width, px = (ev.clientX - r.left) * k, py = (ev.clientY - r.top) * k; let best = -1, bd = 1e9; o.pts.forEach((p, i) => { const d = Math.hypot(x(p.x) - px, y(p.y) - py); if (d < bd) { bd = d; best = i; } }); return bd <= 36 ? best : -1; };
  let last = -1;
  rect.addEventListener('pointermove', ev => { const i = near(ev); if (last >= 0 && dots[last]) dots[last].setAttribute('r', o.pts[last].r); last = i; if (i < 0) { hideTip(); return; } dots[i].setAttribute('r', o.pts[i].r + 3); showTip(ev, o.pts[i].label, o.pts[i].tipRows, o.onClick ? 'Click to filter' : ''); });
  rect.addEventListener('pointerleave', () => { if (last >= 0) dots[last].setAttribute('r', o.pts[last].r); last = -1; hideTip(); });
  if (o.onClick) rect.addEventListener('click', ev => { const i = near(ev); if (i >= 0) o.onClick(o.pts[i]); });
  host.replaceChildren(...(o.legend ? [legendEl(o.legend)] : []), svg);
}

/* ---------- 7. Histogram ---------- */
function histogram(host, o) {
  // o: {bins:[{x0,x1,count}], vlines:[{x,label}], xFmt, height}
  const W = chartHost(host), H = o.height || 230, ml = 52, mr = 14, mt = 18, mb = 34, n = o.bins.length, pw = W - ml - mr, bw = pw / n;
  const mx = Math.max(1, ...o.bins.map(b => b.count)), ty = niceTicks(0, mx, 4), y = lin(0, ty[ty.length - 1], H - mb, mt);
  const lo = o.bins[0].x0, hi = o.bins[n - 1].x1, X = v => ml + (v - lo) / (hi - lo) * pw;
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Histogram' });
  ty.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, numF(t), { 'text-anchor': 'end' })));
  o.bins.forEach((b, i) => {
    const mid = (b.x0 + b.x1) / 2, c = b.band ? 'var(--axis)' : (mid >= 0 ? 'var(--fav)' : 'var(--unf)');
    const r = sv('path', { d: roundedBarPath(ml + i * bw + 1, y(0), y(b.count), Math.max(2, bw - 2), 3, false), fill: c, opacity: b.band ? .9 : .85 });
    const hit = sv('rect', { x: ml + i * bw, y: mt, width: bw, height: H - mb - mt, fill: 'transparent' });
    hit.addEventListener('pointermove', ev => showTip(ev, o.xFmt(b.x0) + ' to ' + o.xFmt(b.x1), [{ label: 'Lines', value: numF(b.count) }, { label: 'Share of lines', value: pctF(b.count / o.total * 100, 1) }]));
    hit.addEventListener('pointerleave', hideTip); svg.append(r, hit);
  });
  for (let t = Math.ceil(lo / 10) * 10; t <= hi; t += 10) svg.append(txt(X(t), H - mb + 15, o.xFmt(t), { 'text-anchor': 'middle' }));
  (o.vlines || []).forEach(v => svg.append(sv('line', { x1: X(v.x), x2: X(v.x), y1: mt - 4, y2: H - mb, stroke: 'var(--ink2)', 'stroke-width': 1, opacity: .6 }), txt(X(v.x), mt - 6, v.label, { 'text-anchor': 'middle', 'font-size': 10 })));
  host.replaceChildren(svg);
}

/* ---------- 8. Dumbbell (two periods) ---------- */
function dumbbell(host, o) {
  // o: {items:[{id,label,v0,v1,selected}], l0, l1, fmt, onClick}
  const W = chartHost(host), rowH = 26, mt = 22, H = mt + o.items.length * rowH + 22, labW = Math.min(190, Math.max(100, Math.max(...o.items.map(i => i.label.length)) * 6.3));
  const all = o.items.flatMap(i => [i.v0, i.v1]).filter(isFinite), lo = Math.min(...all), hi = Math.max(...all), t = niceTicks(lo - (hi - lo) * .05, hi + (hi - lo) * .05, 5);
  const x = lin(t[0], t[t.length - 1], labW + 12, W - 90);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Change between two periods' });
  t.forEach(v => svg.append(sv('line', { x1: x(v), x2: x(v), y1: mt - 6, y2: H - 18, stroke: v === 100 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(x(v), H - 4, o.fmt(v), { 'text-anchor': 'middle', 'font-size': 10.5 })));
  svg.append(txt(labW + 12, 10, 'o ' + o.l0 + '   ● ' + o.l1, { 'font-size': 11 }));
  o.items.forEach((it, k) => {
    const cy = mt + k * rowH + rowH / 2, d = it.v1 - it.v0, c = d >= 0 ? 'var(--fav)' : 'var(--unf)';
    const g = sv('g', { class: o.onClick ? 'clk' : '', opacity: o.anySel && !it.selected ? .4 : 1 });
    g.append(sv('rect', { x: 0, y: cy - rowH / 2, width: W, height: rowH, fill: 'transparent' }), txt(labW, cy + 4, it.label, { 'text-anchor': 'end', 'font-size': 12, style: 'fill:var(--ink)' }),
      sv('line', { x1: x(it.v0), x2: x(it.v1), y1: cy, y2: cy, stroke: c, 'stroke-width': 2.5, 'stroke-linecap': 'round' }),
      sv('circle', { cx: x(it.v0), cy, r: 4.5, fill: 'var(--surface)', stroke: 'var(--muted)', 'stroke-width': 2 }), sv('circle', { cx: x(it.v1), cy, r: 4.5, fill: c, stroke: 'var(--surface)', 'stroke-width': 2 }),
      txt(Math.max(x(it.v0), x(it.v1)) + 10, cy + 4, ppF(d), { 'font-size': 11, 'font-weight': 600, style: 'fill:var(--ink)' }));
    g.addEventListener('pointermove', ev => showTip(ev, it.label, [{ label: o.l0, value: o.fmt(it.v0) }, { label: o.l1, value: o.fmt(it.v1) }, { label: 'Change', value: ppF(d) }], o.onClick ? 'Click to filter' : ''));
    g.addEventListener('pointerleave', hideTip); if (o.onClick) g.addEventListener('click', () => o.onClick(it)); svg.append(g);
  });
  host.replaceChildren(svg);
}

/* ---------- 9. Single stacked bar (parts of 100 %) ---------- */
function stack100(host, segs, o = {}) {
  // segs: [{label,value,color,txt}]
  const W = chartHost(host), H = 44, total = segs.reduce((a, s) => a + s.value, 0) || 1, svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Share breakdown' });
  let x = 0;
  segs.forEach(s => {
    const w = s.value / total * W; if (w <= 0) return;
    const r = sv('rect', { x: x + 1, y: 6, width: Math.max(0, w - 2), height: 30, rx: 4, fill: s.color });
    r.addEventListener('pointermove', ev => showTip(ev, s.label, [{ label: 'Lines', value: numF(s.value) }, { label: 'Share', value: pctF(s.value / total * 100, 1) }].concat(s.extra || [])));
    r.addEventListener('pointerleave', hideTip); svg.append(r);
    const t = s.txt || pctF(s.value / total * 100, 0); if (w > t.length * 7 + 16) svg.append(txt(x + w / 2, 26, t, { 'text-anchor': 'middle', 'font-size': 12, 'font-weight': 650, style: `fill:${/mid|target|wash/.test(s.color) ? 'var(--ink)' : '#fff'}` }));
    x += w;
  });
  host.replaceChildren(svg, legendEl(segs.map(s => ({ label: `${s.label} · ${numF(s.value)}`, color: s.color, type: 'bar' }))));
}

/* ---------- 10. Pareto (bars + cumulative line, one common % axis) ---------- */
function pareto(host, o) {
  // o: {items:[{label,share,cum,selected}], onClick}
  const W = chartHost(host), H = 300, ml = 50, mr = 14, mt = 14, mb = 96, n = o.items.length, pw = W - ml - mr, band = pw / n, X = i => ml + band * (i + .5), y = lin(0, 100, H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Pareto of shortfall by product' });
  [0, 20, 40, 60, 80, 100].forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: t === 80 ? 'var(--ink2)' : 'var(--grid)', 'stroke-width': 1, opacity: t === 80 ? .6 : 1 }), txt(ml - 8, y(t) + 4, t + '%', { 'text-anchor': 'end' })));
  svg.append(txt(W - mr - 4, y(80) - 4, '80% of shortfall', { 'text-anchor': 'end', 'font-size': 10.5 }));
  const bw = Math.min(24, band * .62);
  o.items.forEach((it, i) => {
    const g = sv('g', { class: o.onClick && it.id != null ? 'clk' : '', opacity: o.anySel && !it.selected ? .4 : 1 });
    g.append(sv('rect', { x: X(i) - band / 2, y: mt, width: band, height: H - mb - mt, fill: 'transparent' }));
    g.append(sv('path', { d: roundedBarPath(X(i) - bw / 2, y(0), y(it.share), bw, 4, false), fill: 'var(--unf)' }));
    const lx = X(i), ly = H - mb + 10; g.append(sv('text', { x: lx, y: ly, 'font-size': 10.5, class: 'axis-t', transform: `rotate(-45 ${lx} ${ly})`, 'text-anchor': 'end', text: it.label }));
    g.addEventListener('pointermove', ev => showTip(ev, it.label, [{ label: 'Shortfall', value: usd(it.usd) }, { label: 'Share of total shortfall', value: pctF(it.share, 1) }, { label: 'Cumulative', value: pctF(it.cum, 1) }], o.onClick && it.id != null ? 'Click to filter' : ''));
    g.addEventListener('pointerleave', hideTip); if (o.onClick && it.id != null) g.addEventListener('click', () => o.onClick(it)); svg.append(g);
  });
  const d = o.items.map((it, i) => (i ? 'L' : 'M') + X(i) + ',' + y(it.cum)).join('');
  svg.append(sv('path', { d, fill: 'none', stroke: 'var(--ink)', 'stroke-width': 2, 'stroke-linejoin': 'round' }));
  o.items.forEach((it, i) => svg.append(sv('circle', { cx: X(i), cy: y(it.cum), r: 3.5, fill: 'var(--ink)', stroke: 'var(--surface)', 'stroke-width': 2 })));
  host.replaceChildren(legendEl([{ label: 'Share of total shortfall (bars)', color: 'var(--unf)', type: 'bar' }, { label: 'Cumulative share (line)', color: 'var(--ink)' }]), svg);
}

/* ---------- 11. Sparkline ---------- */
function spark(vals, w = 120, hgt = 32, signed = true) {
  const v = vals.filter(isFinite); if (v.length < 2) return h('span');
  const lo = Math.min(...v), hi = Math.max(...v), x = lin(0, vals.length - 1, 2, w - 6), y = lin(lo, hi === lo ? hi + 1 : hi, hgt - 4, 4);
  let d = ''; vals.forEach((q, i) => { if (isFinite(q)) d += (d ? 'L' : 'M') + x(i) + ',' + y(q); });
  const last = vals.length - 1; const svg = sv('svg', { width: w, height: hgt, viewBox: `0 0 ${w} ${hgt}`, 'aria-hidden': 'true' });
  if (signed && lo < 0 && hi > 0) svg.append(sv('line', { x1: 0, x2: w, y1: y(0), y2: y(0), stroke: 'var(--axis)', 'stroke-width': 1 }));
  svg.append(sv('path', { d, fill: 'none', stroke: 'var(--target)', 'stroke-width': 1.5, 'stroke-linejoin': 'round' }), sv('circle', { cx: x(last), cy: y(vals[last]), r: 3.5, fill: 'var(--s1)', stroke: 'var(--surface)', 'stroke-width': 1.5 }));
  return svg;
}
/* ===========================================================================
   UI: filter bar, cards, tables, theme, export
   =========================================================================== */
const tableMode = new Map();          // card id -> show table instead of chart
let pending = [];                     // draw callbacks to run once the view is attached
function card(o) {                    // {id,title,sub,span,build(body),table(),tools[]}
  const id = o.id || o.title, body = h('div', { class: 'body' });
  const toolbar = h('div', { class: 'tools' });
  const el = h('section', { class: 'card ' + (o.span || 's6'), id: 'card-' + id.replace(/\W+/g, '-') },
    h('div', { class: 'card-h' }, h('div', null, h('h3', null, o.title), o.sub ? h('div', { class: 'sub' }, o.sub) : null), toolbar), body);
  const draw = () => { body.replaceChildren(); if (o.table && tableMode.get(id)) body.append(tableView(o.table())); else o.build(body); };
  if (o.tools) o.tools.forEach(t => toolbar.append(t));
  if (o.table) {
    const b = h('button', { class: 'tool', 'aria-pressed': !!tableMode.get(id), title: 'Show the numbers behind this chart', onclick: () => { tableMode.set(id, !tableMode.get(id)); b.setAttribute('aria-pressed', !!tableMode.get(id)); draw(); } }, '⊞ Table');
    toolbar.append(b);
  }
  pending.push(draw); return el;
}
function flushPending() { const p = pending; pending = []; p.forEach(f => { try { f(); } catch (e) { console.error(e); } }); }
function seg(options, current, onPick, aria) {
  return h('div', { class: 'seg', role: 'group', 'aria-label': aria || '' }, options.map(([v, l]) => h('button', { 'aria-pressed': v === current, onclick: () => onPick(v) }, l)));
}
function emptyMsg(msg) { return h('div', { class: 'empty' }, msg || 'No sales lines match the current filters. Use Reset to clear them.'); }

/* ---------- generic sortable / searchable / paged table ---------- */
function tableView(spec) {
  // spec: {cols:[{k,label,fmt,align,csv}], rows:[], pageSize, search:true, sortKey, sortDir, onRow(row), name}
  const st = { q: '', k: spec.sortKey || null, d: spec.sortDir || -1, page: 0 }, ps = spec.pageSize || 0;
  const root = h('div'), wrap = h('div', { class: 'tbl-wrap' }), pager = h('div', { class: 'pager' });
  const q = h('input', { type: 'search', placeholder: 'Search…', 'aria-label': 'Search table', oninput: e => { st.q = e.target.value.toLowerCase(); st.page = 0; draw(); } });
  const info = h('span', { class: 'sub' });
  const exp = h('button', { class: 'tool', onclick: () => downloadCSV((spec.name || 'table') + '.csv', spec.cols.map(c => c.label), view().map(r => spec.cols.map(c => c.csv ? c.csv(r) : r[c.k]))) }, '⬇ CSV');
  const bar = h('div', { class: 'tbar' }, spec.search === false ? null : q, info, exp);
  function view() {
    let rows = spec.rows;
    if (st.q) rows = rows.filter(r => spec.cols.some(c => String(c.csv ? c.csv(r) : r[c.k]).toLowerCase().includes(st.q)));
    if (st.k) { const c = spec.cols.find(c => c.k === st.k), g = c.sortVal || (r => r[c.k]); rows = rows.slice().sort((a, b) => { const x = g(a), y = g(b); const nx = x == null || (typeof x === 'number' && !isFinite(x)), ny = y == null || (typeof y === 'number' && !isFinite(y)); if (nx || ny) return nx === ny ? 0 : nx ? 1 : -1; return (x < y ? -1 : x > y ? 1 : 0) * st.d; }); }
    return rows;
  }
  function draw() {
    const rows = view(), tot = rows.length, pages = ps ? Math.max(1, Math.ceil(tot / ps)) : 1; st.page = Math.min(st.page, pages - 1);
    const slice = ps ? rows.slice(st.page * ps, st.page * ps + ps) : rows;
    const thead = h('thead', null, h('tr', null, spec.cols.map(c => h('th', { class: c.align === 'r' ? 'r' : '', onclick: () => { if (st.k === c.k) st.d = -st.d; else { st.k = c.k; st.d = c.align === 'r' ? -1 : 1; } draw(); }, title: 'Sort by ' + c.label }, c.label + (st.k === c.k ? (st.d > 0 ? ' ▲' : ' ▼') : '')))));
    const tb = h('tbody', null, slice.map(r => { const tr = h('tr', { class: spec.onRow ? 'clk' : '', onclick: spec.onRow ? () => spec.onRow(r) : null }, spec.cols.map(c => { const v = r[c.k], cell = c.fmt ? c.fmt(v, r) : v; return h('td', { class: c.align === 'r' ? 'r' : '' }, cell == null ? '' : cell); })); return tr; }));
    wrap.replaceChildren(h('table', { class: 't' }, thead, tb));
    info.textContent = `${numF(tot)} row${tot === 1 ? '' : 's'}` + (st.q ? ' (filtered)' : '');
    pager.replaceChildren();
    if (ps && pages > 1) pager.append(h('button', { class: 'tool', disabled: st.page === 0, onclick: () => { st.page--; draw(); } }, '‹ Prev'), `Page ${st.page + 1} of ${pages}`, h('button', { class: 'tool', disabled: st.page >= pages - 1, onclick: () => { st.page++; draw(); } }, 'Next ›'));
  }
  root.append(bar, wrap, pager); draw(); return root;
}
function downloadCSV(name, head, rows) {
  const q = v => { v = v == null ? '' : String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const text = [head.map(q).join(','), ...rows.map(r => r.map(q).join(','))].join('\r\n');
  const blob = new Blob(['\uFEFF' + text], { type: 'text/csv;charset=utf-8' }), a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

/* ---------- filter helpers ---------- */
function toggleSet(set, v) { set.has(v) ? set.delete(v) : set.add(v); }
function change(fn) { fn(); if (S.m0 > S.m1) [S.m0, S.m1] = [S.m1, S.m0]; bump(); renderAll(); }
function resetFilters() { change(() => { S.years.clear(); S.qtrs.clear(); S.regs.clear(); S.cats.clear(); S.prods.clear(); S.stat.clear(); S.mat = 'all'; S.dq = 'all'; S.m0 = 0; S.m1 = NM - 1; }); }
function selectMonth(m) { change(() => { S.years.clear(); S.qtrs.clear(); if (S.m0 === m && S.m1 === m) { S.m0 = 0; S.m1 = NM - 1; } else { S.m0 = S.m1 = m; } }); }
function pickRegion(r, ev) { change(() => { if (ev && (ev.shiftKey || ev.ctrlKey || ev.metaKey)) toggleSet(S.regs, r); else if (S.regs.size === 1 && S.regs.has(r)) S.regs.clear(); else { S.regs.clear(); S.regs.add(r); } }); }
function pickCat(c, ev) { change(() => { if (ev && (ev.shiftKey || ev.ctrlKey || ev.metaKey)) toggleSet(S.cats, c); else if (S.cats.size === 1 && S.cats.has(c)) S.cats.clear(); else { S.cats.clear(); S.cats.add(c); } }); }
function pickProd(p, ev) { change(() => { if (ev && (ev.shiftKey || ev.ctrlKey || ev.metaKey)) toggleSet(S.prods, p); else if (S.prods.size === 1 && S.prods.has(p)) S.prods.clear(); else { S.prods.clear(); S.prods.add(p); } }); }
function isFiltered() { return S.years.size || S.qtrs.size || S.regs.size || S.cats.size || S.prods.size || S.stat.size || S.mat !== 'all' || S.dq !== 'all' || S.m0 !== 0 || S.m1 !== NM - 1; }

let FILT_OPEN = true;
function buildFilters() {
  const root = $('#filters'); root.replaceChildren();
  const chip = (label, on, fn, dot) => h('button', { class: 'chip', 'aria-pressed': on, onclick: fn }, dot ? h('span', { class: 'dot', style: { background: dot } }) : null, label);
  const fg = (lab, ...kids) => h('div', { class: 'fg' }, h('div', { class: 'lbl' }, lab), ...kids);
  const monthSel = (which) => { const s = h('select', { class: 'sel', 'aria-label': which === 0 ? 'From month' : 'To month', onchange: e => change(() => { which === 0 ? S.m0 = +e.target.value : S.m1 = +e.target.value; S.years.clear(); S.qtrs.clear(); }) }, MON.map((m, i) => h('option', { value: i }, MLONG[i]))); s.value = which === 0 ? S.m0 : S.m1; return s; };
  const prodSel = h('select', { class: 'sel', 'aria-label': 'Product', onchange: e => change(() => { S.prods.clear(); if (e.target.value !== '') S.prods.add(+e.target.value); }) }, h('option', { value: '' }, 'All products (' + NP + ')'));
  CAT.forEach((c, ci) => { const og = h('optgroup', { label: c }); PRD.forEach((p, pi) => { if (DIM.prodCat[pi] === ci) og.append(h('option', { value: pi }, p)); }); prodSel.append(og); });
  prodSel.value = S.prods.size === 1 ? [...S.prods][0] : '';
  if (S.prods.size > 1) prodSel.firstChild.textContent = S.prods.size + ' products selected';
  const frow = h('div', { class: 'frow' },
    fg('Year', h('div', { class: 'chips' }, YEARS.map(y => chip(String(y), S.years.has(y), () => change(() => { toggleSet(S.years, y); S.m0 = 0; S.m1 = NM - 1; }))))),
    fg('Quarter', h('div', { class: 'chips' }, [1, 2, 3, 4].map(q => chip('Q' + q, S.qtrs.has(q), () => change(() => { toggleSet(S.qtrs, q); S.m0 = 0; S.m1 = NM - 1; }))))),
    fg('From', monthSel(0)), fg('To', monthSel(1)),
    fg('Region', h('div', { class: 'chips' }, REG.map((r, i) => chip(r, S.regs.has(i), () => change(() => toggleSet(S.regs, i)), `var(--s${i + 1})`)))),
    fg('Category', h('div', { class: 'chips' }, CAT.map((c, i) => chip(c, S.cats.has(i), () => change(() => toggleSet(S.cats, i)), `var(--s${i + 1})`)))),
    fg('Product', prodSel),
    fg('Line status', h('div', { class: 'chips' }, [0, 1, 2].map(s => chip(ST_GLYPH[s] + ' ' + ST_NAME[s], S.stat.has(s), () => change(() => toggleSet(S.stat, s)))))),
    fg('Material', seg([['all', 'All'], ['yes', 'Material'], ['no', 'Not material']], S.mat, v => change(() => { S.mat = v; }), 'Material filter')),
    fg('Data quality', seg([['all', 'All lines'], ['repaired', 'Qty repaired'], ['clean', 'Untouched']], S.dq, v => change(() => { S.dq = v; }), 'Data quality filter')),
    h('div', { class: 'fg' }, h('div', { class: 'lbl' }, ' '), h('button', { class: 'btn', onclick: resetFilters, disabled: !isFiltered(), title: 'Clear every filter' }, '↺ Reset')));
  root.classList.toggle('open', FILT_OPEN); if (FILT_OPEN) root.append(frow);
  // active filter tags + live count
  const tags = h('div', { class: 'active' });
  const tag = (label, fn) => tags.append(h('span', { class: 'tag' }, label, h('button', { 'aria-label': 'Remove ' + label, onclick: fn }, '×')));
  S.years.forEach(y => tag('Year ' + y, () => change(() => S.years.delete(y)))); S.qtrs.forEach(q => tag('Q' + q, () => change(() => S.qtrs.delete(q))));
  if (S.m0 !== 0 || S.m1 !== NM - 1) tag(S.m0 === S.m1 ? MLONG[S.m0] : MLONG[S.m0] + ' → ' + MLONG[S.m1], () => change(() => { S.m0 = 0; S.m1 = NM - 1; }));
  S.regs.forEach(r => tag(REG[r], () => change(() => S.regs.delete(r)))); S.cats.forEach(c => tag(CAT[c], () => change(() => S.cats.delete(c)))); S.prods.forEach(p => tag(PRD[p], () => change(() => S.prods.delete(p))));
  S.stat.forEach(s => tag('Lines: ' + ST_NAME[s], () => change(() => S.stat.delete(s)))); if (S.mat !== 'all') tag(S.mat === 'yes' ? 'Material lines only' : 'Non-material lines only', () => change(() => { S.mat = 'all'; }));
  if (S.dq !== 'all') tag(S.dq === 'repaired' ? 'Quantity-repaired lines' : 'Untouched lines', () => change(() => { S.dq = 'all'; }));
  tags.prepend(h('button', { class: 'ftoggle', 'aria-expanded': FILT_OPEN, title: FILT_OPEN ? 'Collapse the filters into a compact bar that stays on screen' : 'Show all filters', onclick: () => { FILT_OPEN = !FILT_OPEN; buildFilters(); } }, FILT_OPEN ? '▴ Collapse filters' : '▾ Show filters'));
  if (!FILT_OPEN && isFiltered()) tags.append(h('button', { class: 'ftoggle', onclick: resetFilters }, '↺ Clear all'));
  const cnt = maskCount(getMask());
  tags.append(h('span', { style: { marginLeft: 'auto' }, id: 'cnt' }, `${numF(cnt)} of ${numF(N)} sales lines in view` + (isFiltered() ? '' : ' (no filters)')));
  root.append(tags);
}

/* ---------- theme ---------- */
function setTheme(t) { if (t) document.documentElement.setAttribute('data-theme', t); else document.documentElement.removeAttribute('data-theme'); try { localStorage.setItem('svd-theme', t || ''); } catch (e) { } }
function curTheme() { return document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); }
/* ===========================================================================
   SHARED VIEW HELPERS
   =========================================================================== */
const sgnCls = v => (v > 0 ? 'pos' : v < 0 ? 'neg' : '');
function vSpan(v, txt) { return h('span', { class: sgnCls(v) }, txt); }
function monthSeries(mask) { const g = agg(A.M, NM, mask); return Array.from({ length: NM }, (_, i) => dv(g, i)); }
function selTracking(ser, inSel) {          // tracking signal over the months currently in view
  let cum = 0, abs = 0, n = 0; for (let m = 0; m < NM; m++) if (inSel[m] && ser[m].n) { cum += ser[m].vc; abs += Math.abs(ser[m].vc); n++; }
  return abs > 0 ? cum / (abs / n) : NaN;
}
function groupRows(key, K, mask, names) { const g = agg(key, K, mask); return names.map((nm, i) => Object.assign(dv(g, i), { name: nm, id: i })); }
function statusCell(s) { return stBadge(s); }
function matCell(m, trig) { return m ? h('span', { class: 'mat' }, '◆ Yes' + (trig ? ' (' + trig + ')' : '')) : 'No'; }
function tipFor(d, name) {
  return { title: name, rows: [
    { label: 'Target', value: usd(d.target) }, { label: 'Actual', value: usd(d.actual) },
    { label: 'Variance', value: usd(d.v) + '  (' + pctF(d.pct, 2, true) + ')' }, { label: 'Qty attainment', value: pctF(d.att) },
    { label: 'Status (group total)', value: ST_GLYPH[d.st] + ' ' + ST_NAME[d.st] + (d.mat ? '  ◆ material' : '') },
    { label: 'Lines below target', value: pctF(d.pu, 1) + ' of ' + numF(d.n) }] };
}
function itemOf(d, name, id, selected, extra) {
  return Object.assign({ id, label: name, value: d.v, txt: usdS(d.v) + '  ·  ' + pctF(d.pct, 2, true), selected, d }, extra || {});
}

/* ---------- live insight sentences for the current selection ---------- */
function liveInsights() {
  const m = getMask(), T = totals(m), out = [];
  if (!T.n) return [['warn', 'No sales lines match the current filters.']];
  const inSel = monthOk(), ser = monthSeries(m);
  out.push([T.v < 0 ? 'bad' : '', `Sales are ${usdC(Math.abs(T.v))} ${T.v < 0 ? 'below' : 'above'} target (${pctF(T.pct, 2, true)}): ${usdC(T.actual)} actual vs ${usdC(T.target)} planned, with quantity attainment of ${pctF(T.att)}. Group status: ${ST_NAME[T.st]}${T.mat ? ', material' : ''}.`]);
  const dims = [['region', A.R, NR, REG, 'reg'], ['category', A.C, NC, CAT, 'cat']];
  dims.forEach(([lab, key, K, names, ex]) => {
    const rows = groupRows(key, K, getMask(ex), names).filter(r => r.n); if (rows.length < 2) return;
    const worst = rows.reduce((a, b) => (b.v < a.v ? b : a)), best = rows.reduce((a, b) => (b.v > a.v ? b : a)), net = rows.reduce((s, r) => s + r.v, 0);
    let t = `Weakest ${lab}: ${worst.name} (${usdC(worst.v)}, ${pctF(worst.pct, 2, true)})`;
    if (net < 0 && worst.v < 0) t += `, equal to ${Math.round(worst.v / net * 100)}% of the net variance`;
    t += `. Strongest: ${best.name} (${usdS(best.v)}, ${pctF(best.pct, 2, true)}).`;
    out.push([worst.pct < -5 ? 'bad' : worst.v < 0 ? 'warn' : '', t]);
  });
  const nSel = inSel.reduce((a, b) => a + b, 0), below = ser.filter((d, i) => inSel[i] && d.n && d.vc < 0).length, present = ser.filter((d, i) => inSel[i] && d.n).length;
  if (present) { const worstM = ser.reduce((a, b, i) => (inSel[i] && b.n && (a === null || b.pct < ser[a].pct) ? i : a), null); out.push([below === present && present > 3 ? 'bad' : '', `${below} of ${present} month${present > 1 ? 's' : ''} in view finished below target` + (worstM != null ? `; the weakest was ${MLONG[worstM]} (${pctF(ser[worstM].pct, 2, true)}).` : '.')]); }
  if (Math.abs(T.prc) < 0.5) out.push(['warn', `The price effect is ${usd(T.prc, 2)}: all of the ${usdC(T.v)} variance is volume (${numF(T.qv)} units vs plan). Actual unit prices equal the planned price on every line, so price/mix cannot be measured from this data.`]);
  else out.push(['', `Driver split: volume ${usdC(T.pure)}, mix ${usdC(T.mix)}, price ${usdC(T.prc)}, rounding ${usd(T.res, 2)}.`]);
  const pr = groupRows(A.P, NP, getMask('prod'), PRD).filter(r => r.n), neg = pr.filter(r => r.v < 0), gross = neg.reduce((s, r) => s + r.v, 0);
  if (neg.length) { const top = neg.slice().sort((a, b) => a.v - b.v).slice(0, 3); out.push(['', `${neg.length} of ${pr.length} products are below target. The three largest shortfalls (${top.map(r => r.name).join(', ')}) explain ${Math.round(top.reduce((s, r) => s + r.v, 0) / gross * 100)}% of the product-level shortfall.`]); }
  out.push(['', `${numF(T.nm)} of ${numF(T.n)} lines (${pctF(T.nm / T.n * 100, 0)}) are flagged material at line level, so line flags are noisy; use the Exceptions tab (region × category × month) for triage. Volume shortfall cost about ${usdC(Math.abs(T.mls))} of gross margin.`]);
  return out;
}
function insightsList(items) { return h('ul', { class: 'ins' }, items.map(([c, t]) => h('li', { class: c }, t))); }


/* ---------- key takeaway banner (live: follows the filters) ---------- */
function takeawayBanner(T, m, mSer, inSel) {
  const b = t => h('strong', null, t), sp = (c, t) => h('strong', { class: c }, t);
  const miss = T.v < 0, c = miss ? 'neg' : 'pos';
  const present = mSer.filter((d, i) => inSel[i] && d.n), below = present.filter(d => d.vc < 0).length;
  const s1 = [miss ? (T.pct > -3 ? 'A modest miss: ' : T.pct > -6 ? 'A clear miss: ' : 'A large miss: ') + 'sales finished ' : 'Ahead of plan: sales finished ', sp(c, usdC(Math.abs(T.v)) + ' (' + pctF(Math.abs(T.pct), 2) + ')'), miss ? ' below target' : ' above target',
    present.length > 2 ? (below === present.length ? ', and every one of the ' + present.length + ' months in view was under plan.' : below === 0 ? ', and no month in view finished below plan.' : ', with ' + below + ' of ' + present.length + ' months below plan.') : '.'];
  const reg = groupRows(A.R, NR, m, REG).filter(r => r.n), out = [h('p', null, ...s1)];
  if (reg.length >= 2) {
    const worst = reg.reduce((a, x) => (x.v < a.v ? x : a)), best = reg.reduce((a, x) => (x.v > a.v ? x : a)), net = reg.reduce((s, r) => s + r.v, 0), pos = reg.filter(r => r.v > 0);
    const p2 = [];
    if (worst.v < 0) { p2.push(sp('neg', worst.name), ' is the main problem (', sp('neg', pctF(worst.pct, 2, true)), net < 0 ? ', ' + Math.round(worst.v / net * 100) + '% of the net gap)' : ')'); }
    if (best.v > 0) { p2.push(worst.v < 0 ? '; ' : '', sp('pos', best.name), ' (', sp('pos', pctF(best.pct, 2, true)), pos.length === 1 ? ') is the only region above target.' : ') is the strongest region.'); }
    else p2.push('.');
    if (p2.length) { if (typeof p2[p2.length - 1] === 'string' && !/[.]$/.test(p2[p2.length - 1])) p2.push('.'); out.push(h('p', null, ...p2)); }
  }
  const pr = groupRows(A.P, NP, m, PRD).filter(r => r.n), up = pr.filter(r => r.v > 0).length;
  const s3 = [];
  if (Math.abs(T.prc) < 0.5 && T.v !== 0) s3.push('The variance is ', b('100% volume-driven'), ' (price effect ' + usd(T.prc, 0) + ')');
  else if (T.v !== 0) s3.push('Volume explains ', b(usdC(T.pure)), ' of it; price ' + usdC(T.prc) + ', mix ' + usdC(T.mix));
  if (pr.length > 1) { const w = up === pr.length ? ['all ', b(String(pr.length)), ' products beat target.'] : up === 0 ? ['none of the ', b(String(pr.length)), ' products beat target.'] : ['only ', b(up + ' of ' + pr.length), ' products beat target.']; if (s3.length) s3.push(', and ', ...w); else s3.push(w[0].charAt(0).toUpperCase() + w[0].slice(1), ...w.slice(1)); }
  else if (s3.length) s3.push('.');
  if (s3.length) out.push(h('p', null, ...s3));
  const right = h('div', { class: 'tk-r' },
    h('div', { class: 'stat' }, h('b', { class: c }, pctF(T.pct, 2, true)), h('span', null, 'vs target')),
    present.length ? h('div', { class: 'stat' }, h('b', { class: below === present.length ? 'neg' : '' }, below + ' / ' + present.length), h('span', null, 'months below target')) : null,
    pr.length > 1 ? h('div', { class: 'stat' }, h('b', null, up + ' / ' + pr.length), h('span', null, 'products above target')) : null);
  return h('section', { class: 'takeaway ' + (miss ? '' : 'pos'), 'aria-label': 'Key takeaway' }, h('div', { class: 'tk-l' }, h('div', { class: 'tk-k' }, isFiltered() ? 'Key takeaway · current filters' : 'Key takeaway'), ...out), right);
}

/* ===========================================================================
   TAB 1 - OVERVIEW
   =========================================================================== */
function tabOverview(root) {
  const m = getMask(), T = totals(m);
  if (!T.n) { root.append(emptyMsg()); return; }
  const mT = getMask('time'), ser = monthSeries(mT), inSel = monthOk(), ts = selTracking(monthSeries(m), inSel);
  const mSer = monthSeries(m);
  const cells = (() => { const g = agg(KEY_RCM, NR * NC * NM, m); let c = 0, mc = 0; for (let k = 0; k < g.K; k++) if (g.n[k]) { c++; if (isMat(g.ac[k] - g.tc[k], g.tc[k])) mc++; } return [mc, c]; })();
  const monthsBelow = mSer.filter((d, i) => inSel[i] && d.n && d.vc < 0).length, monthsPresent = mSer.filter((d, i) => inSel[i] && d.n).length;
  root.append(takeawayBanner(T, m, mSer, inSel));
  // hero
  const hero = h('section', { class: 'card s4 hero ' + (T.v < 0 ? '' : 'pos-t') },
    h('div', null, h('div', { class: 'sub', style: { fontWeight: 600 } }, 'SALES VARIANCE vs TARGET'),
      h('div', { class: 'big ' + sgnCls(T.v) }, usdC(T.v)),
      h('div', { class: 'row' }, h('span', { class: 'sgn ' + sgnCls(T.v), style: { fontSize: '20px', fontWeight: 650 } }, pctF(T.pct, 2, true)), stBadge(T.st), T.mat ? h('span', { class: 'mat' }, '◆ Material') : null)),
    h('div', { class: 'sub' }, `${usd(T.v)} on a target of ${usd(T.target)}. Favorable above +${CFG.favPct}%, Unfavorable below ${MINUS}${Math.abs(CFG.unfPct)}%; material at ±${CFG.matPctV}% or ${usd(CFG.matUsd)}.`),
    h('div', null, spark(ser.map(d => d.v), 220, 42), h('div', { class: 'sub' }, 'Monthly variance, all months (current selection of other filters)')));
  const tile = (lab, val, dl, cls, pri) => h('div', { class: 'card kpi' + (pri ? ' pri ' + (cls === 'neg' ? 'neg-t' : cls === 'pos' ? 'pos-t' : '') : '') }, h('div', { class: 'lab' }, lab), h('div', { class: 'val ' + (cls || '') }, val), h('div', { class: 'dl' }, dl));
  const tiles = h('div', { class: 's8 kpi-wrap' },
    h('div', { class: 'kpi-pri' },
      tile('Actual sales', usdC(T.actual), `${pctF(T.actual / T.target * 100)} of target`, '', true),
      tile('Quantity attainment', pctF(T.att), `${numF(T.qv)} units vs plan`, sgnCls(T.qv), true),
      tile('Volume effect', usdC(T.vol), `pure ${usdC(T.pure)} · mix ${usdC(T.mix)}`, sgnCls(T.vol), true),
      tile('Price effect', usdC(T.prc), Math.abs(T.prc) < .5 ? 'zero – prices equal plan' : 'actual vs plan price × qty', sgnCls(T.prc), true)),
    h('div', { class: 'kpi-sec' },
      tile('Target sales', usdC(T.target), `${numF(T.n)} lines · ${numF(T.tq)} units`),
      tile('Gross margin impact', usdC(T.mls), `at ${pctF(T.gmp, 1)} gross margin`, sgnCls(T.mls)),
      tile('Lines below target', pctF(T.pu, 1), `${pctF(T.nf / T.n * 100, 1)} favorable · ${pctF(T.no / T.n * 100, 1)} on target`),
      tile('Material lines', numF(T.nm), `${pctF(T.nm / T.n * 100, 1)} of lines · ${cells[0]} of ${cells[1]} cells`),
      tile('Months below target', `${monthsBelow} of ${monthsPresent}`, 'months in view with sales', monthsBelow === monthsPresent && monthsPresent > 3 ? 'neg' : ''),
      tile('Tracking signal', isFinite(ts) ? numF(ts, 1) : '—', `limit ±${TS_LIMIT} (flags systematic bias)`, isFinite(ts) && Math.abs(ts) > TS_LIMIT ? 'neg' : '')));
  root.append(h('div', { class: 'grid' }, hero, tiles));
  const grid = h('div', { class: 'grid', style: { marginTop: '14px' } }); root.append(grid);

  grid.append(card({ id: 'ov-ins', title: 'What the data says right now', sub: 'Updates with every filter', span: 's4', build: b => b.append(insightsList(liveInsights())) }));
  grid.append(card({
    id: 'ov-trend', title: 'Target vs actual, month by month', sub: 'Click a month to filter the whole dashboard · dimmed months are outside the current time filter', span: 's8',
    build: b => trendChart(b, { m: [...Array(NM).keys()], tc: ser.map(d => d.tc), ac: ser.map(d => d.ac), vc: ser.map(d => d.vc), pct: ser.map(d => d.pct), att: ser.map(d => d.att), st: ser.map(d => d.st), mat: ser.map(d => d.mat), inSel: Array.from(inSel), onMonth: selectMonth }),
    table: () => ({ name: 'monthly_target_vs_actual', cols: monthCols(), rows: ser.map((d, i) => Object.assign({ month: MLONG[i] }, d)), sortKey: null })
  }));
  const regItems = () => { const r = groupRows(A.R, NR, getMask('reg'), REG).filter(x => x.n).sort((a, b) => a.v - b.v); return r.map(d => itemOf(d, d.name, d.id, S.regs.has(d.id))); };
  const catItems = () => { const r = groupRows(A.C, NC, getMask('cat'), CAT).filter(x => x.n).sort((a, b) => a.v - b.v); return r.map(d => itemOf(d, d.name, d.id, S.cats.has(d.id))); };
  grid.append(card({ id: 'ov-reg', title: 'Variance by region', sub: 'Blue = above target, red = below · click to filter (shift-click for several)', span: 's4',
    build: b => hbars(b, regItems(), { onClick: (it, ev) => pickRegion(it.id, ev), tip: it => tipFor(it.d, it.label), aria: 'Sales variance by region' }), table: () => dimTable(regItems(), 'Region') }));
  grid.append(card({ id: 'ov-cat', title: 'Variance by category', sub: 'Click to filter (shift-click for several)', span: 's4',
    build: b => hbars(b, catItems(), { onClick: (it, ev) => pickCat(it.id, ev), tip: it => tipFor(it.d, it.label), aria: 'Sales variance by category' }), table: () => dimTable(catItems(), 'Category') }));
  grid.append(card({ id: 'ov-status', title: 'Status of individual sales lines', sub: `Each line classified vs target: Favorable > +${CFG.favPct}%, Unfavorable < ${MINUS}${Math.abs(CFG.unfPct)}%`, span: 's4', build: b => {
    const sh = h('div'); b.append(sh);
    stack100(sh, [
      { label: 'Favorable', value: T.nf, color: 'var(--fav)', extra: [{ label: 'Variance $', value: usd(lineSum(m, 0)) }] }, { label: 'On Target', value: T.no, color: 'var(--axis)', extra: [{ label: 'Variance $', value: usd(lineSum(m, 1)) }] },
      { label: 'Unfavorable', value: T.nu, color: 'var(--unf)', extra: [{ label: 'Variance $', value: usd(lineSum(m, 2)) }] }]);
    b.append(h('div', { class: 'note' }, `Lines swing widely on their own (typical ±${lineSd(m).toFixed(0)}%), so line counts are noisy. The totals above are what matter: ${T.nu > T.nf ? 'more lines miss than beat target' : 'more lines beat than miss target'} (${numF(T.nu)} vs ${numF(T.nf)}).`));
  } }));
  const bd = bridgeDim(); const bm = getMask(bd.ex);
  const gb = groupRows(bd.key, bd.K, bm, bd.names).filter(x => x.n).sort((a, b) => b.v - a.v);
  const steps = [{ label: 'Target', value: T.target, kind: 'total', color: 'var(--target)' }].concat(gb.map(d => ({ label: d.name, value: d.v, kind: 'delta', extra: [{ label: 'Variance %', value: pctF(d.pct, 2, true) }] })), [{ label: 'Actual', value: T.actual, kind: 'total', color: 'var(--s1)' }]);
  grid.append(card({ id: 'ov-bridge', title: `Target-to-actual bridge by ${bd.label}`, sub: 'How each ' + bd.label + ' moves the company from target to actual', span: 's8', build: b => waterfall(b, steps, { aria: 'Bridge from target to actual by ' + bd.label }),
    table: () => ({ name: 'bridge', cols: [{ k: 'label', label: bd.label }, { k: 'value', label: 'USD', align: 'r', fmt: v => usd(v) }], rows: steps, search: false }) }));
  grid.append(card({ id: 'ov-hint', title: 'How to read this dashboard', sub: '', span: 's4', build: b => b.append(h('ul', { class: 'ins' },
    h('li', null, `Every figure is computed in your browser from the ${numF(N)} sales lines, in exact cents.`),
    h('li', null, 'Click any bar, month, heat-map cell or table row to filter everything. Shift-click adds to a selection.'),
    h('li', null, 'Group statuses are re-derived from the group totals, never averaged from line percentages.'),
    h('li', null, 'Use ⊞ Table on any chart to see the exact numbers; ⬇ Export saves the filtered lines.'))) }));
}
const TS_LIMIT = 4;
function lineSum(m, s) { let t = 0; for (let i = 0; i < N; i++) if (m[i] && A.ST[i] === s) t += A.AC[i] - A.TC[i]; return t / 100; }
function lineSd(m) { let n = 0, s = 0, q = 0; for (let i = 0; i < N; i++) if (m[i] && A.TC[i] > 0) { const p = (A.AC[i] - A.TC[i]) / A.TC[i] * 100; n++; s += p; q += p * p; } return n > 1 ? Math.sqrt((q - s * s / n) / (n - 1)) : 0; }
function bridgeDim() {
  if (S.regs.size === 1 && S.cats.size === 1) return { key: A.P, K: NP, names: PRD, ex: 'prod', label: 'product' };
  if (S.regs.size === 1) return { key: A.C, K: NC, names: CAT, ex: 'cat', label: 'category' };
  return { key: A.R, K: NR, names: REG, ex: 'reg', label: 'region' };
}
function dimTable(items, label) {
  return { name: label.toLowerCase() + '_variance', cols: [{ k: 'label', label }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Variance $', align: 'r', fmt: (v) => vSpan(v, usd(v)) }, { k: 'pct', label: 'Variance %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'att', label: 'Qty attainment', align: 'r', fmt: v => pctF(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }], rows: items.map(i => Object.assign({ label: i.label }, i.d)), search: false, sortKey: 'v', sortDir: 1 };
}
function monthCols() {
  return [{ k: 'month', label: 'Month' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Variance %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) },
    { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) }, { k: 'mom', label: 'MoM sales', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'yoy', label: 'YoY sales', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'ytdV', label: 'YTD var $', align: 'r', fmt: v => usd(v) },
    { k: 'r3', label: 'Rolling 3M var $', align: 'r', fmt: v => usd(v) }, { k: 'cum', label: 'Cumulative var $', align: 'r', fmt: v => usd(v) }, { k: 'ts', label: 'Tracking signal', align: 'r', fmt: v => numF(v, 1) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }];
}
/* ===========================================================================
   TAB 2 - DRIVERS (price / volume / mix, margin, what-if)
   =========================================================================== */
const WI = { region: null, pct: null };
function tabDrivers(root) {
  const m = getMask(), T = totals(m);
  if (!T.n) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const steps = [
    { label: 'Target', value: T.target, kind: 'total', color: 'var(--target)' },
    { label: 'Volume (pure)', value: T.pure, kind: 'delta', foot: '(Actual qty – target qty) × average target price of the group' },
    { label: 'Mix', value: T.mix, kind: 'delta', foot: 'Effect of selling a different product / region mix than planned' },
    { label: 'Price', value: T.prc, kind: 'delta', foot: '(Actual unit price – target price) × actual qty' },
    { label: 'Rounding', value: T.res, kind: 'delta', foot: 'Cent rounding of Actual_Sales_USD vs qty × price' },
    { label: 'Actual', value: T.actual, kind: 'total', color: 'var(--s1)' }];
  const rec = Math.abs(T.target + T.pure + T.mix + T.prc + T.res - T.actual);
  grid.append(card({ id: 'dr-pvm', title: 'Price-volume-mix bridge', sub: 'Why actual sales differ from target (current filters)', span: 's7', build: b => {
    waterfall(b, steps, { aria: 'Price volume mix bridge' });
    b.append(h('div', { class: 'note' }, Math.abs(T.prc) < .5 ? `Price effect is ${usd(T.prc, 2)}: every line sold at its planned price, so the whole variance is volume and mix. This is a property of the data (the price column looks like a standard price), not proof that pricing was perfect.` : `Volume ${usdC(T.pure)}, mix ${usdC(T.mix)}, price ${usdC(T.prc)}.`,
      ` Bridge check: target + effects – actual = ${usd(rec, 2)}.`));
  }, table: () => ({ name: 'pvm_bridge', cols: [{ k: 'label', label: 'Step' }, { k: 'value', label: 'USD', align: 'r', fmt: v => usd(v, 2) }], rows: steps, search: false }) }));
  grid.append(card({ id: 'dr-def', title: 'How the bridge is built', span: 's5', build: b => b.append(h('ul', { class: 'ins' },
    h('li', null, h('strong', null, 'Volume (pure): '), 'units sold vs plan, valued at the group’s average target price.'),
    h('li', null, h('strong', null, 'Mix: '), 'the extra (or lost) value from selling a different product/region mix than planned. Zero for a single product.'),
    h('li', null, h('strong', null, 'Price: '), '(actual unit price – target price) × actual units.'),
    h('li', null, h('strong', null, 'Rounding: '), 'sub-cent differences; should be about $0.'),
    h('li', { class: Math.abs(T.prc) < .5 ? 'warn' : '' }, Math.abs(T.prc) < .5 ? 'Here price = 0 on every line: the dataset records one price per line equal to plan.' : 'Price variance is present in this selection.'))) }));
  // by region table (exclude region filter so all regions stay visible)
  const rr = groupRows(A.R, NR, getMask('reg'), REG).filter(d => d.n);
  grid.append(card({ id: 'dr-reg', title: 'Drivers by region', sub: 'Click a row to filter', span: 's12', build: b => b.append(tableView({
    name: 'drivers_by_region', search: false, sortKey: 'v', sortDir: 1, onRow: r => pickRegion(r.id), rows: rr,
    cols: [{ k: 'name', label: 'Region' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'pure', label: 'Volume (pure) $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'mix', label: 'Mix $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'prc', label: 'Price $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'res', label: 'Rounding $', align: 'r', fmt: v => usd(v, 2) },
      { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Variance %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'qv', label: 'Units vs plan', align: 'r', fmt: v => vSpan(v, numF(v)) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }] })) }));
  const mItems = rr.map(d => ({ id: d.id, label: d.name, value: d.mls, txt: usdS(d.mls), selected: S.regs.has(d.id), d })).sort((a, b) => a.value - b.value);
  grid.append(card({ id: 'dr-gm', title: 'Gross margin lost to volume', sub: 'Quantity variance × unit margin (price – cost), per line', span: 's6', build: b => hbars(b, mItems, { onClick: (it, ev) => pickRegion(it.id, ev), tip: it => ({ title: it.label, rows: [{ label: 'Gross margin impact', value: usd(it.d.mls) }, { label: 'Sales variance', value: usd(it.d.v) }, { label: 'Gross margin %', value: pctF(it.d.gmp, 1) }] }), aria: 'Gross margin impact by region' }) }));
  const uItems = rr.map(d => ({ id: d.id, label: d.name, value: d.qv, txt: (d.qv > 0 ? '+' : '') + numF(d.qv) + ' units', selected: S.regs.has(d.id), d })).sort((a, b) => a.value - b.value);
  grid.append(card({ id: 'dr-units', title: 'Units vs plan by region', sub: 'Actual minus target quantity', span: 's6', build: b => hbars(b, uItems, { tickFmt: v => numF(v), onClick: (it, ev) => pickRegion(it.id, ev), tip: it => ({ title: it.label, rows: [{ label: 'Planned units', value: numF(it.d.tq) }, { label: 'Actual units', value: numF(it.d.aq) }, { label: 'Attainment', value: pctF(it.d.att) }] }), aria: 'Units versus plan by region' }) }));
  grid.append(card({ id: 'dr-whatif', title: 'What-if: close the gap in one region', sub: 'Move a region’s variance % and see the effect on the company (uses all current filters except region)', span: 's12', build: b => whatIf(b) }));
}
function whatIf(b) {
  const rows = groupRows(A.R, NR, getMask('reg'), REG).filter(d => d.n); if (!rows.length) { b.append(emptyMsg()); return; }
  if (WI.region == null || !rows.find(r => r.id === WI.region)) { WI.region = rows.reduce((a, c) => (c.pct < a.pct ? c : a)).id; WI.pct = null; }
  const r = rows.find(x => x.id === WI.region), rest = rows.filter(x => x.id !== r.id);
  const restT = rest.reduce((s, x) => s + x.tc, 0), restV = rest.reduce((s, x) => s + x.vc, 0), restPct = restT > 0 ? restV / restT * 100 : 0;
  const tot = rows.reduce((o, x) => ({ tc: o.tc + x.tc, vc: o.vc + x.vc }), { tc: 0, vc: 0 });
  if (WI.pct == null) WI.pct = Math.round(Math.max(r.pct, restPct) * 10) / 10;
  const lo = Math.floor(Math.min(r.pct, -12)), hi = Math.ceil(Math.max(r.pct, 6));
  const out = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(190px,1fr))', gap: '14px', marginTop: '10px' } });
  const sel = h('select', { class: 'sel', 'aria-label': 'Region to simulate', onchange: e => { WI.region = +e.target.value; WI.pct = null; b.replaceChildren(); whatIf(b); } }, rows.map(x => h('option', { value: x.id }, x.name)));
  sel.value = r.id;
  const val = h('strong', null, pctF(WI.pct, 1, true));
  const sl = h('input', { type: 'range', min: lo, max: hi, step: 0.1, value: WI.pct, 'aria-label': 'Target variance percent for the region', oninput: e => { WI.pct = +e.target.value; val.textContent = pctF(WI.pct, 1, true); paint(); } });
  function paint() {
    const revGain = (WI.pct - r.pct) / 100 * r.target, gmGain = revGain * (isFinite(r.gmp) ? r.gmp / 100 : 0), newV = tot.vc / 100 + revGain, newPct = newV / (tot.tc / 100) * 100;
    const box = (l, v, d, c) => h('div', { class: 'card kpi' }, h('div', { class: 'lab' }, l), h('div', { class: 'val ' + (c || '') }, v), h('div', { class: 'dl' }, d));
    out.replaceChildren(box('Revenue recovered', usdC(revGain), `${r.name}: ${pctF(r.pct, 2, true)} → ${pctF(WI.pct, 1, true)}`, sgnCls(revGain)), box('Gross margin recovered', usdC(gmGain), `at ${pctF(r.gmp, 1)} gross margin`, sgnCls(gmGain)),
      box('Company variance', usdC(newV), `was ${usdC(tot.vc / 100)} · now ${pctF(newPct, 2, true)} of target`, sgnCls(newV)), box('Share of total gap closed', tot.vc < 0 ? pctF(revGain / -(tot.vc / 100) * 100, 0) : '—', `rest of company runs at ${pctF(restPct, 2, true)}`));
  }
  b.append(h('div', { class: 'slider' }, h('label', null, 'Region '), sel, h('span', null, 'Target variance %'), sl, val, h('button', { class: 'tool', onclick: () => { WI.pct = Math.round(restPct * 10) / 10; sl.value = WI.pct; val.textContent = pctF(WI.pct, 1, true); paint(); } }, 'Match rest of company (' + pctF(restPct, 1, true) + ')'), h('button', { class: 'tool', onclick: () => { WI.pct = 0; sl.value = 0; val.textContent = pctF(0, 1, true); paint(); } }, 'Hit target (0%)')), out,
    h('div', { class: 'note' }, 'An illustration, not a forecast: it assumes the region’s target stays the same and its gross-margin rate holds. Use it to size the prize before committing to a recovery plan.'));
  paint();
}

/* ===========================================================================
   TAB 3 - TRENDS
   =========================================================================== */
const TRM = {
  v: { label: 'Variance $', get: d => d.v, fmt: v => usd(v), tick: v => usdC(v, 1), bar: true, zero: true },
  pct: { label: 'Variance %', get: d => d.pct, fmt: v => pctF(v, 2, true), tick: v => pctF(v, 1), bar: true, zero: true, refs: [{ y: CFG.favPct, label: '+' + CFG.favPct + '% Favorable above' }, { y: CFG.unfPct, label: MINUS + Math.abs(CFG.unfPct) + '% Unfavorable below' }] },
  att: { label: 'Qty attainment %', get: d => d.att, fmt: v => pctF(v), tick: v => v.toFixed(0) + '%', refs: [{ y: 100, label: '100% = on plan' }] },
  actual: { label: 'Actual sales $', get: d => d.actual, fmt: v => usd(v), tick: v => usdC(v, 0) },
  target: { label: 'Target sales $', get: d => d.target, fmt: v => usd(v), tick: v => usdC(v, 0) },
  mom: { label: 'MoM sales growth %', get: d => d.mom, fmt: v => pctF(v, 2, true), tick: v => v.toFixed(0) + '%', zero: true },
  yoy: { label: 'YoY sales growth %', get: d => d.yoy, fmt: v => pctF(v, 2, true), tick: v => v.toFixed(0) + '%', zero: true },
  ytd: { label: 'YTD variance $', get: d => d.ytdV, fmt: v => usd(v), tick: v => usdC(v, 1), zero: true },
  r3: { label: 'Rolling 3M variance $', get: d => d.r3, fmt: v => usd(v), tick: v => usdC(v, 1), zero: true },
  cum: { label: 'Cumulative variance $', get: d => d.cum, fmt: v => usd(v), tick: v => usdC(v, 1), zero: true },
  ts: { label: 'Tracking signal', get: d => d.ts, fmt: v => numF(v, 1), tick: v => numF(v, 0), zero: true, refs: [{ y: TS_LIMIT, label: '+' + TS_LIMIT + ' limit' }, { y: -TS_LIMIT, label: MINUS + TS_LIMIT + ' limit' }] }
};
const TR = { metric: 'v', split: 'total', seas: 'actual' };
function tabTrends(root) {
  const mAll = getMask('time'), inSel = monthOk();
  if (!maskCount(mAll)) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const mt = TRM[TR.metric];
  const controls = h('div', { class: 'tbar', style: { marginBottom: '4px' } }, h('span', { class: 'sub' }, 'Measure'), h('select', { class: 'sel', 'aria-label': 'Trend measure', onchange: e => { TR.metric = e.target.value; renderView(); } }, Object.entries(TRM).map(([k, v]) => h('option', { value: k, selected: k === TR.metric }, v.label))),
    h('span', { class: 'sub' }, 'Split by'), seg([['total', 'Total'], ['region', 'Region'], ['category', 'Category']], TR.split, v => { TR.split = v; renderView(); }, 'Split'));
  const splitDefs = TR.split === 'region' ? { key: A.R, K: NR, names: REG, col: regColor } : TR.split === 'category' ? { key: A.C, K: NC, names: CAT, col: catColor } : null;
  let series = [];
  if (!splitDefs) { const ser = monthlySeries(mAll, null, 0, -1); series = [{ id: 'all', name: mt.label, color: 'var(--s1)', values: ser.map(d => (d.n ? mt.get(d) : NaN)), ser }]; }
  else { const g = agg(splitDefs.key, splitDefs.K * NM, mAll); for (let i = 0; i < splitDefs.K; i++) { if (!Array.from({ length: NM }, (_, k) => g.n[i * NM + k]).some(Boolean)) continue; const ser = monthlyFrom(g, i * NM); series.push({ id: i, name: splitDefs.names[i], color: splitDefs.col(i), values: ser.map(d => (d.n ? mt.get(d) : NaN)), ser }); } }
  const lineMode = splitDefs || !mt.bar;
  const allSer = monthlySeries(mAll, null, 0, -1);
  grid.append(card({
    id: 'tr-main', title: mt.label + (splitDefs ? ' by ' + TR.split : ' by month'), sub: 'Click a month to filter · hover for all series' + (TR.metric === 'ts' ? ' · limit ±' + TS_LIMIT : ''), span: 's12', tools: [], build: b => {
      b.append(controls);
      const host = h('div'); b.append(host);
      lineChart(host, { x: MLAB, xFull: MLONG, series, fmt: mt.fmt, tickFmt: mt.tick, zero: !!mt.zero, mode: lineMode ? 'line' : 'bar', refs: mt.refs, height: 320, onX: selectMonth, selX: inSel, aria: mt.label + ' trend',
        extraTip: i => (splitDefs ? [] : [{ label: 'Status', value: ST_GLYPH[allSer[i].st] + ' ' + ST_NAME[allSer[i].st] }]) });
      if (TR.metric === 'ts') b.append(h('div', { class: 'note' }, 'Tracking signal = cumulative variance ÷ average absolute monthly variance. If every month so far is below target it equals –(number of months): it is the mathematical floor and simply restates “every month missed”. Read it as a bias alarm, not as an extra proof.'));
    },
    table: () => ({ name: 'trend_' + TR.metric, cols: [{ k: 'month', label: 'Month' }].concat(series.map(s => ({ k: 'v' + s.id, label: s.name, align: 'r', fmt: v => mt.fmt(v), csv: r => r['v' + s.id] }))), rows: MLONG.map((l, i) => Object.assign({ month: l }, ...series.map(s => ({ ['v' + s.id]: s.values[i] })))), search: false })
  }));
  // seasonality
  const sm = { actual: ['Actual sales', d => d.actual, v => usd(v), v => usdC(v, 0)], target: ['Target sales', d => d.target, v => usd(v), v => usdC(v, 0)], pct: ['Variance %', d => d.pct, v => pctF(v, 2, true), v => v.toFixed(1) + '%'] }[TR.seas];
  const yS = YEARS.map((y, yi) => ({ id: y, name: String(y), color: `var(--s${yi + 1})`, values: Array.from({ length: 12 }, (_, k) => { const mi = MON.findIndex((q, i) => MYEAR[i] === y && MNUM[i] === k + 1); return mi >= 0 && allSer[mi].n ? sm[1](allSer[mi]) : NaN; }) }));
  grid.append(card({ id: 'tr-seas', title: 'Same month, different year', sub: 'Seasonality check: is the gap seasonal or constant?', span: 's6', tools: [seg([['actual', 'Actual'], ['target', 'Target'], ['pct', 'Variance %']], TR.seas, v => { TR.seas = v; renderView(); }, 'Seasonality measure')],
    build: b => lineChart(b, { x: MNAME, series: yS, fmt: sm[2], tickFmt: sm[3], zero: TR.seas === 'pct', height: 280, aria: 'Month of year by year' }),
    table: () => ({ name: 'seasonality', cols: [{ k: 'month', label: 'Month' }].concat(yS.map(s => ({ k: 'y' + s.id, label: s.name, align: 'r', fmt: v => sm[2](v), csv: r => r['y' + s.id] }))), rows: MNAME.map((n, i) => Object.assign({ month: n }, ...yS.map(s => ({ ['y' + s.id]: s.values[i] })))), search: false }) }));
  // quarters
  const qIdx = Float64Array.from({ length: N }, (_, i) => { const m = A.M[i]; return (MYEAR[m] - YEARS[0]) * 4 + MQ[m] - 1; });
  const gq = agg(qIdx, YEARS.length * 4, mAll), qRows = [];
  for (let q = 0; q < YEARS.length * 4; q++) { if (!gq.n[q]) continue; const d = dv(gq, q); d.q = `${YEARS[Math.floor(q / 4)]} Q${(q % 4) + 1}`; d.qi = q; qRows.push(d); }
  qRows.forEach(d => { const p1 = qRows.find(x => x.qi === d.qi - 1), p4 = qRows.find(x => x.qi === d.qi - 4); d.qoq = p1 && p1.ac > 0 ? (d.ac / p1.ac - 1) * 100 : NaN; d.yoyq = p4 && p4.ac > 0 ? (d.ac / p4.ac - 1) * 100 : NaN; });
  grid.append(card({ id: 'tr-q', title: 'Quarterly view', sub: 'All months (time filter not applied)', span: 's6', build: b => b.append(tableView({ name: 'quarters', search: false, sortKey: null, rows: qRows,
    cols: [{ k: 'q', label: 'Quarter' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usdC(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usdC(v) }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => vSpan(v, usdC(v)) }, { k: 'pct', label: 'Var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'qoq', label: 'QoQ', align: 'r', fmt: v => pctF(v, 1, true) }, { k: 'yoyq', label: 'YoY', align: 'r', fmt: v => pctF(v, 1, true) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }] })) }));
  grid.append(card({ id: 'tr-tbl', title: 'Monthly scorecard', sub: 'Growth, year-to-date, rolling and tracking measures for the whole company', span: 's12', build: b => b.append(tableView({ name: 'monthly_scorecard', search: false, rows: allSer.map((d, i) => Object.assign({ month: MLONG[i] }, d)), cols: monthCols().concat([{ k: 'mat', label: 'Material', fmt: (v, r) => matCell(v, r.trig), csv: r => (r.mat ? 'Yes' : 'No') }]), onRow: r => selectMonth(MLONG.indexOf(r.month)) })) }));
}
/* ===========================================================================
   TAB 4 - REGION x CATEGORY
   =========================================================================== */
const RC = { metric: 'pct', persSort: null };
function persistenceRows(mask) {
  const g = agg(KEY_RCM, NR * NC * NM, mask), rows = [];
  for (let r = 0; r < NR; r++) for (let c = 0; c < NC; c++) {
    let months = 0, below = 0, unf = 0, sumV = 0, worst = Infinity, worstM = -1, tc = 0, ac = 0, nlines = 0;
    for (let m = 0; m < NM; m++) { const k = (r * NC + c) * NM + m; if (!g.n[k]) continue; const vc = g.ac[k] - g.tc[k]; months++; nlines += g.n[k]; if (vc < 0) below++; if (statusOf(vc, g.tc[k]) === 2) unf++; sumV += vc; tc += g.tc[k]; ac += g.ac[k]; if (vc < worst) { worst = vc; worstM = m; } }
    if (!months) continue;
    rows.push({ r, c, region: REG[r], category: CAT[c], months, below, share: below / months * 100, chronic: below / months >= PERSIST, unf, avg: sumV / 100 / months, worstV: worst / 100, worstM: MLONG[worstM], total: sumV / 100, pct: pct2(sumV, tc), n: nlines });
  }
  return rows;
}
const PERSIST = 0.75;
function tabRegCat(root) {
  const mRC = getMask('reg,cat');
  if (!maskCount(mRC)) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const g = agg(KEY_RC, NR * NC, mRC), pers = persistenceRows(mRC), pmap = {}; pers.forEach(p => pmap[p.r * NC + p.c] = p);
  const tot = totals(mRC);
  const M = {
    pct: { l: 'Variance %', f: d => d.pct, fmt: v => pctF(v, 1, true), lf: v => pctF(v, 1), s: 'diverging' },
    v: { l: 'Variance $', f: d => d.v, fmt: v => usdC(v, 2), fmtShort: v => usdC(v, 1), lf: v => usdC(v, 1), s: 'diverging' },
    att: { l: 'Qty attainment %', f: d => d.att - 100, fmt: v => pctF(v + 100, 1), lf: v => pctF(v + 100, 0), s: 'diverging', show: v => v + 100 },
    below: { l: '% of months below target', f: (d, k) => (pmap[k] ? pmap[k].share : NaN), fmt: v => v.toFixed(0) + '%', lf: v => v.toFixed(0) + '%', s: 'seq' },
    contrib: { l: 'Contribution to total variance (pp)', f: d => (tot.tc > 0 ? d.vc / tot.tc * 100 : NaN), fmt: v => (v > 0 ? '+' : '') + v.toFixed(2), lf: v => v.toFixed(2), s: 'diverging' }
  };
  const mm = M[RC.metric];
  const vals = REG.map((_, r) => CAT.map((__, c) => { const k = r * NC + c; return g.n[k] ? mm.f(dv(g, k), k) : NaN; }));
  const selR = new Set(S.regs), selC = new Set(S.cats);
  grid.append(card({
    id: 'rc-heat', title: 'Region × category heat map', sub: 'Blue above target, red below, grey = on target · click a cell to filter region and category', span: 's7',
    tools: [h('select', { class: 'sel', 'aria-label': 'Heat map measure', onchange: e => { RC.metric = e.target.value; renderView(); } }, Object.entries(M).map(([k, v]) => h('option', { value: k, selected: k === RC.metric }, v.l)))],
    build: b => heatmap(b, { rows: REG, cols: CAT, vals, fmt: mm.fmt, labW: 120, cellH: 52, diverging: mm.s === 'diverging', legendFmt: mm.lf, selR, selC, aria: 'Region by category heat map of ' + mm.l,
      onCell: (r, c) => change(() => { if (S.regs.size === 1 && S.regs.has(r) && S.cats.size === 1 && S.cats.has(c)) { S.regs.clear(); S.cats.clear(); } else { S.regs.clear(); S.cats.clear(); S.regs.add(r); S.cats.add(c); } }),
      cellTip: (r, c) => { const k = r * NC + c, d = dv(g, k), p = pmap[k]; return tipFor(d, REG[r] + ' · ' + CAT[c]).rows ? { title: REG[r] + ' · ' + CAT[c], rows: tipFor(d, '').rows.concat(p ? [{ label: 'Months below target', value: p.below + ' of ' + p.months }] : []) } : { title: '', rows: [] }; } }),
    table: () => ({ name: 'region_category', cols: [{ k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => usd(v) }, { k: 'pct', label: 'Variance %', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }], rows: REG.flatMap((rn, r) => CAT.map((cn, c) => { const k = r * NC + c; return g.n[k] ? Object.assign({ region: rn, category: cn }, dv(g, k)) : null; }).filter(Boolean)) })
  }));
  grid.append(card({ id: 'rc-chronic', title: 'Persistent shortfalls', sub: `A pair is “chronic” if it is below target in at least ${PERSIST * 100}% of the months in view`, span: 's5', build: b => {
    const ch = pers.filter(p => p.chronic).length;
    b.append(h('div', { style: { display: 'flex', alignItems: 'baseline', gap: '10px', margin: '2px 0 10px' } }, h('span', { style: { fontSize: '40px', fontWeight: 650 } }, `${ch}`), h('span', { class: 'sub' }, `of ${pers.length} region × category pairs are chronic`)));
    const top = pers.slice().sort((a, c) => c.share - a.share || a.total - c.total).slice(0, 7);
    b.append(h('ul', { class: 'ins' }, top.map(p => h('li', { class: p.chronic ? 'bad' : '' }, `${p.region} · ${p.category}: below target in ${p.below} of ${p.months} months (${p.share.toFixed(0)}%), total ${usdC(p.total)}`))));
    b.append(h('div', { class: 'note' }, 'Chronic gaps are structural (targets, market, distribution), not one-off events. Full ranking below.'));
  } }));
  // region x month heat strip
  const mRM = getMask('time,reg'), gRM = agg(KEY_RM, NR * NM, mRM), inSel = monthOk();
  const valsM = REG.map((_, r) => MON.map((__, m) => { const k = r * NM + m; return gRM.n[k] ? dv(gRM, k).pct : NaN; }));
  const selCm = new Set(); inSel.forEach((v, i) => { if (v && (S.m0 !== 0 || S.m1 !== NM - 1 || S.years.size || S.qtrs.size)) selCm.add(i); });
  grid.append(card({ id: 'rc-rm', title: 'Region × month: variance %', sub: 'Every cell is a region’s monthly total vs its target · click a cell to pick the region and month', span: 's12',
    build: b => heatmap(b, { rows: REG, cols: MON.map((_, i) => MNAME[MNUM[i] - 1]), colSub: MON.map((_, i) => (MNUM[i] === 1 ? String(MYEAR[i]) : '')), vals: valsM, fmt: v => v.toFixed(1), fmtShort: v => v.toFixed(0), labW: 120, cellH: 34, legendFmt: v => pctF(v, 1), selR: new Set(S.regs), selC: selCm, aria: 'Region by month variance heat map',
      onCell: (r, m) => change(() => { S.regs.clear(); S.regs.add(r); S.years.clear(); S.qtrs.clear(); S.m0 = S.m1 = m; }),
      cellTip: (r, m) => { const k = r * NM + m, d = dv(gRM, k); return { title: REG[r] + ' · ' + MLONG[m], rows: tipFor(d, '').rows }; } }),
    table: () => ({ name: 'region_month', cols: [{ k: 'region', label: 'Region' }].concat(MON.map((m, i) => ({ k: 'm' + i, label: MLAB[i], align: 'r', fmt: v => pctF(v, 1, true), csv: r => r['m' + i] }))), rows: REG.map((rn, r) => Object.assign({ region: rn }, ...MON.map((m, i) => ({ ['m' + i]: valsM[r][i] })))), search: false }) }));
  // dumbbell YoY attainment
  if (YEARS.length >= 2) {
    const yLo = YEARS[YEARS.length - 2], yHi = YEARS[YEARS.length - 1];
    const mT = getMask('time,reg,cat'); const KY = Float64Array.from({ length: N }, (_, i) => KEY_RC[i] * 2 + (MYEAR[A.M[i]] === yHi ? 1 : 0));
    const mY = new Uint8Array(N); for (let i = 0; i < N; i++) mY[i] = mT[i] && (MYEAR[A.M[i]] === yLo || MYEAR[A.M[i]] === yHi) ? 1 : 0;
    const gy = agg(KY, NR * NC * 2, mY);
    const items = []; for (let r = 0; r < NR; r++) for (let c = 0; c < NC; c++) { const k0 = (r * NC + c) * 2, k1 = k0 + 1; if (!gy.tc[k0] || !gy.tc[k1]) continue; items.push({ id: r * NC + c, r, c, label: REG[r] + ' · ' + CAT[c], v0: gy.ac[k0] / gy.tc[k0] * 100, v1: gy.ac[k1] / gy.tc[k1] * 100, selected: (S.regs.size === 0 || S.regs.has(r)) && (S.cats.size === 0 || S.cats.has(c)) && (S.regs.size || S.cats.size) ? true : false }); }
    items.sort((a, b) => (a.v1 - a.v0) - (b.v1 - b.v0));
    grid.append(card({ id: 'rc-db', title: `Is it getting better? Sales attainment ${yLo} → ${yHi}`, sub: 'Actual sales ÷ target sales per region × category, sorted by change · red = deteriorated, blue = improved · click to filter', span: 's6',
      build: b => dumbbell(b, { items, l0: String(yLo), l1: String(yHi), fmt: v => v.toFixed(1) + '%', anySel: items.some(i => i.selected), onClick: it => change(() => { S.regs.clear(); S.cats.clear(); S.regs.add(it.r); S.cats.add(it.c); }), aria: 'Attainment change by region and category' }),
      table: () => ({ name: 'attainment_change', cols: [{ k: 'label', label: 'Region · Category' }, { k: 'v0', label: String(yLo), align: 'r', fmt: v => pctF(v) }, { k: 'v1', label: String(yHi), align: 'r', fmt: v => pctF(v) }, { k: 'd', label: 'Change (pp)', align: 'r', fmt: v => ppF(v), sortVal: r => r.v1 - r.v0, csv: r => (r.v1 - r.v0).toFixed(2) }], rows: items, search: false }) }));
  }
  grid.append(card({ id: 'rc-pers', title: 'Persistence ranking: all region × category pairs', sub: 'Months below target, worst month and total variance · click a row to filter', span: YEARS.length >= 2 ? 's6' : 's12', build: b => b.append(tableView({
    name: 'persistence', rows: pers, sortKey: 'share', sortDir: -1, search: false, onRow: r => change(() => { S.regs.clear(); S.cats.clear(); S.regs.add(r.r); S.cats.add(r.c); }),
    cols: [{ k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'below', label: 'Months below', align: 'r', fmt: (v, r) => v + ' / ' + r.months, sortVal: r => r.share }, { k: 'share', label: 'Share', align: 'r', fmt: v => v.toFixed(0) + '%' }, { k: 'chronic', label: 'Chronic', fmt: v => (v ? h('span', { class: 'chr' }, '● Chronic') : h('span', { style: { color: 'var(--ink2)' } }, 'No')), csv: r => (r.chronic ? 'Yes' : 'No') },
      { k: 'avg', label: 'Avg monthly var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'worstV', label: 'Worst month $', align: 'r', fmt: (v, r) => usd(v) + ' (' + r.worstM + ')' }, { k: 'total', label: 'Total var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }] })) }));
}

/* ===========================================================================
   TAB 5 - PRODUCTS
   =========================================================================== */
const PR = { view: 'worst', n: 10, by: 'v' };
function tabProducts(root) {
  const mP = getMask('prod');
  if (!maskCount(mP)) { root.append(emptyMsg()); return; }
  const rows = groupRows(A.P, NP, mP, PRD).filter(d => d.n).map(d => Object.assign(d, { cat: CAT[DIM.prodCat[d.id]], ci: DIM.prodCat[d.id] }));
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const key = PR.by === 'v' ? 'v' : 'pct';
  let shown = rows.slice().sort((a, b) => a[key] - b[key]);
  if (PR.view === 'worst') shown = shown.slice(0, PR.n); else if (PR.view === 'best') shown = shown.slice(-PR.n).reverse(); 
  if (PR.view === 'best' || PR.view === 'all') shown.sort((a, b) => b[key] - a[key]); else shown.sort((a, b) => a[key] - b[key]);
  const above = rows.filter(d => d.v > 0).length;
  const items = shown.map(d => ({ id: d.id, label: d.name, value: d.v, txt: usdS(d.v) + ' · ' + pctF(d.pct, 2, true), selected: S.prods.has(d.id), d, color: undefined }));
  const ctl = h('div', { class: 'tbar', style: { marginBottom: '4px' } }, seg([['worst', 'Biggest shortfalls'], ['best', 'Best performers'], ['all', 'All']], PR.view, v => { PR.view = v; renderView(); }, 'Product ranking view'),
    h('select', { class: 'sel', 'aria-label': 'How many products', onchange: e => { PR.n = +e.target.value; renderView(); } }, [5, 10, 15, 27].map(n => h('option', { value: n, selected: n === PR.n }, 'Top ' + n))),
    seg([['v', 'by $'], ['pct', 'by %']], PR.by, v => { PR.by = v; renderView(); }, 'Rank by'));
  grid.append(card({ id: 'pr-rank', title: 'Products ranked by variance', sub: `${above} of ${rows.length} products are above target. A ranking is not a verdict: a “best” product can still miss target.`, span: 's7', build: b => {
    b.append(ctl); const host = h('div'); b.append(host);
    hbars(host, items, { onClick: (it, ev) => pickProd(it.id, ev), rowH: 28, tip: it => { const t = tipFor(it.d, it.label + ' (' + it.d.cat + ')'); t.rows.push({ label: 'Above target', value: it.d.v > 0 ? 'Yes' : 'No' }); return t; }, aria: 'Products ranked by sales variance' });
  }, table: () => ({ name: 'products', cols: prodCols(), rows, sortKey: 'v', sortDir: 1 }) }));
  // pareto of shortfall
  const neg = rows.filter(d => d.v < 0).sort((a, b) => a.v - b.v), gross = neg.reduce((s, d) => s + d.v, 0);
  const keep = neg.slice(0, 15); let cum = 0; const pi = keep.map(d => { cum += d.v; return { id: d.id, label: d.name, usd: d.v, share: d.v / gross * 100, cum: cum / gross * 100, selected: S.prods.has(d.id) }; });
  if (neg.length > 15) { const rest = neg.slice(15), rv = rest.reduce((s, d) => s + d.v, 0); cum += rv; pi.push({ id: null, label: `Other (${rest.length})`, usd: rv, share: rv / gross * 100, cum: cum / gross * 100 }); }
  grid.append(card({ id: 'pr-pareto', title: 'Where the shortfall sits (Pareto)', sub: neg.length ? `${neg.length} products below target add up to ${usdC(gross)}; ${pi.filter(p => p.cum - p.share < 80 && p.id != null).length} reach 80%` : 'No product is below target in this selection', span: 's5',
    build: b => (neg.length ? pareto(b, { items: pi, anySel: pi.some(p => p.selected), onClick: it => pickProd(it.id) }) : b.append(emptyMsg('Nothing below target.'))), table: () => ({ name: 'pareto', cols: [{ k: 'label', label: 'Product' }, { k: 'usd', label: 'Shortfall $', align: 'r', fmt: v => usd(v) }, { k: 'share', label: 'Share %', align: 'r', fmt: v => pctF(v, 1) }, { k: 'cum', label: 'Cumulative %', align: 'r', fmt: v => pctF(v, 1) }], rows: pi, search: false }) }));
  // bubble
  const maxAbs = Math.max(1, ...rows.map(d => Math.abs(d.v))), pts = rows.map(d => ({ id: d.id, label: d.name, x: d.target, y: d.pct, r: 5 + 11 * Math.sqrt(Math.abs(d.v) / maxAbs), color: catColor(d.ci), selected: S.prods.has(d.id), tipRows: tipFor(d, '').rows.concat([{ label: 'Category', value: d.cat }]) }));
  const lab = new Set(rows.slice().sort((a, b) => a.pct - b.pct).slice(0, 3).concat(rows.slice().sort((a, b) => b.pct - a.pct).slice(0, 3)).map(d => d.id));
  grid.append(card({ id: 'pr-bubble', title: 'Size vs performance', sub: 'Each bubble is a product: planned sales (x), variance % (y), bubble size = dollar variance · colour = category · click to filter', span: 's7',
    build: b => bubble(b, { pts, xFmt: v => usdC(v, 0), yFmt: v => pctF(v, 1), xTitle: 'Target sales', yTitle: 'Variance %', labelIds: lab, onClick: p => pickProd(p.id), height: 340, bands: [{ y0: CFG.unfPct, y1: CFG.favPct, label: `On target band (±${CFG.favPct}%)` }], legend: CAT.map((c, i) => ({ label: c, color: catColor(i), type: 'bar' })), aria: 'Products by size and variance' }),
    table: () => ({ name: 'products_bubble', cols: prodCols(), rows, sortKey: 'pct', sortDir: 1 }) }));
  // by category within product tab: share of lines? product table
  grid.append(card({ id: 'pr-tbl', title: 'Product scorecard', sub: 'Click a row to filter', span: 's5', build: b => b.append(tableView({ name: 'product_scorecard', rows, sortKey: 'v', sortDir: 1, onRow: r => pickProd(r.id), cols: prodCols().slice(0, 6) })) }));
}
function prodCols() {
  return [{ k: 'name', label: 'Product' }, { k: 'cat', label: 'Category' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) },
    { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }, { k: 'mat', label: 'Material', fmt: (v, r) => matCell(v, r.trig), csv: r => (r.mat ? 'Yes' : 'No') }, { k: 'pu', label: '% lines below', align: 'r', fmt: v => pctF(v, 1) }];
}

/* ===========================================================================
   TAB 6 - EXCEPTIONS & LINES
   =========================================================================== */
const EX = { lineQ: '' };
function tabExceptions(root) {
  const m = getMask(), T = totals(m);
  if (!T.n) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const g = agg(KEY_RCM, NR * NC * NM, m), cells = [];
  for (let r = 0; r < NR; r++) for (let c = 0; c < NC; c++) for (let mo = 0; mo < NM; mo++) { const k = (r * NC + c) * NM + mo; if (!g.n[k]) continue; const d = dv(g, k); d.r = r; d.c = c; d.mo = mo; d.region = REG[r]; d.category = CAT[c]; d.month = MLONG[mo]; cells.push(d); }
  const matCells = cells.filter(d => d.mat).sort((a, b) => Math.abs(b.v) - Math.abs(a.v));
  const fav = matCells.filter(d => d.v > 0).length;
  grid.append(h('section', { class: 'card s4 kpi' }, h('div', { class: 'lab' }, 'Material region × category × month cells'), h('div', { class: 'val' }, `${matCells.length} of ${cells.length}`), h('div', { class: 'dl' }, `${matCells.length - fav} below target, ${fav} above · rule: |var %| ≥ ${CFG.matPctV}% or |var $| ≥ ${usd(CFG.matUsd)}`)));
  grid.append(h('section', { class: 'card s4 kpi' }, h('div', { class: 'lab' }, 'Material individual lines'), h('div', { class: 'val' }, numF(T.nm)), h('div', { class: 'dl' }, `${pctF(T.nm / T.n * 100, 1)} of ${numF(T.n)} lines – too many to review one by one`)));
  grid.append(h('section', { class: 'card s4 kpi' }, h('div', { class: 'lab' }, 'Typical swing of a single line'), h('div', { class: 'val' }, '±' + lineSd(m).toFixed(1) + '%'), h('div', { class: 'dl' }, 'standard deviation of line variance %: why line-level flags are noisy')));
  grid.append(card({ id: 'ex-cells', title: 'Exceptions: material region × category × month cells', sub: 'The short list worth a conversation · click a row to filter to it', span: 's12', build: b => b.append(tableView({ name: 'material_cells', rows: matCells, sortKey: null, pageSize: 15, onRow: r => change(() => { S.regs.clear(); S.cats.clear(); S.regs.add(r.r); S.cats.add(r.c); S.years.clear(); S.qtrs.clear(); S.m0 = S.m1 = r.mo; }),
    cols: [{ k: 'month', label: 'Month', sortVal: r => r.mo }, { k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) },
      { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }, { k: 'trig', label: 'Trigger' }, { k: 'n', label: 'Lines', align: 'r', fmt: v => numF(v) }] })) }));
  // histogram of line variance %
  const bins = []; const lo = -60, hi = 60, bw = 2.5; for (let x = lo; x < hi; x += bw) bins.push({ x0: x, x1: x + bw, count: 0, band: x >= CFG.unfPct - 1e-9 && x + bw <= CFG.favPct + 1e-9 });
  let tot = 0; for (let i = 0; i < N; i++) { if (!m[i] || !(A.TC[i] > 0)) continue; const p = (A.AC[i] - A.TC[i]) / A.TC[i] * 100; let k = Math.floor((p - lo) / bw); k = Math.max(0, Math.min(bins.length - 1, k)); bins[k].count++; tot++; }
  grid.append(card({ id: 'ex-hist', title: 'How individual lines are spread', sub: `Variance % of each line · grey bars = inside the ±${CFG.favPct}% on-target band`, span: 's6', build: b => { histogram(b, { bins, total: tot, xFmt: v => (v > 0 ? '+' : v < 0 ? MINUS : '') + Math.abs(v) + '%', vlines: [{ x: 0, label: '0' }, { x: CFG.unfPct, label: MINUS + CFG.favPct + '%' }, { x: CFG.favPct, label: '+' + CFG.favPct + '%' }, { x: -CFG.matPctV, label: MINUS + CFG.matPctV + '%' }, { x: CFG.matPctV, label: '+' + CFG.matPctV + '%' }], aria: 'Distribution of line variance percent' }); b.append(h('div', { class: 'note' }, 'The spread is wide and roughly symmetric around a small negative centre. The company-level miss is a small shift of a very noisy distribution – which is why conclusions must come from aggregated totals, not from individual lines.')); } }));
  // line explorer
  const lineRows = []; for (let i = 0; i < N; i++) { if (!m[i]) continue; const vc = A.AC[i] - A.TC[i]; lineRows.push({ src: A.SRC[i], month: MLONG[A.M[i]], mo: A.M[i], region: REG[A.R[i]], category: CAT[DIM.prodCat[A.P[i]]], product: PRD[A.P[i]], tq: A.TQ[i], aq: A.AQ[i], price: A.UP[i] / 100, target: A.TC[i] / 100, actual: A.AC[i] / 100, v: vc / 100, pct: pct2(vc, A.TC[i]), st: A.ST[i], mat: A.MT[i] === 1, trig: ['', 'Pct', 'USD', 'Pct and USD'][A.TRG[i]], vol: A.VOL[i] / 1e4, prc: A.PRC[i] / 1e4, note: NOTES[A.NT[i]], rep: A.QR[i] === 1 }); }
  grid.append(card({ id: 'ex-lines', title: 'Line explorer', sub: `${numF(lineRows.length)} sales lines match your filters · sort, search and export`, span: 's12', build: b => b.append(tableView({ name: 'sales_lines_filtered', rows: lineRows, pageSize: 25, sortKey: 'src', sortDir: 1,
    cols: [{ k: 'src', label: 'Source row', align: 'r' }, { k: 'month', label: 'Month', sortVal: r => r.mo }, { k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'product', label: 'Product' }, { k: 'tq', label: 'Target qty', align: 'r', fmt: v => numF(v) }, { k: 'aq', label: 'Actual qty', align: 'r', fmt: v => numF(v) }, { k: 'price', label: 'Unit price', align: 'r', fmt: v => usd(v, 2) },
      { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v, 2) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v, 2) }, { k: 'v', label: 'Variance $', align: 'r', fmt: v => vSpan(v, usd(v, 2)) }, { k: 'pct', label: 'Var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] },
      { k: 'mat', label: 'Material', fmt: (v, r) => matCell(v, r.trig), csv: r => (r.mat ? 'Yes' : 'No') }, { k: 'vol', label: 'Volume effect $', align: 'r', fmt: v => usd(v, 2) }, { k: 'prc', label: 'Price effect $', align: 'r', fmt: v => usd(v, 2) }, { k: 'note', label: 'Note' }, { k: 'rep', label: 'Qty repaired', fmt: v => (v ? 'Yes' : ''), csv: r => (r.rep ? 'Yes' : 'No') }] })) }));
}
/* ===========================================================================
   TAB 7 - INSIGHTS & RECOMMENDATIONS     TAB 8 - DATA & VALIDATION
   =========================================================================== */
let ALL_MASK = null;
function maskAll() { if (!ALL_MASK) { ALL_MASK = new Uint8Array(N).fill(1); } return ALL_MASK; }

function headlineClaims() {
  const all = maskAll(), T = totals(all), ser = monthSeries(all), below = ser.filter(d => d.vc < 0).length;
  const reg = groupRows(A.R, NR, all, REG), cat = groupRows(A.C, NC, all, CAT), prd = groupRows(A.P, NP, all, PRD);
  const R = n => reg.find(r => r.name === n), net = T.v, pers = persistenceRows(all), chronic = pers.filter(p => p.chronic).length;
  const rm = agg(KEY_RM, NR * NM, all), monthsBelow = REG.map((_, r) => Array.from({ length: NM }, (__, m) => { const k = r * NM + m; return rm.ac[k] - rm.tc[k]; }).filter(v => v < 0).length);
  const cells = agg(KEY_RCM, NR * NC * NM, all); let matCells = 0; for (let k = 0; k < cells.K; k++) if (cells.n[k] && isMat(cells.ac[k] - cells.tc[k], cells.tc[k])) matCells++;
  const elx = (() => { const m = new Uint8Array(N); for (let i = 0; i < N; i++) m[i] = (CAT[A.C[i]] === 'Electronics' && REG[A.R[i]] !== 'Asia Pacific') ? 1 : 0; return totals(m).v; })();
  const haY = (y) => { const m = new Uint8Array(N); for (let i = 0; i < N; i++) m[i] = (CAT[A.C[i]] === 'Home Appliances' && MYEAR[A.M[i]] === y) ? 1 : 0; const t = totals(m); return t.ac / t.tc * 100; };
  const la = R('Latin America'), ap = R('Asia Pacific'), above = prd.filter(p => p.v > 0), matLines = Array.from(A.MT).reduce((a, b) => a + b, 0);
  const others = REG.map((n, i) => i).filter(i => REG[i] !== 'Asia Pacific');
  const rows = [
    ['Company variance', '–$9,465,029 (–1.94%)', usd(T.v) + ' (' + pctF(T.pct, 2, true) + ')', usd(T.v) === '–$9,465,029' && T.pct === -1.94],
    ['Quantity attainment', '98.03%', pctF(T.att), pctF(T.att) === '98.03%'],
    ['Months below target', '24 of 24', below + ' of ' + NM, below === NM && NM === 24],
    ['Asia Pacific variance', '+$4,478,728 (+4.57%)', usd(ap.v) + ' (' + pctF(ap.pct, 2, true) + ')', Math.round(ap.v) === 4478728 && ap.pct === 4.57],
    ['Latin America share of net variance', '74%', Math.round(la.v / net * 100) + '%', Math.round(la.v / net * 100) === 74],
    ['Regions below target in 24 of 24 months', 'All four non-Asia Pacific; Asia Pacific above in 24/24', others.filter(i => monthsBelow[i] === NM).length + ' of 4 below every month; Asia Pacific below in ' + monthsBelow[REG.indexOf('Asia Pacific')] + ' months', others.every(i => monthsBelow[i] === NM) && monthsBelow[REG.indexOf('Asia Pacific')] === 0],
    ['Chronic region × category pairs', '17 of 25', chronic + ' of ' + pers.length, chronic === 17 && pers.length === 25],
    ['Products above target', '6 of 27 (all Electronics)', above.length + ' of ' + NP + (above.every(p => CAT[DIM.prodCat[p.id]] === 'Electronics') ? ' (all Electronics)' : ' (not all Electronics)'), above.length === 6 && above.every(p => CAT[DIM.prodCat[p.id]] === 'Electronics')],
    ['Price effect (total)', '$0', usd(T.prc, 2), Math.abs(T.prc) < 0.5],
    ['Electronics outside Asia Pacific', '–$683,122', usd(elx), Math.round(elx) === -683122],
    ['Home Appliances sales attainment 2024 → 2025', '98.01% → 96.72%', haY(2024).toFixed(2) + '% → ' + haY(2025).toFixed(2) + '%', haY(2024).toFixed(2) === '98.01' && haY(2025).toFixed(2) === '96.72'],
    ['Material region × category × month cells / material lines', '57 / 12,968', matCells + ' / ' + numF(matLines), matCells === 57 && matLines === 12968]
  ];
  return rows;
}
function tabInsights(root) {
  const grid = h('div', { class: 'grid' }); root.append(grid);
  grid.append(card({ id: 'in-md', title: 'Insights & recommendations', sub: 'Written analysis for the full period (not affected by the filters)', span: 's8', build: b => {
    if (D.md) {
      const box = h('div', { class: 'md' }); box.innerHTML = D.md;
      const heads = [...box.querySelectorAll('h3')].filter(x => /^\d+\./.test(x.textContent.trim()));
      if (heads.length > 1) { const jb = h('div', { class: 'md-jump' }); heads.forEach(x => { const t = x.textContent.trim().replace(/\s*\(.*$/, ''); jb.append(h('button', { class: 'chip', onclick: () => x.scrollIntoView({ behavior: 'smooth', block: 'start' }) }, t.length > 34 ? t.slice(0, 32) + '…' : t)); }); b.append(jb); }
      b.append(box);
    }       // HTML produced by the Python builder from your .md file (escaped)
    else b.append(h('div', { class: 'note' }, 'Insights_and_Recommendations.md was not found next to the build script, so only the live insights are shown.'));
  } }));
  const side = h('div', { class: 's4', style: { display: 'flex', flexDirection: 'column', gap: '14px' } });
  const rows = headlineClaims(), okN = rows.filter(r => r[3]).length;
  const c1 = card({ id: 'in-claims', title: 'Does the data still say this?', sub: `${okN} of ${rows.length} headline statements re-checked against the line data`, span: '', build: b => b.append(h('div', { class: 'tbl-wrap', style: { maxHeight: '480px' } }, h('table', { class: 't' }, h('thead', null, h('tr', null, ['', 'Statement', 'In the write-up', 'Recomputed now'].map(x => h('th', null, x)))),
    h('tbody', null, rows.map(r => h('tr', null, h('td', { class: r[3] ? 'pass' : 'fail' }, r[3] ? '✔' : '✖'), h('td', { style: { whiteSpace: 'normal' } }, r[0]), h('td', { style: { whiteSpace: 'normal' } }, r[1]), h('td', { style: { whiteSpace: 'normal' } }, r[2])))))))
  });
  const c2 = card({ id: 'in-live', title: 'Live insights for your current filters', sub: 'The same analysis applied to whatever you have selected', span: '', build: b => b.append(insightsList(liveInsights())) });
  side.append(c1, c2); grid.append(side);
}

/* ---------- reconciliation of the in-browser numbers with the Excel workbook ---------- */
function reconcile() {
  const xl = D.xl, out = []; if (!xl) return out;
  const add = (name, items, maxDiff, tol, detail) => out.push({ name, items, maxDiff, tol, pass: maxDiff <= tol + 1e-12, detail: detail || '' });
  const sumBy = (key, K) => { const v = new Float64Array(K); for (let i = 0; i < N; i++) v[key[i]] += A.AC[i] - A.TC[i]; return v; };
  const cmp = (arr, names, ref) => { let mx = 0; names.forEach((nm, i) => { if (ref[nm] == null) return; mx = Math.max(mx, Math.abs(arr[i] / 100 - ref[nm])); }); return mx; };
  add('Variance by region (5)', NR, cmp(sumBy(A.R, NR), REG, xl.region), 0.005);
  add('Variance by category (5)', NC, cmp(sumBy(A.C, NC), CAT, xl.category), 0.005);
  add('Variance by product (27)', NP, cmp(sumBy(A.P, NP), PRD, xl.product), 0.005);
  const all = maskAll(), ser = monthSeries(all), mx = f => MON.reduce((a, m, i) => { const r = xl.month[m]; if (!r || r[f] == null) return a; const v = { var: ser[i].v, pct: ser[i].pct, ts: ser[i].ts, ytd: ser[i].ytdV, r3: ser[i].r3, mom: ser[i].mom, yoy: ser[i].yoy }[f]; return isFinite(v) ? Math.max(a, Math.abs(v - r[f])) : a; }, 0);
  add('Monthly variance $ (24)', NM, mx('var'), 0.005); add('Monthly variance % (24)', NM, mx('pct'), 0.0051);
  add('Tracking signal by month (24)', NM, mx('ts'), 0.0101); add('YTD variance by month (24)', NM, mx('ytd'), 0.005); add('Rolling 3-month variance (22)', NM - 2, mx('r3'), 0.005);
  add('Month-on-month sales growth (23)', NM - 1, mx('mom'), 0.0101); add('Year-on-year sales growth (12)', NM - 12, mx('yoy'), 0.0101);
  const yrs = Object.keys(xl.year || {}); let dy = 0; yrs.forEach(y => { const m = new Uint8Array(N); for (let i = 0; i < N; i++) m[i] = MYEAR[A.M[i]] === +y ? 1 : 0; const t = totals(m); dy = Math.max(dy, Math.abs(t.v - xl.year[y].var), Math.abs(t.pct - xl.year[y].pct) / 100); }); add('Variance by year', yrs.length, dy, 0.0051);
  const pers = persistenceRows(all); let dp = 0; pers.forEach(p => { const r = xl.pers[p.region + '|' + p.category]; if (r) dp = Math.max(dp, Math.abs(p.below - r.below)); });
  add('Months below target per region × category (25)', pers.length, dp, 0); add('Chronic pairs count', 1, Math.abs(pers.filter(p => p.chronic).length - xl.chronic), 0);
  const g = agg(KEY_RCM, NR * NC * NM, all); let mc = 0; for (let k = 0; k < g.K; k++) if (g.n[k] && isMat(g.ac[k] - g.tc[k], g.tc[k])) mc++;
  add('Material region × category × month cells', 1, Math.abs(mc - xl.exc_cells), 0, mc + ' vs ' + xl.exc_cells);
  let ml = 0; for (let i = 0; i < N; i++) ml += A.MT[i]; add('Material lines', 1, Math.abs(ml - xl.mat_lines), 0, numF(ml) + ' vs ' + numF(xl.mat_lines));
  const T = totals(all); add('Price effect total ($)', 1, Math.abs(T.prc - xl.price_effect), 0.5); add('Volume effect total ($)', 1, Math.abs(T.vol - xl.volume_effect), 0.5);
  // independent re-classification of every line from cents vs the status stored in the CSV
  let bad = 0; for (let i = 0; i < N; i++) { if (statusOf(A.AC[i] - A.TC[i], A.TC[i]) !== A.ST[i] || (isMat(A.AC[i] - A.TC[i], A.TC[i]) ? 1 : 0) !== A.MT[i]) bad++; }
  add('Line status & materiality re-derived from cents (' + numF(N) + ' lines)', N, bad, 0, bad + ' differences');
  return out;
}
let RECON = [];
function tabData(root) {
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const okN = RECON.filter(r => r.pass).length;
  grid.append(card({ id: 'dt-rec', title: 'Reconciliation: dashboard vs Excel workbook', sub: D.xl ? `${okN} of ${RECON.length} checks pass · recomputed from the line data when this page loaded` : 'Workbook not found - nothing to reconcile', span: 's7', build: b => {
    if (!D.xl) { b.append(h('div', { class: 'note' }, 'Place Sales_Variance_Analysis.xlsx next to the build script (and pip install openpyxl) to enable reconciliation.')); return; }
    b.append(h('div', { class: 'tbl-wrap' }, h('table', { class: 't' }, h('thead', null, h('tr', null, ['', 'Check', 'Items', 'Largest difference'].map(x => h('th', null, x)))), h('tbody', null, RECON.map(r => h('tr', null, h('td', { class: r.pass ? 'pass' : 'fail' }, r.pass ? '✔' : '✖'), h('td', { style: { whiteSpace: 'normal' } }, r.name), h('td', { class: 'r' }, numF(r.items)), h('td', { class: 'r' }, r.detail || (r.maxDiff === 0 ? 'identical' : r.maxDiff.toFixed(4)))))))));
  } }));
  const cnt = k => { const o = {}; D.xl && D.xl.validation.forEach(r => o[r[0]] = (o[r[0]] || 0) + 1); return o[k] || 0; };
  grid.append(card({ id: 'dt-prof', title: 'Dataset profile', span: 's5', build: b => {
    const rep = Array.from(A.QR).reduce((a, c) => a + c, 0), notes = Array.from(A.NT).filter(v => v > 0).length;
    const kv = [['Source file', D.source], ['Sales lines', numF(N)], ['Period', MLONG[0] + ' – ' + MLONG[NM - 1] + ` (${NM} months)`], ['Regions / categories / products', `${NR} / ${NC} / ${NP}`], ['Lines with a quantity repaired by the cleaning step', `${numF(rep)} (${pctF(rep / N * 100, 1)})`], ['Lines carrying an operational note', numF(notes)],
      ['Favorable / Unfavorable thresholds', `> +${CFG.favPct}% / < ${MINUS}${Math.abs(CFG.unfPct)}%`], ['Material rule', `|var %| ≥ ${CFG.matPctV}% or |var $| ≥ ${usd(CFG.matUsd)}`], ['Arithmetic', 'Integer cents; percentages rounded half away from zero'], ['Python analysis checks', D.xl ? `${cnt('PASS')} pass · ${cnt('WARN')} warnings · ${cnt('FAIL')} fail` : 'n/a'], ['Dashboard built', D.built], ['Data nature', D.note]];
    b.append(h('table', { class: 't' }, h('tbody', null, kv.map(([k, v]) => h('tr', null, h('td', { style: { color: 'var(--ink2)', whiteSpace: 'normal' } }, k), h('td', { style: { whiteSpace: 'normal', fontWeight: 600 } }, v))))));
  } }));
  if (D.xl) {
    grid.append(card({ id: 'dt-val', title: 'Validation log from the Python analysis', sub: 'Every check the analysis script ran (PASS / WARN / FAIL / INFO)', span: 's7', build: b => b.append(tableView({ name: 'validation_log', rows: D.xl.validation.map(r => ({ s: r[0], c: r[1], d: r[2] })), search: true, pageSize: 20,
      cols: [{ k: 's', label: 'Status', fmt: v => h('span', { class: v === 'PASS' ? 'pass' : v === 'FAIL' ? 'fail' : v === 'WARN' ? 'warnc' : '' }, v) }, { k: 'c', label: 'Check', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }, { k: 'd', label: 'Detail', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }] })) }));
    grid.append(card({ id: 'dt-dic', title: 'Metric dictionary', sub: 'How every measure is defined', span: 's5', build: b => b.append(tableView({ name: 'metric_dictionary', rows: D.xl.dictionary.map(r => ({ m: r[0], d: r[1], n: r[2] })), search: true, pageSize: 12,
      cols: [{ k: 'm', label: 'Metric', fmt: v => h('strong', null, v) }, { k: 'd', label: 'Definition', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }, { k: 'n', label: 'Notes', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }] })) }));
  }
}
/* ===========================================================================
   BOOT
   =========================================================================== */
const TABS = [['overview', 'Overview', tabOverview], ['drivers', 'Drivers', tabDrivers], ['trends', 'Trends', tabTrends], ['regcat', 'Region × Category', tabRegCat], ['products', 'Products', tabProducts], ['exceptions', 'Exceptions & Lines', tabExceptions], ['insights', 'Insights & Recommendations', tabInsights], ['data', 'Data & Validation', tabData]];
let ACTIVE = 'overview';
function renderTabsNav() {
  const nav = $('#tabs'); nav.replaceChildren(...TABS.map(([id, label]) => h('button', { class: 'tab', role: 'tab', 'aria-selected': id === ACTIVE, onclick: () => { ACTIVE = id; renderTabsNav(); renderView(); window.scrollTo(0, Math.min(scrollY, $('#tabs').offsetTop - 60)); } }, label)));
}
function renderView() {
  const y = scrollY; pending = []; hideTip();
  const root = h('div'); const t = TABS.find(x => x[0] === ACTIVE);
  try { t[2](root); } catch (e) { console.error(e); root.append(h('div', { class: 'note' }, 'Something went wrong drawing this tab: ' + e.message)); }
  $('#view').replaceChildren(root); flushPending(); window.scrollTo(0, y);
  const c = $('#cnt'); if (c) c.textContent = `${numF(maskCount(getMask()))} of ${numF(N)} sales lines in view` + (isFiltered() ? '' : ' (no filters)');
}
function renderAll() { buildFilters(); renderView(); }
function exportLines() {
  const m = getMask(), head = ['Source_Excel_Row', 'Month', 'Region', 'Category', 'Product', 'Target_Qty', 'Actual_Qty', 'Unit_Price_USD', 'Target_Sales_USD', 'Actual_Sales_USD', 'Sales_Variance_USD', 'Sales_Variance_Pct', 'Sales_Variance_Status', 'Material_Variance_Flag', 'Material_Trigger', 'Volume_Effect_USD', 'Price_Effect_USD', 'Notes'], rows = [];
  for (let i = 0; i < N; i++) { if (!m[i]) continue; const vc = A.AC[i] - A.TC[i]; rows.push([A.SRC[i], MON[A.M[i]], REG[A.R[i]], CAT[DIM.prodCat[A.P[i]]], PRD[A.P[i]], A.TQ[i], A.AQ[i], (A.UP[i] / 100).toFixed(2), (A.TC[i] / 100).toFixed(2), (A.AC[i] / 100).toFixed(2), (vc / 100).toFixed(2), pct2(vc, A.TC[i]), ST_NAME[A.ST[i]], A.MT[i] ? 'Yes' : 'No', ['', 'Pct', 'USD', 'Pct and USD'][A.TRG[i]], (A.VOL[i] / 1e4).toFixed(4), (A.PRC[i] / 1e4).toFixed(4), NOTES[A.NT[i]]]); }
  downloadCSV('sales_variance_filtered_lines.csv', head, rows);
}
function init() {
  let th = null; try { th = localStorage.getItem('svd-theme'); } catch (e) { } if (th) document.documentElement.setAttribute('data-theme', th);
  document.title = D.title; $('#ttl').textContent = D.title;
  $('#subttl').textContent = `${numF(N)} sales lines · ${NR} regions · ${NC} categories · ${NP} products · ${MLONG[0]} – ${MLONG[NM - 1]} · ${D.note}`;
  $('#foot').textContent = `Built ${D.built} from ${D.source}. All figures are computed in your browser in exact cents; nothing leaves your computer.`;
  $('#btnExport').addEventListener('click', exportLines); $('#btnPrint').addEventListener('click', () => window.print());
  $('#btnTheme').addEventListener('click', () => { setTheme(curTheme() === 'dark' ? 'light' : 'dark'); renderAll(); });
  RECON = reconcile(); const rb = $('#recon');
  if (!D.xl) { rb.textContent = 'Workbook not found: no reconciliation'; rb.className = 'badge note'; }
  else { const ok = RECON.filter(r => r.pass).length; rb.textContent = (ok === RECON.length ? '✔ ' : '✖ ') + `Reconciled to workbook: ${ok}/${RECON.length} checks`; rb.className = 'badge ' + (ok === RECON.length ? 'ok' : 'bad'); rb.style.cursor = 'pointer'; rb.title = 'Open the Data & Validation tab'; rb.addEventListener('click', () => { ACTIVE = 'data'; renderTabsNav(); renderView(); }); }
  renderTabsNav(); renderAll();
  let rt = null; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(renderView, 160); });
  addEventListener('keydown', e => { if (e.key === 'Escape') hideTip(); });
  window.__dash = { S, getMask, totals, agg, dv, A, N, REG, CAT, PRD, MON, change, resetFilters, selectMonth, pickRegion, pickCat, pickProd, RECON, headlineClaims, liveInsights, setTab: t => { ACTIVE = t; renderTabsNav(); renderView(); }, renderAll, statusOf, isMat, pct2, persistenceRows, monthSeries, maskAll };
}
try { init(); } catch (e) { console.error(e); document.body.prepend(h('pre', { style: { color: 'red', padding: '20px', whiteSpace: 'pre-wrap' } }, 'Dashboard failed to start: ' + e.message + '\n' + e.stack)); }
</script>
</body>
</html>
'''




# ============================================================================
#  MAIN
# ============================================================================
def main():
    print("=" * 78 + "\n SALES VARIANCE DASHBOARD BUILDER\n" + "=" * 78)
    folder = find_dir()
    print(f"   Folder: {folder}")
    csv_path = folder / CSV_NAME
    if not csv_path.is_file():
        die(f"{CSV_NAME} not found in {folder}")
    df = pd.read_csv(csv_path, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
    df.columns = [c.strip() for c in df.columns]
    for _c in ("Notes", "Material_Trigger", "Cleaning_Notes"):
        if _c in df.columns:
            df[_c] = df[_c].fillna("").astype(str)
    miss = [c for c in REQUIRED if c not in df.columns]
    if miss:
        die(f"{CSV_NAME} is missing columns {miss}. Run sales_variance_analysis.py first (it creates this file).")
    n = len(df)
    log("INFO", f"Loaded {n:,} sales lines from {CSV_NAME}")
    if n == 0:
        die("the CSV has no rows")
    num_cols = ["Target_Qty", "Actual_Qty", "Unit_Price_USD", "Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD",
                "Gross_Margin_USD", "Volume_Effect_USD", "Price_Effect_USD", "PV_Residual_USD", "Source_Excel_Row"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    bad = {c: int(df[c].isna().sum()) for c in num_cols + ["Region", "Category", "Product", "Year_Month"] if df[c].isna().any()}
    if bad:
        die(f"nulls / non-numeric values in {bad}")
    for c in ("Region", "Category", "Product", "Year_Month", "Notes", "Sales_Variance_Status",
              "Material_Variance_Flag", "Material_Trigger"):
        df[c] = df[c].astype(str).str.strip()
    log("PASS", "No nulls in any column the dashboard uses")

    # ---- dimensions -----------------------------------------------------------
    regions = sorted(df["Region"].unique())
    cats = sorted(df["Category"].unique())
    prod_cat = df.groupby("Product")["Category"].agg(lambda s: sorted(set(s)))
    if (prod_cat.map(len) != 1).any():
        die("a product appears in more than one category - cannot build the product filter")
    prods = sorted(df["Product"].unique(), key=lambda p: (prod_cat[p][0], p))
    months = sorted(df["Year_Month"].unique())
    notes = [""] + sorted(x for x in df["Notes"].unique() if x != "")
    ridx = {v: i for i, v in enumerate(regions)}; cidx = {v: i for i, v in enumerate(cats)}
    pidx = {v: i for i, v in enumerate(prods)}; midx = {v: i for i, v in enumerate(months)}
    nidx = {v: i for i, v in enumerate(notes)}
    mm = np.array([int(m[:4]) * 12 + int(m[5:7]) - 1 for m in months])
    if len(months) > 1 and not (np.diff(mm) == 1).all():
        log("WARN", "months are not consecutive - month-over-month metrics will skip the gaps")

    # ---- line-level arrays (all money as integers) -----------------------------------
    tc = to_cents(df["Target_Sales_USD"]); ac = to_cents(df["Actual_Sales_USD"]); vc = ac - tc
    st_py, mat_py = classify(vc, tc)
    st_csv = df["Sales_Variance_Status"].map({s: i for i, s in enumerate(STATUSES)})
    if st_csv.isna().any():
        die("unknown values in Sales_Variance_Status: " + str(sorted(df.loc[st_csv.isna(), 'Sales_Variance_Status'].unique())))
    bad_st = int((st_csv.values != st_py).sum())
    mat_csv = (df["Material_Variance_Flag"] == "Yes").values
    bad_mat = int((mat_csv != mat_py).sum())
    if bad_st or bad_mat:
        log("WARN", f"CSV vs independent re-classification: {bad_st} status and {bad_mat} materiality differences "
                    f"(the dashboard recomputes its own classification from cents)")
    else:
        log("PASS", f"Status and materiality of all {n:,} lines re-derived from cents = CSV flags")
    trg = df["Material_Trigger"].map(TRIGGERS)
    if trg.isna().any():
        die("unknown Material_Trigger values")
    q_rep = df["Cleaning_Notes"].astype(str).str.contains("quantity", case=False).astype(int) if "Cleaning_Notes" in df.columns else pd.Series(0, index=df.index)
    mls = np.rint((df["Actual_Qty"] - df["Target_Qty"]) * (df["Unit_Price_USD"] - df["Unit_Cost_USD"]) * 1e4).astype("int64")
    rows = {
        "R": df["Region"].map(ridx).tolist(), "C": df["Category"].map(cidx).tolist(),
        "P": df["Product"].map(pidx).tolist(), "M": df["Year_Month"].map(midx).tolist(),
        "TQ": df["Target_Qty"].astype(int).tolist(), "AQ": df["Actual_Qty"].astype(int).tolist(),
        "UP": to_cents(df["Unit_Price_USD"]).tolist(), "UC": to_cents(df["Unit_Cost_USD"]).tolist(),
        "TC": tc.tolist(), "AC": ac.tolist(),
        "VOL": np.rint(df["Volume_Effect_USD"] * 1e4).astype("int64").tolist(),
        "PRC": np.rint(df["Price_Effect_USD"] * 1e4).astype("int64").tolist(),
        "RES": np.rint(df["PV_Residual_USD"] * 1e4).astype("int64").tolist(),
        "MLS": mls.tolist(), "GM": to_cents(df["Gross_Margin_USD"]).tolist(),
        "ST": st_csv.astype(int).tolist(), "MT": mat_csv.astype(int).tolist(), "TRG": trg.astype(int).tolist(),
        "NT": df["Notes"].map(nidx).tolist(), "SRC": df["Source_Excel_Row"].astype(int).tolist(),
        "QR": q_rep.tolist(),
    }
    # reconcile the integer encodings back to dollars
    chk_t = sum(rows["TC"]); chk_a = sum(rows["AC"])
    if abs(chk_t / 100 - df["Target_Sales_USD"].sum()) > 0.5 or abs(chk_a / 100 - df["Actual_Sales_USD"].sum()) > 0.5:
        die("integer encoding does not reproduce the sales totals")
    log("PASS", f"Totals reproduce from the encoded data: target ${chk_t/100:,.2f}, actual ${chk_a/100:,.2f}, variance ${(chk_a-chk_t)/100:,.2f}")

    # ---- workbook (optional) -----------------------------------------------------------
    xl = None
    xpath = locate(XLSX_NAME, folder)
    if xpath.is_file():
        try:
            X = pd.ExcelFile(xpath)
            P = lambda s: X.parse(s, keep_default_na=False)
            by_reg = P("By_Region"); by_cat = P("By_Category"); by_prod = P("By_Product"); by_m = P("By_Year_Month")
            pers = P("Persistence_Region_Category"); exc = P("Exceptions_RegCat_Month"); matv = P("Material_Variances")
            by_year = P("By_Year"); pvm = P("PVM_Bridge")
            val = P("Validation_Summary"); dic = P("Metric_Dictionary"); exs = P("Executive_Summary")
            tot = pvm[pvm["Level"] == "Total"].iloc[0]
            xl = {
                "region": {r.Region: round(float(r.Sales_Variance_USD), 2) for r in by_reg.itertuples()},
                "category": {r.Category: round(float(r.Sales_Variance_USD), 2) for r in by_cat.itertuples()},
                "product": {r.Product: round(float(r.Sales_Variance_USD), 2) for r in by_prod.itertuples()},
                "month": {r.Year_Month: {"var": round(float(r.Sales_Variance_USD), 2), "pct": float(r.Sales_Variance_Pct),
                                         "ts": float(r.Tracking_Signal), "ytd": float(r.YTD_Variance_USD),
                                         "r3": (None if r.Rolling3M_Variance_USD == "" else float(r.Rolling3M_Variance_USD)),
                                         "mom": (None if r.MoM_Actual_Growth_Pct == "" else float(r.MoM_Actual_Growth_Pct)),
                                         "yoy": (None if r.YoY_Actual_Growth_Pct == "" else float(r.YoY_Actual_Growth_Pct))}
                          for r in by_m.itertuples()},
                "year": {str(int(r.Year)): {"var": round(float(r.Sales_Variance_USD), 2), "pct": float(r.Sales_Variance_Pct)} for r in by_year.itertuples()},
                "chronic": int((pers["Chronic_Shortfall"] == "Yes").sum()), "pairs": int(len(pers)),
                "exc_cells": int(len(exc)), "mat_lines": int(len(matv)),
                "price_effect": round(float(tot["Price_Effect_USD"]), 2), "volume_effect": round(float(tot["Volume_Effect_USD"]), 2),
                "pers": {f"{r.Region}|{r.Category}": {"below": int(r.Months_Below_Target), "worst": r.Worst_Month} for r in pers.itertuples()},
                "validation": [[str(a), str(b), str(c)] for a, b, c in val[["Status", "Check", "Detail"]].itertuples(index=False)],
                "dictionary": [[str(a), str(b), str(c)] for a, b, c in dic[["Metric", "Definition", "Notes"]].itertuples(index=False)],
                "summary": [[str(a), str(b), str(c)] for a, b, c in exs[["Item", "Value", "Comment"]].itertuples(index=False)],
            }
            log("PASS", f"Read {XLSX_NAME}: {len(X.sheet_names)} sheets (reconciliation + validation log enabled)")
        except Exception as e:  # noqa
            log("WARN", f"Could not read {XLSX_NAME} ({type(e).__name__}: {e}). Install openpyxl (pip install openpyxl) for the workbook tabs.")
    else:
        log("WARN", f"{XLSX_NAME} not found - workbook reconciliation and validation tabs will be hidden")

    # ---- markdown (optional) -------------------------------------------------------------
    md_html = None
    mpath = locate(MD_NAME, folder)
    if mpath.is_file():
        md_html = md_to_html(mpath.read_text(encoding="utf-8"))
        log("PASS", f"Read {MD_NAME} ({len(md_html):,} characters of HTML)")
    else:
        log("WARN", f"{MD_NAME} not found - Insights tab will show live insights only")

    payload = {
        "title": DASHBOARD_TITLE, "built": _dt.datetime.now().strftime("%d %b %Y %H:%M"),
        "source": CSV_NAME, "note": PRIVACY_NOTE,
        "cfg": {"fav": int(round(FAV_PCT * 1000)), "unf": int(round(UNF_PCT * 1000)), "matPct": int(round(MAT_PCT * 1000)),
                "matUsdC": int(round(MAT_USD * 100)), "favPct": FAV_PCT, "unfPct": UNF_PCT, "matPctV": MAT_PCT, "matUsd": MAT_USD},
        "dims": {"regions": regions, "cats": cats, "prods": prods, "prodCat": [cidx[prod_cat[p][0]] for p in prods],
                 "months": months, "notes": notes, "statuses": STATUSES},
        "rows": rows, "xl": xl, "md": md_html,
    }
    js = json.dumps(payload, separators=(",", ":"), ensure_ascii=True).replace("</", "<\\/")
    page = TEMPLATE.replace("__TITLE__", html.escape(DASHBOARD_TITLE)).replace("__DATA_JSON__", js)
    out = out_dir_for(folder) / OUT_NAME
    out.write_text(page, encoding="utf-8")
    kb = out.stat().st_size / 1024
    n_pass = sum(1 for k, _ in BUILD_LOG if k == "PASS"); n_warn = sum(1 for k, _ in BUILD_LOG if k == "WARN")
    print("\n" + "=" * 78)
    print(f" DONE: {out}   ({kb/1024:,.1f} MB)   build checks: {n_pass} PASS, {n_warn} WARN")
    print(" Open the file in Chrome / Edge / Firefox / Safari (works offline).")
    print("=" * 78)


if __name__ == "__main__":
    main()
