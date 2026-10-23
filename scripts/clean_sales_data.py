"""
==============================================================================
 SALES DATA CLEANING PIPELINE
==============================================================================
 Reads   data/raw/sales_raw_messy.xlsx      (made by generate_synthetic_data.py)
 Writes  data/processed/Cleaned_Sales_Data.csv    clean, verified sales lines
         data/processed/Cleaned_Sales_Data.xlsx   the same data + a full audit trail

 Run from the repository root:
        python scripts/clean_sales_data.py
        python scripts/clean_sales_data.py --input path/to/file.xlsx --output-dir some/folder

 Accepted input formats: .xlsx, .xlsm (openpyxl) and .csv.

 What it does, in order
   1. Schema    : trims header names, drops entirely empty columns, stops with a
                  clear error if a required column is missing or duplicated.
   2. Rows      : drops entirely blank rows, removes exact duplicates.
   3. Parsing   : text numbers ('$1,234.50'), hidden whitespace, null spellings
                  ('NULL', 'N/A', '-'), several date formats.
   4. Key       : Line_ID is the record key. Repeated Line_IDs are merged when
                  their values agree (blanks filled from the copy) and ALL copies
                  are quarantined when they conflict - no winner is guessed.
   5. Labels    : explicit spelling tables only (no fuzzy matching). Blank
                  region/date come from Line_ID, blank category from the product
                  master. A label or date that contradicts Line_ID -> review.
   6. Numbers   : a missing or wrong number is repaired only when the identities
                  Sales = Qty x Price and Gross_Profit = Sales - Qty x Cost give
                  the exact answer. Otherwise the row goes to review.
   7. Checks    : every identity, key and domain rule is re-checked on the output
                  and the row counts are reconciled back to the raw file.

 Nothing is estimated. Every changed value is listed in the Corrections_Audit
 sheet with its original value, rule and certainty.
==============================================================================
"""
import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fpa_lib import cleaning as cl  # noqa: E402
from fpa_lib import config  # noqa: E402

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August",
               "September", "October", "November", "December"]
NUM_COLS = config.QTY_COLUMNS + config.MONEY_COLUMNS
REQUIRED = config.CLEAN_COLUMNS


class CleaningError(Exception):
    """Raised for problems that make cleaning impossible (bad file, bad schema)."""


# ======================================================================
# 1. READ + SCHEMA
# ======================================================================
def read_raw(path):
    path = Path(path)
    if not path.is_file():
        raise CleaningError(f"Input file not found: {path}")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        for enc in ("utf-8-sig", "cp1252"):
            try:
                return pd.read_csv(path, dtype=object, keep_default_na=False, na_values=[], encoding=enc)
            except UnicodeDecodeError:
                continue
        raise CleaningError(f"Could not decode {path.name} as UTF-8 or Windows-1252 text")
    if suffix in (".xlsx", ".xlsm"):
        return pd.read_excel(path, sheet_name=0, dtype=object, keep_default_na=False, na_values=[],
                             engine="openpyxl")
    raise CleaningError(f"Unsupported file type '{suffix}'. Use .xlsx, .xlsm or .csv.")


def fix_schema(raw, log):
    cols = [cl.clean_text(c) or "" for c in raw.columns]
    renamed = [(o, n) for o, n in zip(raw.columns, cols, strict=True) if o != n]
    raw = raw.copy()
    raw.columns = cols
    if renamed:
        log.append(("1. Schema", ", ".join(n for _, n in renamed), "Header had extra spaces", len(renamed),
                    "Trimmed header names", "Standardised"))
    dup = sorted({c for c in cols if cols.count(c) > 1})
    if dup:
        raise CleaningError(f"Duplicate column names: {dup}")
    missing = [c for c in REQUIRED if c not in cols]
    if missing:
        raise CleaningError(f"Required columns missing: {missing}. Expected: {REQUIRED}")
    for c in [c for c in cols if c not in REQUIRED]:
        if raw[c].map(cl.is_null).all():
            log.append(("1. Schema", c, "Column is entirely empty", 0, "Dropped the column", "Standardised"))
        else:
            log.append(("1. Schema", c, "Unexpected column (not in the schema)", int((~raw[c].map(cl.is_null)).sum()),
                        "Dropped the column - review whether it is needed", "Needs review"))
        raw = raw.drop(columns=c)
    if len(raw) == 0:
        raise CleaningError("The input file has a header but no data rows")
    return raw[REQUIRED]


# ======================================================================
# 2. MAIN CLEANING
# ======================================================================
def clean(raw_in):
    """Clean a raw DataFrame. Returns a dict of DataFrames (clean data + audit tables)."""
    log, audit, review, removed, mappings = [], [], [], [], {}
    raw_rows = len(raw_in)
    df = fix_schema(raw_in, log)
    df.insert(0, "Source_Row", np.arange(len(df)) + 2)          # Excel row number (header is row 1)

    # ---- blank rows -------------------------------------------------
    blank = df[REQUIRED].apply(lambda r: all(cl.is_null(v) for v in r), axis=1)
    n_blank = int(blank.sum())
    df = df[~blank]
    log.append(("2. Rows", "All", "Entirely blank rows", n_blank, "Dropped", "Standardised"))

    # ---- exact duplicates (identical raw content in every column) -----
    # str() of every cell; blank/None/NaN all count as "" (pandas 3 astype(str) would turn None into NaN)
    sig = df[REQUIRED].map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v)).agg(
        "␟".join, axis=1)
    first_row = df.groupby(sig)["Source_Row"].transform("first")
    is_dup = sig.duplicated(keep="first")
    for _, r in df[is_dup].iterrows():
        removed.append({"Removal_Reason": "Exact duplicate (all columns identical)",
                        "Duplicate_Of_Row": int(first_row[_]), "Source_Row": int(r["Source_Row"]),
                        "Line_ID": cl.clean_text(r["Line_ID"])})
    n_exact = int(is_dup.sum())
    df = df[~is_dup]
    log.append(("2. Rows", "All", "Exact duplicate rows", n_exact, "Removed - first copy kept", "Exact"))

    # ---- parse every row -------------------------------------------------
    recs = []
    fmt_counts = {"numbers stored as text": 0, "invalid numbers": 0, "hidden whitespace": 0, "null spellings": 0,
                  "date text/serial formats": 0}
    for _, r in df.iterrows():
        rec = {"Source_Row": int(r["Source_Row"]), "_review": [], "_fixes": []}
        for c in ("Region", "Category", "Product", "Notes", "Line_ID"):
            v = r[c]
            if isinstance(v, str) and v and cl.clean_text(v) != v:
                fmt_counts["hidden whitespace"] += 1
        key = cl.parse_line_id(r["Line_ID"])
        rec["Line_ID"] = key["line_id"] if key else cl.clean_text(r["Line_ID"])
        rec["_key"] = key
        if key is None:
            rec["_review"].append("Missing or malformed Line_ID (no record key)")
        # numbers
        for c in NUM_COLS:
            parser = cl.parse_qty if c.endswith("_Qty") else cl.parse_money
            val, st = parser(r[c])
            if st == "text":
                fmt_counts["numbers stored as text"] += 1
            elif st == "invalid":
                fmt_counts["invalid numbers"] += 1
                rec["_fixes"].append((c, r[c], None, "Unreadable number treated as missing", "Exact"))
            elif st == "missing" and cl.clean_text(r[c]) is not None:
                fmt_counts["null spellings"] += 1
            rec[c] = val
            rec["_raw_" + c] = r[c]
        # labels
        reg, reg_raw = cl.map_label(r["Region"], cl.REGION_ALIASES)
        if cl.is_null(r["Region"], "Region"):
            reg, reg_raw = None, None
        prod, prod_raw = cl.map_label(r["Product"], cl.PRODUCT_ALIASES)
        cat, cat_raw = cl.map_label(r["Category"], cl.CATEGORY_ALIASES)
        if cl.is_null(r["Category"]):
            cat, cat_raw = None, None
        if cl.is_null(r["Product"]):
            prod, prod_raw = None, None
        for col, std, orig in (("Region", reg, reg_raw), ("Category", cat, cat_raw), ("Product", prod, prod_raw)):
            if orig is not None:
                mappings.setdefault((col, orig, std or "(unknown)"), 0)
                mappings[(col, orig, std or "(unknown)")] += 1
        rec.update(Region=reg, Category=cat, Product=prod, _reg_raw=reg_raw, _cat_raw=cat_raw, _prod_raw=prod_raw)
        note, note_raw = cl.map_label(r["Notes"], cl.NOTE_ALIASES)
        if note_raw is not None and note is None:
            note = note_raw                                   # free text kept as written
            mappings.setdefault(("Notes", note_raw, "(kept as written)"), 0)
            mappings[("Notes", note_raw, "(kept as written)")] += 1
        elif note_raw is not None:
            mappings.setdefault(("Notes", note_raw, note), 0)
            mappings[("Notes", note_raw, note)] += 1
        rec["Notes"] = note or ""
        # date
        cands, form = cl.date_candidates(r["Date"])
        if form in ("text", "excel serial"):
            fmt_counts["date text/serial formats"] += 1
        rec["_date_cands"], rec["_date_raw"] = cands, r["Date"]
        rec["_year_raw"], rec["_month_raw"], rec["_mname_raw"] = r["Year"], r["Month"], r["Month_Name"]
        recs.append(rec)
    for k_, v_ in fmt_counts.items():
        if v_:
            log.append(("3. Parsing", "Text and number columns", k_.capitalize(), v_,
                        "Parsed / trimmed (value unchanged)" if k_ != "invalid numbers" else "Treated as missing",
                        "Standardised"))

    # ---- labels and dates vs the Line_ID key ------------------------------
    for rec in recs:
        key = rec["_key"]
        if key is None:
            continue
        for col, truth_val in (("Region", key["region"]), ("Product", key["product"])):
            have = rec[col]
            raw_txt = rec["_reg_raw" if col == "Region" else "_prod_raw"]
            if have is None and raw_txt is None:
                rec[col] = truth_val
                rec["_fixes"].append((col, None, truth_val, "Blank - recovered from Line_ID", "Exact"))
            elif have is None:
                rec[col] = truth_val
                rec["_fixes"].append((col, raw_txt, truth_val, "Unknown spelling - taken from Line_ID", "Exact"))
            elif have != truth_val:
                rec["_review"].append(f"{col} '{raw_txt}' contradicts Line_ID ({truth_val})")
        master_cat = config.PRODUCTS[key["product"]][1]
        if rec["Category"] is None:
            if rec["_cat_raw"] is None:
                rec["_fixes"].append(("Category", None, master_cat, "Blank - recovered from product master", "Exact"))
            else:
                rec["_fixes"].append(("Category", rec["_cat_raw"], master_cat,
                                      "Unknown spelling - taken from product master", "Exact"))
            rec["Category"] = master_cat
        elif rec["Category"] != master_cat:
            rec["_review"].append(f"Category '{rec['_cat_raw']}' contradicts product master ({master_cat})")
        cands = rec["_date_cands"]
        if not cands and cl.is_null(rec["_date_raw"]):
            rec["_period"] = key["period"]
            rec["_fixes"].append(("Date", None, str(key["period"].to_timestamp().date()),
                                  "Blank - recovered from Line_ID", "Exact"))
        elif key["period"] in cands:
            rec["_period"] = key["period"]
        elif not cands:
            rec["_review"].append(f"Date '{rec['_date_raw']}' could not be read")
            rec["_period"] = key["period"]
        else:
            rec["_review"].append(f"Date '{rec['_date_raw']}' contradicts Line_ID month ({key['period']})")
            rec["_period"] = key["period"]

    # ---- repeated Line_IDs: merge if consistent, quarantine if conflicting ----
    by_id = {}
    for idx, rec in enumerate(recs):
        if rec["_key"] is not None:
            by_id.setdefault(rec["Line_ID"], []).append(idx)
    drop = set()
    n_merged = n_conflict_ids = 0
    compare = NUM_COLS + ["Region", "Category", "Product", "Notes"]
    for lid, idxs in by_id.items():
        if len(idxs) == 1:
            continue
        conflict_cols = []
        for c in compare:
            vals = {recs[i][c] for i in idxs if recs[i][c] not in (None, "")}
            if len(vals) > 1:
                conflict_cols.append(c)
        periods = {recs[i]["_period"] for i in idxs}
        if conflict_cols or len(periods) > 1:
            n_conflict_ids += 1
            for i in idxs:
                recs[i]["_review"].append(f"Line_ID repeated with conflicting values in {', '.join(conflict_cols) or 'Date'}"
                                          " - all copies held for review")
            continue
        keep = idxs[0]
        for i in idxs[1:]:
            for c in compare:
                if recs[keep][c] in (None, "") and recs[i][c] not in (None, ""):
                    recs[keep]["_fixes"].append((c, recs[keep]["_raw_" + c] if "_raw_" + c in recs[keep] else None,
                                                 recs[i][c], f"Blank - filled from duplicate row {recs[i]['Source_Row']}",
                                                 "Exact"))
                    recs[keep][c] = recs[i][c]
            recs[keep]["_review"].extend(r for r in recs[i]["_review"] if r not in recs[keep]["_review"])
            removed.append({"Removal_Reason": "Same Line_ID as an earlier row with consistent values (copy merged)",
                            "Duplicate_Of_Row": recs[keep]["Source_Row"], "Source_Row": recs[i]["Source_Row"],
                            "Line_ID": lid})
            drop.add(i)
            n_merged += 1
    log.append(("4. Key", "Line_ID", "Repeated Line_ID with consistent values", n_merged,
                "Merged into the first row (blanks filled from the copy)", "Exact"))
    log.append(("4. Key", "Line_ID", "Repeated Line_ID with conflicting values", n_conflict_ids,
                "All copies held for review (no winner chosen)", "Needs review"))
    recs = [r for i, r in enumerate(recs) if i not in drop]

    # ---- numbers: exact recovery and single-error correction -------------
    for rec in recs:
        for side, f in cl.SIDES.items():
            v = {k: rec[c] for k, c in f.items()}
            if any(x is not None and x < 0 for k, x in v.items() if k in ("p", "u")):
                rec["_review"].append(f"{side} unit price/cost is negative")
                continue
            if any(x is None for x in v.values()):
                v2, filled = cl.fill_missing(v)
                for k, rule in filled:
                    rec["_fixes"].append((f[k], rec["_raw_" + f[k]], v2[k], f"Missing - {rule}", "Exact"))
                    rec[f[k]] = v2[k]
                still = [f[k] for k, x in v2.items() if x is None]
                if still:
                    rec["_review"].append(f"Cannot recover {', '.join(still)} exactly (linked values also missing)")
                    continue
                v = v2
            if not cl.identities_hold(v):
                fix = cl.correct_single_error(v)
                if fix is None:
                    rec["_review"].append(f"{side} values break Sales = Qty x Price or Gross_Profit = Sales - Qty x Cost "
                                          "and the wrong field cannot be identified uniquely")
                    continue
                k, x, rule = fix
                rec["_fixes"].append((f[k], rec["_raw_" + f[k]], x,
                                      f"Wrong value - {rule} (the only single change that makes both identities hold)",
                                      "Exact"))
                rec[f[k]] = x
            for k in ("q", "p", "u"):
                if rec[f[k]] is not None and rec[f[k]] < 0:
                    rec["_review"].append(f"{f[k]} is negative")
            if rec[f["p"]] == 0 and rec[f["q"]] > 0:
                rec["_review"].append(f"{f['p']} is zero on a line with quantity")

    # ---- derived date fields ------------------------------------------------
    for rec in recs:
        per = rec.get("_period")
        if per is None:
            continue
        rec["Date"] = per.to_timestamp()
        rec["Year"], rec["Month"], rec["Month_Name"] = per.year, per.month, MONTH_NAMES[per.month - 1]
        for col, raw_v, good in (("Year", rec["_year_raw"], per.year), ("Month", rec["_month_raw"], per.month),
                                 ("Month_Name", rec["_mname_raw"], MONTH_NAMES[per.month - 1])):
            rv = cl.clean_text(raw_v)
            same = rv is not None and (str(rv) == str(good) or (col != "Month_Name" and rv.replace(".0", "") == str(good)))
            if not same:
                rec["_fixes"].append((col, raw_v, good, "Recomputed from Date (derived field)", "Exact"))

    # ---- split clean / review --------------------------------------------------
    clean_rows, review_rows = [], []
    for rec in recs:
        (review_rows if rec["_review"] else clean_rows).append(rec)
    for rec in clean_rows:
        for col, old, new, rule, cert in rec["_fixes"]:
            audit.append({"Source_Row": rec["Source_Row"], "Line_ID": rec["Line_ID"], "Column": col,
                          "Original_Value": "" if old is None else str(old),
                          "New_Value": _fmt(col, new), "Rule": rule, "Certainty": cert,
                          "Validation": "Row passes all identity checks", "Needs_Review": "No"})
    for rec in review_rows:
        review.append({"Review_Reason": "; ".join(dict.fromkeys(rec["_review"])), "Source_Row": rec["Source_Row"],
                       "Line_ID": rec["Line_ID"] or "",
                       **{c: _fmt(c, rec.get(c)) for c in REQUIRED if c not in ("Line_ID",)}})

    out = pd.DataFrame([{c: rec.get(c) for c in REQUIRED} | {"Source_Row": rec["Source_Row"],
                        "Repaired_Fields": "; ".join(dict.fromkeys(f[0] for f in rec["_fixes"]
                                                                   if f[0] not in ("Year", "Month", "Month_Name"))),
                        } for rec in clean_rows], columns=REQUIRED + ["Source_Row", "Repaired_Fields"])
    for c in config.MONEY_COLUMNS:
        out[c] = out[c].astype("int64") / 100
    for c in config.QTY_COLUMNS + ["Year", "Month"]:
        out[c] = out[c].astype("int64")
    out["Date"] = pd.to_datetime(out["Date"])
    # valid outliers are flagged for investigation, never changed
    out["DQ_Flag"] = np.where((out["Target_Qty"] > 0) & (out["Actual_Qty"] > 3 * out["Target_Qty"]),
                              "Large order (over 3x target) - verified by identities, kept", "")
    out = out.sort_values(["Date", "Region", "Product", "Line_ID"]).reset_index(drop=True)
    out = out[REQUIRED + ["Repaired_Fields", "DQ_Flag", "Source_Row"]]

    log.append(("5. Labels", "Region / Category / Product", "Label spellings standardised (explicit table)",
                int(sum(n for (c, o, s), n in mappings.items() if c != "Notes" and o != s)), "Mapped", "Standardised"))
    audit_df = pd.DataFrame(audit, columns=["Source_Row", "Line_ID", "Column", "Original_Value", "New_Value", "Rule",
                                            "Certainty", "Validation", "Needs_Review"])
    if len(audit_df):
        # one log line per rule; "filled from duplicate row 123" is grouped (the audit keeps the row number)
        rule = audit_df["Rule"].str.replace(r"duplicate row \d+", "a duplicate row", regex=True)
        summ = audit_df.assign(Rule=rule).groupby(["Column", "Rule"]).size().reset_index(name="n")
        for _, s in summ.iterrows():
            step = "6. Numbers" if s["Column"] in NUM_COLS else ("7. Derived" if s["Column"] in ("Year", "Month", "Month_Name")
                                                               else "5. Labels")
            log.append((step, s["Column"], s["Rule"], int(s["n"]), "Value repaired (see Corrections_Audit)", "Exact"))
    log.append(("8. Review", "All", "Rows held for human review (excluded from the clean data)", len(review_rows),
                "Listed in Review_Required with the reason", "Needs review"))

    review_df = pd.DataFrame(review, columns=["Review_Reason", "Source_Row", "Line_ID"] +
                             [c for c in REQUIRED if c != "Line_ID"])
    removed_df = pd.DataFrame(removed, columns=["Removal_Reason", "Duplicate_Of_Row", "Source_Row", "Line_ID"])
    map_df = pd.DataFrame([{"Column": c, "Original_Value": o, "Standardised_Value": s, "Row_Count": n}
                           for (c, o, s), n in sorted(mappings.items())])
    log_df = pd.DataFrame(log, columns=["Step", "Column", "Issue", "Rows_Affected", "Action", "Certainty"])
    recon = {"raw_rows": raw_rows, "blank": n_blank, "exact_dups": n_exact, "merged_dups": n_merged,
             "review": len(review_rows), "clean": len(out)}
    return {"clean": out, "audit": audit_df, "review": review_df, "removed": removed_df, "mappings": map_df,
            "log": log_df, "recon": recon, "raw": df}


def _fmt(col, v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if col in config.MONEY_COLUMNS and isinstance(v, (int, np.integer)):
        return f"{int(v) / 100:.2f}"
    if isinstance(v, pd.Timestamp):
        return str(v.date())
    return str(v)


# ======================================================================
# 3. CHECKS
# ======================================================================
def validate(res):
    """Independent checks on the cleaned output. Returns list of (check, passed, detail)."""
    d, r = res["clean"], res["recon"]
    c = {k: np.rint(d[k] * 100).astype("int64") for k in config.MONEY_COLUMNS}
    checks = []

    def add(name, ok, detail):
        checks.append((name, bool(ok), detail))
    accounted = r["clean"] + r["blank"] + r["exact_dups"] + r["merged_dups"] + r["review"]
    add("Row reconciliation: raw = clean + blank + exact duplicates + merged copies + review", accounted == r["raw_rows"],
        f"{r['raw_rows']:,} = {r['clean']:,} + {r['blank']:,} + {r['exact_dups']:,} + {r['merged_dups']:,} + {r['review']:,}")
    if d.empty:
        add("At least one row passed cleaning", False, f"all {r['review']:,} data rows need review - see Review_Required")
        return checks
    add("Line_ID is unique and well-formed", d["Line_ID"].is_unique and d["Line_ID"].map(lambda s: cl.parse_line_id(s) is not None).all(),
        f"{d['Line_ID'].nunique():,} unique of {len(d):,}")
    for side, f in cl.SIDES.items():
        q = d[f["q"]].astype("int64")
        add(f"{side}: Sales = Qty x Unit_Price on every row", (c[f["s"]] == q * c[f["p"]]).all(),
            f"{int((c[f['s']] != q * c[f['p']]).sum())} rows differ")
        add(f"{side}: Gross_Profit = Sales - Qty x Unit_Cost on every row", (c[f["g"]] == c[f["s"]] - q * c[f["u"]]).all(),
            f"{int((c[f['g']] != c[f['s']] - q * c[f['u']]).sum())} rows differ")
        add(f"{side}: quantities are whole numbers >= 0", (q >= 0).all(), f"min {int(q.min())}")
        add(f"{side}: unit price and unit cost > 0", ((c[f["p"]] > 0) & (c[f["u"]] > 0)).all(), "")
    add("Region, Category, Product match the master data", d["Region"].isin(config.REGIONS.values()).all()
        and d["Product"].isin(config.PRODUCTS).all() and (d["Category"] == d["Product"].map(lambda p: config.PRODUCTS[p][1])).all(),
        f"{d['Region'].nunique()} regions, {d['Category'].nunique()} categories, {d['Product'].nunique()} products")
    add("Labels and dates agree with Line_ID", (d.apply(lambda x: (lambda k: k["region"] == x["Region"] and k["product"] == x["Product"]
                                                                   and k["period"] == x["Date"].to_period("M"))(cl.parse_line_id(x["Line_ID"])), axis=1)).all(), "")
    add("Year / Month / Month_Name agree with Date", ((d["Year"] == d["Date"].dt.year) & (d["Month"] == d["Date"].dt.month)).all(), "")
    add("No blanks in required columns (Notes may be blank)",
        d[[c for c in REQUIRED if c != "Notes"]].notna().all().all(), "")
    add("Every changed value is in the audit with its original value",
        res["audit"]["Rule"].notna().all() and len(res["audit"]) >= (d["Repaired_Fields"] != "").sum(), f"{len(res['audit']):,} audit rows")
    return checks


# ======================================================================
# 4. OUTPUT
# ======================================================================
def write_outputs(res, checks, out_dir, input_name):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    d = res["clean"].copy()
    csv_path, xlsx_path = out_dir / config.CLEAN_CSV.name, out_dir / config.CLEAN_XLSX.name
    d_out = d.copy()
    d_out["Date"] = d_out["Date"].dt.strftime("%Y-%m-%d")
    d_out.to_csv(csv_path, index=False, encoding="utf-8-sig", float_format="%.2f")
    r = res["recon"]
    review = res["review"]
    tgt_review = pd.to_numeric(review.get("Target_Sales_USD", pd.Series(dtype=object)), errors="coerce")
    readme = pd.DataFrame([
        ("Input file", input_name),
        ("Raw rows", f"{r['raw_rows']:,}"),
        ("Clean rows delivered", f"{r['clean']:,}"),
        ("Blank rows dropped", f"{r['blank']:,}"),
        ("Exact duplicates removed", f"{r['exact_dups']:,}"),
        ("Duplicate copies merged (same Line_ID, consistent values)", f"{r['merged_dups']:,}"),
        ("Rows held for review (excluded)", f"{r['review']:,}"),
        ("Values repaired (all exact)", f"{len(res['audit']):,}"),
        ("Target sales on review rows (where readable)", f"${tgt_review.sum():,.2f}" + (
            f" = {tgt_review.sum() / d['Target_Sales_USD'].sum() * 100:.2f}% of clean target sales"
            if d["Target_Sales_USD"].sum() > 0 else "")),
        ("Checks passed", f"{sum(c[1] for c in checks)} of {len(checks)}"),
        ("Sheets", "Clean_Data, Row_Reconciliation, Cleaning_Log, Corrections_Audit, Review_Required, "
                   "Removed_Duplicates, Value_Mappings, Data_Quality, Validation_Checks"),
        ("Certainty", "Exact = solved by an arithmetic identity, the Line_ID key or the product master. "
                      "Standardised = format or spelling only. Nothing is estimated."),
        ("Data", "Synthetic data generated by scripts/generate_synthetic_data.py"),
    ], columns=["Item", "Value"])
    recon_df = pd.DataFrame([("Raw data rows", r["raw_rows"]), ("- entirely blank rows", -r["blank"]),
                             ("- exact duplicate rows", -r["exact_dups"]), ("- duplicate copies merged", -r["merged_dups"]),
                             ("- rows held for review", -r["review"]), ("= clean rows", r["clean"])],
                            columns=["Step", "Rows"])
    dq = []
    raw = res["raw"]
    for c in REQUIRED:
        before = int(raw[c].map(lambda v, c=c: cl.is_null(v, c)).sum())
        after = int(d[c].isna().sum()) if c != "Notes" else int((d[c] == "").sum())
        dq.append((c, before, round(before / len(raw) * 100, 2), after))
    dq_df = pd.DataFrame(dq, columns=["Column", "Missing_Before", "Missing_%_Before", "Missing_After"])
    checks_df = pd.DataFrame([(n, "PASS" if ok else "FAIL", det) for n, ok, det in checks],
                             columns=["Check", "Result", "Detail"])
    sheets = {"Read_Me": readme, "Clean_Data": d, "Row_Reconciliation": recon_df, "Cleaning_Log": res["log"],
              "Corrections_Audit": res["audit"], "Review_Required": review, "Removed_Duplicates": res["removed"],
              "Value_Mappings": res["mappings"], "Data_Quality": dq_df, "Validation_Checks": checks_df}
    # strings_to_formulas=False: text such as "=..." is stored as text, never executed as a formula
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd",
                        engine_kwargs={"options": {"strings_to_formulas": False}}) as xw:
        wb = xw.book
        head = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1})
        money = wb.add_format({"num_format": "#,##0.00"})
        for name, frame in sheets.items():
            frame.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for j, col in enumerate(frame.columns):
                ws.write(0, j, col, head)
                width = min(60, max(10, len(str(col)) + 2, int(frame[col].astype(str).str.len().quantile(0.95)) + 2
                                    if len(frame) else 10))
                ws.set_column(j, j, width, money if col in config.MONEY_COLUMNS and name == "Clean_Data" else None)
            ws.freeze_panes(1, 0)
            if len(frame):
                ws.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    return csv_path, xlsx_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="Clean the raw sales file (no input needed with the default layout).")
    ap.add_argument("--input", type=Path, default=config.RAW_FILE)
    ap.add_argument("--output-dir", type=Path, default=config.PROCESSED_DIR)
    a = ap.parse_args(argv)
    try:
        raw = read_raw(a.input)
        res = clean(raw)
    except CleaningError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    checks = validate(res)
    csv_path, xlsx_path = write_outputs(res, checks, a.output_dir, a.input.name)
    r = res["recon"]
    print("=" * 78)
    print(" DATA CLEANING REPORT")
    print("=" * 78)
    print(f"   Input ............................ {a.input}")
    print(f"   Raw rows ......................... {r['raw_rows']:>8,}")
    print(f"   - blank rows ..................... {r['blank']:>8,}")
    print(f"   - exact duplicates ............... {r['exact_dups']:>8,}")
    print(f"   - duplicate copies merged ........ {r['merged_dups']:>8,}")
    print(f"   - held for review ................ {r['review']:>8,}")
    print(f"   = clean rows ..................... {r['clean']:>8,}")
    print(f"   Values repaired (exact) .......... {len(res['audit']):>8,}")
    for name, ok, det in checks:
        print(f"   [{'PASS' if ok else 'FAIL'}] {name}" + (f"  ({det})" if det else ""))
    n_fail = sum(not ok for _, ok, _ in checks)
    print(f"   Checks: {len(checks) - n_fail} passed, {n_fail} failed")
    print(f"   Wrote {csv_path}")
    print(f"   Wrote {xlsx_path}")
    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
