"""
==============================================================================
 SYNTHETIC SALES DATA GENERATOR  (deterministic - fixed random seed)
==============================================================================
 Creates the project's raw source file from scratch, so the whole project is
 reproducible without any external or hand-made dataset:

   data/raw/sales_raw_messy.xlsx     the messy file the cleaning step must fix
   data/raw/injected_defects.csv     every defect that was deliberately injected

 Run from the repository root:
        python scripts/generate_synthetic_data.py
        python scripts/generate_synthetic_data.py --seed 7 --output-dir some/folder

 Step 1 builds a CLEAN, internally consistent dataset ("the truth").
 Step 2 copies it and injects documented data-quality defects.
 The same seed always produces the same data (tested in tests/).

 THE DATA IS SYNTHETIC. The business scenario below is an assumption used to
 generate it. The analysis must rediscover the results from the data; nothing
 in the reporting layer reads these settings.

 One row = one customer account's monthly sales of one product in one region,
 identified by Line_ID = <region code>-<product code>-<YYYYMM>-<account no.>
 e.g. LAT-HOM01-202403-004.

 Scenario assumptions (ground truth of the simulation)
 ----------------------------------------------------
 * Targets: every region was planned with the same +5 % annual growth and a
   flattened seasonal profile (planners under-estimated seasonal peaks).
 * Real demand trend differs by region (Asia Pacific and Middle East grow
   faster than plan, Latin America shrinks, Europe is flat).
 * Price: 2025 list-price increases were planned by category. Europe passed on
   only 40 % of them. Promotions (6-12 % discounts in two random months a year
   per region x category, plus an 8-12 % November promotion for Electronics and
   Home Appliances) lift volume (price elasticity 1.6).
 * Cost: Home Appliances unit costs rise ~8 % from March 2025 (freight and
   components); Electronics costs fall ~3 % in 2025.
 * Supply: an Electronics stock-out in North America, Aug-Oct 2024.
 * New accounts opened during the year have no target until the next plan.
==============================================================================
"""
import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fpa_lib import config  # noqa: E402

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August",
               "September", "October", "November", "December"]

# ----------------------------------------------------------- scenario --
REGION_PARAMS = {  # accounts per product (mean), price index, real annual trend, demand level vs plan
    "NAM": dict(accounts=8, price_index=1.05, trend=0.035, level=1.000),
    "EUR": dict(accounts=7, price_index=1.02, trend=0.010, level=0.995),
    "APA": dict(accounts=7, price_index=0.97, trend=0.090, level=1.015),
    "LAT": dict(accounts=4, price_index=0.90, trend=-0.020, level=0.955),
    "MEA": dict(accounts=3, price_index=1.00, trend=0.060, level=1.000),
}
PLANNED_GROWTH = 0.05            # the same planned growth for every region
TARGET_STRETCH = 1.06            # targets were set 6 % above expected baseline demand (stretch)
PLAN_SEASONALITY_SHRINK = 0.6    # planners used 60 % of the true seasonal swing

LIST_PRICE_2024 = {
    "Bluetooth Speaker": 59, "Laptop Stand": 39, "Power Bank": 35, "Smart Watch": 229, "Tablet": 349,
    "Wireless Earbuds": 99, "Air Fryer": 119, "Air Purifier": 189, "Blender": 69, "Coffee Maker": 89,
    "Robot Vacuum": 399, "Toaster": 45, "Ergonomic Chair": 279, "Keyboard": 59, "Monitor Arm": 89,
    "Standing Desk": 449, "Webcam": 79, "Electric Toothbrush": 69, "Face Massager": 49, "Hair Dryer": 79,
    "Straightener": 59, "Trimmer": 39, "Dumbbell Set": 129, "Foam Roller": 29, "Jump Rope": 15,
    "Resistance Bands": 25, "Yoga Mat": 35,
}
CATEGORY_MARGIN = {"Electronics": 0.30, "Home Appliances": 0.33, "Office Supplies": 0.40,
                   "Personal Care": 0.46, "Sports Equipment": 0.42}
PLANNED_PRICE_INCREASE_2025 = {"Electronics": 0.00, "Home Appliances": 0.03, "Office Supplies": 0.04,
                               "Personal Care": 0.03, "Sports Equipment": 0.02}
PLANNED_COST_INFLATION_2025 = 0.02
SEASONALITY = {  # true monthly demand profile, Jan..Dec (normalised to mean 1 below)
    "Electronics": [0.85, 0.80, 0.90, 0.92, 0.95, 0.92, 0.90, 0.95, 1.00, 1.05, 1.35, 1.55],
    "Home Appliances": [0.85, 0.85, 1.00, 1.05, 1.05, 0.95, 0.90, 0.92, 0.98, 1.05, 1.25, 1.35],
    "Office Supplies": [1.15, 1.05, 1.05, 0.95, 0.90, 0.85, 0.85, 1.20, 1.20, 1.00, 0.95, 0.80],
    "Personal Care": [0.90, 0.95, 0.95, 1.00, 1.10, 0.95, 0.90, 0.92, 0.95, 1.00, 1.08, 1.35],
    "Sports Equipment": [1.40, 1.15, 1.10, 1.10, 1.10, 1.05, 0.95, 0.90, 0.85, 0.80, 0.80, 0.85],
}
PROMO_ELASTICITY = 1.6


def month_list():
    start = pd.Period(config.FIRST_MONTH, freq="M")
    return [start + i for i in range(config.N_MONTHS)]


# ======================================================================
# STEP 1 - CLEAN TRUTH
# ======================================================================
def build_truth(seed=config.RANDOM_SEED):
    """Return the clean, internally consistent dataset as a DataFrame."""
    rng = np.random.default_rng(seed)
    months = month_list()
    n_m = len(months)
    seas = {c: np.array(v) / np.mean(v) for c, v in SEASONALITY.items()}
    plan_seas = {c: 1 + PLAN_SEASONALITY_SHRINK * (v - 1) for c, v in seas.items()}
    products = list(config.PRODUCTS)                      # fixed order -> deterministic
    regions = list(config.REGIONS)

    # promotions: per region x category, 2 random promo months per year + November for
    # Electronics and Home Appliances everywhere (holiday promotion)
    promo = {}
    for r in regions:
        for c in config.CATEGORIES:
            depth = np.zeros(n_m)
            for y in (0, 1):
                for m in rng.choice(np.arange(12), size=2, replace=False):
                    depth[y * 12 + m] = rng.uniform(0.06, 0.12)
                if c in ("Electronics", "Home Appliances"):
                    depth[y * 12 + 10] = rng.uniform(0.08, 0.12)
            promo[(r, c)] = depth

    rows = []
    for p in products:
        pcode, cat = config.PRODUCTS[p]
        list24 = LIST_PRICE_2024[p]
        margin_p = CATEGORY_MARGIN[cat] + rng.uniform(-0.03, 0.03)
        base_units = 7000 / list24 ** 0.75                 # units per account per month
        for r in regions:
            rp = REGION_PARAMS[r]
            n_acc = int(rng.poisson(rp["accounts"])) + 1
            for a in range(1, n_acc + 1):
                start = 0 if rng.random() < 0.92 else int(rng.integers(1, 19))
                end = n_m if rng.random() < 0.9 else int(rng.integers(start + 6, n_m + 1))
                tier = rng.choice([0.97, 1.00, 1.02], p=[0.3, 0.5, 0.2])
                size = rng.lognormal(0.0, 0.45)
                for mi in range(start, end):
                    per = months[mi]
                    y = per.year
                    m0 = per.month - 1
                    # ---------------- prices & costs (targets) -------------
                    plan_inc = PLANNED_PRICE_INCREASE_2025[cat] if y == 2025 else 0.0
                    tp = round(list24 * (1 + plan_inc) * rp["price_index"] * tier, 2)
                    tcu = round(list24 * (1 - margin_p) * (1 + (PLANNED_COST_INFLATION_2025 if y == 2025 else 0))
                                * (1.02 if r == "LAT" else 1.0), 2)
                    # ---------------- target quantity ---------------------
                    has_target = start == 0 or (start < 12 and y == 2025)  # new accounts enter next plan
                    t_units = base_units * size * plan_seas[cat][m0] * TARGET_STRETCH * (1 + PLANNED_GROWTH) ** (mi / 12)
                    tq = max(1, int(round(t_units))) if has_target else 0
                    # ---------------- actual price -------------------------
                    d = promo[(r, cat)][mi]
                    realised_inc = plan_inc * (0.4 if r == "EUR" else 1.0)
                    ap = tp / (1 + plan_inc) * (1 + realised_inc) * (1 - d) * (1 + rng.normal(0, 0.006))
                    if r == "APA" and y == 2025:
                        ap *= 1.01
                    ap = round(max(ap, 0.5), 2)
                    # ---------------- actual cost --------------------------
                    shock = 0.0
                    if cat == "Home Appliances" and (y == 2025 and per.month >= 3):
                        shock = 0.08
                    if cat == "Electronics" and y == 2025:
                        shock = -0.03
                    acu = round(tcu * (1 + shock + rng.normal(0.003, 0.006)), 2)
                    # ---------------- actual quantity ----------------------
                    demand = (base_units * size * seas[cat][m0] * rp["level"]
                              * (1 + rp["trend"]) ** (mi / 12) * (1 + PROMO_ELASTICITY * d))
                    note = ""
                    if r == "NAM" and cat == "Electronics" and str(per) in ("2024-08", "2024-09", "2024-10"):
                        demand *= 0.70
                        if rng.random() < 0.5:
                            note = "Stock-out"
                    noise = rng.lognormal(0.0, 0.10)
                    if rng.random() < 0.003:                        # genuine large order (valid outlier)
                        noise *= 5
                    aq = int(round(demand * noise))
                    u = rng.random()
                    if not note:
                        if d > 0 and u < 0.6:
                            note = "Promotion"
                        elif u < 0.005:
                            note, aq = "Late shipment", int(round(aq * 0.85))
                        elif u < 0.009:
                            note, aq = "Returns", int(round(aq * 0.9))
                        elif y == 2025 and per.month == 1 and plan_inc > 0 and u < 0.05:
                            note = "Price increase"
                        elif shock > 0.05 and u < 0.1:
                            note = "Cost increase"
                    rows.append((f"{r}-{pcode}-{per.year}{per.month:02d}-{a:03d}", per, r, cat, p,
                                 tq, aq, tp, ap, tcu, acu, note))

    df = pd.DataFrame(rows, columns=["Line_ID", "Period", "Region", "Category", "Product", "Target_Qty",
                                     "Actual_Qty", "Target_Unit_Price_USD", "Actual_Unit_Price_USD",
                                     "Target_Unit_Cost_USD", "Actual_Unit_Cost_USD", "Notes"])
    df["Date"] = df["Period"].dt.to_timestamp()
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    df["Month_Name"] = df["Month"].map(lambda m: MONTH_NAMES[m - 1])
    df["Region"] = df["Region"].map(config.REGIONS)
    cents = {c: np.rint(df[c] * 100).astype("int64") for c in
             ["Target_Unit_Price_USD", "Actual_Unit_Price_USD", "Target_Unit_Cost_USD", "Actual_Unit_Cost_USD"]}
    ts = df["Target_Qty"] * cents["Target_Unit_Price_USD"]
    as_ = df["Actual_Qty"] * cents["Actual_Unit_Price_USD"]
    df["Target_Sales_USD"] = ts / 100
    df["Actual_Sales_USD"] = as_ / 100
    df["Target_Gross_Profit_USD"] = (ts - df["Target_Qty"] * cents["Target_Unit_Cost_USD"]) / 100
    df["Actual_Gross_Profit_USD"] = (as_ - df["Actual_Qty"] * cents["Actual_Unit_Cost_USD"]) / 100
    df = df.sort_values(["Date", "Region", "Product", "Line_ID"]).reset_index(drop=True)
    return df[config.CLEAN_COLUMNS]


# ======================================================================
# STEP 2 - INJECT DATA-QUALITY DEFECTS
# ======================================================================
REGION_VARIANTS = {  # deliberately messy spellings; every one is in the cleaner's explicit alias table
    "North America": ["North America", "north america", "N. America", "NA", "North-America", "NAM", "NorthAmerica"],
    "Europe": ["Europe", "EUROPE", "EU", "Europa", "Eur", "europe"],
    "Asia Pacific": ["Asia Pacific", "APAC", "Asia-Pacific", "AsiaPac", "asia pacific", "AP"],
    "Latin America": ["Latin America", "LATAM", "LatAm", "Latin-America", "LatinAmerica", "latin america"],
    "Middle East": ["Middle East", "ME", "Middle-East", "MiddleEast", "MIDDLE EAST", "Mid East"],
}
NULL_SPELLINGS = ["", "NULL", "N/A", "-", "nan", "n/a"]
HIDDEN = [" ", "​", " ", "  "]
NOTE_VARIANTS = {"Promotion": ["Promotion", "promo", "PROMOTION", "Promo "], "Stock-out": ["Stock-out", "stock out", "Stockout"],
                 "Late shipment": ["Late shipment", "late shipment", "Late Shipment"], "Returns": ["Returns", "returns", "Return"],
                 "Price increase": ["Price increase", "price increase"], "Cost increase": ["Cost increase", "cost increase"]}

NUMERIC = ["Target_Qty", "Actual_Qty", "Target_Unit_Price_USD", "Actual_Unit_Price_USD", "Target_Unit_Cost_USD",
           "Actual_Unit_Cost_USD", "Target_Sales_USD", "Actual_Sales_USD", "Target_Gross_Profit_USD",
           "Actual_Gross_Profit_USD"]


def _money_text(v, rng):
    s = f"{v:,.2f}" if rng.random() < 0.5 else f"{v:.2f}"
    return rng.choice([f"${s}", s, f" {s} ", f"{s} "])


def corrupt(truth, seed=config.RANDOM_SEED):
    """Return (raw DataFrame, defect manifest DataFrame). Deterministic for a seed."""
    rng = np.random.default_rng(seed + 1)
    raw = truth.copy().astype(object)
    raw["Date"] = truth["Date"].dt.to_pydatetime()
    n = len(raw)
    log = []

    def note(i, col, defect, clean, dirty, expected):
        log.append((raw.at[i, "Line_ID"], col, defect, clean, dirty, expected))

    order = rng.permutation(n)
    # disjoint row pools so that every injected numeric error stays exactly recoverable
    pool = iter(order)

    def take(k):
        return [next(pool) for _ in range(k)]

    k = lambda share: int(round(n * share))  # noqa: E731

    # ---- single missing numeric values (recoverable by an arithmetic identity) ----
    for col, share in [("Actual_Qty", 0.025), ("Actual_Unit_Price_USD", 0.015), ("Target_Unit_Price_USD", 0.010),
                       ("Target_Qty", 0.010), ("Actual_Unit_Cost_USD", 0.015), ("Target_Unit_Cost_USD", 0.010),
                       ("Actual_Sales_USD", 0.010), ("Target_Sales_USD", 0.008),
                       ("Actual_Gross_Profit_USD", 0.015), ("Target_Gross_Profit_USD", 0.010)]:
        for i in take(k(share)):
            if col == "Target_Qty" and truth.at[i, "Target_Qty"] == 0:
                continue                                   # 0 target cannot be rebuilt from 0 sales
            if col in ("Actual_Qty", "Actual_Unit_Price_USD", "Actual_Unit_Cost_USD") and truth.at[i, "Actual_Qty"] == 0:
                continue                                   # a unit value cannot be rebuilt when quantity is 0
            if col in ("Target_Unit_Price_USD", "Target_Unit_Cost_USD") and truth.at[i, "Target_Qty"] == 0:
                continue
            dirty = str(rng.choice(NULL_SPELLINGS))
            note(i, col, "Missing value", truth.at[i, col], dirty, "Recovered (exact)")
            raw.at[i, col] = dirty if dirty else None
    # ---- wrong numbers (one error per row, detectable with two identities) ----
    for i in take(k(0.006)):
        if truth.at[i, "Actual_Qty"] > 0:
            v = -int(truth.at[i, "Actual_Qty"])
            note(i, "Actual_Qty", "Sign error", truth.at[i, "Actual_Qty"], v, "Recovered (exact)")
            raw.at[i, "Actual_Qty"] = v
    for i in take(k(0.005)):
        if truth.at[i, "Actual_Qty"] > 0:
            v = int(truth.at[i, "Actual_Qty"]) * 10
            note(i, "Actual_Qty", "Typo (x10)", truth.at[i, "Actual_Qty"], v, "Recovered (exact)")
            raw.at[i, "Actual_Qty"] = v
    for i in take(k(0.002)):
        if truth.at[i, "Actual_Qty"] > 0:
            note(i, "Actual_Unit_Price_USD", "Zero price", truth.at[i, "Actual_Unit_Price_USD"], 0.0, "Recovered (exact)")
            raw.at[i, "Actual_Unit_Price_USD"] = 0.0
    for i in take(k(0.004)):
        v = truth.at[i, "Target_Sales_USD"]
        s = f"{v:.2f}"
        digits = [j for j in range(len(s) - 1) if s[j].isdigit() and s[j + 1].isdigit() and s[j] != s[j + 1]]
        if digits and truth.at[i, "Target_Qty"] > 1:
            j = digits[0]
            bad = float(s[:j] + s[j + 1] + s[j] + s[j + 2:])
            note(i, "Target_Sales_USD", "Transposed digits", v, bad, "Recovered (exact) or review if ambiguous")
            raw.at[i, "Target_Sales_USD"] = bad
    for i in take(k(0.004)):
        v = truth.at[i, "Actual_Gross_Profit_USD"]
        bad = round(v + float(rng.choice([-1, 1])) * float(rng.integers(1, 9)) * 10 ** int(rng.integers(0, 3)) + 0.07, 2)
        if truth.at[i, "Actual_Qty"] > 1:
            note(i, "Actual_Gross_Profit_USD", "Wrong derived value", v, bad, "Recovered (exact) or review if ambiguous")
            raw.at[i, "Actual_Gross_Profit_USD"] = bad
    # ---- unrecoverable: two linked values missing -> must go to review ----
    for i in take(25):
        for col in ("Actual_Qty", "Actual_Sales_USD"):
            note(i, col, "Missing value (two linked fields)", truth.at[i, col], None, "Review (unrecoverable)")
            raw.at[i, col] = None
    for i in take(15):
        for col in ("Actual_Unit_Cost_USD", "Actual_Gross_Profit_USD"):
            note(i, col, "Missing value (two linked fields)", truth.at[i, col], None, "Review (unrecoverable)")
            raw.at[i, col] = None
    # ---- key / date / label conflicts that need a human -> review ----
    for i in take(12):
        right = truth.at[i, "Region"]
        wrong = str(rng.choice([r for r in config.REGIONS.values() if r != right]))
        note(i, "Region", "Label conflicts with Line_ID", right, wrong, "Review (conflict)")
        raw.at[i, "Region"] = wrong
    for i in take(8):
        d = truth.at[i, "Date"]
        bad = (d + pd.DateOffset(months=1)).to_pydatetime()
        note(i, "Date", "Date conflicts with Line_ID", d.date(), bad.date(), "Review (conflict)")
        raw.at[i, "Date"] = bad
    for i in take(6):
        note(i, "Line_ID", "Missing Line_ID", truth.at[i, "Line_ID"], None, "Review (no key)")
        raw.at[i, "Line_ID"] = None
    others = list(config.PRODUCTS)
    for i in take(8):
        right = truth.at[i, "Product"]
        wrong = str(rng.choice([p for p in others if p != right]))
        note(i, "Product", "Label conflicts with Line_ID", right, wrong, "Review (conflict)")
        raw.at[i, "Product"] = wrong
    # ---- missing labels / dates recoverable from the key or the product master ----
    for i in rng.choice(n, size=k(0.03), replace=False):
        if raw.at[i, "Line_ID"] is None or truth.at[i, "Region"] != raw.at[i, "Region"]:
            continue
        note(i, "Region", "Missing value", truth.at[i, "Region"], None, "Recovered from Line_ID")
        raw.at[i, "Region"] = None
    for i in rng.choice(n, size=k(0.02), replace=False):
        if raw.at[i, "Line_ID"] is None or truth.at[i, "Product"] != raw.at[i, "Product"]:
            continue
        note(i, "Category", "Missing value", truth.at[i, "Category"], None, "Recovered from product master")
        raw.at[i, "Category"] = None
    for i in rng.choice(n, size=k(0.01), replace=False):
        if raw.at[i, "Line_ID"] is None or raw.at[i, "Date"] != truth.at[i, "Date"]:
            continue
        note(i, "Date", "Missing value", truth.at[i, "Date"].date(), None, "Recovered from Line_ID")
        raw.at[i, "Date"] = None
    # ---- formatting noise (standardised by the cleaner, no information change) ----
    for i in range(n):
        reg = raw.at[i, "Region"]
        if isinstance(reg, str) and reg == truth.at[i, "Region"] and rng.random() < 0.55:
            raw.at[i, "Region"] = str(rng.choice(REGION_VARIANTS[reg]))
        for col in ("Category", "Product"):
            v = raw.at[i, col]
            if isinstance(v, str) and rng.random() < 0.12:
                raw.at[i, col] = rng.choice([v.lower(), v.upper(), v.title()])
        nv = raw.at[i, "Notes"]
        if isinstance(nv, str) and nv and rng.random() < 0.5:
            raw.at[i, "Notes"] = str(rng.choice(NOTE_VARIANTS[nv]))
        for col in ("Region", "Category", "Product", "Notes"):
            v = raw.at[i, col]
            if isinstance(v, str) and v and rng.random() < 0.06:
                h = str(rng.choice(HIDDEN))
                raw.at[i, col] = (h + v) if rng.random() < 0.5 else (v + h)
        dt = raw.at[i, "Date"]
        if dt is not None and dt == truth.at[i, "Date"]:
            u = rng.random()
            if u < 0.10:
                raw.at[i, "Date"] = dt.strftime("%Y-%m-%d")
            elif u < 0.15:
                raw.at[i, "Date"] = dt.strftime("%b-%Y")
            elif u < 0.20:
                raw.at[i, "Date"] = dt.strftime("%d/%m/%Y")
            elif u < 0.24:
                raw.at[i, "Date"] = f"{dt.month}/{dt.day}/{dt.year}"
            elif u < 0.27:
                raw.at[i, "Date"] = int((pd.Timestamp(dt) - pd.Timestamp("1899-12-30")).days)
            elif u < 0.29:
                raw.at[i, "Date"] = dt.strftime("%B %Y")
        for col in NUMERIC:
            v = raw.at[i, col]
            if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) and rng.random() < 0.05:
                if col.endswith("_Qty"):
                    raw.at[i, col] = rng.choice([f"{int(v):,}", f" {int(v)} ", f"{int(v)} "])
                else:
                    raw.at[i, col] = _money_text(float(v), rng)
        if rng.random() < 0.01:
            mi = int(truth.at[i, "Month"])
            wrong = MONTH_NAMES[mi % 12]
            note(i, "Month_Name", "Wrong derived value", truth.at[i, "Month_Name"], wrong, "Recomputed from Date")
            raw.at[i, "Month_Name"] = wrong

    raw["_truth_row"] = np.arange(n)
    # ---- duplicates ----
    extra = []
    for i in rng.choice(n, size=k(0.025), replace=False):
        if raw.at[i, "Line_ID"] is None:
            continue
        extra.append(raw.loc[i].copy())                     # exact copy
        log.append((raw.at[i, "Line_ID"], "(row)", "Exact duplicate", "", "", "Removed (exact duplicate)"))
    logical_rows = [i for i in rng.choice(n, size=k(0.012), replace=False)]
    for i in logical_rows:
        if raw.at[i, "Line_ID"] is None:
            continue
        cp = raw.loc[i].copy()
        reg = truth.at[i, "Region"]
        cp["Region"] = str(rng.choice(REGION_VARIANTS[reg]))
        cp["Notes"] = None if rng.random() < 0.5 else cp["Notes"]
        col = str(rng.choice(["Unit_Price_x", "Actual_Unit_Cost_USD", "Target_Gross_Profit_USD"]))
        if col == "Unit_Price_x":
            col = "Target_Unit_Price_USD"
        for c2 in NUMERIC:                                   # the copy still has what the original lost
            v2 = cp[c2]
            if v2 is None or (isinstance(v2, str) and v2.strip().lower() in ("", "null", "n/a", "-", "nan")):
                cp[c2] = truth.at[i, c2]
        if cp[col] is not None and not (isinstance(raw.at[i, col], str) and raw.at[i, col].strip().lower() in ("", "null", "n/a", "-", "nan")) and raw.at[i, col] is not None:
            cp[col] = None                                   # ...and is missing a field the original has
        extra.append(cp)
        log.append((raw.at[i, "Line_ID"], "(row)", "Logical duplicate (complementary)", "", "", "Merged into first copy"))
    used = set(logical_rows)
    conflict_rows = [i for i in rng.choice(n, size=40, replace=False) if i not in used][:20]
    for i in conflict_rows:
        if raw.at[i, "Line_ID"] is None or truth.at[i, "Actual_Qty"] == 0:
            continue
        cp = truth.loc[i].copy().astype(object)
        cp["Date"] = truth.at[i, "Date"].to_pydatetime()
        bump = int(rng.integers(3, 15))
        cp["Actual_Qty"] = int(truth.at[i, "Actual_Qty"]) + bump
        price_c = int(round(truth.at[i, "Actual_Unit_Price_USD"] * 100))
        cost_c = int(round(truth.at[i, "Actual_Unit_Cost_USD"] * 100))
        cp["Actual_Sales_USD"] = cp["Actual_Qty"] * price_c / 100
        cp["Actual_Gross_Profit_USD"] = (cp["Actual_Qty"] * (price_c - cost_c)) / 100
        cp["_truth_row"] = i
        extra.append(cp)
        log.append((raw.at[i, "Line_ID"], "(row)", "Conflicting duplicate", truth.at[i, "Actual_Qty"],
                    cp["Actual_Qty"], "Review (conflict) - all copies"))
    raw = pd.concat([raw, pd.DataFrame(extra)], ignore_index=True)
    # blank rows and shuffle (keep months roughly in order like an export, but interleave copies)
    blanks = pd.DataFrame([{c: None for c in raw.columns} for _ in range(25)])
    raw = pd.concat([raw, blanks], ignore_index=True)
    key = raw["_truth_row"].astype(float).to_numpy().copy()
    nan = np.isnan(key)
    key[nan] = rng.integers(0, n, size=int(nan.sum()))
    key = key + rng.uniform(0, 1, size=len(raw)) * 40
    raw = raw.iloc[np.argsort(key, kind="stable")].reset_index(drop=True)
    raw = raw.drop(columns="_truth_row")
    # header noise + an entirely empty column
    raw = raw.rename(columns={"Actual_Qty": " Actual_Qty", "Target_Sales_USD": "Target_Sales_USD "})
    raw.insert(len(raw.columns), "Comments", None)
    manifest = pd.DataFrame(log, columns=["Line_ID", "Column", "Defect", "Clean_Value", "Dirty_Value", "Expected_Handling"])
    return raw, manifest


def write_raw(raw, path):
    """Write the messy frame to .xlsx exactly as-is (mixed types per column)."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "sales_raw"
    ws.append(list(raw.columns))
    for row in raw.itertuples(index=False):
        out = []
        for v in row:
            if v is None or (isinstance(v, float) and np.isnan(v)) or (isinstance(v, str) and v == ""):
                out.append(None)
            elif isinstance(v, (np.integer,)):
                out.append(int(v))
            elif isinstance(v, (np.floating,)):
                out.append(float(v))
            elif isinstance(v, pd.Timestamp):
                out.append(v.to_pydatetime())
            else:
                out.append(v)
        ws.append(out)
    for c in ws["B"][1:]:
        if hasattr(c.value, "year"):
            c.number_format = "yyyy-mm-dd"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Generate the synthetic raw sales file (deterministic).")
    ap.add_argument("--seed", type=int, default=config.RANDOM_SEED)
    ap.add_argument("--output-dir", type=Path, default=config.RAW_DIR)
    a = ap.parse_args(argv)
    truth = build_truth(a.seed)
    raw, manifest = corrupt(truth, a.seed)
    raw_path = a.output_dir / config.RAW_FILE.name
    write_raw(raw, raw_path)
    manifest.to_csv(a.output_dir / config.DEFECT_MANIFEST_FILE.name, index=False)
    print(f"Seed {a.seed}: {len(truth):,} true sales lines -> {len(raw):,} raw rows "
          f"({len(manifest):,} injected defects)")
    print(f"Wrote {raw_path}")
    print(f"Wrote {a.output_dir / config.DEFECT_MANIFEST_FILE.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
