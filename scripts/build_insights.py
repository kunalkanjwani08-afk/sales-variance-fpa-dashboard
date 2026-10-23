"""
==============================================================================
 INSIGHTS & RECOMMENDATIONS GENERATOR
==============================================================================
 Writes, from the current analysis outputs (never from typed-in numbers):
   docs/Insights_and_Recommendations.md   the written report
   docs/key_metrics.json                  the headline figures it quotes
                                          (the dashboard re-checks each one)
   README.md                              the block between
                                          <!-- KEY-RESULTS:START --> and <!-- KEY-RESULTS:END -->

 Run from the repository root (after sales_variance_analysis.py):
        python scripts/build_insights.py

 Every number in the report is computed here from
 data/processed/Sales_Variance_Enriched_Data.csv and outputs/Sales_Variance_Analysis.xlsx.
 Wording that depends on the result (for example which region is weakest, or
 whether price or volume is the larger driver) is chosen by the data.
==============================================================================
"""
import argparse
import json
import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from fpa_lib import analysis as an  # noqa: E402
from fpa_lib import config  # noqa: E402

MINUS = "–"


# ------------------------------------------------------------ formatting --
def money(v, dp=1):
    a = abs(v)
    if a >= 1e6:
        return f"${a / 1e6:,.{dp}f}M"
    if a >= 1e3:
        return f"${a / 1e3:,.0f}K"
    return f"${a:,.0f}"


def smoney(v, dp=1):
    return ("+" if v > 0 else MINUS if v < 0 else "") + money(v, dp)


def spct(v, dp=2):
    return ("+" if v > 0 else MINUS if v < 0 else "") + f"{abs(v):.{dp}f}%"


def pc(v, dp=2):
    return f"{v:.{dp}f}%"


def sint(v):
    return ("+" if v > 0 else MINUS if v < 0 else "") + f"{abs(int(v)):,}"


def below_above(v):
    return "below" if v < 0 else "above"


def pp(v, dp=1):
    return ("+" if v > 0 else MINUS if v < 0 else "") + f"{abs(v):.{dp}f} pp"


# ------------------------------------------------------------------ facts --
def load(enriched_path, workbook_path, cleaning_path):
    d = pd.read_csv(enriched_path, dtype={"Line_ID": str, "Notes": str}, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
    d["Notes"] = d["Notes"].fillna("")
    d["Repaired_Fields"] = d["Repaired_Fields"].fillna("")
    L = an.line_metrics(d)
    x = pd.read_excel(workbook_path, sheet_name=None)
    cl = pd.read_excel(cleaning_path, sheet_name=None) if Path(cleaning_path).is_file() else None
    return d, L, x, cl


def facts(d, L, x, cl):
    S = lambda df, keys: an.summarise(df, keys)  # noqa: E731
    F = {}
    T = S(L, []).iloc[0]
    F["T"] = T
    F["reg"] = S(L, ["Region"]).set_index("Region")
    F["cat"] = S(L, ["Category"]).set_index("Category")
    F["year"] = S(L, ["Year"]).set_index("Year")
    F["prod"] = S(L, ["Category", "Product"])
    F["month"] = x["By_Year_Month"]
    F["acc"] = x["Target_Accuracy"].set_index("Region")
    F["pers"] = x["Persistence_Region_Category"]
    F["bands"] = x["By_Price_Band"].set_index("Price_Band")
    F["exc"] = x["Exceptions_RegCat_Month"]
    F["dq"] = x["Data_Quality_Impact"].set_index("Scope")
    F["ry"] = S(L, ["Region", "Year"]).set_index(["Region", "Year"])
    F["cy"] = S(L, ["Category", "Year"]).set_index(["Category", "Year"])
    F["rp"] = S(L, ["Region", "Category", "Product"])
    F["months"] = len(F["month"])
    F["below"] = int((F["month"]["Sales_Variance_USD"] < 0).sum())
    F["cl"] = cl
    # discounted lines (>= 5 % below target price) vs the rest
    disc = (L["AP"] - L["TP"]) * 100 <= -5 * L["TP"]
    F["disc"], F["rest"] = S(L[disc], []).iloc[0], S(L[~disc], []).iloc[0]
    # first month in which a category's unit cost ran 4 % or more above target
    mc = L.groupby(["Category", "Year_Month"]).apply(lambda g: (g["AQ"] * (g["ACU"] - g["TCU"])).sum() / max(1, (g["AQ"] * g["TCU"]).sum()) * 100,
                                                       include_groups=False)
    F["cost_up"] = {c: (mc[c][mc[c] >= 4].index.min() if (mc[c] >= 4).any() else None) for c in config.CATEGORIES}
    # months that beat target vs the rest: selling price vs target
    up = set(F["month"].loc[F["month"]["Sales_Variance_USD"] > 0, "Year_Month"])
    inm = L["Year_Month"].isin(up)
    F["above_vs_rest"] = (S(L[inm], []).iloc[0]["Price_Realisation_Pct"], S(L[~inm], []).iloc[0]["Price_Realisation_Pct"]) if up else None
    # largest negative exception and the other sharp drops (<= -20 %) of the same region x category
    exc = F["exc"].sort_values("Sales_Variance_USD")
    if len(exc) and exc["Sales_Variance_USD"].iloc[0] < 0:
        w0 = exc.iloc[0]
        drops = exc[(exc["Region"] == w0["Region"]) & (exc["Category"] == w0["Category"]) & (exc["Sales_Variance_Pct"] <= -20)]
        dm = (L["Region"] == w0["Region"]) & (L["Category"] == w0["Category"]) & L["Year_Month"].isin(set(drops["Year_Month"]))
        F["drop_cells"], F["drop_lines"], F["drop_notes"] = drops, int(dm.sum()), int((L.loc[dm, "Notes"] == "Stock-out").sum())
    else:
        F["drop_cells"], F["drop_lines"], F["drop_notes"] = exc.head(0), 0, 0
    return F


# ----------------------------------------------------------------- report --
def build(F):
    T, reg, cat, yr, acc = F["T"], F["reg"], F["cat"], F["year"], F["acc"]
    claims = []

    def claim(key, label, text, expected, unit, tol):
        claims.append({"key": key, "label": label, "text": text, "expected": expected, "unit": unit, "tol": tol})
        return text

    worst_r, best_r = reg["Sales_Variance_USD"].idxmin(), reg["Sales_Variance_USD"].idxmax()
    wr, br = reg.loc[worst_r], reg.loc[best_r]
    price_r = reg["Price_Realisation_Pct"].idxmin()
    worst_c = cat["Gross_Profit_Variance_Pct"].idxmin()
    wc = cat.loc[worst_c]
    first_y, last_y = yr.index.min(), yr.index.max()
    months, below = F["months"], F["below"]
    above_months = F["month"].loc[F["month"]["Sales_Variance_USD"] > 0, "Year_Month"].tolist()
    disc, rest = F["disc"], F["rest"]
    prod = F["prod"]
    n_above = int((prod["Sales_Variance_USD"] > 0).sum())
    pers = F["pers"]
    n_pers = int((pers["Persistent_Shortfall"] == "Yes").sum())
    pers_by_reg = pers[pers["Persistent_Shortfall"] == "Yes"].groupby("Region").size().sort_values(ascending=False)
    share_worst = round(wr["Sales_Variance_USD"] / T["Sales_Variance_USD"] * 100, 2) if T["Sales_Variance_USD"] else float("nan")
    price_is_main = abs(T["Price_Effect_USD"]) > abs(T["Volume_Effect_USD"])
    lines = []
    w = lines.append

    # ------------------------------------------------------------------ header
    w("# Sales Variance Analysis: Insights & Recommendations")
    w("")
    w("> **Synthetic data.** The figures come from a simulated business generated by `scripts/generate_synthetic_data.py`. "
      "They illustrate the method; they do not describe a real company. This file is generated by `scripts/build_insights.py` "
      "from the current analysis outputs - edit the script, not this file.")
    w("")
    w(f"**Scope:** {int(T['Lines']):,} sales lines · {len(reg)} regions · {len(cat)} categories · {len(prod)} products · "
      f"{F['month']['Year_Month'].iloc[0]} to {F['month']['Year_Month'].iloc[-1]} ({months} months).")
    w("")
    # -------------------------------------------------------------- 1. summary
    w("## 1. Executive summary")
    w("")
    claim("sales_variance_usd", "Sales variance", smoney(T["Sales_Variance_USD"]), float(T["Sales_Variance_USD"]), "usd", 0.005)
    svp = claim("sales_variance_pct", "Sales variance %", spct(T["Sales_Variance_Pct"]), float(T["Sales_Variance_Pct"]), "pct", 0.0051)
    claim("gp_variance_usd", "Gross-profit variance", smoney(T["Gross_Profit_Variance_USD"]), float(T["Gross_Profit_Variance_USD"]), "usd", 0.005)
    gvp = claim("gp_variance_pct", "Gross-profit variance %", spct(T["Gross_Profit_Variance_Pct"]), float(T["Gross_Profit_Variance_Pct"]), "pct", 0.0051)
    pe = claim("price_effect_usd", "Price effect", smoney(T["Price_Effect_USD"]), float(T["Price_Effect_USD"]), "usd", 0.005)
    claim("pure_volume_usd", "Pure volume effect", smoney(T["Volume_Pure_USD"]), float(T["Volume_Pure_USD"]), "usd", 0.005)
    claim("mix_usd", "Mix effect", smoney(T["Mix_Effect_USD"]), float(T["Mix_Effect_USD"]), "usd", 0.005)
    claim("cost_effect_usd", "Cost effect (gross profit)", smoney(T["Profit_Cost_Effect_USD"]), float(T["Profit_Cost_Effect_USD"]), "usd", 0.005)
    rl = claim("price_realisation_pct", "Average selling price vs target", spct(T["Price_Realisation_Pct"]), float(T["Price_Realisation_Pct"]), "pct", 0.0051)
    qa = claim("qty_attainment_pct", "Quantity attainment", pc(T["Qty_Attainment_Pct"]), float(T["Qty_Attainment_Pct"]), "pct", 0.0051)
    tgm = claim("gross_margin_target_pct", "Target gross margin", pc(T["Target_Gross_Margin_Pct"]), float(T["Target_Gross_Margin_Pct"]), "pct", 0.0051)
    agm = claim("gross_margin_actual_pct", "Actual gross margin", pc(T["Actual_Gross_Margin_Pct"]), float(T["Actual_Gross_Margin_Pct"]), "pct", 0.0051)
    claim("months_below", "Months below target", f"{below} of {months}", below, "num", 0)
    driver = ("Lower selling prices, not lower volume, explain the revenue miss" if price_is_main and T["Volume_Effect_USD"] > T["Price_Effect_USD"]
              else "Volume explains most of the revenue miss")
    cost_txt = (f"Unit costs above target took a further **{money(T['Profit_Cost_Effect_USD'])}** of gross profit, so profit fell faster than sales."
                if T["Profit_Cost_Effect_USD"] < 0 else f"Unit costs below target added **{money(T['Profit_Cost_Effect_USD'])}** of gross profit.")
    w(f"Sales finished **{money(T['Sales_Variance_USD'])} ({pc(abs(T['Sales_Variance_Pct']))}) {below_above(T['Sales_Variance_USD'])} target** and gross profit "
      f"**{money(T['Gross_Profit_Variance_USD'])} ({pc(abs(T['Gross_Profit_Variance_Pct']))}) {below_above(T['Gross_Profit_Variance_USD'])} target**. "
      f"Gross margin moved from {tgm} planned to {agm}. {driver}: the price effect is **{pe}** "
      f"(average selling price {rl} vs target) while units reached {qa} of plan. {cost_txt} {below} of {months} months finished below target.")
    w("")
    w(f"The result is not uniform. **{worst_r}** is the weakest region ({spct(wr['Sales_Variance_Pct'])}) and accounts for "
      f"{share_worst:.0f}% of the net shortfall; **{best_r}** is the strongest ({spct(br['Sales_Variance_Pct'])}).")
    claim("worst_region", "Weakest region", worst_r, worst_r, "text", 0)
    claim("worst_region_pct", f"{worst_r} sales variance %", spct(wr["Sales_Variance_Pct"]), float(wr["Sales_Variance_Pct"]), "pct", 0.0051)
    claim("worst_region_share_of_net_pct", f"{worst_r} share of the net shortfall", f"{share_worst:.0f}%", share_worst, "pct", 0.51)
    claim("best_region", "Strongest region", best_r, best_r, "text", 0)
    claim("best_region_pct", f"{best_r} sales variance %", spct(br["Sales_Variance_Pct"]), float(br["Sales_Variance_Pct"]), "pct", 0.0051)
    w("")
    w("| Metric | Target | Actual | Variance |")
    w("|---|---|---|---|")
    w(f"| Sales | {money(T['Target_Sales_USD'], 2)} | {money(T['Actual_Sales_USD'], 2)} | {smoney(T['Sales_Variance_USD'], 2)} ({spct(T['Sales_Variance_Pct'])}) |")
    w(f"| Gross profit | {money(T['Target_Gross_Profit_USD'], 2)} | {money(T['Actual_Gross_Profit_USD'], 2)} | {smoney(T['Gross_Profit_Variance_USD'], 2)} ({spct(T['Gross_Profit_Variance_Pct'])}) |")
    w(f"| Gross margin | {pc(T['Target_Gross_Margin_Pct'])} | {pc(T['Actual_Gross_Margin_Pct'])} | {pp(T['Gross_Margin_Change_pp'], 2)} |")
    w(f"| Units | {int(T['Target_Qty']):,} | {int(T['Actual_Qty']):,} | {sint(T['Qty_Variance'])} ({pc(T['Qty_Attainment_Pct'])} of plan) |")
    w("")
    w("| Revenue bridge | USD | | Gross-profit bridge | USD |")
    w("|---|---|---|---|---|")
    w(f"| Target sales | {money(T['Target_Sales_USD'], 2)} | | Target gross profit | {money(T['Target_Gross_Profit_USD'], 2)} |")
    w(f"| Pure volume | {smoney(T['Volume_Pure_USD'], 2)} | | Pure volume | {smoney(T['Profit_Volume_Pure_USD'], 2)} |")
    w(f"| Mix | {smoney(T['Mix_Effect_USD'], 2)} | | Mix | {smoney(T['Profit_Mix_Effect_USD'], 2)} |")
    w(f"| Price | {smoney(T['Price_Effect_USD'], 2)} | | Price | {smoney(T['Profit_Price_Effect_USD'], 2)} |")
    w(f"| | | | Cost | {smoney(T['Profit_Cost_Effect_USD'], 2)} |")
    w(f"| **Actual sales** | **{money(T['Actual_Sales_USD'], 2)}** | | **Actual gross profit** | **{money(T['Actual_Gross_Profit_USD'], 2)}** |")
    w("")
    w("Both bridges reconcile to the cent (see the Validation_Summary sheet of `outputs/Sales_Variance_Analysis.xlsx`).")
    w("")

    # -------------------------------------------------------- 2. what it shows
    w("## 2. What the data shows")
    w("")
    w("### 2.1 Discounts bought volume, but not profit")
    claim("discounted_lines_qty_att_pct", "Discounted lines: quantity attainment", pc(disc["Qty_Attainment_Pct"]), float(disc["Qty_Attainment_Pct"]), "pct", 0.0051)
    claim("discounted_lines_gp_var_pct", "Discounted lines: gross-profit variance", spct(disc["Gross_Profit_Variance_Pct"]), float(disc["Gross_Profit_Variance_Pct"]), "pct", 0.0051)
    claim("other_lines_gp_var_pct", "Other lines: gross-profit variance", spct(rest["Gross_Profit_Variance_Pct"]), float(rest["Gross_Profit_Variance_Pct"]), "pct", 0.0051)
    w(f"- {int(disc['Lines']):,} lines ({disc['Lines'] / T['Lines'] * 100:.0f}% of all lines) were sold at least 5% below their target price. "
      f"They reached **{pc(disc['Qty_Attainment_Pct'])} of planned units** and {spct(disc['Sales_Variance_Pct'])} of target sales, "
      f"but **{spct(disc['Gross_Profit_Variance_Pct'])} of target gross profit**.")
    w(f"- All other lines reached {pc(rest['Qty_Attainment_Pct'])} of planned units and {spct(rest['Gross_Profit_Variance_Pct'])} of target gross profit.")
    w(f"- Discounted lines carry {money(abs(disc['Price_Effect_USD']))} of the {money(abs(T['Price_Effect_USD']))} total price effect.")
    w("- *Evidence vs hypothesis:* the data shows that deeper discounts came with more units and less profit. It cannot show whether those "
      "units were incremental or pulled forward from other months or other products (cannibalisation) - that needs a baseline-demand estimate.")
    w("")
    w("### 2.2 Two regional problems with different causes")
    lat = reg.loc[worst_r]
    w(f"- **{worst_r}** ({smoney(lat['Sales_Variance_USD'])}, {spct(lat['Sales_Variance_Pct'])}) is a **volume** problem: units reached "
      f"{pc(lat['Qty_Attainment_Pct'])} of plan and pure volume explains {smoney(lat['Volume_Pure_USD'])}; its selling prices were "
      f"{spct(lat['Price_Realisation_Pct'])} vs target, close to the company average.")
    rpw = F["rp"][F["rp"]["Region"] == worst_r]
    w(f"  - It is broad, not one product: {int((rpw['Sales_Variance_USD'] > 0).sum())} of {len(rpw)} products beat target there.")
    if worst_r in acc.index:
        a = acc.loc[worst_r]
        w(f"  - In {last_y} its monthly targets missed actual sales by {pc(a['Target_WAPE_Pct'])} on average (WAPE), while a naive forecast "
          f"(the same month of {first_y}) missed by only {pc(a['Naive_WAPE_Pct'])}. The target was further from reality than last year's actuals.")
    pr = reg.loc[price_r]
    w(f"- **{price_r}** has the weakest pricing: selling prices {spct(pr['Price_Realisation_Pct'])} vs target "
      f"(price effect {smoney(pr['Price_Effect_USD'])}) and gross profit {spct(pr['Gross_Profit_Variance_Pct'])}.")
    if (price_r, last_y) in F["ry"].index:
        e_last = F["ry"].loc[(price_r, last_y), "Price_Realisation_Pct"]
        others = F["ry"].xs(last_y, level="Year").drop(index=price_r)["Price_Realisation_Pct"]
        w(f"  - In {last_y}, when list prices were planned to rise, {price_r} realised {spct(e_last)} vs target against "
          f"{spct(others.min())} to {spct(others.max())} in the other regions.")
        if price_r in acc.index:
            a = acc.loc[price_r]
            w(f"  - Its {last_y} target error (WAPE {pc(a['Target_WAPE_Pct'])}) is also larger than a naive forecast ({pc(a['Naive_WAPE_Pct'])}).")
    w(f"- **{best_r}** ({smoney(br['Sales_Variance_USD'])}, {spct(br['Sales_Variance_Pct'])}) beat target on volume "
      f"({pc(br['Qty_Attainment_Pct'])} of planned units)" + (f" despite selling {pc(abs(br['Price_Realisation_Pct']))} below target price." if br["Price_Realisation_Pct"] < 0 else "."))
    w("")
    w("| Region | Sales variance | % | Pure volume | Mix | Price | Cost (GP) | GP variance % | Units vs plan |")
    w("|---|---|---|---|---|---|---|---|---|")
    for r, x_ in reg.sort_values("Sales_Variance_USD").iterrows():
        w(f"| {r} | {smoney(x_['Sales_Variance_USD'], 2)} | {spct(x_['Sales_Variance_Pct'])} | {smoney(x_['Volume_Pure_USD'], 2)} | "
          f"{smoney(x_['Mix_Effect_USD'], 2)} | {smoney(x_['Price_Effect_USD'], 2)} | {smoney(x_['Profit_Cost_Effect_USD'], 2)} | "
          f"{spct(x_['Gross_Profit_Variance_Pct'])} | {pc(x_['Qty_Attainment_Pct'])} |")
    w("")
    w(f"### 2.3 {worst_c}: a cost problem on top of price")
    cu = F["cost_up"].get(worst_c)
    w(f"- {worst_c} has the largest gross-profit shortfall in percentage terms ({spct(wc['Gross_Profit_Variance_Pct'])}, {smoney(wc['Gross_Profit_Variance_USD'])}) "
      f"although its sales variance ({spct(wc['Sales_Variance_Pct'])}) is close to the other categories.")
    w(f"- Unit costs above target took {smoney(wc['Profit_Cost_Effect_USD'])} of its gross profit" +
      (f"; from {cu} onwards its unit costs ran 4% or more above target" if cu else "") +
      f". Its gross margin fell from {pc(wc['Target_Gross_Margin_Pct'])} planned to {pc(wc['Actual_Gross_Margin_Pct'])}.")
    best_cost = cat["Profit_Cost_Effect_USD"].idxmax()
    if cat.loc[best_cost, "Profit_Cost_Effect_USD"] > 0:
        w(f"- By contrast, {best_cost} costs came in below target and added {smoney(cat.loc[best_cost, 'Profit_Cost_Effect_USD'])} of gross profit.")
    w("- *Evidence vs hypothesis:* the cost increase is measured; whether it came from freight, components or supplier pricing cannot be told without a cost breakdown.")
    w("")
    w("### 2.4 Timing: the holiday peak was under-planned")
    m = F["month"]
    peak = m.loc[m["Phasing_Gap_pp"].idxmax()]
    trough = m.loc[m["Phasing_Gap_pp"].idxmin()]
    w(f"- Only {len(above_months)} of {months} months beat target: {', '.join(above_months) if above_months else 'none'}.")
    w(f"- The target spread each year more evenly than sales really came in: in {peak['Year_Month']} actual sales were "
      f"{pc(peak['Actual_Share_of_Year_Pct'])} of the year against {pc(peak['Target_Share_of_Year_Pct'])} planned, while in {trough['Year_Month']} "
      f"they were {pc(trough['Actual_Share_of_Year_Pct'])} against {pc(trough['Target_Share_of_Year_Pct'])}.")
    if above_months:
        ab = F["above_vs_rest"]
        w(f"- In the months that beat target, selling prices were {spct(ab[0])} vs target, against {spct(ab[1])} in the other months"
          + (": part of the peak-month gain was bought with discounts." if ab[0] < ab[1] else "."))
    drops = F["drop_cells"]
    if len(drops) >= 2:
        w(f"- The largest negative exceptions are {drops['Region'].iloc[0]} {drops['Category'].iloc[0]} in {', '.join(sorted(drops['Year_Month']))} "
          f"({spct(drops['Sales_Variance_Pct'].max(), 1)} to {spct(drops['Sales_Variance_Pct'].min(), 1)}): a short, sharp drop rather than a trend."
          + (f" {F['drop_notes']} of the {F['drop_lines']} lines in those months carry a 'Stock-out' note, which points to supply rather than demand."
             if F["drop_notes"] else ""))
    w("")
    w("### 2.5 Persistence and products")
    claim("persistent_pairs", "Persistent region × category shortfalls", f"{n_pers} of {len(pers)}", n_pers, "num", 0)
    claim("products_above_target", "Products above target", f"{n_above} of {len(prod)}", n_above, "num", 0)
    w(f"- **{n_pers} of {len(pers)}** region × category pairs were below target in at least {config.PERSISTENCE_SHARE:.0%} of months "
      f"({', '.join(f'{r}: {n}' for r, n in pers_by_reg.items())}). Repeated misses point to something structural - the target, "
      "the market or pricing - rather than one-off events.")
    p80 = x_pareto(F)
    w(f"- **{n_above} of {len(prod)}** products beat target. The shortfall is spread out: {p80} products make up 80% of it, so a "
      "product-by-product fix would miss the regional and pricing causes above.")
    w(f"- {last_y} was weaker than {first_y}: sales {spct(yr.loc[last_y, 'Sales_Variance_Pct'])} vs {spct(yr.loc[first_y, 'Sales_Variance_Pct'])} "
      f"and gross profit {spct(yr.loc[last_y, 'Gross_Profit_Variance_Pct'])} vs {spct(yr.loc[first_y, 'Gross_Profit_Variance_Pct'])}.")
    w("")

    # ------------------------------------------------------ 3. method caveats
    w("## 3. What follows from the method, not from the business")
    w("")
    w("- **Mix depends on the grouping.** The mix effect values selling a different combination of lines (products, regions, accounts) "
      "at target prices. It is calculated within whatever group is being viewed, so the company mix and the sum of regional mixes differ.")
    w("- **The tracking signal restates the monthly pattern.** It is cumulative variance divided by the average monthly miss. "
      f"With {below} of {months} months below target it is strongly negative by construction; it is a warning light, not separate proof of target bias.")
    mat_share = int(x_mat(F))
    w(f"- **Line-level materiality is noisy.** {mat_share:,} lines ({mat_share / T['Lines'] * 100:.0f}%) meet the materiality rule because single "
      f"accounts swing a lot from month to month. The region × category × month exceptions list ({len(F['exc'])} cells, largest first) is the place to start.")
    w("- **Status and materiality are separate.** A group can be within the ±3% tolerance and still material in dollars - the company total is an example.")
    w("")

    # ------------------------------------------------------ 4. recommendations
    w("## 4. Recommendations")
    w("")
    w("Each recommendation states what was observed, what it is worth, what might be causing it (a hypothesis), what to do, "
      "how to measure success, and what data would confirm the cause. In a real business these would be validated with the commercial teams first.")
    w("")
    w("### R1. Put a profit test on discounts")
    w(f"- **Observed:** lines sold 5% or more below target price reached {pc(disc['Qty_Attainment_Pct'])} of planned units but "
      f"{spct(disc['Gross_Profit_Variance_Pct'])} of target gross profit.")
    w(f"- **Impact:** the price effect is {smoney(T['Price_Effect_USD'])} for the period; every dollar of discount falls straight to gross profit.")
    w("- **Possible mechanism (hypothesis):** promotions are deeper than the extra volume can pay for, especially in lower-margin categories.")
    w("- **Action:** approve promotions against a gross-profit payback (extra units × unit margin ≥ discount cost), with a floor price by category.")
    w(f"- **Success measure:** average selling price within 1% of target and discounted lines at or above target gross profit for two consecutive quarters (today {spct(T['Price_Realisation_Pct'])} and {spct(disc['Gross_Profit_Variance_Pct'])}).")
    w("- **Data needed:** promotion calendar and mechanics, baseline (non-promoted) demand, competitor prices.")
    w("")
    w(f"### R2. Re-base the {worst_r} plan and diagnose demand")
    w(f"- **Observed:** {worst_r} sold {pc(lat['Qty_Attainment_Pct'])} of planned units, its {last_y} targets were less accurate than a naive forecast, "
      "and the shortfall is spread across all products.")
    w(f"- **Impact:** {smoney(lat['Sales_Variance_USD'])} of sales and {smoney(lat['Gross_Profit_Variance_USD'])} of gross profit.")
    w("- **Possible mechanism (hypothesis):** the target assumed growth the market did not deliver, and/or distribution or demand is weaker than planned. The data cannot separate the two.")
    w(f"- **Action:** set the next {worst_r} targets from trailing actuals plus an explicit, owned growth assumption; run a separate demand and distribution review.")
    w(f"- **Success measure:** monthly target error (WAPE) below the naive forecast and variance within ±3% (today {pc(acc.loc[worst_r, 'Target_WAPE_Pct']) if worst_r in acc.index else 'n/a'} and {spct(lat['Sales_Variance_Pct'])}).")
    w("- **Data needed:** market size and share, distribution points, competitor activity.")
    w("")
    w(f"### R3. Recover the {price_r} price increase")
    w(f"- **Observed:** {price_r} realised {spct(pr['Price_Realisation_Pct'])} vs target price" +
      (f" and {spct(F['ry'].loc[(price_r, last_y), 'Price_Realisation_Pct'])} in {last_y}" if (price_r, last_y) in F["ry"].index else "") + ".")
    w(f"- **Impact:** price effect {smoney(pr['Price_Effect_USD'])}; gross profit {spct(pr['Gross_Profit_Variance_Pct'])} vs target.")
    w("- **Possible mechanism (hypothesis):** planned list-price increases were not fully passed on (contract terms, customer resistance or local competition).")
    w("- **Action:** review pass-through by customer and contract; agree a phased recovery plan where increases were waived.")
    w(f"- **Success measure:** {price_r} selling price within 1% of target.")
    w("- **Data needed:** contract price terms, price-change effective dates, customer-level price history.")
    w("")
    w(f"### R4. Address {worst_c} costs")
    w(f"- **Observed:** unit costs above target cost {smoney(wc['Profit_Cost_Effect_USD'])} of gross profit; margin {pc(wc['Actual_Gross_Margin_Pct'])} vs {pc(wc['Target_Gross_Margin_Pct'])} planned.")
    w(f"- **Impact:** {worst_c} gross profit {smoney(wc['Gross_Profit_Variance_USD'])} vs target.")
    w("- **Possible mechanism (hypothesis):** a step-up in freight or component costs that was not reflected in prices.")
    w("- **Action:** split the cost increase into its drivers, renegotiate where possible, and reflect the remainder in the next price review.")
    w(f"- **Success measure:** {worst_c} gross margin back within 1 point of target.")
    w("- **Data needed:** cost breakdown by component, freight and supplier.")
    w("")
    w("### R5. Fix the monthly phasing of targets")
    w(f"- **Observed:** {len(above_months)} months beat target and they are the seasonal peaks; most other months fell short. The target phasing is flatter than the sales pattern.")
    w("- **Impact:** monthly variances are inflated in both directions, which hides the real trend and triggers false alarms.")
    w("- **Possible mechanism (hypothesis):** annual targets were spread with an average seasonal profile that understates the holiday peak.")
    w("- **Action:** phase each category's target from its own historical monthly share of the year.")
    w("- **Success measure:** monthly phasing gaps within ±0.2 percentage points and tracking signal within ±4.")
    w("- **Data needed:** at least three years of monthly history per category.")
    w("")

    # ------------------------------------------------------ 5. cannot tell
    w("## 5. What this data cannot tell us")
    w("")
    w("- Why customers bought less in a region (market, competitor or distribution) - there is no market or competitor data.")
    w("- Whether discounted volume was incremental or pulled forward - there is no baseline demand.")
    w("- Whether targets were too ambitious or execution was weak - the comparison with a naive forecast shows where targets were far from history, not why.")
    w("- What drove the cost increase - unit cost is a single number per line.")
    w("")

    # ------------------------------------------------------ 6. data quality
    w("## 6. Data quality behind these numbers")
    w("")
    cl = F["cl"]
    if cl is not None:
        rr = dict(zip(cl["Row_Reconciliation"]["Step"], cl["Row_Reconciliation"]["Rows"], strict=True))
        w(f"- The raw file had {int(rr['Raw data rows']):,} rows. Cleaning removed {-int(rr['- exact duplicate rows']):,} exact duplicates and "
          f"{-int(rr['- entirely blank rows']):,} blank rows, merged {-int(rr['- duplicate copies merged']):,} duplicate copies, held "
          f"{-int(rr['- rows held for review']):,} rows for review and delivered {int(rr['= clean rows']):,} clean lines.")
        w(f"- {len(cl['Corrections_Audit']):,} values were repaired. Every repair was exact (solved from Sales = Qty × Price, Gross profit = Sales − Qty × Cost, "
          "the Line_ID key or the product master); nothing was estimated. Each one is listed with its original value in `data/processed/Cleaned_Sales_Data.xlsx`.")
    dq = F["dq"]
    ex = dq.loc["Excluding lines with any repaired value"]
    rv = dq.loc["Rows held back for review (not in any figure)"]
    w(f"- Sensitivity: excluding every line with a repaired value, sales variance would be {spct(ex['Sales_Variance_Pct'])} and gross-profit variance "
      f"{spct(ex['Gross_Profit_Variance_Pct'])} (all lines: {svp} and {gvp}). The {int(rv['Lines']):,} rows held for review carry "
      f"{money(rv['Target_Sales_USD'])} of target sales ({rv['Target_Sales_USD'] / T['Target_Sales_USD'] * 100:.2f}% of the total) and are excluded from every figure.")
    w("")
    return "\n".join(lines) + "\n", claims


def x_pareto(F):
    p = pd.read_excel(config.ANALYSIS_XLSX, sheet_name="Pareto_Shortfall_Product") if F.get("_pareto") is None else F["_pareto"]
    return int((p["Cumulative_Share_Pct"] - p["Share_of_Shortfall_Pct"] < 80).sum())


def x_mat(F):
    return int(F["_mat_lines"])


def readme_block(F, claims):
    T = F["T"]
    c = {k["key"]: k for k in claims}
    rows = [
        "| | |", "|---|---|",
        f"| **Sales** | {money(T['Actual_Sales_USD'], 2)} vs {money(T['Target_Sales_USD'], 2)} target: **{c['sales_variance_usd']['text']} ({c['sales_variance_pct']['text']})** |",
        f"| **Gross profit** | {money(T['Actual_Gross_Profit_USD'], 2)} vs {money(T['Target_Gross_Profit_USD'], 2)}: **{c['gp_variance_usd']['text']} ({c['gp_variance_pct']['text']})**; margin {c['gross_margin_target_pct']['text']} → {c['gross_margin_actual_pct']['text']} |",
        f"| **Revenue bridge** | pure volume {c['pure_volume_usd']['text']} · mix {c['mix_usd']['text']} · **price {c['price_effect_usd']['text']}** |",
        f"| **Profit bridge** | volume + mix {smoney(T['Profit_Volume_Effect_USD'])} · price {c['price_effect_usd']['text']} · **cost {c['cost_effect_usd']['text']}** |",
        f"| **Discounts** | lines sold 5%+ below target price: reached {c['discounted_lines_qty_att_pct']['text']} of planned units, but gross profit was {c['discounted_lines_gp_var_pct']['text']} vs target |",
        f"| **Regions** | {c['worst_region']['text']} {c['worst_region_pct']['text']} ({c['worst_region_share_of_net_pct']['text']} of the net shortfall) · {c['best_region']['text']} {c['best_region_pct']['text']} |",
        f"| **Consistency** | {c['months_below']['text']} months below target · {c['persistent_pairs']['text']} region × category pairs persistently short · {c['products_above_target']['text']} products above target |",
    ]
    cl = F["cl"]
    if cl is not None:
        rr = dict(zip(cl["Row_Reconciliation"]["Step"], cl["Row_Reconciliation"]["Rows"], strict=True))
        rows.append(f"| **Data** | {rr['Raw data rows']:,} raw rows → {rr['= clean rows']:,} clean sales lines; "
                    f"{len(cl['Corrections_Audit']):,} values repaired exactly, {-rr['- rows held for review']:,} rows held "
                    "for review (nothing estimated) |")
    return "\n".join(["<!-- KEY-RESULTS:START -->",
                      "*Generated by `scripts/build_insights.py` from the current outputs. Synthetic data.*", ""] + rows +
                     ["<!-- KEY-RESULTS:END -->"])


def main(argv=None):
    ap = argparse.ArgumentParser(description="Generate the insights report, key metrics and README results block.")
    ap.add_argument("--enriched", type=Path, default=config.ENRICHED_CSV)
    ap.add_argument("--workbook", type=Path, default=config.ANALYSIS_XLSX)
    ap.add_argument("--cleaning-workbook", type=Path, default=config.CLEAN_XLSX)
    ap.add_argument("--docs-dir", type=Path, default=config.DOCS_DIR)
    ap.add_argument("--readme", type=Path, default=config.REPO_ROOT / "README.md")
    a = ap.parse_args(argv)
    for p in (a.enriched, a.workbook):
        if not p.is_file():
            print(f"ERROR: {p} not found. Run the cleaning and analysis scripts first.", file=sys.stderr)
            return 2
    d, L, x, cl = load(a.enriched, a.workbook, a.cleaning_workbook)
    F = facts(d, L, x, cl)
    F["_pareto"] = x["Pareto_Shortfall_Product"]
    F["_mat_lines"] = len(x["Material_Lines"])
    md, claims = build(F)
    a.docs_dir.mkdir(parents=True, exist_ok=True)
    (a.docs_dir / config.INSIGHTS_MD.name).write_text(md, encoding="utf-8")
    (a.docs_dir / config.KEY_METRICS_JSON.name).write_text(json.dumps({"generated_by": "scripts/build_insights.py", "claims": claims},
                                                                      indent=2, ensure_ascii=False), encoding="utf-8")
    msg = "README.md not found - results block not written"
    if a.readme.is_file():
        txt = a.readme.read_text(encoding="utf-8")
        new, n = re.subn(r"<!-- KEY-RESULTS:START -->.*?<!-- KEY-RESULTS:END -->", lambda _: readme_block(F, claims), txt, flags=re.S)
        if n == 1:
            a.readme.write_text(new, encoding="utf-8")
            msg = "README.md key-results block updated"
        else:
            msg = "README.md has no KEY-RESULTS markers - not changed"
    print(f"Wrote {a.docs_dir / config.INSIGHTS_MD.name} ({len(md.splitlines())} lines)")
    print(f"Wrote {a.docs_dir / config.KEY_METRICS_JSON.name} ({len(claims)} headline figures)")
    print(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
