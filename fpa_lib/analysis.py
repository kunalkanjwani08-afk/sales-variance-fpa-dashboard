"""Line-level and group-level variance calculations (all money in integer cents).

line_metrics(df)        -> adds the cent columns and line-level effects
summarise(lines, keys)  -> one row per group with revenue and gross-profit
                           bridges, status, materiality and margins
"""
import numpy as np
import pandas as pd

from . import config
from . import finance as fin

# short names for the cent columns
C_COLS = {"TP": "Target_Unit_Price_USD", "AP": "Actual_Unit_Price_USD", "TCU": "Target_Unit_Cost_USD",
          "ACU": "Actual_Unit_Cost_USD", "TS": "Target_Sales_USD", "AS": "Actual_Sales_USD",
          "TG": "Target_Gross_Profit_USD", "AG": "Actual_Gross_Profit_USD"}


def line_metrics(df):
    """Return a copy of the clean data with integer-cent columns and line-level effects.

    Raises ValueError if any line breaks Sales = Qty x Price or
    Gross_Profit = Sales - Qty x Cost (the analysis relies on both).
    """
    L = df.copy()
    for k, col in C_COLS.items():
        L[k] = fin.to_cents(L[col])
    L["TQ"] = L["Target_Qty"].astype("int64")
    L["AQ"] = L["Actual_Qty"].astype("int64")
    for side, (q, p, u, s, g) in {"Target": ("TQ", "TP", "TCU", "TS", "TG"), "Actual": ("AQ", "AP", "ACU", "AS", "AG")}.items():
        if (L[s] != L[q] * L[p]).any() or (L[g] != L[s] - L[q] * L[u]).any():
            raise ValueError(f"{side} figures break the sales or gross-profit identity on some lines")
    L["VAR"] = L["AS"] - L["TS"]
    L["GVAR"] = L["AG"] - L["TG"]
    L["VOL"], L["PRC"] = fin.revenue_bridge(L["TQ"], L["AQ"], L["TP"], L["AP"])
    L["PVOL"], _, L["COST"] = fin.profit_bridge(L["TQ"], L["AQ"], L["TP"], L["AP"], L["TCU"], L["ACU"])
    L["AQ_TP"] = L["AQ"] * L["TP"]                      # actual quantity valued at target price
    L["AQ_TM"] = L["AQ"] * (L["TP"] - L["TCU"])         # actual quantity valued at target unit margin
    L["STATUS"] = fin.status(L["VAR"], L["TS"])
    L["MAT"], L["TRIG"] = fin.material(L["VAR"], L["TS"])
    L["BELOW"] = (L["VAR"] < 0).astype("int64")
    L["UNF"] = (L["STATUS"] == config.STATUS_UNFAVORABLE).astype("int64")
    L["FAV"] = (L["STATUS"] == config.STATUS_FAVORABLE).astype("int64")
    L["WITHIN"] = (L["STATUS"] == config.STATUS_WITHIN).astype("int64")
    L["NOTGT"] = (L["STATUS"] == config.STATUS_NO_TARGET).astype("int64")
    L["MATN"] = L["MAT"].astype("int64")
    return L


SUM_COLS = ["TQ", "AQ", "TS", "AS", "TG", "AG", "VAR", "GVAR", "VOL", "PRC", "PVOL", "COST", "AQ_TP", "AQ_TM",
            "BELOW", "UNF", "FAV", "WITHIN", "NOTGT", "MATN"]


def summarise(L, keys):
    """Group the line table by `keys` (list; [] = company total) and compute
    every group metric from summed cents (percentages are sum/sum, never an
    average of line percentages)."""
    if keys:
        g = L.groupby(keys, sort=True, observed=True)[SUM_COLS].sum()
        g["Lines"] = L.groupby(keys, sort=True, observed=True).size()
        g = g.reset_index()
    else:
        g = L[SUM_COLS].sum().to_frame().T
        g["Lines"] = len(L)
    rows = []
    for r in g.itertuples(index=False):
        r = r._asdict()
        vol, pure, mix = fin.split_volume(r["TQ"], r["AQ"], r["TS"], r["AQ_TP"])
        pvol, ppure, pmix = fin.split_volume(r["TQ"], r["AQ"], r["TG"], r["AQ_TM"])
        assert vol == r["VOL"] and pvol == r["PVOL"], "volume effect does not reconcile"
        st = fin.status(r["VAR"], r["TS"])
        mat, trig = fin.material(r["VAR"], r["TS"])
        out = {k: r[k] for k in keys}
        out.update({
            "Lines": int(r["Lines"]),
            "Target_Qty": int(r["TQ"]), "Actual_Qty": int(r["AQ"]), "Qty_Variance": int(r["AQ"] - r["TQ"]),
            "Qty_Attainment_Pct": fin.pct(r["AQ"], r["TQ"]) if r["TQ"] > 0 else np.nan,
            "Target_Sales_USD": r["TS"] / 100, "Actual_Sales_USD": r["AS"] / 100,
            "Sales_Variance_USD": r["VAR"] / 100, "Sales_Variance_Pct": fin.pct(r["VAR"], r["TS"]),
            "Sales_Variance_Status": st, "Material_Variance_Flag": "Yes" if mat else "No", "Material_Trigger": trig,
            "Volume_Effect_USD": vol / 100, "Volume_Pure_USD": pure / 100, "Mix_Effect_USD": mix / 100,
            "Price_Effect_USD": r["PRC"] / 100,
            "Target_Gross_Profit_USD": r["TG"] / 100, "Actual_Gross_Profit_USD": r["AG"] / 100,
            "Gross_Profit_Variance_USD": r["GVAR"] / 100, "Gross_Profit_Variance_Pct": fin.pct(r["GVAR"], r["TG"]),
            "Profit_Volume_Effect_USD": pvol / 100, "Profit_Volume_Pure_USD": ppure / 100, "Profit_Mix_Effect_USD": pmix / 100,
            "Profit_Price_Effect_USD": r["PRC"] / 100, "Profit_Cost_Effect_USD": r["COST"] / 100,
            "Target_Gross_Margin_Pct": fin.pct(r["TG"], r["TS"]), "Actual_Gross_Margin_Pct": fin.pct(r["AG"], r["AS"]),
            "Price_Realisation_Pct": fin.pct(r["PRC"], r["AQ_TP"]) if r["AQ_TP"] > 0 else np.nan,
            "Avg_Target_Price_USD": round(r["TS"] / r["TQ"] / 100, 2) if r["TQ"] > 0 else np.nan,
            "Avg_Actual_Price_USD": round(r["AS"] / r["AQ"] / 100, 2) if r["AQ"] > 0 else np.nan,
            "Lines_Below_Target_Pct": fin.pct(r["BELOW"], r["Lines"]),
            "Lines_Unfavorable_Pct": fin.pct(r["UNF"], r["Lines"]),
            "Material_Lines": int(r["MATN"]),
        })
        out["Gross_Margin_Change_pp"] = (round(out["Actual_Gross_Margin_Pct"] - out["Target_Gross_Margin_Pct"], 2)
                                         if r["TS"] > 0 and r["AS"] > 0 else np.nan)
        rows.append(out)
    res = pd.DataFrame(rows)
    total_var = int(L["VAR"].sum())
    if keys and total_var:
        sign = 1 if total_var > 0 else -1          # share of the NET variance; the shares add up to 100 %
        res["Share_of_Net_Variance_Pct"] = [fin.pct(sign * int(round(v * 100)), abs(total_var)) for v in res["Sales_Variance_USD"]]
        res["Share_of_Target_Pct"] = [fin.pct(int(round(v * 100)), int(L["TS"].sum())) for v in res["Target_Sales_USD"]]
    return res
