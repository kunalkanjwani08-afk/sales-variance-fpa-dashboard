"""Financial arithmetic used by every stage of the project.

Conventions
-----------
* Money is held as integer CENTS (numpy int64 or Python int). Totals and
  identities are therefore exact; there is no floating-point drift.
* Percentages are rounded HALF AWAY FROM ZERO (like Excel), using integer
  arithmetic, never binary floating point.
* Revenue bridge, line level (exact identity):
      Volume_Effect = (Actual_Qty - Target_Qty) x Target_Unit_Price
      Price_Effect  =  Actual_Qty x (Actual_Unit_Price - Target_Unit_Price)
      Volume_Effect + Price_Effect = Actual_Sales - Target_Sales
* Gross-profit bridge, line level (exact identity):
      Profit_Volume_Effect = (Actual_Qty - Target_Qty) x (Target_Unit_Price - Target_Unit_Cost)
      Profit_Price_Effect  =  Actual_Qty x (Actual_Unit_Price - Target_Unit_Price)
      Profit_Cost_Effect   = -Actual_Qty x (Actual_Unit_Cost - Target_Unit_Cost)
      sum of the three     =  Actual_Gross_Profit - Target_Gross_Profit
* Mix (group level only): the volume effect of a group is split into
      Pure volume = (sum Actual_Qty - sum Target_Qty) x group average target price
      Mix         = Volume effect - Pure volume
  i.e. mix is the value of selling a different combination of lines
  (products, regions, accounts) than planned, measured at target prices.
"""
from decimal import Decimal, InvalidOperation

import numpy as np

from . import config

# -------------------------------------------------------------- cents --


def to_cents(values):
    """Convert dollar amounts (floats/ints/strings, max 2 decimals) to int64 cents.

    Raises ValueError if a value has more than 2 decimals or is not finite, so
    a precision problem in the input is reported instead of silently rounded.
    """
    arr = np.asarray(values, dtype="float64")
    if not np.all(np.isfinite(arr)):
        raise ValueError("to_cents: input contains missing or non-finite values")
    scaled = arr * 100.0
    cents = np.rint(scaled)
    if np.any(np.abs(scaled - cents) > 1e-6):
        bad = arr[np.abs(scaled - cents) > 1e-6][:3]
        raise ValueError(f"to_cents: values with more than 2 decimals, e.g. {bad.tolist()}")
    return cents.astype("int64")


def money_text_to_cents(text):
    """Exact Decimal parse of one money string such as '$1,234.50' -> 123450 (int).

    Returns None if the text is not a valid amount with at most 2 decimals.
    """
    s = str(text).strip().replace(",", "").replace("$", "").replace(" ", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    if not d.is_finite() or d != d.quantize(Decimal("0.01")):
        return None
    return int(d * 100)


def cents_to_dollars(cents):
    """int cents -> float dollars (for writing CSV/Excel). Exact for |x| < 2**53 cents."""
    return np.asarray(cents, dtype="int64") / 100.0


def div_round(num, den):
    """Integer division rounded half away from zero (works on Python ints)."""
    if den == 0:
        raise ZeroDivisionError("div_round: zero denominator")
    sign = -1 if (num < 0) != (den < 0) else 1
    num, den = abs(num), abs(den)
    return sign * ((2 * num + den) // (2 * den))


# -------------------------------------------------------- percentages --


def pct(num_cents, den_cents, dp=2):
    """Percentage num/den*100 rounded half away from zero to `dp` decimals.

    Works on scalars or numpy int64 arrays (exact integer arithmetic).
    Returns NaN where den <= 0: a percentage of a zero or negative target is
    not meaningful, and it is reported as missing rather than hidden.
    """
    scale = 100 * 10 ** dp
    if np.ndim(num_cents) == 0 and np.ndim(den_cents) == 0:
        num, den = int(num_cents), int(den_cents)
        if den <= 0:
            return float("nan")
        q = (2 * abs(num) * scale + den) // (2 * den)
        return (q if num >= 0 else -q) / 10 ** dp
    num_b, den_b = np.broadcast_arrays(np.asarray(num_cents, dtype="int64"),
                                       np.asarray(den_cents, dtype="int64"))
    out = np.full(num_b.shape, np.nan)
    ok = den_b > 0
    n = np.abs(num_b[ok]).astype(object) * scale
    d = den_b[ok].astype(object)
    q = np.array((2 * n + d) // (2 * d), dtype="float64")
    out[ok] = np.where(num_b[ok] < 0, -q, q) / 10 ** dp
    return out


# ---------------------------------------------------- status & materiality --

def status(var_cents, target_cents, fav_pct=None, unf_pct=None):
    """Classify variance vs target. Exact integer comparison (no rounding).

    Favorable if var/target > +fav_pct %, Unfavorable if < -unf_pct %,
    Within tolerance otherwise (the boundary itself counts as within),
    No target if target <= 0.
    """
    fav_pct = config.FAVORABLE_PCT if fav_pct is None else fav_pct
    unf_pct = config.UNFAVORABLE_PCT if unf_pct is None else unf_pct
    v = np.asarray(var_cents, dtype="int64")
    t = np.asarray(target_cents, dtype="int64")
    fav_bp, unf_bp = int(round(fav_pct * 100)), int(round(unf_pct * 100))
    vv = v.astype(object) * 10_000
    out = np.where(t <= 0, config.STATUS_NO_TARGET,
                   np.where(vv > fav_bp * t.astype(object), config.STATUS_FAVORABLE,
                            np.where(vv < -unf_bp * t.astype(object), config.STATUS_UNFAVORABLE,
                                     config.STATUS_WITHIN)))
    return out.item() if out.ndim == 0 else out


def material(var_cents, target_cents, pct_threshold=None, usd_threshold=None):
    """Return (is_material, trigger). Material if |var %| >= pct or |var $| >= usd.

    trigger is '' / 'Percent' / 'Amount' / 'Percent and amount'.
    The percent test is skipped when target <= 0.
    """
    pct_threshold = config.MATERIAL_PCT if pct_threshold is None else pct_threshold
    usd_threshold = config.MATERIAL_USD if usd_threshold is None else usd_threshold
    v = np.abs(np.asarray(var_cents, dtype="int64")).astype(object)
    t = np.asarray(target_cents, dtype="int64")
    by_pct = (t > 0) & (v * 10_000 >= int(round(pct_threshold * 100)) * t.astype(object))
    by_usd = v >= int(round(usd_threshold * 100))
    by_pct, by_usd = np.asarray(by_pct, dtype=bool), np.asarray(by_usd, dtype=bool)
    trig = np.where(by_pct & by_usd, "Percent and amount",
                    np.where(by_pct, "Percent", np.where(by_usd, "Amount", "")))
    is_mat = by_pct | by_usd
    if is_mat.ndim == 0:
        return bool(is_mat), str(trig)
    return is_mat, trig


# ------------------------------------------------------------- bridges --

def revenue_bridge(tq, aq, tp_c, ap_c):
    """Line-level revenue bridge in cents: returns (volume_effect, price_effect)."""
    tq, aq = np.asarray(tq, dtype="int64"), np.asarray(aq, dtype="int64")
    tp_c, ap_c = np.asarray(tp_c, dtype="int64"), np.asarray(ap_c, dtype="int64")
    return (aq - tq) * tp_c, aq * (ap_c - tp_c)


def profit_bridge(tq, aq, tp_c, ap_c, tcu_c, acu_c):
    """Line-level gross-profit bridge in cents: (volume, price, cost) effects."""
    tq, aq = np.asarray(tq, dtype="int64"), np.asarray(aq, dtype="int64")
    tp_c, ap_c = np.asarray(tp_c, dtype="int64"), np.asarray(ap_c, dtype="int64")
    tcu_c, acu_c = np.asarray(tcu_c, dtype="int64"), np.asarray(acu_c, dtype="int64")
    return (aq - tq) * (tp_c - tcu_c), aq * (ap_c - tp_c), -aq * (acu_c - tcu_c)


def split_volume(sum_tq, sum_aq, target_value_c, actual_qty_at_target_c):
    """Split a group's volume effect into pure volume and mix (cents, exact).

    target_value_c          = sum(Target_Qty x target unit value)   e.g. Target_Sales
    actual_qty_at_target_c  = sum(Actual_Qty x target unit value)
    volume effect           = actual_qty_at_target_c - target_value_c
    pure volume             = (sum_aq - sum_tq) x (target_value_c / sum_tq), rounded to the cent
    mix                     = volume effect - pure volume
    The unit value is the target price (revenue bridge) or the target unit
    margin (gross-profit bridge). Returns (volume, pure_volume, mix).
    """
    sum_tq, sum_aq = int(sum_tq), int(sum_aq)
    target_value_c, actual_qty_at_target_c = int(target_value_c), int(actual_qty_at_target_c)
    volume = actual_qty_at_target_c - target_value_c
    pure = div_round((sum_aq - sum_tq) * target_value_c, sum_tq) if sum_tq > 0 else 0
    return volume, pure, volume - pure
