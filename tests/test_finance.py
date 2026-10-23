"""Unit tests for fpa_lib.finance: cents, rounding, status, materiality and bridges."""
import math

import numpy as np
import pytest

from fpa_lib import config
from fpa_lib import finance as fin


# ---------------------------------------------------------------- cents
def test_to_cents_exact():
    assert fin.to_cents([1.1, 0.29, 1234.56, -0.01]).tolist() == [110, 29, 123456, -1]


@pytest.mark.parametrize("bad", [[1.005], [float("nan")], [float("inf")]])
def test_to_cents_rejects_precision_loss_and_missing(bad):
    with pytest.raises(ValueError):
        fin.to_cents(bad)


@pytest.mark.parametrize("text, cents", [("$1,234.50", 123450), (" 12 ", 1200), ("(12.30)", -1230),
                                         ("0.1", 10), ("12.345", None), ("abc", None), ("", None)])
def test_money_text_to_cents(text, cents):
    assert fin.money_text_to_cents(text) == cents


# ------------------------------------------------------------- rounding
@pytest.mark.parametrize("num, den, out", [(5, 2, 3), (-5, 2, -3), (7, 3, 2), (-7, 3, -2), (1, 2, 1), (0, 9, 0)])
def test_div_round_half_away_from_zero(num, den, out):
    assert fin.div_round(num, den) == out


def test_div_round_zero_denominator():
    with pytest.raises(ZeroDivisionError):
        fin.div_round(1, 0)


def test_pct_rounds_half_away_from_zero_not_bankers():
    # 0.5% rounded to 0 dp: Python's round() would give 0, Excel gives 1.
    assert fin.pct(5, 1000, dp=0) == 1
    assert fin.pct(-5, 1000, dp=0) == -1
    assert fin.pct(1, 3) == 33.33
    assert fin.pct(2, 3) == 66.67


@pytest.mark.parametrize("den", [0, -100])
def test_pct_of_zero_or_negative_target_is_missing(den):
    assert math.isnan(fin.pct(10, den))


def test_pct_vector_matches_scalar():
    rng = np.random.default_rng(1)
    num = rng.integers(-10**9, 10**9, 500)
    den = rng.integers(-10**3, 10**9, 500)
    vec = fin.pct(num, den)
    for n, d, v in zip(num, den, vec, strict=True):
        s = fin.pct(int(n), int(d))
        assert (math.isnan(s) and math.isnan(v)) or s == v


# --------------------------------------------------------------- status
@pytest.mark.parametrize("var, expected", [
    (301, config.STATUS_FAVORABLE),      # +3.01%
    (300, config.STATUS_WITHIN),         # exactly +3% is inside the tolerance band
    (0, config.STATUS_WITHIN),
    (-300, config.STATUS_WITHIN),        # exactly -3%
    (-301, config.STATUS_UNFAVORABLE),   # -3.01%
])
def test_status_boundaries(var, expected):
    assert fin.status(var, 10_000) == expected


@pytest.mark.parametrize("target", [0, -5])
def test_status_no_target(target):
    assert fin.status(100, target) == config.STATUS_NO_TARGET


def test_status_vector():
    out = fin.status(np.array([301, -301, 0, 5]), np.array([10_000, 10_000, 10_000, 0]))
    assert out.tolist() == [config.STATUS_FAVORABLE, config.STATUS_UNFAVORABLE,
                            config.STATUS_WITHIN, config.STATUS_NO_TARGET]


def test_status_labels_are_not_on_target():
    # "On Target" was misleading for a +/-3% band; the label must say tolerance.
    assert config.STATUS_WITHIN == "Within tolerance"


# ----------------------------------------------------------- materiality
@pytest.mark.parametrize("var, target, expected", [
    (10_000, 100_000, (True, "Percent")),                 # exactly 10%
    (9_999, 100_000, (False, "")),                        # just under 10%
    (-10_000, 100_000, (True, "Percent")),                # sign does not matter
    (7_500_000, 10**12, (True, "Amount")),                # exactly $75,000
    (7_499_999, 10**12, (False, "")),
    (7_500_000, 1_000_000, (True, "Percent and amount")),
    (100, 0, (False, "")),                                # no target: % test skipped
    (7_500_000, 0, (True, "Amount")),
])
def test_material(var, target, expected):
    assert fin.material(var, target) == expected


# --------------------------------------------------------------- bridges
@pytest.fixture
def lines():
    rng = np.random.default_rng(7)
    n = 2_000
    tq, aq = rng.integers(0, 500, n), rng.integers(0, 500, n)
    tp = rng.integers(100, 200_000, n)
    ap = tp + rng.integers(-5_000, 5_000, n)
    tcu = (tp * rng.uniform(0.3, 0.8, n)).astype("int64")
    acu = tcu + rng.integers(-2_000, 2_000, n)
    return tq, aq, tp, ap, tcu, acu


def test_revenue_bridge_is_exact(lines):
    tq, aq, tp, ap, _, _ = lines
    vol, prc = fin.revenue_bridge(tq, aq, tp, ap)
    np.testing.assert_array_equal(vol + prc, aq * ap - tq * tp)


def test_profit_bridge_is_exact(lines):
    tq, aq, tp, ap, tcu, acu = lines
    pvol, prc, cost = fin.profit_bridge(tq, aq, tp, ap, tcu, acu)
    np.testing.assert_array_equal(pvol + prc + cost, aq * (ap - acu) - tq * (tp - tcu))


def test_split_volume_adds_back(lines):
    tq, aq, tp, _, _, _ = lines
    vol, pure, mix = fin.split_volume(tq.sum(), aq.sum(), (tq * tp).sum(), (aq * tp).sum())
    assert pure + mix == vol == int(((aq - tq) * tp).sum())


def test_split_volume_no_mix_when_prices_are_equal():
    tq, aq = np.array([10, 20, 30]), np.array([12, 15, 40])
    price = 2_500
    vol, pure, mix = fin.split_volume(tq.sum(), aq.sum(), (tq * price).sum(), (aq * price).sum())
    assert mix == 0 and pure == vol == (aq.sum() - tq.sum()) * price


def test_split_volume_without_target_quantity_is_all_mix():
    vol, pure, mix = fin.split_volume(0, 10, 0, 50_000)
    assert (vol, pure, mix) == (50_000, 0, 50_000)
