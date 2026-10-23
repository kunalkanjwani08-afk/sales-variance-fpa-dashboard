"""Unit tests for fpa_lib.cleaning and the clean() step of scripts/clean_sales_data.py."""
from datetime import datetime

import pandas as pd
import pytest

from fpa_lib import cleaning as cl
from fpa_lib import config


# ------------------------------------------------------------------ text
def test_clean_text_strips_hidden_characters():
    assert cl.clean_text("​Europe  ") == "Europe"
    assert cl.clean_text("  a   b ") == "a b"
    assert cl.clean_text("   ") is None
    assert cl.clean_text(float("nan")) is None


def test_na_means_north_america_only_in_region():
    assert cl.is_null("NA", "Region") is False
    assert cl.is_null("NA", "Actual_Qty") is True
    assert cl.is_null(" n/a ") is True
    assert cl.map_label("NA", cl.REGION_ALIASES)[0] == "North America"


def test_explicit_mapping_never_guesses():
    assert cl.map_label(" N. America ", cl.REGION_ALIASES) == ("North America", "N. America")
    assert cl.map_label("LATAM", cl.REGION_ALIASES)[0] == "Latin America"
    assert cl.map_label("Narnia", cl.REGION_ALIASES) == (None, "Narnia")
    assert cl.map_label("", cl.REGION_ALIASES) == (None, None)


# --------------------------------------------------------------- Line_ID
def test_line_id_keeps_leading_zeros_and_decodes_key():
    code, product = next(iter(config.PRODUCT_BY_CODE.items()))
    key = cl.parse_line_id(f" LAT-{code}-202403-004 ")
    assert key["line_id"] == f"LAT-{code}-202403-004"
    assert key["region"] == "Latin America"
    assert key["product"] == product
    assert key["period"] == pd.Period("2024-03", freq="M")


@pytest.mark.parametrize("bad", ["LAT-XXX99-202403-004", "LAT-{code}-202413-004", "lat-{code}-202403-004",
                                 "LAT-{code}-202403-04", "", None])
def test_malformed_line_id(bad):
    code = next(iter(config.PRODUCT_BY_CODE))
    assert cl.parse_line_id(bad.format(code=code) if bad else bad) is None


# --------------------------------------------------------------- numbers
@pytest.mark.parametrize("value, expected", [
    (5, (5, "ok")), (5.0, (5, "ok")), (5.5, (None, "invalid")), ("1,200", (1200, "text")),
    ("7.0", (7, "text")), ("abc", (None, "invalid")), (None, (None, "missing")), ("N/A", (None, "missing")),
    (True, (None, "invalid")),
])
def test_parse_qty(value, expected):
    assert cl.parse_qty(value) == expected


@pytest.mark.parametrize("value, expected", [
    ("$1,234.50", (123450, "text")), (" 12.3 ", (1230, "text")), (12.3, (1230, "ok")), (0.1 + 0.2, (30, "ok")),
    (7, (700, "ok")), ("12.345", (None, "invalid")), (1.005, (None, "invalid")), ("-", (None, "missing")),
])
def test_parse_money_never_rounds_silently(value, expected):
    assert cl.parse_money(value) == expected


# ----------------------------------------------------------------- dates
M = lambda s: pd.Period(s, freq="M")  # noqa: E731


@pytest.mark.parametrize("value, periods, form", [
    ("2024-03-01", {M("2024-03")}, "text"),
    ("Mar-2024", {M("2024-03")}, "text"),
    ("March 2024", {M("2024-03")}, "text"),
    ("01/03/2024", {M("2024-03")}, "text"),         # day-first reading is the only one with day == 1
    ("03/01/2024", {M("2024-03")}, "text"),         # month-first reading
    ("01/01/2024", {M("2024-01")}, "text"),
    (45292, {M("2024-01")}, "excel serial"),         # Excel serial for 2024-01-01
    (datetime(2024, 5, 1), {M("2024-05")}, "date"),
    ("2024-03-15", set(), "invalid"),               # not a month start
    ("missing", set(), "missing"),
])
def test_date_candidates(value, periods, form):
    assert cl.date_candidates(value) == (periods, form)


# ------------------------------------------------- identities and repair
GOOD = dict(q=10, p=25_000, u=10_000, s=250_000, g=150_000)


def test_identities_hold():
    assert cl.identities_hold(GOOD)
    assert not cl.identities_hold(GOOD | {"g": 1})


def test_fill_missing_uses_exact_identities_only():
    v, filled = cl.fill_missing(GOOD | {"s": None, "g": None})
    assert v == GOOD and [k for k, _ in filled] == ["s", "g"]
    v, _ = cl.fill_missing(GOOD | {"q": None})
    assert v == GOOD
    # Sales not divisible by price: quantity cannot be solved exactly -> stays missing
    v, filled = cl.fill_missing({"q": None, "p": 30_000, "u": None, "s": 250_000, "g": None})
    assert v["q"] is None and filled == []


def test_two_linked_values_missing_are_not_estimated():
    v, _ = cl.fill_missing(GOOD | {"q": None, "s": None})
    assert v["q"] is None and v["s"] is None


@pytest.mark.parametrize("field, wrong", [("s", 2_500_000), ("q", 100), ("p", 2_500)])
def test_correct_single_error_finds_the_one_wrong_field(field, wrong):
    key, value, _ = cl.correct_single_error(GOOD | {field: wrong})
    assert (key, value) == (field, GOOD[field])


@pytest.mark.parametrize("field, wrong", [("g", 15_000), ("u", 1_000)])
def test_wrong_profit_or_cost_alone_is_ambiguous(field, wrong):
    # With only Sales = Qty x Price and Profit = Sales - Qty x Cost, a wrong profit and a
    # wrong unit cost explain the failure equally well, so no correction is guessed.
    assert cl.correct_single_error(GOOD | {field: wrong}) is None


def test_correct_single_error_refuses_when_not_unique():
    assert cl.correct_single_error(GOOD | {"s": 999_999, "g": 7}) is None


# --------------------------------------------- clean() on a small frame
def _row(n, **over):
    product, (code, category) = next(iter(config.PRODUCTS.items()))
    tq, aq, tp, ap, tc, ac = 10, 12, 250.0, 240.0, 100.0, 105.0
    r = {"Line_ID": f"NAM-{code}-202403-{n:03d}", "Date": datetime(2024, 3, 1), "Year": 2024, "Month": 3,
         "Month_Name": "March", "Region": "North America", "Category": category, "Product": product,
         "Target_Qty": tq, "Actual_Qty": aq, "Target_Unit_Price_USD": tp, "Actual_Unit_Price_USD": ap,
         "Target_Unit_Cost_USD": tc, "Actual_Unit_Cost_USD": ac, "Target_Sales_USD": tq * tp,
         "Actual_Sales_USD": aq * ap, "Target_Gross_Profit_USD": tq * (tp - tc),
         "Actual_Gross_Profit_USD": aq * (ap - ac), "Notes": ""}
    r.update(over)
    return r


@pytest.fixture(scope="module")
def small(scripts):
    rows = [
        _row(1),                                                       # clean
        _row(2), _row(2),                                              # exact duplicate
        _row(3, Actual_Qty=None), _row(3, Actual_Sales_USD=None),      # complementary copies -> merged
        _row(4), _row(4, Actual_Qty=13),                               # conflicting copies -> review
        _row(5, Actual_Qty=None, Actual_Sales_USD=None),               # two linked values missing -> review
        _row(6, Region="Europe"),                                      # label contradicts Line_ID -> review
        _row(7, Region="NA", Actual_Sales_USD="$2,880.00 "),           # NA region + money text
        _row(8, Actual_Sales_USD=28_800.0),                            # typo x10 -> corrected
        _row(9, Line_ID=None),                                         # no key -> review
        {c: None for c in config.CLEAN_COLUMNS},                       # blank row
    ]
    return scripts["clean"].clean(pd.DataFrame(rows, columns=config.CLEAN_COLUMNS, dtype=object))


def test_small_clean_reconciles(small):
    r = small["recon"]
    assert r == {"raw_rows": 13, "blank": 1, "exact_dups": 1, "merged_dups": 1, "review": 5, "clean": 5}
    assert r["raw_rows"] == r["blank"] + r["exact_dups"] + r["merged_dups"] + r["review"] + r["clean"]


def test_small_clean_rows_and_values(small):
    c = small["clean"].set_index("Line_ID")
    assert sorted(i.split("-")[-1] for i in c.index) == ["001", "002", "003", "007", "008"]
    assert (c["Actual_Sales_USD"] == 2_880.0).all()                   # text parsed, typo and blank repaired
    assert (c["Actual_Qty"] == 12).all()
    assert (c["Region"] == "North America").all()
    assert c.loc[c.index.str.endswith("-003"), "Repaired_Fields"].iloc[0] != ""


def test_small_review_reasons(small):
    rv = small["review"]
    reasons = " | ".join(rv["Review_Reason"])
    assert (rv["Line_ID"].str.endswith("-004")).sum() == 2            # all conflicting copies held
    assert "conflicting values" in reasons
    assert "Cannot recover" in reasons
    assert "contradicts Line_ID" in reasons
    assert "no record key" in reasons


def test_small_audit_has_original_values(small):
    a = small["audit"]
    typo = a[(a["Line_ID"].str.endswith("-008")) & (a["Column"] == "Actual_Sales_USD")]
    assert len(typo) == 1 and typo.iloc[0]["Original_Value"] == "28800.0"
    assert typo.iloc[0]["New_Value"] in ("2880.00", "2,880.00", "2880.0")


# ------------------------------------------------------- schema / files
def test_missing_required_column(scripts):
    df = pd.DataFrame([_row(1)]).drop(columns=["Actual_Qty"])
    with pytest.raises(scripts["clean"].CleaningError, match="Actual_Qty"):
        scripts["clean"].clean(df)


def test_header_without_rows(scripts):
    with pytest.raises(scripts["clean"].CleaningError, match="no data rows"):
        scripts["clean"].clean(pd.DataFrame(columns=config.CLEAN_COLUMNS))


def test_duplicate_column_names(scripts):
    df = pd.DataFrame([_row(1)])
    df.columns = [c if c != "Notes" else "Region " for c in df.columns]   # trims to a second "Region"
    with pytest.raises(scripts["clean"].CleaningError, match="Duplicate column"):
        scripts["clean"].clean(df)


def test_padded_headers_and_extra_columns_are_handled(scripts):
    df = pd.DataFrame([_row(1)])
    df.columns = [f" {c} " for c in df.columns]
    df["Unnamed: 20"] = ""
    res = scripts["clean"].clean(df)
    assert len(res["clean"]) == 1
    assert "Column is entirely empty" in set(res["log"]["Issue"])


def test_every_row_held_for_review_gives_empty_clean_table(scripts):
    res = scripts["clean"].clean(pd.DataFrame([_row(1, Line_ID=None)], columns=config.CLEAN_COLUMNS, dtype=object))
    assert res["recon"]["clean"] == 0 and res["recon"]["review"] == 1
    assert list(res["clean"].columns)[: len(config.CLEAN_COLUMNS)] == config.CLEAN_COLUMNS


@pytest.mark.parametrize("name", ["sales.xls", "sales.txt"])
def test_unsupported_file_types_are_rejected(scripts, tmp_path, name):
    (tmp_path / name).write_bytes(b"not a workbook")
    with pytest.raises(scripts["clean"].CleaningError, match="Unsupported file type"):
        scripts["clean"].read_raw(tmp_path / name)


def test_csv_input_matches_excel_semantics(scripts, tmp_path):
    path = tmp_path / "in.csv"
    pd.DataFrame([_row(1, Region="NA", Date="2024-03-01")]).to_csv(path, index=False)
    res = scripts["clean"].clean(scripts["clean"].read_raw(path))
    assert res["clean"].loc[0, "Region"] == "North America"           # 'NA' not read as missing


def test_main_returns_error_code_for_missing_file(scripts, tmp_path, capsys):
    assert scripts["clean"].main(["--input", str(tmp_path / "nope.xlsx"), "--output-dir", str(tmp_path)]) == 2
    assert "not found" in capsys.readouterr().err


def test_main_flags_a_file_where_nothing_passes(scripts, tmp_path):
    pd.DataFrame([_row(1, Line_ID=None)]).to_csv(tmp_path / "in.csv", index=False)
    assert scripts["clean"].main(["--input", str(tmp_path / "in.csv"), "--output-dir", str(tmp_path)]) == 1
    review = pd.read_excel(tmp_path / config.CLEAN_XLSX.name, sheet_name="Review_Required")
    assert len(review) == 1
