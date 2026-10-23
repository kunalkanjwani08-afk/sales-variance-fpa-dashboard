"""Open the freshly built dashboard in headless Chromium and check it behaves.

Skipped automatically when Playwright or its browser is not installed.
Locally:  pip install -r requirements-dev.txt  then  python -m playwright install chromium
Set PLAYWRIGHT_CHROMIUM_EXECUTABLE to use an already-installed Chromium instead.
"""
import os

import pandas as pd
import pytest

pytestmark = [pytest.mark.browser, pytest.mark.slow]
sync_api = pytest.importorskip("playwright.sync_api")

TABS = ["overview", "drivers", "trends", "regcat", "products", "exceptions", "insights", "data"]


@pytest.fixture(scope="module")
def browser():
    with sync_api.sync_playwright() as p:
        exe = os.environ.get("PLAYWRIGHT_CHROMIUM_EXECUTABLE") or None
        try:
            b = p.chromium.launch(executable_path=exe)
        except Exception as exc:  # browser binary not installed
            pytest.skip(f"Chromium not available: {exc}".splitlines()[0])
        yield b
        b.close()


def _open(browser, pipeline, width=1440, height=900):
    page = browser.new_page(viewport={"width": width, "height": height})
    errors = []
    page.on("pageerror", lambda e: errors.append(f"pageerror: {e}"))
    page.on("console", lambda m: errors.append(f"console.{m.type}: {m.text}") if m.type == "error" else None)
    page.goto(pipeline["html"].as_uri())
    page.wait_for_function("window.__dash && window.__dash.RECON")
    return page, errors


@pytest.fixture(scope="module")
def desktop(browser, pipeline):
    page, errors = _open(browser, pipeline)
    yield page, errors
    page.close()


def test_every_tab_renders_without_errors(desktop):
    page, errors = desktop
    for tab in TABS:
        page.evaluate(f"__dash.setTab('{tab}')")
        broken = page.evaluate("[...document.querySelectorAll('#view .note')].map(n => n.textContent)"
                               ".filter(t => t.startsWith('Something went wrong'))")
        assert broken == [], (tab, broken)
        assert page.evaluate("document.querySelector('#view').children.length") > 0, tab
    assert errors == []


def test_status_label_shows_the_tolerance(desktop):
    page, _ = desktop
    page.evaluate("__dash.setTab('overview')")
    assert "Within ±3%" in page.inner_text("body")
    assert "On Target" not in page.inner_text("body")


def test_in_browser_reconciliation_and_claims_pass(desktop):
    page, _ = desktop
    page.evaluate("__dash.resetFilters()")
    failed = page.evaluate("__dash.RECON.filter(r => !r.pass).map(r => r.name)")
    assert page.evaluate("__dash.RECON.length") >= 20 and failed == []
    bad_claims = page.evaluate("__dash.headlineClaims().filter(r => !r[3]).map(r => r[0])")
    assert bad_claims == []


@pytest.mark.parametrize("region", ["Latin America", "Asia Pacific"])
def test_filtered_totals_match_pandas(desktop, pipeline, region):
    page, _ = desktop
    e = pd.read_csv(pipeline["enriched"], usecols=["Region", "Target_Sales_USD", "Actual_Sales_USD",
                                                   "Target_Gross_Profit_USD", "Actual_Gross_Profit_USD"])
    sub = e[e["Region"] == region]
    cents = lambda s: int(round(s.mul(100).round().sum()))  # noqa: E731
    page.evaluate("__dash.resetFilters()")
    page.evaluate(f"__dash.pickRegion(__dash.REG.indexOf({region!r}))")
    t = page.evaluate("(() => { const t = __dash.totals(__dash.getMask()); return [t.tc, t.ac, t.tgc, t.agc]; })()")
    page.evaluate("__dash.resetFilters()")
    assert [round(x) for x in t] == [cents(sub["Target_Sales_USD"]), cents(sub["Actual_Sales_USD"]),
                                     cents(sub["Target_Gross_Profit_USD"]), cents(sub["Actual_Gross_Profit_USD"])]


def test_chart_marks_work_from_the_keyboard(desktop):
    page, _ = desktop
    page.evaluate("__dash.resetFilters(); __dash.setTab('overview')")
    marks = page.locator("#view svg [tabindex='0']")
    assert marks.count() > 0
    marks.first.focus()
    page.keyboard.press("Enter")
    assert page.evaluate("__dash.S.regs.size + __dash.S.cats.size + __dash.S.prods.size + "
                         "(__dash.S.m0 !== 0 || __dash.S.m1 !== __dash.MON.length - 1 ? 1 : 0)") >= 1
    page.evaluate("__dash.resetFilters()")


def test_phone_width_has_no_page_overflow(browser, pipeline):
    page, errors = _open(browser, pipeline, width=390, height=844)
    try:
        for tab in TABS:
            page.evaluate(f"__dash.setTab('{tab}')")
            width = page.evaluate("document.documentElement.scrollWidth")
            assert width <= 391, (tab, width)
        assert errors == []
    finally:
        page.close()
