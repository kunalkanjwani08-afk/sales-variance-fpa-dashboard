# Methodology

How the data is made, cleaned, analysed and checked. Thresholds and file paths are set in `fpa_lib/config.py`; the arithmetic is in `fpa_lib/finance.py`.

## 1. Synthetic data

`scripts/generate_synthetic_data.py` builds a simulated distributor with a fixed random seed (`20240101`), so the same files are produced on every run.

**Grain.** One row is one **sales line**: one account (customer) buying one product in one region in one month. Each region × product has a varying number of accounts of different sizes; most accounts trade for the whole period, some start later (new accounts) and some stop early (lost accounts). So a region × product × month can have several lines, or none. Over 24 months there are about 21,000 true lines.

**How targets are set.** Target quantity = the account's baseline demand × a planned seasonal profile × a 6% stretch × 5% planned annual growth (the same growth for every region). Planners used only 60% of the true seasonal swing, so peak months are under-planned and quiet months over-planned. New accounts have no target until the next year's plan. Target price = 2024 list price × regional price index × the account's price tier (−3%, 0 or +2%) × the planned 2025 increase for its category. Target unit cost = list price × (1 − category margin, varied ±3 points per product) × 2% planned cost inflation in 2025 (Latin America costs 2% more).

**What actually happens** (the "story" the analysis has to discover):

| Built-in driver | Effect |
|---|---|
| Regional demand trends differ (Asia Pacific and Middle East grow faster than plan, Europe is flat, Latin America shrinks and starts below plan) | Volume variance by region; a plan that ignores the trend |
| Promotions: 6-12% discounts in two random months a year per region × category, plus a November holiday promotion (8-12%) for Electronics and Home Appliances; demand is price-elastic (1.6) | More units, lower price, lower gross profit |
| Europe passes on only 40% of the planned 2025 price increase; Asia Pacific achieves about 1% above target price in 2025 | Negative price effect in Europe, small positive one in Asia Pacific |
| Home Appliances unit costs rise about 8% from March 2025 | Negative cost effect on gross profit |
| Electronics unit costs fall about 3% in 2025 | Positive cost effect |
| North America Electronics stock-out, August-October 2024 (demand × 0.7) | Lost volume in those months |
| A few genuine large orders (about 0.3% of lines, 5× normal size) | Valid outliers the cleaner must keep, not "fix" |
| New accounts with no target | Sales with target = 0 ("No target" status) |

**Messy raw file.** The clean "truth" is then corrupted on purpose: padded headers, an empty extra column, blank rows, exact duplicates, complementary and conflicting duplicate Line_IDs, 30+ spellings of region names, product and category spellings, hidden characters, numbers stored as text (`$1,234.50`), null spellings (`NULL`, `N/A`, `-`), dates as text, Excel serial numbers and day-first or month-first strings, missing values, typos (×10), transposed digits, sign errors, wrong derived values, zero prices, and labels or dates that contradict the Line_ID. Every planted defect is written to `data/raw/injected_defects.csv` with the expected handling. **The cleaner never reads this file**; only the tests use it.

## 2. Cleaning

`scripts/clean_sales_data.py`, using the rules in `fpa_lib/cleaning.py`:

1. **Schema.** Trim header names, drop empty unexpected columns (and report non-empty ones), stop with a clear error if a required column is missing or duplicated, or the file has no rows. Only `.xlsx`, `.xlsm` and `.csv` are accepted.
2. **Rows.** Drop entirely blank rows; remove exact duplicates (first copy kept, the removed copy is listed).
3. **Parsing.** Hidden whitespace, null spellings, text numbers and money, several date formats. A number with more than two decimals is treated as unreadable, never rounded.
4. **Key.** `Line_ID` (`REG-PRODCODE-YYYYMM-NNN`) is the record key. Repeated Line_IDs are merged when the copies agree (blanks filled from the copy). When they conflict, **all copies go to review**: no winner is chosen.
5. **Labels and dates.** Explicit mapping tables only. Blank region and date come from the Line_ID, a blank category from the product master. A label or date that contradicts the Line_ID goes to review.
6. **Numbers.** Two identities must hold on each side (target and actual), in cents:
   `Sales = Qty × Unit_Price` and `Gross_Profit = Sales − Qty × Unit_Cost`.
   - Missing values are filled only when an identity gives the exact answer (for example `Qty = Sales / Unit_Price` with no remainder).
   - If the identities fail, the cleaner tests "exactly one field is wrong" for each field. The fix is kept only if **exactly one** field explains the failure. A wrong gross profit and a wrong unit cost cannot be told apart with these two identities, so such rows go to review rather than being guessed.
7. **Checks.** Row reconciliation, unique and well-formed keys, both identities on every row, master-data consistency, no blanks, and every change present in the audit.

Large orders (actual above 3× target) pass every identity, so they are **kept and flagged**, not changed.

The cleaning workbook has the clean data plus `Row_Reconciliation`, `Cleaning_Log`, `Corrections_Audit` (one row per changed value: original, new value, rule, certainty), `Review_Required` (with the reason), `Removed_Duplicates`, `Value_Mappings`, `Data_Quality` and `Validation_Checks`.

## 3. Variance analysis

`scripts/sales_variance_analysis.py`, using `fpa_lib/analysis.py`.

**Exact arithmetic.** All money is converted to integer cents (`to_cents` refuses values with more than two decimals). Sums, bridges and identities are therefore exact. Percentages are rounded half away from zero with integer arithmetic, as Excel does (Python's `round` would round 0.5 to even).

**Line-level bridges** (each adds back to the line variance exactly):

```
Revenue:   Volume = (AQ − TQ) × TP           Price = AQ × (AP − TP)
           Volume + Price = Actual sales − Target sales

Profit:    Volume = (AQ − TQ) × (TP − TC)    Price = AQ × (AP − TP)    Cost = −AQ × (AC − TC)
           Volume + Price + Cost = Actual GP − Target GP
```
(TQ/AQ = target/actual quantity, TP/AP = price, TC/AC = unit cost.)

**Pure volume and mix** (any group of lines). The group's volume effect is split:

```
pure volume = (Σ AQ − Σ TQ) × (Σ Target sales / Σ TQ)     rounded to the cent
mix         = volume effect − pure volume
```

Mix is the value of selling a different combination of lines (products, regions, accounts) than planned, measured at target prices. The same split is applied to the profit bridge using target unit margin. Mix depends on the grouping: the split for "company" differs from the sum of the splits for each region, while volume + price (and the total) never change. A group with no target quantity has no pure volume, so its whole volume effect is mix.

**Status and materiality** (exact integer comparisons, no rounding):

| Rule | Definition |
|---|---|
| Favorable | variance / target > +3% |
| Unfavorable | variance / target < −3% |
| Within tolerance | between −3% and +3% inclusive (shown as "Within ±3%") |
| No target | target ≤ 0 (the percentage is left blank, not shown as 0) |
| Material | \|variance %\| ≥ 10% **or** \|variance\| ≥ $75,000 |

Status and materiality are separate on purpose: a region can be within ±3% and still be a material dollar miss.

**Other analysis**
- Monthly view: month-on-month and year-on-year growth, year-to-date, rolling 3 months, cumulative variance, tracking signal (cumulative variance / mean absolute monthly variance), and target vs actual share of the year (phasing gap).
- Price bands: lines grouped by actual price vs target price (discount ≥ 5%, 1-5%, within 1%, above target by more than 1%).
- Target accuracy: 2025 targets by region and month compared with a naive forecast (the same month of 2024): WAPE = Σ|actual − forecast| / Σ actual, and bias.
- Persistence: region × category pairs below target in at least 75% of months.
- Pareto of the product shortfall, an exceptions list (region × category × month cells that are both Unfavorable and material), and the material lines.

**Validation** (`Validation_Summary` sheet): line count and key; both bridges on every line; the four money totals re-added from the CSV **text** with `Decimal` (an independent method); for every summary table, pure volume + mix + price = sales variance and the profit bridge adds back; tables add back to the company total; status re-derived with Decimal arithmetic. The script exits with an error if any check fails.

## 4. Report and dashboard

`scripts/build_insights.py` writes the insights report, `docs/key_metrics.json` (each headline claim with its expected value and tolerance) and the results table in the README, all from the analysis outputs. `dashboard/build_variance_dashboard.py` packs the line data, the workbook tables and the report into one HTML file. In the browser, every figure is recomputed from the lines in integer cents; when the page opens it reconciles its totals against the workbook and checks every headline claim, and shows the result in the header.

## 5. Testing

`tests/` (pytest) covers the money and rounding rules, threshold boundaries, parsers and repair rules, schema errors, and a full pipeline run in a temporary folder. That run checks the cleaned data **cell by cell against the generator's ground truth**, checks that the only missing lines are the planned review cases, reads both workbooks' validation sheets, recomputes the headline numbers independently and confirms the committed outputs are up to date. A headless-browser test opens the dashboard and checks that there are no errors, that reconciliation passes, that filtered totals match pandas, that the page works from the keyboard, and that nothing overflows at phone width.
