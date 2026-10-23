# Data dictionary

All amounts are in US dollars with at most two decimals. Internally every calculation uses integer cents. Percentages are rounded half away from zero to two decimals. A blank percentage means the target is zero (a percentage of a zero target is not meaningful).

## 1. Clean sales lines: `data/processed/Cleaned_Sales_Data.csv`

One row = one sales line (one account, one product, one region, one month).

| Column | Type | Description |
|---|---|---|
| `Line_ID` | text | Record key `REG-PRODCODE-YYYYMM-NNN`, e.g. `LAT-HOM01-202403-004`. Always read as text (leading zeros matter). |
| `Date` | date | First day of the month (`YYYY-MM-DD`). |
| `Year`, `Month`, `Month_Name` | int, int, text | Derived from `Date`. |
| `Region` | text | North America, Europe, Asia Pacific, Latin America, Middle East. |
| `Category` | text | Electronics, Home Appliances, Office Supplies, Personal Care, Sports Equipment. |
| `Product` | text | One of 27 products; the category is fixed by the product master in `fpa_lib/config.py`. |
| `Target_Qty`, `Actual_Qty` | int ≥ 0 | Units. `Target_Qty = 0` means the line had no target (new account). |
| `Target_Unit_Price_USD`, `Actual_Unit_Price_USD` | money > 0 | Selling price per unit. |
| `Target_Unit_Cost_USD`, `Actual_Unit_Cost_USD` | money > 0 | Cost per unit. |
| `Target_Sales_USD`, `Actual_Sales_USD` | money | Always exactly `Qty × Unit_Price`. |
| `Target_Gross_Profit_USD`, `Actual_Gross_Profit_USD` | money | Always exactly `Sales − Qty × Unit_Cost`. |
| `Notes` | text | Promotion, Stock-out, Late shipment, Returns, Price increase, Cost increase, or blank. |
| `Repaired_Fields` | text | Columns whose value was repaired on this line (`;`-separated); blank if untouched. Details in the `Corrections_Audit` sheet. |
| `DQ_Flag` | text | `Large order (over 3x target) - verified by identities, kept` or blank. Flagged, never changed. |
| `Source_Row` | int | Row number in the raw Excel file (header = row 1), for tracing back. |

`Cleaned_Sales_Data.xlsx` has the same data plus the audit sheets: `Read_Me`, `Row_Reconciliation`, `Cleaning_Log`, `Corrections_Audit`, `Review_Required`, `Removed_Duplicates`, `Value_Mappings`, `Data_Quality`, `Validation_Checks`.

## 2. Enriched lines: `data/processed/Sales_Variance_Enriched_Data.csv`

All clean columns, plus:

| Column | Definition |
|---|---|
| `Year_Month` | `YYYY-MM` |
| `Quarter` | `Q1`-`Q4` |
| `Qty_Variance` | `Actual_Qty − Target_Qty` |
| `Qty_Attainment_Pct` | `Actual_Qty / Target_Qty × 100` |
| `Sales_Variance_USD` | `Actual_Sales − Target_Sales` (positive = above target) |
| `Sales_Variance_Pct` | `Sales_Variance_USD / Target_Sales × 100` |
| `Sales_Variance_Status` | `Favorable` (> +3%), `Unfavorable` (< −3%), `Within tolerance` (−3% to +3% inclusive), `No target` (target = 0) |
| `Material_Variance_Flag` | `Yes` if \|variance %\| ≥ 10 or \|variance\| ≥ $75,000 |
| `Material_Trigger` | `Percent`, `Amount`, `Percent and amount`, or blank |
| `Volume_Effect_USD` | `(Actual_Qty − Target_Qty) × Target_Unit_Price` |
| `Price_Effect_USD` | `Actual_Qty × (Actual_Unit_Price − Target_Unit_Price)` |
| `Price_Realisation_Pct` | `Price_Effect / (Actual_Qty × Target_Unit_Price) × 100`: average price vs target |
| `Target_COGS_USD`, `Actual_COGS_USD` | `Qty × Unit_Cost` |
| `Gross_Profit_Variance_USD`, `_Pct` | `Actual_Gross_Profit − Target_Gross_Profit`, and as % of target gross profit |
| `Profit_Volume_Effect_USD` | `(Actual_Qty − Target_Qty) × (Target_Unit_Price − Target_Unit_Cost)` |
| `Profit_Price_Effect_USD` | Same as `Price_Effect_USD` |
| `Profit_Cost_Effect_USD` | `−Actual_Qty × (Actual_Unit_Cost − Target_Unit_Cost)` (negative = cost above target) |
| `Unit_Cost_Change_Pct` | `(Actual_Unit_Cost − Target_Unit_Cost) / Target_Unit_Cost × 100` |
| `Target_Gross_Margin_Pct`, `Actual_Gross_Margin_Pct` | `Gross_Profit / Sales × 100` |

Identities that hold on every line: `Volume + Price = Sales variance` and `Profit volume + Profit price + Profit cost = Gross-profit variance`.

**A note on line-level materiality.** Individual lines are small and noisy, so about half of them move by 10% or more against target and are flagged material. The line flag is for drilling down. The review list for management is the `Exceptions_RegCat_Month` sheet: region × category × month cells that are both Unfavorable and material.

## 3. Analysis workbook: `outputs/Sales_Variance_Analysis.xlsx`

| Sheet | Contents |
|---|---|
| `Executive_Summary` | Headline figures with comments |
| `Validation_Summary` | Every check with PASS/FAIL and detail (the script fails if any check fails) |
| `Metric_Dictionary` | Formula and notes for each metric |
| `Company`, `By_Year`, `By_Quarter`, `By_Region`, `By_Category`, `By_Region_Category`, `By_Product`, `By_Region_Product`, `By_Region_Month` | Summary tables (columns below) |
| `By_Year_Month` | Summary columns plus month-on-month and year-on-year growth, year-to-date, rolling 3 months, cumulative variance, tracking signal and phasing gap |
| `By_Price_Band` | Lines grouped by actual vs target price (discount ≥ 5%, 1-5%, within 1%, above target by more than 1%) |
| `Target_Accuracy` | 2025 targets vs a naive same-month-last-year forecast, by region: WAPE, bias, months where the target was closer |
| `Persistence_Region_Category` | Months below target for each region × category, and whether it is persistent (≥ 75% of months) |
| `Pareto_Shortfall_Product` | Products below target, ranked, with cumulative share of the shortfall |
| `Exceptions_RegCat_Month` | Region × category × month cells that are Unfavorable and material |
| `Material_Lines` | Every material line, for drill-down |
| `Data_Quality_Impact` | Variance for all lines, excluding repaired lines, and only repaired lines (to show the repairs do not drive the result), plus the target value of rows held for review |

**Summary table columns** (each group): `Lines`, `Target_Qty`, `Actual_Qty`, `Qty_Variance`, `Qty_Attainment_Pct`, `Target_Sales_USD`, `Actual_Sales_USD`, `Sales_Variance_USD`, `Sales_Variance_Pct`, `Sales_Variance_Status`, `Material_Variance_Flag`, `Material_Trigger`, `Volume_Effect_USD`, `Volume_Pure_USD`, `Mix_Effect_USD`, `Price_Effect_USD`, `Target_Gross_Profit_USD`, `Actual_Gross_Profit_USD`, `Gross_Profit_Variance_USD`, `Gross_Profit_Variance_Pct`, `Profit_Volume_Effect_USD`, `Profit_Volume_Pure_USD`, `Profit_Mix_Effect_USD`, `Profit_Price_Effect_USD`, `Profit_Cost_Effect_USD`, `Target_Gross_Margin_Pct`, `Actual_Gross_Margin_Pct`, `Price_Realisation_Pct`, average target and actual prices, `Lines_Below_Target_Pct` (share of lines with variance < 0), `Lines_Unfavorable_Pct` (share below −3%), `Material_Lines`, `Gross_Margin_Change_pp`, and for split tables `Share_of_Net_Variance_Pct` and `Share_of_Target_Pct`.

Group percentages are always `Σ variance / Σ target`. In every table `Volume_Pure_USD + Mix_Effect_USD + Price_Effect_USD = Sales_Variance_USD` exactly.

## 4. Raw data and answer key: `data/raw/`

| File | Contents |
|---|---|
| `sales_raw_messy.xlsx` | The messy source: the clean columns above (no `Repaired_Fields`, `DQ_Flag`, `Source_Row`), plus an empty `Comments` column, with planted defects. |
| `injected_defects.csv` | One row per planted defect: `Line_ID`, `Column`, `Defect`, `Clean_Value`, `Dirty_Value`, `Expected_Handling`. Used only by the tests and notebook 01; the cleaner never reads it. |

## 5. `docs/key_metrics.json`

Every headline number used in the insights report and the README, as `{key, label, text, expected, unit, tol}`. The dashboard re-computes each one from the line data and shows whether it still matches. The tests compare them with the workbook.
