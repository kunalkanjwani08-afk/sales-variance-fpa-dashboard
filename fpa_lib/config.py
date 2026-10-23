"""Project configuration: paths, thresholds and master data.

Every threshold used anywhere in the project lives here, so the scripts, the
tests and the dashboard all apply the same rules.
"""
from pathlib import Path

# ----------------------------------------------------------------- paths --
REPO_ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = REPO_ROOT / "data" / "raw"
PROCESSED_DIR = REPO_ROOT / "data" / "processed"
OUTPUTS_DIR = REPO_ROOT / "outputs"
DOCS_DIR = REPO_ROOT / "docs"
DASHBOARD_DIR = REPO_ROOT / "dashboard"

RAW_FILE = RAW_DIR / "sales_raw_messy.xlsx"
DEFECT_MANIFEST_FILE = RAW_DIR / "injected_defects.csv"
CLEAN_CSV = PROCESSED_DIR / "Cleaned_Sales_Data.csv"
CLEAN_XLSX = PROCESSED_DIR / "Cleaned_Sales_Data.xlsx"
ENRICHED_CSV = PROCESSED_DIR / "Sales_Variance_Enriched_Data.csv"
ANALYSIS_XLSX = OUTPUTS_DIR / "Sales_Variance_Analysis.xlsx"
INSIGHTS_MD = DOCS_DIR / "Insights_and_Recommendations.md"
KEY_METRICS_JSON = DOCS_DIR / "key_metrics.json"
DASHBOARD_HTML = DASHBOARD_DIR / "Sales_Variance_Dashboard.html"

# ------------------------------------------------------- synthetic data --
RANDOM_SEED = 20240101          # fixed seed: the same seed always produces the same dataset
FIRST_MONTH = "2024-01"
N_MONTHS = 24

# ---------------------------------------------------------- thresholds --
# Status of a line or group, from variance % = (actual - target) / target:
#   Favorable         variance % >  +FAVORABLE_PCT
#   Unfavorable       variance % <  -UNFAVORABLE_PCT
#   Within tolerance  otherwise (boundaries included)
#   No target         target <= 0
FAVORABLE_PCT = 3
UNFAVORABLE_PCT = 3
# Materiality is separate from status: a variance is material if
#   |variance %| >= MATERIAL_PCT   or   |variance $| >= MATERIAL_USD
# so a group can be "Within tolerance" and still material in dollars.
MATERIAL_PCT = 10
MATERIAL_USD = 75_000
# A region x category pair is a "persistent shortfall" if it is below target
# in at least this share of the months in view.
PERSISTENCE_SHARE = 0.75

# ----------------------------------------------------------- master data --
STATUS_FAVORABLE = "Favorable"
STATUS_WITHIN = "Within tolerance"
STATUS_UNFAVORABLE = "Unfavorable"
STATUS_NO_TARGET = "No target"
STATUSES = [STATUS_FAVORABLE, STATUS_WITHIN, STATUS_UNFAVORABLE, STATUS_NO_TARGET]

REGIONS = {  # code -> name (the code is the first part of every Line_ID)
    "NAM": "North America",
    "EUR": "Europe",
    "APA": "Asia Pacific",
    "LAT": "Latin America",
    "MEA": "Middle East",
}

# product -> (code, category). One category per product.
PRODUCTS = {
    "Bluetooth Speaker": ("ELE01", "Electronics"),
    "Laptop Stand": ("ELE02", "Electronics"),
    "Power Bank": ("ELE03", "Electronics"),
    "Smart Watch": ("ELE04", "Electronics"),
    "Tablet": ("ELE05", "Electronics"),
    "Wireless Earbuds": ("ELE06", "Electronics"),
    "Air Fryer": ("HOM01", "Home Appliances"),
    "Air Purifier": ("HOM02", "Home Appliances"),
    "Blender": ("HOM03", "Home Appliances"),
    "Coffee Maker": ("HOM04", "Home Appliances"),
    "Robot Vacuum": ("HOM05", "Home Appliances"),
    "Toaster": ("HOM06", "Home Appliances"),
    "Ergonomic Chair": ("OFF01", "Office Supplies"),
    "Keyboard": ("OFF02", "Office Supplies"),
    "Monitor Arm": ("OFF03", "Office Supplies"),
    "Standing Desk": ("OFF04", "Office Supplies"),
    "Webcam": ("OFF05", "Office Supplies"),
    "Electric Toothbrush": ("PER01", "Personal Care"),
    "Face Massager": ("PER02", "Personal Care"),
    "Hair Dryer": ("PER03", "Personal Care"),
    "Straightener": ("PER04", "Personal Care"),
    "Trimmer": ("PER05", "Personal Care"),
    "Dumbbell Set": ("SPO01", "Sports Equipment"),
    "Foam Roller": ("SPO02", "Sports Equipment"),
    "Jump Rope": ("SPO03", "Sports Equipment"),
    "Resistance Bands": ("SPO04", "Sports Equipment"),
    "Yoga Mat": ("SPO05", "Sports Equipment"),
}
CATEGORIES = sorted({c for _, c in PRODUCTS.values()})
PRODUCT_BY_CODE = {code: name for name, (code, _) in PRODUCTS.items()}

NOTE_VALUES = ["Promotion", "Price increase", "Stock-out", "Late shipment", "Returns", "Cost increase"]

# The clean schema (column order of Cleaned_Sales_Data.csv). Money is in USD with 2 decimals.
CLEAN_COLUMNS = [
    "Line_ID", "Date", "Year", "Month", "Month_Name", "Region", "Category", "Product",
    "Target_Qty", "Actual_Qty",
    "Target_Unit_Price_USD", "Actual_Unit_Price_USD",
    "Target_Unit_Cost_USD", "Actual_Unit_Cost_USD",
    "Target_Sales_USD", "Actual_Sales_USD",
    "Target_Gross_Profit_USD", "Actual_Gross_Profit_USD",
    "Notes",
]
MONEY_COLUMNS = [c for c in CLEAN_COLUMNS if c.endswith("_USD")]
QTY_COLUMNS = ["Target_Qty", "Actual_Qty"]
