# config.py
# Central configuration untuk KPI Validator.
# Edit file ini jika ada perubahan nama file, segment, atau threshold.

from datetime import datetime

# ── Segment Mapping ───────────────────────────────────────────────────────────
# Maps internal codes di source file → display names di laporan
SEGMENT_MAP = {
    "CODE_A": "SEG_A",
    "CODE_B": "SEG_B",
    "CODE_C": "SEG_C",
    "CODE_D": "SEG_D",
    "ALL":    "TOTAL",
}

# Reverse map: display name → internal code
SEGMENT_MAP_REVERSE = {v: k for k, v in SEGMENT_MAP.items()}

SEGMENTS_DETAIL = ["CODE_A", "CODE_B", "CODE_C", "CODE_D"]  # individual segments
SEGMENT_BSGS    = ["CODE_A", "CODE_B"]                       # untuk derived SEG_A+SEG_B

# ── Month Config ──────────────────────────────────────────────────────────────
MONTH_COLS_MTD = ["JAN","FEB","MAR","APR","MAY","JUN",
                  "JUL","AUG","SEP","OCT","NOV","DEC"]

MONTH_NUMBER_MAP = {
    1:"JAN", 2:"FEB",  3:"MAR", 4:"APR",  5:"MAY",  6:"JUN",
    7:"JUL", 8:"AUG",  9:"SEP", 10:"OCT", 11:"NOV", 12:"DEC",
}

# ── Source File Names ─────────────────────────────────────────────────────────
# Taruh semua file di data/input/
# Update nama file di sini, sesuaikan nama file karna ini akan berpengaruh pada loader
# source data input / data yang akan diolah

SOURCE_REVENUE              = "REVENUE FILE.xlsx"
SOURCE_PRODA_SRCA           = "PRODA FILE.csv"
SOURCE_PRODB_SRCA           = "PRODB FILE.csv"
SOURCE_PRODA_SRCB           = "PRODA FILE 2.csv"
SOURCE_PRODB_SRCB           = "PRODB FILE 2.csv"
SOURCE_PRODC_SRCB           = "PRODC FILE.csv"

SOURCE_COLLECTION              = "data_collection.xlsx"

# source data target dan historical
SOURCE_TARGET               = "data target.xlsx"
SOURCE_HISTORICAL           = "data historical.xlsx"


# ── Revenue Source Sheet Config ───────────────────────────────────────────────
REVENUE_SHEET      = "REVENUE_DATA"
REVENUE_HEADER_ROW = 2   # 0-indexed

REVENUE_FILTER = {"INDICATOR": "REVENUE", "TYPE": "ALL"}
GROWTH_FILTER  = {"INDICATOR": "GROWTH",  "TYPE": "GROWTH"}

# ── Target File Config ────────────────────────────────────────────────────────
TARGET_SHEET      = "Revenue"
TARGET_HEADER_ROW = 1    # 0-indexed (row ke-2 di Excel)

TARGET_SALES_CONNECTIVITY = "Sales Connectivity"
TARGET_SALES_CONNECTIVITY_HEADER_ROW = 1

COLLECTION_DATA_SHEET = "collection"

TARGET_COLLECTION = "Collection"
TARGET_COLLECTION_HEADER_ROW = 1

# ── Historical File Config ────────────────────────────────────────────────────
HISTORICAL_SHEET      = "Revenue"
HISTORICAL_HEADER_ROW = 1   # 0-indexed

HISTORICAL_SALES_CONNECTIVITY     = "Sales Connectivity"
HISTORICAL_SALES_CONNECTIVITY_HEADER_ROW = 1

HISTORICAL_COLLECTION     = "Collection"
HISTORICAL_COLLECTION_HEADER_ROW = 1


# Nama INDICATOR di file historical
HIST_REVENUE = "REVENUE"
HIST_GROWTH  = "GROWTH"
HIST_PRODA   = "PROD_A"
HIST_PRODC   = "PROD_C"
HIST_PRODB   = "PROD_B"


# ── filter mapping source A PROD_A
SRCA_PRODUCT_MAPPING = {
    "PROD_A": ["PRODUCT_A_1", "PRODUCT_A_2", "PRODUCT_A_3", "PRODUCT_A_4"],
    "PROD_C": ["PRODUCT_C_1", "PRODUCT_C_2", "PRODUCT_C_3"]
}

# ── filter mapping source A PROD_B
SRCA_REGION_MAPPING = {
    "REGION_1": "REGION_1_FULL",
    "REGION_2": "REGION_2_FULL",
    "REGION_3": "REGION_3_FULL",
    "REGION_4": "REGION_4_FULL",
    "REGION_5": "REGION_5_FULL",
    "REGION_6": "REGION_6_FULL",
}

# ── Mapping divisi source B
SRCB_DIVISION_MAPPING = {
    "DIV_B": "SEG_B",
    "DIV_C": "SEG_C",
    "DIV_D": "SEG_D"
}

# ── Mapping data collection performance excel
COLLECTION_SEGMENT_MAPPING = {
    "SEG_A": "SEG_A",
    "SEG_B": "SEG_B",
    "SEG_C": "SEG_C",
    "SEG_D": "SEG_D",
    "SEG_A+SEG_B": "SEG_A+SEG_B",
    "TOTAL": "TOTAL"
}
ACTIVE_COLS_COLLECTION = ['INDICATOR', 'SEGMENT', 'TERRITORY_NAME', 'YEAR']


# ── Validation Thresholds ─────────────────────────────────────────────────────
MOM_ANOMALY_THRESHOLD = 0.50   # flag jika MoM change > 50%
ACH_LOW_THRESHOLD     = 0.70   # flag jika achievement < 70%

# ── Output Config ─────────────────────────────────────────────────────────────
OUTPUT_DIR          = "data/output"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILENAME_FMT = "kpi_summary_{territory}_{month}_{year}_{timestamp}.xlsx"
