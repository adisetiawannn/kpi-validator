# src/processors/segments/_historical_helper.py
# Load data historis dan hitung YtD dari akumulasi MtD.
#
# Support struktur file:
#   YEAR | TERRITORY | INDICATOR | SEGMENT | JAN | FEB | ... | DEC

import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from config import (
    MONTH_COLS_MTD, MONTH_NUMBER_MAP,
    HISTORICAL_SHEET, HISTORICAL_HEADER_ROW,
    HIST_REVENUE, HIST_GROWTH, HIST_PRODA, HIST_PRODC, HIST_PRODB,
    HISTORICAL_SALES_CONNECTIVITY, HISTORICAL_SALES_CONNECTIVITY_HEADER_ROW,
    TARGET_SALES_CONNECTIVITY, TARGET_SALES_CONNECTIVITY_HEADER_ROW,
)

# ---- LOAD HISTORICAL REVENUE ----
def load_historical_revenue(
    filepath  : str,
    territory : str,
    segment   : str,
) -> dict:
    """
    Load data historis Revenue MtD untuk segment tertentu.
    Hitung YtD dari akumulasi MtD.

    Mendukung dua struktur file:
    1. TERRITORY | INDICATOR | SEGMENT | JAN | ...
    2. YEAR | TERRITORY | INDICATOR | SEGMENT | JAN | ...

    Returns:
        {
          "mtd": { "JAN": float, "FEB": float, ... },
          "ytd": { "JAN": float, "FEB": float, ... },
        }
    """
    df = pd.read_excel(
        filepath,
        sheet_name=HISTORICAL_SHEET,
        header=HISTORICAL_HEADER_ROW,
    )

    # Normalize kolom — strip whitespace
    df.columns = [str(c).strip() for c in df.columns]

    # Drop baris kosong atau baris note (ciri: kolom pertama berisi teks panjang)
    df = df.dropna(subset=["INDICATOR", "SEGMENT"]).copy()

    # Filter: territory, indicator, segment
    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == HIST_REVENUE) &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]

    # Jika tidak ditemukan → return None semua
    if row.empty:
        empty = {m: None for m in MONTH_COLS_MTD}
        return {"mtd": empty, "ytd": empty}

    # Ambil nilai MtD per bulan
    mtd = {}
    for month in MONTH_COLS_MTD:
        if month not in df.columns:
            mtd[month] = None
            continue
        val = row.iloc[0][month]
        mtd[month] = None if pd.isna(val) else float(val)

    # Hitung YtD = akumulasi MtD dari JAN
    ytd = {}; running = 0.0
    for idx, month in MONTH_NUMBER_MAP.items():
        val = mtd.get(month)
        if val is not None: running += val
        ytd[month] = running  # ← selalu carry forward

    return {"mtd": mtd, "ytd": ytd}


# ---- LOAD HISTORICAL SALES CONNECTIVITY ----
def load_historical_sales_connectivity(
    filepath  : str,
    territory : str,
    segment   : str,
    indicator : str,
) -> dict:
    """
    Load data historis Sales Connectivity MtD untuk segment tertentu.
    Hitung YtD dari akumulasi MtD.

    Mendukung dua struktur file:
    1. TERRITORY | INDICATOR | SEGMENT | JAN | ...
    2. YEAR | TERRITORY | INDICATOR | SEGMENT | JAN | ...

    Returns:
        {
          "mtd": { "JAN": float, "FEB": float, ... },
          "ytd": { "JAN": float, "FEB": float, ... },
        }
    """
    df = pd.read_excel(
        filepath,
        sheet_name=HISTORICAL_SALES_CONNECTIVITY,
        header=HISTORICAL_SALES_CONNECTIVITY_HEADER_ROW,
    )

    # Normalize kolom — strip whitespace
    df.columns = [str(c).strip() for c in df.columns]

    # Drop baris kosong atau baris note (ciri: kolom pertama berisi teks panjang)
    df = df.dropna(subset=["INDICATOR", "SEGMENT"]).copy()

    # Filter: territory, indicator, segment
    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == indicator) &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]

    # Jika tidak ditemukan → return None semua
    if row.empty:
        empty = {m: None for m in MONTH_COLS_MTD}
        return {"mtd": empty, "ytd": empty}

    # Ambil nilai MtD per bulan
    mtd = {}
    for month in MONTH_COLS_MTD:
        if month not in df.columns:
            mtd[month] = None
            continue
        val = row.iloc[0][month]
        mtd[month] = None if pd.isna(val) else float(val)

    # Hitung YtD = akumulasi MtD dari JAN
    ytd = {}; running = 0.0
    for idx, month in MONTH_NUMBER_MAP.items():
        val = mtd.get(month)
        if val is not None: running += val
        ytd[month] = running  # ← selalu carry forward

    return {"mtd": mtd, "ytd": ytd}


# ---- LOAD HISTORICAL UNIVERSAL SESUAI TEMPLATE FILE HISTORICAL ----
def load_historical(filepath, territory, segment, indicator):
    """Universal historical loader."""
    if indicator in ["PROD_A", "PROD_B", "PROD_C"]:
        sheet  = HISTORICAL_SALES_CONNECTIVITY
        header = HISTORICAL_SALES_CONNECTIVITY_HEADER_ROW
    else:
        sheet  = HISTORICAL_SHEET
        header = HISTORICAL_HEADER_ROW

    df = pd.read_excel(filepath, sheet_name=sheet, header=header)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(subset=["INDICATOR", "SEGMENT"]).copy()

    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == indicator) &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]

    if row.empty:
        empty = {m: None for m in MONTH_COLS_MTD}
        return {"mtd": empty, "ytd": empty}

    mtd = {}
    for month in MONTH_COLS_MTD:
        if month not in df.columns: mtd[month] = None; continue
        val = row.iloc[0][month]
        mtd[month] = None if pd.isna(val) else float(val)

    ytd = {}; running = 0.0
    for idx, month in MONTH_NUMBER_MAP.items():
        val = mtd.get(month)
        if val is not None: running += val
        ytd[month] = running  # ← selalu carry forward

    return {"mtd": mtd, "ytd": ytd}
