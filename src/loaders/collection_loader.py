# src/loaders/collection_loader.py
# Load Data Collection untuk indicator : COLL_D - COLL_C SEG_A, COLL_A-COLL_B SEG_A,SEG_B, SEG_C, SEG_D & ALL
# Logic:
#   - Load data dilakukan pada worksheet excel : collection performance sheet collection
#   - Data yang diload terdiri dari 4 indicator : COLL_D, COLL_C, COLL_A dan COLL_B
#   - Data COLL_D dan COLL_C selalu berupa data Monthly, Sedangkan data COLL_A dan COLL_B berupa data Yearly


# ========================================================
# =========== LOADER REALISASI COLLECTION DATA ===========
# ========================================================

# packages yang dipakai

import pandas as pd
from pathlib import Path
import sys

from config import (
    SOURCE_COLLECTION, COLLECTION_SEGMENT_MAPPING,
    MONTH_COLS_MTD, ACTIVE_COLS_COLLECTION,
    COLLECTION_DATA_SHEET,
    TARGET_COLLECTION,TARGET_COLLECTION_HEADER_ROW,
    HISTORICAL_COLLECTION,HISTORICAL_COLLECTION_HEADER_ROW
)

def load_collection_data(territory:str, period: int, report_month:int) -> dict[str, pd.DataFrame]:
    # load data collection
    path = f"data/input/{SOURCE_COLLECTION}"
    raw_df = pd.read_excel(path, sheet_name=COLLECTION_DATA_SHEET, header=0)

    # config df_raw
    bulan = int(str(report_month)[4:6]) #ex : 202606 -> slice to 06 -> 6
    active_month = MONTH_COLS_MTD[:(bulan)]


    # TARIK DATA COLL_D SEG_A -------------------------------
    mask_colld_sega = (
        (raw_df['INDICATOR'] == 'COLL_D') &
        (raw_df['SEGMENT'] == 'SEG_A') &
        (raw_df['TERRITORY_NAME'] == territory) &
        (raw_df['YEAR'] == int(str(period)[:4]))
    )
    df_colld_sega = raw_df[mask_colld_sega][ACTIVE_COLS_COLLECTION+active_month].set_index('INDICATOR').copy()

    # TARIK DATA COLL_C SEG_A -------------------------------

    mask_collc_sega = (
        (raw_df['INDICATOR'] == 'COLL_C') &
        (raw_df['SEGMENT'] == 'SEG_A') &
        (raw_df['TERRITORY_NAME'] == territory) &
        (raw_df['YEAR'] == int(str(period)[:4]))
    )
    df_collc_sega = raw_df[mask_collc_sega][ACTIVE_COLS_COLLECTION+active_month].set_index('INDICATOR').copy()

    # TARIK DATA COLL_A : SEG_A, SEG_B, SEG_C, SEG_D & ALL -------------------------------
    df_colla = pd.DataFrame()
    for segment in COLLECTION_SEGMENT_MAPPING.values():
        mask_colla = (
            (raw_df['INDICATOR'] == 'COLL_A')&
            (raw_df['SEGMENT'] == segment)&
            (raw_df['TERRITORY_NAME'] == territory)&
            (raw_df['YEAR'] == int(str(period)[:4]))
            )
        data_new = raw_df[mask_colla][ACTIVE_COLS_COLLECTION+active_month].set_index('INDICATOR').copy()
        df_colla = pd.concat([df_colla,data_new])
        df_colla['YEAR'] = df_colla['YEAR'].astype('Int64')  # nullable integer


    # TARIK DATA COLL_B : SEG_A, SEG_B, SEG_C, SEG_D & ALL -------------------------------
    df_collb = pd.DataFrame()
    for segment in COLLECTION_SEGMENT_MAPPING.values():
        mask_collb = (
            (raw_df['INDICATOR'] == 'COLL_B')&
            (raw_df['SEGMENT'] == segment)&
            (raw_df['TERRITORY_NAME'] == territory)&
            (raw_df['YEAR'] == int(str(period)[:4]))
            )
        data_new = raw_df[mask_collb][ACTIVE_COLS_COLLECTION+active_month].set_index('INDICATOR').copy()
        df_collb = pd.concat([df_collb,data_new])
        df_collb['YEAR'] = df_collb['YEAR'].astype('Int64')  # nullable integer


    return {'coll_d' : df_colld_sega,
            'coll_c' : df_collc_sega,
            'coll_a' : df_colla,
            'coll_b' : df_collb
            }


# ========================================================
# ============ LOADER TARGET COLLECTION DATA =============
# ========================================================
def _load_target(filepath, territory, segment, indicator):

    df = pd.read_excel(filepath, sheet_name=TARGET_COLLECTION,
                       header=TARGET_COLLECTION_HEADER_ROW)
    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == indicator) &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]
    if row.empty:
        return pd.Series({m: 0.0 for m in MONTH_COLS_MTD})

    return row.iloc[0][MONTH_COLS_MTD].fillna(0).astype(float)

def _load_historical (filepath, territory, segment, indicator) :

    df = pd.read_excel(filepath, sheet_name=HISTORICAL_COLLECTION,
                       header=HISTORICAL_COLLECTION_HEADER_ROW)

    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == indicator) &
        (df["SEGMENT"]   == segment)
    )

    row = df[mask]
    if row.empty :
        return pd.Series({m:0.0 for m in MONTH_COLS_MTD})

    return row.iloc[0][MONTH_COLS_MTD].fillna(0).astype(float)
