# src/processors/segments/collection_processor.py
# Processor Collection Performance : COLL_A, COLL_B untuk semua segment (SEG_A, SEG_B, SEG_C, SEG_D, TOTAL)
#                                     COLL_C, COLL_D untuk segment SEG_A only.
#
# Logic:
#   - Collection terdiri dari 4 indikator: COLL_A, COLL_B, COLL_C, COLL_D
#   - Data realisasi diambil dari SOURCE_COLLECTION sheet collection
#   - COLL_A & COLL_B mencakup semua segment (SEG_A, SEG_B, SEG_C, SEG_D, TOTAL) — point-in-time per bulan
#   - COLL_C & COLL_D hanya segment SEG_A — point-in-time per bulan
#   - Target diambil dari SOURCE_TARGET sheet Collection
#   - Historical diambil dari SOURCE_HISTORICAL sheet Collection
#   - Tidak ada outlook, tidak ada YtD carry-forward
#   - COLL_A & COLL_B : Achievement (ach_ytd) + YoY (yoy_ytd) — mtd, mom, yoy_mtd = None
#   - COLL_C & COLL_D : Achievement (ach_ytd) + MoM — mtd, yoy_mtd, yoy_ytd = None
#
# Satu processor untuk semua indikator — dipanggil dengan parameter indikator dan segment.

import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[3])) # untuk import config
from config import (
    MONTH_COLS_MTD, MONTH_NUMBER_MAP,
    HISTORICAL_COLLECTION, HISTORICAL_COLLECTION_HEADER_ROW,
    TARGET_COLLECTION, TARGET_COLLECTION_HEADER_ROW,
    COLLECTION_SEGMENT_MAPPING,
)

# import loader data collection performance
from ...loaders.collection_loader import (
    _load_target, _load_historical,
    load_collection_data
)

# ── Helpers ───────────────────────────────────────────────────────────────────
# membuat key dan value setara pada dictionary
def month_idx (month):
    return {v:k for k,v in MONTH_NUMBER_MAP.items()}[month.upper()]

def _safe_div(n, d): # untuk div target
    if n is None or d is None or d == 0: return None
    return n / d

def _safe_growth(current, previous): # hitung MoM dan YoY
    if current is None:  return None   # tampil - di Excel via _fmt_pct
    if not previous:     return 1.0    # bulan ini ada, tahun/bulan lalu tidak ada → +100%
    return (current - previous) / previous


def _aggregate_colla_collb(bs_monthly: list, gs_monthly: list) -> list:
    """Average SEG_A dan SEG_B values untuk SEG_A+SEG_B aggregation."""
    aggregated = []
    for i, month_bs in enumerate(bs_monthly):
        month_gs = gs_monthly[i]
        month = month_bs["month"]

        # Average ytd, target_ytd, ach_ytd, yoy_ytd
        ytd = None
        if month_bs["ytd"] is not None and month_gs["ytd"] is not None:
            ytd = (month_bs["ytd"] + month_gs["ytd"]) / 2

        tytd = None
        if month_bs["target_ytd"] is not None and month_gs["target_ytd"] is not None:
            tytd = (month_bs["target_ytd"] + month_gs["target_ytd"]) / 2

        ach_ytd = _safe_div(ytd, tytd)

        yoy_ytd = None
        if month_bs["yoy_ytd"] is not None and month_gs["yoy_ytd"] is not None:
            yoy_ytd = (month_bs["yoy_ytd"] + month_gs["yoy_ytd"]) / 2

        aggregated.append({
            'month'      : month,
            'mtd'        : None,
            'ytd'        : ytd,
            'target_mtd' : None,
            'target_ytd' : tytd,
            'ach_mtd'    : None,
            'ach_ytd'    : ach_ytd,
            'mom'        : None,
            'yoy_mtd'    : None,
            'yoy_ytd'    : yoy_ytd,
        })

    return aggregated


def calculate_colla_collb (
    territory:str,
    report_month,
    period,
    filepath,
    segment,
    indicator,
    valid_until
) -> list:

    # load data
    df_collection = load_collection_data(territory, period, report_month)

    # load target
    target = _load_target(filepath, territory, segment, indicator)

    # load historical
    historical = _load_historical(filepath, territory, segment, indicator)

    # pilih DataFrame sesuai indicator
    realisasi_df = df_collection['coll_a'] if indicator == 'COLL_A' else df_collection['coll_b']

    # filter row
    row = realisasi_df[realisasi_df['SEGMENT'] == segment]

    bulan = int(str(report_month)[4:6])
    active_month = MONTH_COLS_MTD[:(bulan)]
    monthly = []

    for month in active_month:
        # 1. ambil realisasi
        ytd = None
        if not row.empty:
            try:
                ytd = row.iloc[0][month]
                ytd = None if pd.isna(ytd) else float(ytd)
            except (KeyError, IndexError):
                ytd = None

        # 2. ambil target & historical bulan ini
        tytd = target[month]
        hist = historical[month]

        # 3. hitung achievement & yoy
        ach_ytd = _safe_div(ytd, tytd)
        yoy_ytd = _safe_growth(ytd, hist)

        monthly.append({
            'month'      : month,
            'mtd'        : None,
            'ytd'        : ytd,
            'target_mtd' : None,
            'target_ytd' : tytd,
            'ach_mtd'    : None,
            'ach_ytd'    : ach_ytd,
            'mom'        : None,
            'yoy_mtd'    : None,
            'yoy_ytd'    : yoy_ytd,
        })

    return monthly


def calculate_collc_colld(
    territory: str,
    period: int,
    report_month,
    filepath,
    segment,
    indicator,
) -> list:

    # load data
    df_collection = load_collection_data(territory, period, report_month)

    # load target
    target = _load_target(filepath, territory, segment, indicator)

    # load historical
    historical = _load_historical(filepath, territory, segment, indicator)

    # pilih DataFrame sesuai indicator
    realisasi_df = df_collection['coll_d'] if indicator == 'COLL_D' else df_collection['coll_c']

    # filter row segment SEG_A
    row = realisasi_df[realisasi_df['SEGMENT'] == "SEG_A"]

    bulan = int(str(report_month)[4:6])
    active_month = MONTH_COLS_MTD[:bulan]
    monthly = []
    for i, month in enumerate(active_month):
        # 1. ambil realisasi
        try:
            if row.empty:
                mtd = None
            else:
                mtd = row.iloc[0][month]
                mtd = None if pd.isna(mtd) else float(mtd)
        except(KeyError, IndexError):
            mtd = None

        # 2. ambil target bulan ini
        tmtd = target[month]

        # ambil data bulan sebelumnya
        prev = None
        if i > 0 and not row.empty:
            try:
                prev = row.iloc[0][active_month[i-1]]
                prev = None if pd.isna(prev) else float(prev)
            except (KeyError, IndexError):
                pass

        # 3. hitung achievement & MoM
        ach_mtd = _safe_div(mtd, tmtd)


        mom_mtd = _safe_growth(mtd, prev)

        monthly.append({
            'month'      : month,
            'mtd'        : mtd,
            'ytd'        : None,
            'target_mtd' : tmtd,
            'target_ytd' : None,
            'ach_mtd'    : ach_mtd,
            'ach_ytd'    : None,
            'mom'        : mom_mtd,
            'yoy_mtd'    : None,
            'yoy_ytd'    : None,
        })

    return monthly

def build_collection_results(
    territory: str,
    period: int,
    report_month: int,
    filepath: str,
) -> dict:

    result = {
        'COLL_A':  {},
        'COLL_B':  {},
        'COLL_C':  {},
        'COLL_D':  {},
    }

    # COLL_A & COLL_B → semua segment (termasuk SEG_A+SEG_B aggregate)
    for indicator in ["COLL_A", "COLL_B"]:
        for segment in COLLECTION_SEGMENT_MAPPING:
            if segment == "SEG_A+SEG_B":
                # Aggregate SEG_A dan SEG_B
                bs_monthly = calculate_colla_collb(
                    territory=territory,
                    period=period,
                    report_month=report_month,
                    filepath=filepath,
                    segment="SEG_A",
                    indicator=indicator,
                    valid_until=None,
                )
                gs_monthly = calculate_colla_collb(
                    territory=territory,
                    period=period,
                    report_month=report_month,
                    filepath=filepath,
                    segment="SEG_B",
                    indicator=indicator,
                    valid_until=None,
                )
                result[indicator][segment] = _aggregate_colla_collb(bs_monthly, gs_monthly)
            else:
                result[indicator][segment] = calculate_colla_collb(
                    territory=territory,
                    period=period,
                    report_month=report_month,
                    filepath=filepath,
                    segment=segment,
                    indicator=indicator,
                    valid_until=None,
                )

    # COLL_C & COLL_D → SEG_A only (tidak ada SEG_A+SEG_B aggregation)
    for indicator in ["COLL_C", "COLL_D"]:
        result[indicator]["SEG_A"] = calculate_collc_colld(
            territory=territory,
            period=period,
            report_month=report_month,
            filepath=filepath,
            segment="SEG_A",
            indicator=indicator,
        )

    return result
