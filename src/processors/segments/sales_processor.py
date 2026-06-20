# src/processors/segments/sales_processor.py
# Processor Sales Connectivity : PROD_A, PROD_B & PROD_C untuk semua segment (SEG_A, SEG_B, SEG_C, SEG_D).
#
# Logic:
#   - Sales Connectivity terdiri dari 3 sub-indikator: PROD_A, PROD_B, PROD_C
#   - Data aktual diambil dari 5 file source: srca_proda, srca_prodb, srcb_proda, srcb_prodb, srcb_prodc
#   - MtD diambil dari file source aktual, YtD dikalkulasi via akumulasi
#   - Target diambil dari SOURCE_TARGET sheet Sales Connectivity
#   - Historical diambil dari file historical sheet Sales Connectivity
#   - Tidak ada outlook — semua bulan ambil dari source aktual
#   - MOM dan YOY dihitung untuk semua sub-indikator
#
# Satu processor untuk semua segment dan sub-indikator — dipanggil dengan parameter segment dan indikator.


import pandas as pd
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[3])) # untuk import config
from config import (
    MONTH_COLS_MTD, MONTH_NUMBER_MAP,
    HISTORICAL_SALES_CONNECTIVITY, HISTORICAL_SALES_CONNECTIVITY_HEADER_ROW,
    TARGET_SALES_CONNECTIVITY, TARGET_SALES_CONNECTIVITY_HEADER_ROW,
)

# loader data sales connectivity ---------------------------------------------
from ...loaders.sales_loader import (
    load_srca_proda, load_srca_prodb, load_srcb_proda,
    load_srcb_prodb, load_srcb_prodc
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _month_idx(month): return {v: k for k, v in MONTH_NUMBER_MAP.items()}[month.upper()]

def _safe_div(n, d):
    if n is None or d is None or d == 0: return None
    return n / d

def _safe_growth(current, previous):
    if current is None:  return None   # tampil - di Excel via _fmt_pct
    if not previous:     return 1.0    # bulan ini ada, lalu tidak ada → +100%
    return (current - previous) / previous

def _load_target(filepath, territory, segment, indicator):
    """Load target untuk segment tertentu."""
    df = pd.read_excel(filepath, sheet_name=TARGET_SALES_CONNECTIVITY,
                       header=TARGET_SALES_CONNECTIVITY_HEADER_ROW)
    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == indicator) &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]
    if row.empty:
        return pd.Series({m: 0.0 for m in MONTH_COLS_MTD})
    return row.iloc[0][MONTH_COLS_MTD].fillna(0).astype(float)


from ._historical_helper import (
    load_historical
)


def calculate_sales_connectivity(
    territory       : str,
    period          : int,
    year            : int,
    report_month    : str,
    valid_until     : str,
    segment         : str,
    path_target     : str,
    path_historical : str,
) -> dict:

    # Load semua source data — dibaca sekali di luar loop
    df_srca_proda = load_srca_proda(territory)
    df_srca_prodb = load_srca_prodb(territory)
    df_srcb_proda = load_srcb_proda(territory)
    df_srcb_prodb = load_srcb_prodb(territory)
    df_srcb_prodc = load_srcb_prodc(territory)

    # Load semua source data
    SOURCE_MAP = {
        ("SEG_A", "PROD_A"): df_srca_proda['proda_seg_a'],
        ("SEG_A", "PROD_C"): df_srca_proda['prodc_seg_a'],
        ("SEG_A", "PROD_B"): df_srca_prodb['prodb_seg_a'],
        ("SEG_B", "PROD_A"): df_srcb_proda['proda_seg_b'],
        ("SEG_B", "PROD_B"): df_srcb_prodb['prodb_seg_b'],
        ("SEG_B", "PROD_C"): df_srcb_prodc['prodc_seg_b'],
        ("SEG_C", "PROD_A"): df_srcb_proda['proda_seg_c'],
        ("SEG_C", "PROD_B"): df_srcb_prodb['prodb_seg_c'],
        ("SEG_C", "PROD_C"): df_srcb_prodc['prodc_seg_c'],
        ("SEG_D", "PROD_A"): df_srcb_proda['proda_seg_d'],
        ("SEG_D", "PROD_B"): df_srcb_prodb['prodb_seg_d'],
        ("SEG_D", "PROD_C"): df_srcb_prodc['prodc_seg_d'],
        }

    report_month = report_month.upper()
    report_idx   = _month_idx(report_month)

    results = {}

    INDICATORS = ["PROD_A", "PROD_B", "PROD_C"]
    for indicator in INDICATORS:
        s_tgt  = _load_target(path_target, territory, segment, indicator)
        hist   = load_historical(path_historical, territory, segment, indicator)
        df_source = SOURCE_MAP.get((segment, indicator))
        if df_source is None:
            continue


        # ==========================================================================
        # Hitung MtD, YtD, target YtD, lalu metrik per bulan
        # ==========================================================================

        # hitung realisasi mtd ----------------------------------------------------------
        mtd_values = {}
        for idx, month in MONTH_NUMBER_MAP.items() :
            if idx >= report_idx : break
            period_val = year * 100 + idx
            mtd_val = df_source[df_source['period'] == period_val]['value'].sum()
            mtd_values[month] = float(mtd_val) if mtd_val != 0 else None

        # hitung realisasi ytd ----------------------------------------------------------
        # ── YtD ──────────────────────────────────────────────────────────────────
        ytd_values = {}; running = 0.0
        for idx, month in MONTH_NUMBER_MAP.items():
            if idx >= report_idx: break
            val = mtd_values.get(month)
            if val is not None: running += val
            ytd_values[month] = running  # ← selalu carry forward

        # ── Target YtD ───────────────────────────────────────────────────────────
        tgt_ytd = {}; running_t = 0.0
        for idx, month in MONTH_NUMBER_MAP.items():
            if idx >= report_idx: break
            t_val = s_tgt.get(month)
            if t_val is not None: running_t += t_val
            tgt_ytd[month] = running_t  # ← selalu carry forward

        # ── Metrik per bulan ─────────────────────────────────────────────────────
        monthly = []; prev_mtd = None
        for idx, month in MONTH_NUMBER_MAP.items():
            if idx >= report_idx: break
            mtd = mtd_values.get(month)
            ytd = ytd_values.get(month)
            tgt  = float(s_tgt.get(month, 0) or 0)
            tytd = tgt_ytd.get(month)

            ach_mtd = _safe_div(mtd, tgt)
            ach_ytd = _safe_div(ytd, tytd)

            mom = _safe_growth(mtd, prev_mtd)

            h_mtd = hist["mtd"].get(month)  # biarkan None tetap None
            h_ytd = hist["ytd"].get(month)  # biarkan None tetap None

            yoy_mtd = _safe_growth(mtd, h_mtd)
            yoy_ytd = _safe_growth(ytd, h_ytd)


            monthly.append({
                "month"      : month,
                "mtd"        : mtd,
                "ytd"        : ytd,
                "target_mtd" : tgt,
                "target_ytd" : tytd,
                "ach_mtd"    : ach_mtd,
                "ach_ytd"    : ach_ytd,
                "mom"        : mom,
                "yoy_mtd"    : yoy_mtd,
                "yoy_ytd"    : yoy_ytd,
            })

            prev_mtd = mtd
        results[indicator] = {"monthly": monthly}

    return {
        "territory"    : territory,
        "segment"      : segment,
        "year"         : year,
        "report_month" : report_month,
        "PROD_A"       : results.get("PROD_A", {"monthly": []}),
        "PROD_B"       : results.get("PROD_B",  {"monthly": []}),
        "PROD_C"       : results.get("PROD_C", {"monthly": []})
        }
