# src/processors/aggregator.py
# Menggabungkan hasil 4 segment (SEG_A, SEG_B, SEG_C, SEG_D) menjadi SEG_A+SEG_B dan TOTAL.
#
# Logic:
#   MtD SEG_A+SEG_B = MtD SEG_A + MtD SEG_B
#   MtD TOTAL       = MtD SEG_A + MtD SEG_B + MtD SEG_C + MtD SEG_D
#   YtD             = akumulasi MtD dari JAN
#   Achievement = MtD atau YtD / target dari file (row SEG_A+SEG_B / TOTAL)
#   MoM, YoY   = formula sama seperti segment lain

import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2])) #base path
from config import (
    MONTH_COLS_MTD, MONTH_NUMBER_MAP,
    TARGET_SHEET, TARGET_HEADER_ROW,
    TARGET_SALES_CONNECTIVITY, TARGET_SALES_CONNECTIVITY_HEADER_ROW
)
from src.processors.segments._historical_helper import load_historical


def _safe_div(n, d):
    # Return None jika pembagi nol atau ada None — avoid ZeroDivisionError
    if n is None or d is None or d == 0:
        return None
    return n / d

# hitung growth rate untuk mtd dan ytd
def _safe_growth(current, previous):
    if current is None:  return None   # tampil - di Excel via _fmt_pct
    if not previous:     return 1.0    # bulan ini ada, lalu tidak ada → +100%
    return (current - previous) / previous


def _load_target(filepath: str, territory: str, segment: str, indicator: str="REVENUE") -> pd.Series:
    """Load target untuk segment SEG_A+SEG_B atau TOTAL."""
    # Pilih sheet berdasarkan indicator
    if indicator in ["PROD_A","PROD_C","PROD_B"] :
        sheet = TARGET_SALES_CONNECTIVITY #sheet target untuk sales connectivity
        header = TARGET_SALES_CONNECTIVITY_HEADER_ROW #row pertama data target sales connectivity
    elif indicator in ["REVENUE","GROWTH"]:
        sheet =  TARGET_SHEET # sheet target untuk revenue dan GROWTH
        header =  TARGET_HEADER_ROW #row pertama data target revenue dan GROWTH

    df = pd.read_excel(filepath, sheet_name=sheet,header=header)

    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == indicator) &
        (df["SEGMENT"]   == segment)
    )

    row = df[mask]
    if row.empty:
        return pd.Series({m: 0.0 for m in MONTH_COLS_MTD})
    return row.iloc[0][MONTH_COLS_MTD].fillna(0).astype(float)


def _build_aggregated(
    label         : str,
    segments      : list,
    results       : dict,
    target_file   : str,
    historical_file: str,
    territory     : str,
    report_month  : str,
    indicator      : str = "REVENUE",
    monthly_key = None,
) -> dict:
    """
    Build aggregated result untuk satu label (SEG_A+SEG_B atau TOTAL).

    Args:
        label    : "SEG_A+SEG_B" atau "TOTAL"
        segments : list segment yang dijumlah, e.g. ["SEG_A","SEG_B"] atau TOTAL = ["SEG_A","SEG_B","SEG_C","SEG_D"]
        results  : dict berisi result per segment {"SEG_A": r_bs, "SEG_B": r_gs, ...}
        indicator   : "REVENUE", "GROWTH", "PROD_A", "PROD_B", "PROD_C"
        monthly_key : None untuk Revenue/GROWTH, "PROD_A"/"PROD_B"/"PROD_C" untuk Sales Connectivity
    """
    report_month = report_month.upper()
    report_idx   = {v: k for k, v in MONTH_NUMBER_MAP.items()}[report_month]

    # Load target dan historical untuk label ini
    s_tgt = _load_target(target_file, territory, label, indicator)
    hist  = load_historical(historical_file, territory, label, indicator)

    # ── Hitung MtD per bulan = sum segment ───────────────────────────────────
    monthly_data = {
        seg: results[seg][monthly_key]["monthly"] if monthly_key else results[seg]["monthly"]
        for seg in segments}

    mtd_values = {}
    for idx, month in MONTH_NUMBER_MAP.items():
        if idx >= report_idx:
            break
        vals = [
            monthly_data[seg][idx - 1]["mtd"]
            for seg in segments
            if monthly_data[seg][idx - 1]["mtd"] is not None
            ]

        mtd_values[month] = sum(vals) if vals else None

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

    # ── Hitung semua metrik per bulan ─────────────────────────────────────────
    monthly  = []
    prev_mtd = None

    for idx, month in MONTH_NUMBER_MAP.items():
        if idx >= report_idx:
            break

        mtd  = mtd_values.get(month)
        ytd  = ytd_values.get(month)
        tgt  = float(s_tgt.get(month, 0) or 0)
        tytd = tgt_ytd.get(month)

        ach_mtd  = _safe_div(mtd, tgt)
        ach_ytd  = _safe_div(ytd, tytd)
        mom = _safe_growth(mtd, prev_mtd)

        h_mtd = hist["mtd"].get(month)  # biarkan None tetap None
        h_ytd = hist["ytd"].get(month)  # biarkan None tetap None

        yoy_mtd = _safe_growth(mtd, h_mtd)
        yoy_ytd = _safe_growth(ytd, h_ytd)

        monthly.append({
            "month"      : month,
            "is_outlook" : any(
                monthly_data[seg][idx - 1].get("is_outlook", False)
                for seg in segments
            ),
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

    return {
        "territory"   : territory,
        "segment"     : label,
        "report_month": report_month,
        "monthly"     : monthly,
    }

# ================= AGGREGATE REVENUE =====================
def aggregate_revenue(
    results_bs    : dict,
    results_gs    : dict,
    results_dss   : dict,
    results_dps   : dict,
    target_file   : str,
    historical_file: str,
    territory     : str,
    report_month  : str,
) -> dict:
    """
    Aggregate semua segment menjadi SEG_A+SEG_B dan TOTAL.

    Returns:
        {
          "bsgs" : dict result SEG_A+SEG_B,
          "total": dict result TOTAL,
        }
    """
    results = {
        "SEG_A": results_bs,
        "SEG_B": results_gs,
        "SEG_C": results_dss,
        "SEG_D": results_dps,
    }

    print(f"  [AGG] Computing SEG_A+SEG_B...")
    bsgs = _build_aggregated(
        label          = "SEG_A+SEG_B",
        segments       = ["SEG_A", "SEG_B"],
        results        = results,
        target_file    = target_file,
        historical_file= historical_file,
        territory      = territory,
        report_month   = report_month,
    )

    print(f"  [AGG] Computing TOTAL...")
    total = _build_aggregated(
        label          = "TOTAL",
        segments       = ["SEG_A", "SEG_B", "SEG_C", "SEG_D"],
        results        = results,
        target_file    = target_file,
        historical_file= historical_file,
        territory      = territory,
        report_month   = report_month,
    )

    print(f"  [AGG] ✓ Aggregation complete")
    return {"bsgs": bsgs, "total": total}


def print_aggregated_result(result: dict) -> None:
    seg     = result["segment"]
    rep     = result["report_month"]
    terr    = result["territory"]
    monthly = result["monthly"]

    print(f"\n{'='*72}")
    print(f"  REVENUE {seg} — {terr} | Laporan: {rep}")
    print(f"{'='*72}")

    print(f"\n  {'Bln':<5} {'MtD':>16} {'Target MtD':>16} {'Ach MtD':>8} "
          f"{'MoM':>8} {'YoY MtD':>8} {'YtD':>16} {'Ach YtD':>8} {'YoY YtD':>8}  Ket")
    print(f"  {'─'*106}")

    for r in monthly:
        m       = r["month"]
        mtd     = f"{r['mtd']:>16,.0f}"        if r["mtd"]     is not None else f"{'—':>16}"
        tgt     = f"{r['target_mtd']:>16,.0f}"
        ach_mtd = f"{r['ach_mtd']*100:>7.1f}%" if r["ach_mtd"] is not None else f"{'—':>8}"
        mom     = f"{r['mom']*100:>+7.1f}%"    if r["mom"]     is not None else f"{'—':>8}"
        yoy_mtd = f"{r['yoy_mtd']*100:>+7.1f}%" if r["yoy_mtd"] is not None else f"{'—':>8}"
        ytd     = f"{r['ytd']:>16,.0f}"        if r["ytd"]     is not None else f"{'—':>16}"
        ach_ytd = f"{r['ach_ytd']*100:>7.1f}%" if r["ach_ytd"] is not None else f"{'—':>8}"
        yoy_ytd = f"{r['yoy_ytd']*100:>+7.1f}%" if r["yoy_ytd"] is not None else f"{'—':>8}"
        ket     = " ← OUTLOOK" if r["is_outlook"] else ""

        print(f"  {m:<5} {mtd} {tgt} {ach_mtd} {mom} {yoy_mtd} "
              f"{ytd} {ach_ytd} {yoy_ytd}{ket}")

    print(f"\n{'='*72}\n")

# ================= AGGREGATE GROWTH =====================
# ── GROWTH Aggregator ──────────────────────────────────────────────────────────

def aggregate_ngtma(
    results_ngtma_bs  : dict,
    results_ngtma_gs  : dict,
    results_ngtma_dss : dict,
    results_ngtma_dps : dict,
    target_file       : str,
    historical_file   : str,
    territory         : str,
    report_month      : str,
) -> dict:

    ngtma_results = {
        "SEG_A": results_ngtma_bs,
        "SEG_B": results_ngtma_gs,
        "SEG_C": results_ngtma_dss,
        "SEG_D": results_ngtma_dps,
    }

    print(f"  [AGG-GROWTH] Computing SEG_A+SEG_B...")
    ngtma_bsgs = _build_aggregated(
        label           = "SEG_A+SEG_B",
        segments        = ["SEG_A", "SEG_B"],
        results         = ngtma_results,
        target_file     = target_file,
        historical_file = historical_file,
        territory       = territory,
        report_month    = report_month,
        indicator       = "GROWTH",
    )

    print(f"  [AGG-GROWTH] Computing TOTAL...")
    ngtma_total = _build_aggregated(
        label           = "TOTAL",
        segments        = ["SEG_A", "SEG_B", "SEG_C", "SEG_D"],
        results         = ngtma_results,
        target_file     = target_file,
        historical_file = historical_file,
        territory       = territory,
        report_month    = report_month,
        indicator       = "GROWTH",
    )

    print(f"  [AGG-GROWTH] ✓ Aggregation complete")
    return {"bsgs": ngtma_bsgs, "total": ngtma_total}


# ================= AGGREGATE SALES CONNECTIVITY =====================
def aggregate_sales_connectivity(
    results_bs  : dict,
    results_gs  : dict,
    results_dss : dict,
    results_dps : dict,
    target_file    : str,
    historical_file: str,
    territory      : str,
    report_month   : str,
) -> dict:
    """
    Aggregate Sales Connectivity (PROD_A, PROD_B, PROD_C) menjadi SEG_A+SEG_B dan TOTAL.
    Returns: {"PROD_A": {"bsgs": ..., "total": ...}, "PROD_B": {...}, "PROD_C": {...}}
    """
    results = {
        "SEG_A": results_bs,
        "SEG_B": results_gs,
        "SEG_C": results_dss,
        "SEG_D": results_dps,
    }

    output = {}
    for indicator in ["PROD_A", "PROD_B", "PROD_C"]:
        print(f"  [AGG-SC] Computing {indicator} SEG_A+SEG_B...")
        bsgs = _build_aggregated(
            label           = "SEG_A+SEG_B",
            segments        = ["SEG_A", "SEG_B"],
            results         = results,
            target_file     = target_file,
            historical_file = historical_file,
            territory       = territory,
            report_month    = report_month,
            indicator       = indicator,
            monthly_key     = indicator,
        )
        print(f"  [AGG-SC] Computing {indicator} TOTAL...")
        total = _build_aggregated(
            label           = "TOTAL",
            segments        = ["SEG_A", "SEG_B", "SEG_C", "SEG_D"],
            results         = results,
            target_file     = target_file,
            historical_file = historical_file,
            territory       = territory,
            report_month    = report_month,
            indicator       = indicator,
            monthly_key     = indicator,
        )
        output[indicator] = {"bsgs": bsgs, "total": total}

    print(f"  [AGG-SC] ✓ Aggregation complete")
    return output
