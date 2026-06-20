# src/processors/segments/ngtma_processor.py
# Processor GROWTH untuk semua segment (SEG_A, SEG_B, SEG_C, SEG_D).
#
# Logic:
#   - GROWTH tidak punya formula komponen
#   - Valid maupun outlook → ambil nilai TYPE='GROWTH' dari Excel langsung
#   - Cut off per segment independent (sama seperti Revenue)
#
# Satu processor untuk semua segment — dipanggil dengan parameter segment.

import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from config import (
    MONTH_COLS_MTD, MONTH_NUMBER_MAP,
    REVENUE_SHEET, REVENUE_HEADER_ROW,
    TARGET_SHEET, TARGET_HEADER_ROW,
    HISTORICAL_SHEET,
)
from ._historical_helper import load_historical

# Mapping segment display name → internal code dan target label
SEGMENT_CONFIG = {
    "SEG_A": {"code": "CODE_A", "hist_label": "SEG_A"},
    "SEG_B": {"code": "CODE_B", "hist_label": "SEG_B"},
    "SEG_C": {"code": "CODE_C", "hist_label": "SEG_C"},
    "SEG_D": {"code": "CODE_D", "hist_label": "SEG_D"},
}

TYPE_NGTMA = "GROWTH"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _month_idx(month): return {v: k for k, v in MONTH_NUMBER_MAP.items()}[month.upper()]

def _safe_div(n, d):
    if n is None or d is None or d == 0: return None
    return n / d

# hitung growth rate untuk mtd dan ytd
def _safe_growth(current, previous):
    if current is None:  return None   # tampil - di Excel via _fmt_pct
    if not previous:     return 1.0    # bulan ini ada, lalu tidak ada → +100%
    return (current - previous) / previous

def _get(df, type_val, month):
    if type_val not in df.index: return 0.0
    val = df.loc[type_val, month.upper()]
    if pd.isna(val): return 0.0
    try: return float(str(val))
    except: return 0.0


# ── Loaders ───────────────────────────────────────────────────────────────────

def _load_ngtma_source(filepath, territory, segment, year):
    """Load GROWTH data untuk segment tertentu. Return DataFrame index=TYPE."""
    df = pd.read_excel(filepath, sheet_name=REVENUE_SHEET,
                       header=REVENUE_HEADER_ROW)
    code = SEGMENT_CONFIG[segment]["code"]
    mask = (
        (df["YEAR"]      == year) &
        (df["INDICATOR"] == "GROWTH") &
        (df["SEGMENT"]   == code) &
        (df["TERRITORY"] == territory)
    )
    return df[mask].copy().set_index("TYPE")


def _load_target(filepath, territory, segment):
    """Load target GROWTH untuk segment tertentu."""
    df = pd.read_excel(filepath, sheet_name=TARGET_SHEET,
                       header=TARGET_HEADER_ROW)
    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == "GROWTH") &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]
    if row.empty:
        return pd.Series({m: 0.0 for m in MONTH_COLS_MTD})
    return row.iloc[0][MONTH_COLS_MTD].fillna(0).astype(float)


def _load_historical(filepath, territory, segment):
    """Load historical GROWTH untuk segment tertentu."""
    from config import HIST_GROWTH
    df = pd.read_excel(filepath, sheet_name=HISTORICAL_SHEET,
                       header=1)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(subset=["INDICATOR", "SEGMENT"]).copy()

    mask = (
        (df["TERRITORY"] == territory) &
        (df["INDICATOR"] == HIST_GROWTH) &
        (df["SEGMENT"]   == segment)
    )
    row = df[mask]

    if row.empty:
        empty = {m: None for m in MONTH_COLS_MTD}
        return {"mtd": empty, "ytd": empty}

    mtd = {}
    for month in MONTH_COLS_MTD:
        if month not in df.columns:
            mtd[month] = None
            continue
        val = row.iloc[0][month]
        mtd[month] = None if pd.isna(val) else float(val)

    ytd = {}; running = 0.0
    for idx, month in MONTH_NUMBER_MAP.items():
        val = mtd.get(month)
        if val is not None:
            running += val; ytd[month] = running
        else:
            ytd[month] = None

    return {"mtd": mtd, "ytd": ytd}


# ── Core Calculator ───────────────────────────────────────────────────────────

def calculate_ngtma(
    revenue_file   : str,
    target_file    : str,
    historical_file: str,
    territory      : str,
    year           : int,
    report_month   : str,
    segment        : str,
    valid_until    : str,
) -> dict:
    """
    Hitung GROWTH untuk satu segment.

    Args:
        segment     : "SEG_A", "SEG_B", "SEG_C", atau "SEG_D"
        valid_until : bulan terakhir data valid, e.g. "APR"
                      Bulan setelah valid_until → ambil dari Excel juga
                      (tidak ada formula manual untuk GROWTH)
    """
    report_month = report_month.upper()
    valid_until  = valid_until.upper()
    report_idx   = _month_idx(report_month)

    print(f"  [GROWTH-{segment}] Loading source...")
    df_ngtma = _load_ngtma_source(revenue_file, territory, segment, year)
    print(f"  [GROWTH-{segment}] Loading target...")
    s_tgt    = _load_target(target_file, territory, segment)
    print(f"  [GROWTH-{segment}] Loading historical...")
    hist     = load_historical(historical_file, territory, segment,"GROWTH")
    print(f"  [GROWTH-{segment}] ✓ All data loaded")

    # ── Hitung MtD per bulan ──────────────────────────────────────────────────
    # GROWTH: semua bulan ambil dari Excel — tidak ada formula manual
    mtd_values     = {}
    outlook_months = set()

    for idx, month in MONTH_NUMBER_MAP.items():
        if idx >= report_idx: break
        val = _get(df_ngtma, TYPE_NGTMA, month)
        mtd_values[month] = val if val != 0.0 else None

        # Flag bulan outlook untuk display
        if idx > _month_idx(valid_until):
            outlook_months.add(month)

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
        mtd  = mtd_values.get(month)
        ytd  = ytd_values.get(month)
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
            "is_outlook" : month in outlook_months,
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
        "territory"    : territory,
        "segment"      : segment,
        "kpi"          : "GROWTH",
        "year"         : year,
        "report_month" : report_month,
        "valid_until"  : valid_until,
        "outlook_months": list(outlook_months),
        "monthly"      : monthly,
    }


# ── Pretty Printer ────────────────────────────────────────────────────────────

def print_ngtma_result(result: dict) -> None:
    seg     = result["segment"]
    rep     = result["report_month"]
    valid   = result["valid_until"]
    terr    = result["territory"]
    outl    = ", ".join(result.get("outlook_months", [])) or "Tidak perlu"
    monthly = result["monthly"]

    print(f"\n{'='*72}")
    print(f"  GROWTH {seg} — {terr} | Laporan: {rep} | "
          f"Valid s/d: {valid} | Outlook: {outl}")
    print(f"{'='*72}")
    print(f"\n  {'Bln':<5} {'MtD':>16} {'Target MtD':>16} {'Ach MtD':>8} "
          f"{'MoM':>8} {'YoY MtD':>8} {'YtD':>16} {'Ach YtD':>8} {'YoY YtD':>8}  Ket")
    print(f"  {'─'*106}")

    for r in monthly:
        m       = r["month"]
        mtd     = f"{r['mtd']:>16,.0f}"         if r["mtd"]     is not None else f"{'—':>16}"
        tgt     = f"{r['target_mtd']:>16,.0f}"
        ach_mtd = f"{r['ach_mtd']*100:>7.1f}%"  if r["ach_mtd"] is not None else f"{'—':>8}"
        mom     = f"{r['mom']*100:>+7.1f}%"     if r["mom"]     is not None else f"{'—':>8}"
        yoy_mtd = f"{r['yoy_mtd']*100:>+7.1f}%" if r["yoy_mtd"] is not None else f"{'—':>8}"
        ytd     = f"{r['ytd']:>16,.0f}"         if r["ytd"]     is not None else f"{'—':>16}"
        ach_ytd = f"{r['ach_ytd']*100:>7.1f}%"  if r["ach_ytd"] is not None else f"{'—':>8}"
        yoy_ytd = f"{r['yoy_ytd']*100:>+7.1f}%" if r["yoy_ytd"] is not None else f"{'—':>8}"
        ket     = " ← dari Excel" if r["is_outlook"] else ""

        print(f"  {m:<5} {mtd} {tgt} {ach_mtd} {mom} {yoy_mtd} "
              f"{ytd} {ach_ytd} {yoy_ytd}{ket}")

    print(f"\n{'='*72}\n")
