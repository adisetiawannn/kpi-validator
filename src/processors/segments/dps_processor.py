# src/processors/segments/dps_processor.py
# Processor Revenue SEG_D (Segment CODE_D).

import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from config import (
    MONTH_COLS_MTD, MONTH_NUMBER_MAP,
    REVENUE_SHEET, REVENUE_HEADER_ROW,
    TARGET_SHEET, TARGET_HEADER_ROW,
    HIST_REVENUE,
)
from ._historical_helper import load_historical_revenue, load_historical

TYPE_ALL      = "ALL"
TYPE_SUSTAIN  = "TYPE_1"
TYPE_KOREKSI  = "TYPE_7"
TYPE_JURBAL   = "TYPE_8"
TYPE_SCAL_OTC = "TYPE_4"

def _month_idx(month): return {v:k for k,v in MONTH_NUMBER_MAP.items()}[month.upper()]
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

def _load_revenue_source(filepath, territory, year):
    df = pd.read_excel(filepath, sheet_name=REVENUE_SHEET, header=REVENUE_HEADER_ROW)
    mask = ((df["YEAR"]==year)&(df["INDICATOR"]=="REVENUE")&
            (df["SEGMENT"]=="CODE_D")&(df["TERRITORY"]==territory))
    return df[mask].copy().set_index("TYPE")

def _load_target(filepath, territory):
    df = pd.read_excel(filepath, sheet_name=TARGET_SHEET, header=TARGET_HEADER_ROW)
    mask = ((df["TERRITORY"]==territory)&(df["INDICATOR"]=="REVENUE")&(df["SEGMENT"]=="SEG_D"))
    row = df[mask]
    if row.empty: return pd.Series({m:0.0 for m in MONTH_COLS_MTD})
    return row.iloc[0][MONTH_COLS_MTD].fillna(0).astype(float)

def _load_historical(filepath, territory):
    return load_historical_revenue(filepath, territory, "SEG_D")

def calculate_dps(
    revenue_file, target_file, historical_file,
    territory, year, report_month,
    valid_until, scaling_map: dict,
) -> dict:
    report_month = report_month.upper()
    valid_until  = valid_until.upper()
    report_idx   = _month_idx(report_month)
    valid_idx    = _month_idx(valid_until)

    print(f"  [SEG_D] Loading revenue source...")
    df_rev = _load_revenue_source(revenue_file, territory, year)
    print(f"  [SEG_D] Loading target...")
    s_tgt  = _load_target(target_file, territory)
    print(f"  [SEG_D] Loading historical...")
    hist   = load_historical(historical_file, territory,"SEG_D","REVENUE") # update to load_historical(...)
    print(f"  [SEG_D] ✓ All data loaded")

    mtd_values = {}; outlook_detail = {}
    for idx, month in MONTH_NUMBER_MAP.items():
        if idx >= report_idx: break
        if idx <= valid_idx:
            mtd_values[month] = _get(df_rev, TYPE_ALL, month)
        else:
            scaling = float(scaling_map.get(month, 0) or 0)
            sustain      = _get(df_rev, TYPE_SUSTAIN, month)
            corr         = _get(df_rev, TYPE_KOREKSI, month)
            jrn_bal      = _get(df_rev, TYPE_JURBAL, month)
            scal_otc     = _get(df_rev, TYPE_SCAL_OTC, month)
            total   = sustain + corr + jrn_bal + scaling + scal_otc
            mtd_values[month] = total
            outlook_detail[month] = {
                TYPE_SUSTAIN           : sustain,
                TYPE_KOREKSI           : corr,
                TYPE_JURBAL            : jrn_bal,
                "SCALING (manual)"     : scaling,
                TYPE_SCAL_OTC          : scal_otc,
                "TOTAL OUTLOOK": total,
            }

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

    monthly = []; prev_mtd = None
    outlook_months = set(outlook_detail.keys())
    for idx, month in MONTH_NUMBER_MAP.items():
        if idx >= report_idx: break
        mtd  = mtd_values.get(month); ytd  = ytd_values.get(month)
        tgt  = float(s_tgt.get(month, 0) or 0); tytd = tgt_ytd.get(month)
        ach_mtd = _safe_div(mtd, tgt); ach_ytd = _safe_div(ytd, tytd)

        mom = _safe_growth(mtd, prev_mtd)

        h_mtd = hist["mtd"].get(month)  # biarkan None tetap None
        h_ytd = hist["ytd"].get(month)  # biarkan None tetap None

        yoy_mtd = _safe_growth(mtd, h_mtd)
        yoy_ytd = _safe_growth(ytd, h_ytd)

        monthly.append({"month":month,"is_outlook":month in outlook_months,
            "mtd":mtd,"ytd":ytd,"target_mtd":tgt,"target_ytd":tytd,
            "ach_mtd":ach_mtd,"ach_ytd":ach_ytd,"mom":mom,
            "yoy_mtd":yoy_mtd,"yoy_ytd":yoy_ytd})

        prev_mtd = mtd

    return {"territory":territory,"segment":"SEG_D","year":year,
            "report_month":report_month,"valid_until":valid_until,
            "outlook_months":list(outlook_months),
            "outlook_detail":outlook_detail,"monthly":monthly}

def print_dps_result(result):
    seg=result["segment"]; rep=result["report_month"]
    valid=result["valid_until"]; terr=result["territory"]
    outl_str = ", ".join(result.get("outlook_months",[])) or "Tidak perlu"
    monthly=result["monthly"]
    print(f"\n{'='*72}")
    print(f"  REVENUE {seg} — {terr} | Laporan: {rep} | Valid s/d: {valid} | Outlook: {outl_str}")
    print(f"{'='*72}")
    print(f"\n  {'Bln':<5} {'MtD':>16} {'Target MtD':>16} {'Ach MtD':>8} {'MoM':>8} {'YoY MtD':>8} {'YtD':>16} {'Ach YtD':>8} {'YoY YtD':>8}  Ket")
    print(f"  {'─'*106}")
    for r in monthly:
        m=r["month"]
        mtd    =f"{r['mtd']:>16,.0f}"        if r["mtd"]     is not None else f"{'—':>16}"
        tgt    =f"{r['target_mtd']:>16,.0f}"
        ach_mtd=f"{r['ach_mtd']*100:>7.1f}%" if r["ach_mtd"] is not None else f"{'—':>8}"
        mom    =f"{r['mom']*100:>+7.1f}%"    if r["mom"]     is not None else f"{'—':>8}"
        yoy_mtd=f"{r['yoy_mtd']*100:>+7.1f}%" if r["yoy_mtd"] is not None else f"{'—':>8}"
        ytd    =f"{r['ytd']:>16,.0f}"        if r["ytd"]     is not None else f"{'—':>16}"
        ach_ytd=f"{r['ach_ytd']*100:>7.1f}%" if r["ach_ytd"] is not None else f"{'—':>8}"
        yoy_ytd=f"{r['yoy_ytd']*100:>+7.1f}%" if r["yoy_ytd"] is not None else f"{'—':>8}"
        ket    =" ← OUTLOOK" if r["is_outlook"] else ""
        print(f"  {m:<5} {mtd} {tgt} {ach_mtd} {mom} {yoy_mtd} {ytd} {ach_ytd} {yoy_ytd}{ket}")
    if result["outlook_detail"]:
        for month, detail in result["outlook_detail"].items():
            print(f"\n  Detail Komponen Outlook {month}:")
            print(f"  {'─'*55}")
            for k,v in detail.items():
                if k=="TOTAL OUTLOOK": print(f"  {'─'*55}"); print(f"  {'TOTAL OUTLOOK':<40} {v:>14,.0f}")
                else:
                    marker="  ← manual" if "manual" in k else ""
                    print(f"  {k:<40} {v:>14,.0f}{marker}")
    print(f"\n{'='*72}\n")
