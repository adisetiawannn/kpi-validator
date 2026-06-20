# src/exporters/excel_exporter.py
# Generate output Excel sesuai template kpi_summary.xlsx
#
# Sheet 1: summary_bulan_report  — ringkasan bulan N-1 vs N-2
# Sheet 2: all_detail_kpi        — detail semua bulan YtD

import os
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from config import MONTH_NUMBER_MAP, OUTPUT_DIR, OUTPUT_FILENAME_FMT, HISTORICAL_SHEET


# ── Color constants (sesuai template) ────────────────────────────────────────
COLOR_NAVY   = "002060"   # header MTD bulan N-1
COLOR_MAROON = "5F0F40"   # header YTD year lalu
COLOR_WHITE  = "FFFFFF"
COLOR_GREY   = "F2F2F2"
FONT_NAME    = "Aptos Narrow"

# ── Segment order (fixed) ─────────────────────────────────────────────────────
SEGMENTS_ORDER = ["SEG_A", "SEG_B", "SEG_C", "SEG_D", "SEG_A+SEG_B", "TOTAL"]
SEGMENTS_ORDER_COLLECTION = ["SEG_A", "SEG_B", "SEG_C", "SEG_D", "TOTAL"]

# ── Helpers ───────────────────────────────────────────────────────────────────

def _month_idx(month: str) -> int:
    return {v: k for k, v in MONTH_NUMBER_MAP.items()}[month.upper()]

# format angka untuk excel
def _fmt_num(val) -> any:  #type: ignore
    """Return numeric value or 0."""
    if val is None: return 0
    try: return round(float(val), 0)
    except: return 0

def _fmt_pct(val) -> any: #type: ignore
    """Return percentage as float (0.xx) or None."""
    if val is None: return 0
    try: return round(float(val), 4)
    except: return 0

def _thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(top=s, bottom=s, left=s, right=s)

def _style(cell, bold=False, bg=None, font_color="000000",
           h_align="center", v_align="center", wrap=False,
           size=12, border=True, num_format=None):
    cell.font      = Font(name=FONT_NAME, bold=bold,
                          color=font_color, size=size)
    cell.alignment = Alignment(horizontal=h_align,
                               vertical=v_align, wrap_text=wrap)
    if bg:
        cell.fill = PatternFill("solid", start_color=bg)
    if border:
        cell.border = _thin_border()
    if num_format:
        cell.number_format = num_format

def _hdr(ws, row, col, value, bg, fg=COLOR_WHITE,
         bold=True, wrap=False, h_align="center"):
    cell = ws.cell(row=row, column=col, value=value)
    _style(cell, bold=bold, bg=bg, font_color=fg,
           h_align=h_align, wrap=wrap)
    return cell

def _data(ws, row, col, value, num_format=None, bold=False, bg=None):
    cell = ws.cell(row=row, column=col, value=value)
    _style(cell, bold=bold, bg=bg,
           h_align="center", num_format=num_format)
    return cell


# ── Lookup helpers ────────────────────────────────────────────────────────────

# lookup value untuk revenue dan GROWTH = ambil nilai
def _get_monthly(results: dict, segment: str, month: str, field: str):
    """Get field value dari monthly list untuk segment dan bulan tertentu."""
    result = results.get(segment)
    if not result: return None
    for r in result["monthly"]:
        if r["month"] == month:
            return r.get(field)
    return None

# lookup value sales connectivity = ambil nilai
def _get_sc_monthly(sc_results, segment, indicator, month, field):
    """Ambil field dari SC result untuk segment+indicator+bulan tertentu."""
    if not sc_results: return None
    result = sc_results.get(segment)
    if not result: return None
    monthly = result.get(indicator, {}).get("monthly", [])
    for r in monthly:
        if r["month"] == month:
            return r.get(field)
    return None

# lookup value collection = ambil nilai
def _get_collection_monthly(collection_results, segment,indicator, month, field):
    if not collection_results: return None
    kpi_collection = collection_results.get(indicator)
    if not kpi_collection: return None
    monthly = kpi_collection.get(segment, [])
    for r in monthly :
        if r["month"] == month:
            return r.get(field)
    return None


# Normalisasi shape hasil kalkulasi sales connectivity
def _build_sc_results(sc_bs, sc_gs, sc_dss, sc_dps, agg_sc):
    """Normalize SC results ke struktur seragam per segment."""
    def _wrap_agg(agg_result):
        """Convert agg result {bsgs/total} ke format yang sama dengan segment."""
        if not agg_result: return None
        return {
            "PROD_A": {"monthly": agg_result["PROD_A"]["bsgs"]["monthly"]}  if agg_result.get("PROD_A") else {"monthly": []},
            "PROD_B": {"monthly": agg_result["PROD_B"]["bsgs"]["monthly"]}  if agg_result.get("PROD_B") else {"monthly": []},
            "PROD_C": {"monthly": agg_result["PROD_C"]["bsgs"]["monthly"]}  if agg_result.get("PROD_C") else {"monthly": []},
        }
    def _wrap_total(agg_result):
        if not agg_result: return None
        return {
            "PROD_A": {"monthly": agg_result["PROD_A"]["total"]["monthly"]}  if agg_result.get("PROD_A") else {"monthly": []},
            "PROD_B": {"monthly": agg_result["PROD_B"]["total"]["monthly"]}  if agg_result.get("PROD_B") else {"monthly": []},
            "PROD_C": {"monthly": agg_result["PROD_C"]["total"]["monthly"]}  if agg_result.get("PROD_C") else {"monthly": []},
        }
    return {
        "SEG_A":       sc_bs,
        "SEG_B":       sc_gs,
        "SEG_C":       sc_dss,
        "SEG_D":       sc_dps,
        "SEG_A+SEG_B": _wrap_agg(agg_sc),
        "TOTAL":       _wrap_total(agg_sc),
    }

def _build_results_map(
    r_bs, r_gs, r_dss, r_dps, r_bsgs, r_total
) -> dict:
    """Map segment label → result dict."""
    return {
        "SEG_A"      : r_bs,
        "SEG_B"      : r_gs,
        "SEG_C"      : r_dss,
        "SEG_D"      : r_dps,
        "SEG_A+SEG_B": r_bsgs,
        "TOTAL"      : r_total,
    }

# ── Sheet 1: summary_bulan_report ─────────────────────────────────────────────

def _build_summary_sheet(ws, results: dict, report_month: str,
                         year: int, hist_ytd: dict,
                         ngtma_results: dict = None, #type: ignore
                         ngtma_hist_ytd: dict = None, #type: ignore
                         sc_results=None,
                         sc_hist_ytd=None,
                         collection_hist_ytd=None,
                         collection_results=None): #type: ignore
    """
    Build sheet summary_bulan_report.

    Columns:
    A=SEGMENT, B=KPI
    C=Real MtD N-2
    D=TGT MtD N-1, E=REAL MtD N-1, F=ACH MtD, G=MoM
    H=Real YtD N-1 year lalu (dari historical)
    I=TGT YtD N-1, J=REAL YtD N-1, K=ACH YtD, L=YoY
    """
    report_month = report_month.upper()
    report_idx   = _month_idx(report_month)

    # Bulan N-2 dan N-1
    n2_idx   = report_idx - 2
    n1_idx   = report_idx - 1
    n2_month = MONTH_NUMBER_MAP.get(n2_idx)  # APR jika laporan JUN
    n1_month = MONTH_NUMBER_MAP.get(n1_idx)  # MAY jika laporan JUN

    # ── Row 1: Header grup ────────────────────────────────────────────────────
    # A1:A2 merge — SEGMENT
    ws.merge_cells("A1:A2")
    _hdr(ws, 1, 1, "SEGMENT", bg="1A3A5C")

    # B1:B2 merge — KPI
    ws.merge_cells("B1:B2")
    _hdr(ws, 1, 2, "KPI", bg="1A3A5C")

    # C1:C2 merge — Real MtD N-2
    n2_label = f"REAL MTD - {n2_month} {year}" if n2_month else "REAL MTD N-2"
    ws.merge_cells("C1:C2")
    _hdr(ws, 1, 3, n2_label, bg="1A3A5C", wrap=True)

    # D1:G1 merge — MTD N-1
    n1_label = f"MTD - {n1_month} {year}" if n1_month else "MTD N-1"
    ws.merge_cells("D1:G1")
    _hdr(ws, 1, 4, n1_label, bg=COLOR_NAVY)

    # H1:H2 merge — Real YtD N-1 year lalu
    ytd_ly_label = f"REAL YTD - {n1_month} {year - 1}" if n1_month else "REAL YTD N-1 LY"
    ws.merge_cells("H1:H2")
    _hdr(ws, 1, 8, ytd_ly_label, bg="1A3A5C", wrap=True)

    # I1:L1 merge — YTD N-1 year ini
    ytd_n1_label = f"YTD - {n1_month} {year}" if n1_month else "YTD N-1"
    ws.merge_cells("I1:L1")
    _hdr(ws, 1, 9, ytd_n1_label, bg=COLOR_MAROON)

    # ── Row 2: Sub-header ─────────────────────────────────────────────────────
    for col, label, bg in [
        (4, "TGT",  COLOR_NAVY),
        (5, "REAL", COLOR_NAVY),
        (6, "ACH",  COLOR_NAVY),
        (7, "MoM",  COLOR_NAVY),
        (9, "TGT",  COLOR_MAROON),
        (10,"REAL", COLOR_MAROON),
        (11,"ACH",  COLOR_MAROON),
        (12,"YoY",  COLOR_MAROON),
    ]:
        _hdr(ws, 2, col, label, bg=bg)

    # ── Rows 3-14: Data ───────────────────────────────────────────────────────
    kpi_rows = (
        [(seg, "REVENUE") for seg in SEGMENTS_ORDER] +
        [(seg, "GROWTH")  for seg in SEGMENTS_ORDER] +
        [(seg, "PROD_A")  for seg in SEGMENTS_ORDER] +
        [(seg, "PROD_B")  for seg in SEGMENTS_ORDER] +
        [(seg, "PROD_C")  for seg in SEGMENTS_ORDER] +
        [(seg, "COLL_A")  for seg in SEGMENTS_ORDER] +  # ← tambah
        [(seg, "COLL_B")  for seg in SEGMENTS_ORDER] +  # ← tambah
        [("SEG_A", "COLL_C")]                          + # ← SEG_A only
        [("SEG_A", "COLL_D")]                            # ← SEG_A only
    )

    for i, (seg, kpi) in enumerate(kpi_rows):
        row = i + 3
        bg  = COLOR_GREY if i % 2 == 0 else None

        # A: Segment, B: KPI
        _data(ws, row, 1, seg,  bold=True, bg=bg)
        _data(ws, row, 2, kpi,  bold=True, bg=bg)

        if kpi in ["PROD_A", "PROD_B", "PROD_C"]:
            _data(ws, row, 3,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, n2_month, "mtd")),        bg=bg, num_format="#,##0")
            _data(ws, row, 4,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, n1_month, "target_mtd")), bg=bg, num_format="#,##0")
            _data(ws, row, 5,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, n1_month, "mtd")),        bg=bg, num_format="#,##0")
            _data(ws, row, 6,  _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, n1_month, "ach_mtd")),    bg=bg, num_format="0.0%")
            _data(ws, row, 7,  _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, n1_month, "mom")),     bg=bg, num_format="+0.0%;-0.0%;0.0%")

            # Kolom H — Real YtD year lalu dari historical SC
            h_val = sc_hist_ytd.get(seg, {}).get(kpi, {}).get(n1_month) if (sc_hist_ytd and n1_month) else None
            _data(ws, row, 8,  _fmt_num(h_val),                                                          bg=bg, num_format="#,##0")
            _data(ws, row, 9,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, n1_month, "target_ytd")), bg=bg, num_format="#,##0")
            _data(ws, row, 10, _fmt_num(_get_sc_monthly(sc_results, seg, kpi, n1_month, "ytd")),        bg=bg, num_format="#,##0")
            _data(ws, row, 11, _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, n1_month, "ach_ytd")),    bg=bg, num_format="0.0%")
            _data(ws, row, 12, _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, n1_month, "yoy_ytd")), bg=bg, num_format="+0.0%;-0.0%;0.0%")
            continue

        if kpi in ["COLL_A", "COLL_B", "COLL_C", "COLL_D"] :
            if kpi in ["COLL_A", "COLL_B"] and seg !='SEG_A+SEG_B':
                _data(ws, row, 3,  None)
                _data(ws, row, 4,  None)
                _data(ws, row, 5,  None)
                _data(ws, row, 6,  None)
                _data(ws, row, 7,  None)
                h_val = collection_hist_ytd.get(seg, {}).get(kpi, {}).get(n1_month) #type:ignore
                _data(ws, row, 8,  _fmt_pct(h_val),bg=bg, num_format="0.00%")
                _data(ws, row, 9,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "target_ytd")), bg=bg, num_format="0.00%")
                _data(ws, row, 10,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "ytd")),        bg=bg, num_format="0.00%")
                _data(ws, row, 11, _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "ach_ytd")),    bg=bg, num_format="0.00%")
                _data(ws, row, 12, _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "yoy_ytd")),    bg=bg, num_format="+0.0%;-0.0%;0.0%")

            elif kpi in ["COLL_D", "COLL_C"] and seg == "SEG_A" :
                _data(ws, row, 3,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n2_month, "mtd")),bg=bg, num_format="0.00%")
                _data(ws, row, 4,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "target_mtd")), bg=bg, num_format="0.00%")
                _data(ws, row, 5,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "mtd")),        bg=bg, num_format="0.00%")
                _data(ws, row, 6, _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "ach_mtd")),    bg=bg, num_format="0.00%")
                _data(ws, row, 7, _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, n1_month, "mom")),    bg=bg, num_format="+0.0%;-0.0%;0.0%")
                _data(ws, row, 8,  None)
                _data(ws, row, 9,  None)
                _data(ws, row, 10, None)
                _data(ws, row, 11, None)
            continue

        if kpi == "GROWTH":
            if not ngtma_results or not ngtma_results.get(seg):
                for col in range(3, 13):
                    _data(ws, row, col, None, bg=bg)

            else:
                ngtma_hist = ngtma_hist_ytd or {}
                c_val = _get_monthly(ngtma_results, seg, n2_month, "mtd") if n2_month else None
                _data(ws, row, 3,  _fmt_num(c_val),  bg=bg, num_format="#,##0")
                d_val = _get_monthly(ngtma_results, seg, n1_month, "target_mtd") if n1_month else None
                _data(ws, row, 4,  _fmt_num(d_val),  bg=bg, num_format="#,##0")
                e_val = _get_monthly(ngtma_results, seg, n1_month, "mtd") if n1_month else None
                _data(ws, row, 5,  _fmt_num(e_val),  bg=bg, num_format="#,##0")
                f_val = _get_monthly(ngtma_results, seg, n1_month, "ach_mtd") if n1_month else None
                _data(ws, row, 6,  _fmt_pct(f_val),  bg=bg, num_format="0.0%")
                g_val = _get_monthly(ngtma_results, seg, n1_month, "mom") if n1_month else None
                _data(ws, row, 7,  _fmt_pct(g_val),  bg=bg, num_format="+0.0%;-0.0%;0.0%")
                h_val = ngtma_hist.get(seg, {}).get(n1_month) if n1_month else None
                _data(ws, row, 8,  _fmt_num(h_val),  bg=bg, num_format="#,##0")
                i_val = _get_monthly(ngtma_results, seg, n1_month, "target_ytd") if n1_month else None
                _data(ws, row, 9,  _fmt_num(i_val),  bg=bg, num_format="#,##0")
                j_val = _get_monthly(ngtma_results, seg, n1_month, "ytd") if n1_month else None
                _data(ws, row, 10, _fmt_num(j_val),  bg=bg, num_format="#,##0")
                k_val = _get_monthly(ngtma_results, seg, n1_month, "ach_ytd") if n1_month else None
                _data(ws, row, 11, _fmt_pct(k_val),  bg=bg, num_format="0.0%")
                l_val = _get_monthly(ngtma_results, seg, n1_month, "yoy_ytd") if n1_month else None
                _data(ws, row, 12, _fmt_pct(l_val),  bg=bg, num_format="+0.0%;-0.0%;0.0%")
            continue

        # C: Real MtD N-2
        c_val = _get_monthly(results, seg, n2_month, "mtd") if n2_month else None
        _data(ws, row, 3, _fmt_num(c_val), bg=bg, num_format="#,##0")

        # D: Target MtD N-1
        d_val = _get_monthly(results, seg, n1_month, "target_mtd") if n1_month else None
        _data(ws, row, 4, _fmt_num(d_val), bg=bg, num_format="#,##0")

        # E: Real MtD N-1
        e_val = _get_monthly(results, seg, n1_month, "mtd") if n1_month else None
        _data(ws, row, 5, _fmt_num(e_val), bg=bg, num_format="#,##0")

        # F: ACH MtD N-1
        f_val = _get_monthly(results, seg, n1_month, "ach_mtd") if n1_month else None
        _data(ws, row, 6, _fmt_pct(f_val), bg=bg, num_format="0.0%")

        # G: MoM
        g_val = _get_monthly(results, seg, n1_month, "mom") if n1_month else None
        _data(ws, row, 7, _fmt_pct(g_val), bg=bg, num_format="+0.0%;-0.0%;0.0%")

        # H: Real YtD N-1 year lalu (dari historical)
        h_val = hist_ytd.get(seg, {}).get(n1_month) if n1_month else None
        _data(ws, row, 8, _fmt_num(h_val), bg=bg, num_format="#,##0")

        # I: Target YtD N-1 year ini
        i_val = _get_monthly(results, seg, n1_month, "target_ytd") if n1_month else None
        _data(ws, row, 9, _fmt_num(i_val), bg=bg, num_format="#,##0")

        # J: Real YtD N-1 year ini
        j_val = _get_monthly(results, seg, n1_month, "ytd") if n1_month else None
        _data(ws, row, 10, _fmt_num(j_val), bg=bg, num_format="#,##0")

        # K: ACH YtD
        k_val = _get_monthly(results, seg, n1_month, "ach_ytd") if n1_month else None
        _data(ws, row, 11, _fmt_pct(k_val), bg=bg, num_format="0.0%")

        # L: YoY MtD N-1
        l_val = _get_monthly(results, seg, n1_month, "yoy_ytd") if n1_month else None
        _data(ws, row, 12, _fmt_pct(l_val), bg=bg, num_format="+0.0%;-0.0%;0.0%")

    # ── Column widths ─────────────────────────────────────────────────────────
    widths = {"A":10,"B":10,"C":16,"D":16,"E":16,"F":9,
              "G":9,"H":16,"I":16,"J":16,"K":9,"L":9}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 20
    ws.freeze_panes = "C3"


# ── Sheet 2: all_detail_kpi ───────────────────────────────────────────────────

def _build_detail_sheet(ws, results: dict,
                        report_month: str, year: int,
                        ngtma_results: dict = None, # type:ignore
                        sc_results=None,
                        collection_results = None): #type: ignore
    """
    Build sheet all_detail_kpi.
    Flat rows: SEGMENT × KPI × BULAN (YYYYMM)
    """
    report_month = report_month.upper()
    report_idx   = _month_idx(report_month)

    # ── Header ────────────────────────────────────────────────────────────────
    headers = ["SEGMENT","KPI","BULAN",
               "TGT MTD","REAL MTD","ACH MTD","MOM",
               "TGT YTD","REAL YTD","ACH YTD","YOY"]
    bg_hdr  = "1A3A5C"
    for col, h in enumerate(headers, start=1):
        _hdr(ws, 1, col, h, bg=bg_hdr)

    ws.row_dimensions[1].height = 22

    # ── Data rows ─────────────────────────────────────────────────────────────
    row = 2
    for month_idx in range(1, report_idx):
        month    = MONTH_NUMBER_MAP[month_idx]
        bulan_id = year * 100 + month_idx   # e.g. 202601

        for seg in SEGMENTS_ORDER:
            for kpi in ["REVENUE", "GROWTH","PROD_A", "PROD_B", "PROD_C", "COLL_B","COLL_A","COLL_C","COLL_D"]:
                bg = COLOR_GREY if row % 2 == 0 else None

                _data(ws, row, 1, seg,      bold=True, bg=bg)
                _data(ws, row, 2, kpi,      bold=True, bg=bg)
                _data(ws, row, 3, bulan_id, bg=bg,
                      num_format="000000")

                # Sales Connectivity handler
                if kpi in ["PROD_A", "PROD_B", "PROD_C"]:
                    _data(ws, row, 4,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, month, "target_mtd")), bg=bg, num_format="#,##0")
                    _data(ws, row, 5,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, month, "mtd")),        bg=bg, num_format="#,##0")
                    _data(ws, row, 6,  _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, month, "ach_mtd")),    bg=bg, num_format="0.0%")
                    _data(ws, row, 7,  _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, month, "mom")),        bg=bg, num_format="+0.0%;-0.0%;0.0%")
                    _data(ws, row, 8,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, month, "target_ytd")), bg=bg, num_format="#,##0")
                    _data(ws, row, 9,  _fmt_num(_get_sc_monthly(sc_results, seg, kpi, month, "ytd")),        bg=bg, num_format="#,##0")
                    _data(ws, row, 10, _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, month, "ach_ytd")),    bg=bg, num_format="0.0%")
                    _data(ws, row, 11, _fmt_pct(_get_sc_monthly(sc_results, seg, kpi, month, "yoy_ytd")),    bg=bg, num_format="+0.0%;-0.0%;0.0%")
                    row += 1
                    continue

                # GROWTH handler
                if kpi == "GROWTH":
                    nr = None
                    if ngtma_results and ngtma_results.get(seg):
                        for m in ngtma_results[seg].get("monthly", []):
                            if m["month"] == month:
                                nr = m; break
                    if nr:
                        _data(ws, row, 4,  _fmt_num(nr.get("target_mtd")), bg=bg, num_format="#,##0")
                        _data(ws, row, 5,  _fmt_num(nr.get("mtd")),        bg=bg, num_format="#,##0")
                        _data(ws, row, 6,  _fmt_pct(nr.get("ach_mtd")),    bg=bg, num_format="0.0%")
                        _data(ws, row, 7,  _fmt_pct(nr.get("mom")),        bg=bg, num_format="+0.0%;-0.0%;0.0%")
                        _data(ws, row, 8,  _fmt_num(nr.get("target_ytd")), bg=bg, num_format="#,##0")
                        _data(ws, row, 9,  _fmt_num(nr.get("ytd")),        bg=bg, num_format="#,##0")
                        _data(ws, row, 10, _fmt_pct(nr.get("ach_ytd")),    bg=bg, num_format="0.0%")
                        _data(ws, row, 11, _fmt_pct(nr.get("yoy_ytd")),    bg=bg, num_format="+0.0%;-0.0%;0.0%")
                    row += 1
                    continue

                # COLLECTION handler
                if kpi in ["COLL_A", "COLL_B", "COLL_C", "COLL_D"] :
                    if kpi in ["COLL_A", "COLL_B"] and seg != 'SEG_A+SEG_B' :
                        _data(ws, row, 4,  None)
                        _data(ws, row, 5,  None)
                        _data(ws, row, 6,  None)
                        _data(ws, row, 7,  None)
                        _data(ws, row, 8,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "target_ytd")), bg=bg, num_format="0.00%")
                        _data(ws, row, 9,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "ytd")),        bg=bg, num_format="0.00%")
                        _data(ws, row, 10, _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "ach_ytd")),    bg=bg, num_format="0.00%")
                        _data(ws, row, 11, _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "yoy_ytd")),    bg=bg, num_format="+0.0%;-0.0%;0.0%")

                    elif kpi in ["COLL_D", "COLL_C"] and seg == 'SEG_A':
                        _data(ws, row, 4,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "target_mtd")), bg=bg, num_format="0.00%")
                        _data(ws, row, 5,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "mtd")),        bg=bg, num_format="0.00%")
                        _data(ws, row, 6,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "ach_mtd")),    bg=bg, num_format="0.00%")
                        _data(ws, row, 7,  _fmt_pct(_get_collection_monthly(collection_results, seg, kpi, month, "mom")),        bg=bg, num_format="+0.0%;-0.0%;0.0%")
                        _data(ws, row, 8,  None)
                        _data(ws, row, 9,  None)
                        _data(ws, row, 10, None)
                        _data(ws, row, 11, None)
                    row += 1
                    continue

                # REVENUE handler (default case)
                r = results.get(seg)
                monthly_row = None
                if r:
                    for m in r["monthly"]:
                        if m["month"] == month:
                            monthly_row = m
                            break

                tgt_mtd = _fmt_num(monthly_row["target_mtd"]) if monthly_row else None
                real_mtd= _fmt_num(monthly_row["mtd"])        if monthly_row else None
                ach_mtd = _fmt_pct(monthly_row["ach_mtd"])    if monthly_row else None
                mom     = _fmt_pct(monthly_row["mom"])        if monthly_row else None
                tgt_ytd = _fmt_num(monthly_row["target_ytd"]) if monthly_row else None
                real_ytd= _fmt_num(monthly_row["ytd"])        if monthly_row else None
                ach_ytd = _fmt_pct(monthly_row["ach_ytd"])    if monthly_row else None
                yoy     = _fmt_pct(monthly_row["yoy_ytd"])    if monthly_row else None

                _data(ws, row, 4,  tgt_mtd,  bg=bg, num_format="#,##0")
                _data(ws, row, 5,  real_mtd, bg=bg, num_format="#,##0")
                _data(ws, row, 6,  ach_mtd,  bg=bg, num_format="0.0%")
                _data(ws, row, 7,  mom,      bg=bg, num_format="+0.0%;-0.0%;0.0%")
                _data(ws, row, 8,  tgt_ytd,  bg=bg, num_format="#,##0")
                _data(ws, row, 9,  real_ytd, bg=bg, num_format="#,##0")
                _data(ws, row, 10, ach_ytd,  bg=bg, num_format="0.0%")
                _data(ws, row, 11, yoy,      bg=bg, num_format="+0.0%;-0.0%;0.0%")
                row += 1

    # ── Column widths ─────────────────────────────────────────────────────────
    widths = {"A":10,"B":10,"C":10,"D":16,"E":16,
              "F":9,"G":9,"H":16,"I":16,"J":9,"K":9}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "D2"


# ── Main export function ──────────────────────────────────────────────────────

def export_to_excel(
    r_bs, r_gs, r_dss, r_dps,
    agg_bsgs, agg_total,
    n_bs=None, n_gs=None, n_dss=None, n_dps=None,
    agg_ngtma_bsgs=None, agg_ngtma_total=None,
    report_month   : str = "",
    year          : int = 0,
    territory      : str = "",
    historical_file: str = "",
    output_dir     : str = None,    #type: ignore
    sc_bs                = None,
    sc_gs                = None,
    sc_dss               = None,
    sc_dps               = None,
    agg_sc               = None,
    collection_results = None,
) -> str:
    """
    Generate output Excel dari semua hasil kalkulasi Revenue dan GROWTH.

    Returns:
        str: path ke file yang di-generate
    """
    from src.processors.segments._historical_helper import load_historical


    # Revenue results map
    results = {
        "SEG_A"      : r_bs,
        "SEG_B"      : r_gs,
        "SEG_C"      : r_dss,
        "SEG_D"      : r_dps,
        "SEG_A+SEG_B": agg_bsgs,
        "TOTAL"      : agg_total,
    }

    # GROWTH results map
    ngtma_results = {
        "SEG_A"      : n_bs,
        "SEG_B"      : n_gs,
        "SEG_C"      : n_dss,
        "SEG_D"      : n_dps,
        "SEG_A+SEG_B": agg_ngtma_bsgs,
        "TOTAL"      : agg_ngtma_total,
    }

     # Sales Connectivity results map
    sc_results = _build_sc_results(sc_bs, sc_gs, sc_dss, sc_dps, agg_sc)

    # Load historical YtD untuk kolom H - Revenue dan GROWTH
    # hist_ytd: {segment: {"JAN": val, "FEB": val, ...}}
    hist_ytd = {}
    for seg in ["SEG_A", "SEG_B", "SEG_C", "SEG_D", "SEG_A+SEG_B", "TOTAL"]:
        h = load_historical(historical_file, territory, seg,"REVENUE")
        hist_ytd[seg] = h["ytd"]

    # Load historical YtD untuk kolom H - Sales Connectivity
    sc_hist_ytd = {}
    for seg in ["SEG_A", "SEG_B", "SEG_C", "SEG_D", "SEG_A+SEG_B", "TOTAL"]:
        sc_hist_ytd[seg] = {}
        for indicator in ["PROD_A", "PROD_B", "PROD_C"]:
            h = load_historical(historical_file, territory, seg, indicator)
            sc_hist_ytd[seg][indicator] = h["ytd"]

    # Load historical YtD untuk kolom H - Collection : COLL_B & COLL_A
    from src.loaders.collection_loader import _load_historical as _load_collection_hist

    collection_hist_ytd = {}
    for seg in ["SEG_A", "SEG_B", "SEG_C", "SEG_D", "TOTAL"]:
        collection_hist_ytd[seg] = {}
        for indicator in ["COLL_B", "COLL_A"]:
            h = _load_collection_hist(historical_file, territory, seg, indicator)
            collection_hist_ytd[seg][indicator] = h.to_dict()

    # Setup output path
    out_dir = output_dir or OUTPUT_DIR
    Path(out_dir).mkdir(parents=True, exist_ok=True)



    filename = OUTPUT_FILENAME_FMT.format(
        territory = territory,
        month     = report_month.upper(),
        year      = year,
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    filepath = os.path.join(out_dir, filename)

    # Build workbook
    wb = Workbook()

    # Sheet 1: summary
    ws1 = wb.active
    ws1.title = "summary_bulan_report" #type:ignore
    _build_summary_sheet(ws1, results, report_month, year, hist_ytd,
                        ngtma_results=ngtma_results,
                        ngtma_hist_ytd=_build_ngtma_hist_ytd(historical_file, territory, report_month, year),
                        sc_results=sc_results,
                        sc_hist_ytd=sc_hist_ytd,
                        collection_results=collection_results,        # ← tambah
                        collection_hist_ytd=collection_hist_ytd,        # ← tambah
                        )

    # Sheet 2: detail
    ws2 = wb.create_sheet("all_detail_kpi")
    _build_detail_sheet(ws2, results, report_month, year,
                        ngtma_results=ngtma_results,sc_results=sc_results,
                        collection_results=collection_results)

    wb.save(filepath)
    print(f"  [EXPORT] ✓ File saved: {filepath}")
    return filepath


# ── GROWTH historical helper ───────────────────────────────────────────────────

def _build_ngtma_hist_ytd(historical_file: str, territory: str,
                           report_month: str, year: int) -> dict:
    """Build GROWTH historical YtD untuk kolom H di summary sheet."""
    import pandas as pd
    from config import (MONTH_COLS_MTD, MONTH_NUMBER_MAP,
                        HISTORICAL_HEADER_ROW, HIST_GROWTH)

    result = {}
    for seg in SEGMENTS_ORDER:
        df = pd.read_excel(historical_file, sheet_name=HISTORICAL_SHEET,
                           header=HISTORICAL_HEADER_ROW)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(subset=["INDICATOR","SEGMENT"]).copy()
        mask = ((df["TERRITORY"]==territory) &
                (df["INDICATOR"]==HIST_GROWTH) &
                (df["SEGMENT"]==seg))
        row = df[mask]

        ytd = {}; running = 0.0
        for month in MONTH_COLS_MTD:
            if not row.empty and month in df.columns:
                val = row.iloc[0][month]
                v   = None if pd.isna(val) else float(val)
            else:
                v = None
            if v is not None: running += v
            ytd[month] = running  # ← carry forward
        result[seg] = ytd

    return result
