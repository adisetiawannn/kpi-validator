# tools/generate_sample_data.py
# ----------------------------------------------------------------------------
# Generator DATA DUMMY untuk showcase KPI Validator.
#
# Semua angka di sini FIKTIF dan dihasilkan dari formula deterministik sederhana
# (base * faktor bulan) — TIDAK diambil, disalin, atau dimiripkan dari data nyata
# manapun. Tujuannya hanya agar siapa pun yang clone repo bisa menjalankan
# `python main.py --month 6 --year 2026` dan melihat output Excel-nya.
#
# Jalankan:  python tools/generate_sample_data.py
# Output  :  9 file di data/input/ sesuai nama di config.py
# ----------------------------------------------------------------------------

from pathlib import Path
import pandas as pd

ROOT      = Path(__file__).resolve().parents[1]
INPUT_DIR = ROOT / "data" / "input"
INPUT_DIR.mkdir(parents=True, exist_ok=True)

TERRITORY = "AREA_01"   # territory generik — hindari istilah identitas perusahaan
YEAR      = 2026
MONTHS    = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
             "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


# ── Helper: bikin deret 12 bulan deterministik dari sebuah base ───────────────
def wave(base: float, step: float = 0.015, ndigits_round=None):
    """base * (1 + step*i) untuk i=0..11. Naik landai, jelas sintetis."""
    vals = [base * (1 + step * i) for i in range(12)]
    if ndigits_round is not None:
        vals = [round(v / ndigits_round) * ndigits_round for v in vals]
    return vals


def money_row(base, factor=1.0):
    """Deret rupiah, dibulatkan ke jutaan terdekat."""
    return wave(base * factor, ndigits_round=1_000_000)


def month_map(vals):
    return {m: v for m, v in zip(MONTHS, vals)}


# ============================================================================
# 1. REVENUE FILE.xlsx  — sheet REVENUE_DATA, header di baris index 2
#    Kolom: YEAR, INDICATOR, SEGMENT, TERRITORY, TYPE, JAN..DEC
# ============================================================================
REV_BASE = {"CODE_A": 5_000_000_000, "CODE_B": 3_000_000_000,
            "CODE_C": 1_500_000_000, "CODE_D": 800_000_000}
GROWTH_BASE = {"CODE_A": 250_000_000, "CODE_B": 150_000_000,
               "CODE_C": 80_000_000,  "CODE_D": 40_000_000}

# Proporsi komponen TYPE_x terhadap nilai ALL (untuk perhitungan outlook).
TYPE_FRACTION = {
    "TYPE_1": 0.80, "TYPE_2": 0.05, "TYPE_3": 0.03, "TYPE_4": 0.06,
    "TYPE_5": 0.06, "TYPE_6": 0.10, "TYPE_7": 0.10, "TYPE_8": 0.02,
}


def build_revenue_file():
    rows = []
    for code, base in REV_BASE.items():
        all_vals = money_row(base)
        # baris ALL (dipakai untuk bulan valid)
        rows.append({"YEAR": YEAR, "INDICATOR": "REVENUE", "SEGMENT": code,
                     "TERRITORY": TERRITORY, "TYPE": "ALL", **month_map(all_vals)})
        # baris komponen TYPE_1..TYPE_8 (dipakai untuk bulan outlook)
        for tname, frac in TYPE_FRACTION.items():
            comp = [round(v * frac / 1_000_000) * 1_000_000 for v in all_vals]
            rows.append({"YEAR": YEAR, "INDICATOR": "REVENUE", "SEGMENT": code,
                         "TERRITORY": TERRITORY, "TYPE": tname, **month_map(comp)})
    # GROWTH: satu baris TYPE=GROWTH per segment
    for code, base in GROWTH_BASE.items():
        gvals = money_row(base)
        rows.append({"YEAR": YEAR, "INDICATOR": "GROWTH", "SEGMENT": code,
                     "TERRITORY": TERRITORY, "TYPE": "GROWTH", **month_map(gvals)})

    df = pd.DataFrame(rows, columns=["YEAR", "INDICATOR", "SEGMENT",
                                     "TERRITORY", "TYPE"] + MONTHS)
    path = INPUT_DIR / "REVENUE FILE.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        # header di baris index 2 (REVENUE_HEADER_ROW = 2) → startrow=2
        ws = xw.book.create_sheet("REVENUE_DATA")
        # tulis dua baris judul agar header jatuh di row index 2
        df.to_excel(xw, sheet_name="REVENUE_DATA", index=False, startrow=2)
        ws_obj = xw.sheets["REVENUE_DATA"]
        ws_obj["A1"] = "DUMMY DATA — REVENUE & GROWTH (fiktif, untuk showcase)"
        # buang sheet default kosong jika ada
        if "Sheet" in xw.book.sheetnames:
            del xw.book["Sheet"]
    print(f"  ✓ {path.name}")


# ============================================================================
# Display-segment helper untuk Revenue/GROWTH (target & historical)
# ============================================================================
DISPLAY_SEGMENTS = ["SEG_A", "SEG_B", "SEG_C", "SEG_D", "SEG_A+SEG_B", "TOTAL"]
# nilai "actual" revenue per display-segment (untuk turunkan target/historical)
REV_DISPLAY_BASE = {
    "SEG_A": 5_000_000_000, "SEG_B": 3_000_000_000,
    "SEG_C": 1_500_000_000, "SEG_D": 800_000_000,
    "SEG_A+SEG_B": 8_000_000_000, "TOTAL": 10_300_000_000,
}
GROWTH_DISPLAY_BASE = {
    "SEG_A": 250_000_000, "SEG_B": 150_000_000,
    "SEG_C": 80_000_000,  "SEG_D": 40_000_000,
    "SEG_A+SEG_B": 400_000_000, "TOTAL": 520_000_000,
}


def _revenue_growth_sheet(kind: str, factor: float):
    """kind = 'target' / 'historical'. factor menskala nilai actual."""
    rows = []
    for seg, base in REV_DISPLAY_BASE.items():
        rows.append({"TERRITORY": TERRITORY, "INDICATOR": "REVENUE", "SEGMENT": seg,
                     **month_map(money_row(base, factor))})
    for seg, base in GROWTH_DISPLAY_BASE.items():
        rows.append({"TERRITORY": TERRITORY, "INDICATOR": "GROWTH", "SEGMENT": seg,
                     **month_map(money_row(base, factor))})
    return pd.DataFrame(rows, columns=["TERRITORY", "INDICATOR", "SEGMENT"] + MONTHS)


# ── Sales Connectivity actual (unit/qty & speed) per segment per indikator ────
# Nilai MtD bulanan (konstan antar bulan, jelas sintetis).
SC_ACTUAL = {
    "SEG_A": {"PROD_A": 200, "PROD_B": 1500, "PROD_C": 90},
    "SEG_B": {"PROD_A": 120, "PROD_B": 1200, "PROD_C": 60},
    "SEG_C": {"PROD_A": 80,  "PROD_B": 800,  "PROD_C": 40},
    "SEG_D": {"PROD_A": 40,  "PROD_B": 400,  "PROD_C": 20},
}
# turunan agregat
SC_ACTUAL["SEG_A+SEG_B"] = {k: SC_ACTUAL["SEG_A"][k] + SC_ACTUAL["SEG_B"][k]
                            for k in ["PROD_A", "PROD_B", "PROD_C"]}
SC_ACTUAL["TOTAL"] = {k: sum(SC_ACTUAL[s][k] for s in ["SEG_A", "SEG_B", "SEG_C", "SEG_D"])
                      for k in ["PROD_A", "PROD_B", "PROD_C"]}


def _sales_connectivity_sheet(factor: float):
    rows = []
    for seg in DISPLAY_SEGMENTS:
        for ind in ["PROD_A", "PROD_B", "PROD_C"]:
            base = SC_ACTUAL[seg][ind]
            vals = wave(base * factor, ndigits_round=1)
            rows.append({"TERRITORY": TERRITORY, "INDICATOR": ind, "SEGMENT": seg,
                         **month_map(vals)})
    return pd.DataFrame(rows, columns=["TERRITORY", "INDICATOR", "SEGMENT"] + MONTHS)


# ── Collection (persentase 0..1) ──────────────────────────────────────────────
COLL_BASE = {  # nilai "actual" rate per indikator per segment
    "COLL_A": {"SEG_A": 0.90, "SEG_B": 0.88, "SEG_C": 0.86, "SEG_D": 0.84,
               "SEG_A+SEG_B": 0.89, "TOTAL": 0.87},
    "COLL_B": {"SEG_A": 0.92, "SEG_B": 0.90, "SEG_C": 0.88, "SEG_D": 0.86,
               "SEG_A+SEG_B": 0.91, "TOTAL": 0.89},
    "COLL_C": {"SEG_A": 0.88},
    "COLL_D": {"SEG_A": 0.95},
}


def _collection_pct_series(base):
    """Deret persentase landai, dijaga <= 0.99, dibulatkan 4 desimal."""
    vals = [min(0.99, round(base + 0.004 * i, 4)) for i in range(12)]
    return month_map(vals)


def _collection_sheet(factor: float):
    """Sheet Collection untuk target/historical file. factor menskala rate."""
    rows = []
    for ind, segmap in COLL_BASE.items():
        for seg, base in segmap.items():
            rows.append({"TERRITORY": TERRITORY, "INDICATOR": ind, "SEGMENT": seg,
                         **_collection_pct_series(base * factor)})
    return pd.DataFrame(rows, columns=["TERRITORY", "INDICATOR", "SEGMENT"] + MONTHS)


# ============================================================================
# 2. data target.xlsx  &  3. data historical.xlsx
#    3 sheet: Revenue, Sales Connectivity, Collection — header di baris index 1
# ============================================================================
def build_target_historical(filename: str, factor: float):
    path = INPUT_DIR / filename
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        sheets = {
            "Revenue":            _revenue_growth_sheet("rg", factor),
            "Sales Connectivity": _sales_connectivity_sheet(factor),
            "Collection":         _collection_sheet(min(factor, 1.0)),
        }
        for name, df in sheets.items():
            # header di baris index 1 (HEADER_ROW = 1) → startrow=1
            df.to_excel(xw, sheet_name=name, index=False, startrow=1)
            xw.sheets[name]["A1"] = f"DUMMY DATA — {name} (fiktif, untuk showcase)"
        if "Sheet" in xw.book.sheetnames:
            del xw.book["Sheet"]
    print(f"  ✓ {path.name}")


# ============================================================================
# 4. data_collection.xlsx  — sheet 'collection', header di baris index 0
#    Kolom: INDICATOR, SEGMENT, TERRITORY_NAME, YEAR, JAN..DEC
# ============================================================================
def build_collection_data():
    rows = []
    # COLL_A & COLL_B untuk SEG_A,SEG_B,SEG_C,SEG_D,SEG_A+SEG_B,TOTAL
    for ind in ["COLL_A", "COLL_B"]:
        for seg in DISPLAY_SEGMENTS:
            base = COLL_BASE[ind].get(seg)
            if base is None:  # SEG_A+SEG_B di-handle agregat oleh processor; tetap isi
                base = (COLL_BASE[ind]["SEG_A"] + COLL_BASE[ind]["SEG_B"]) / 2
            rows.append({"INDICATOR": ind, "SEGMENT": seg,
                         "TERRITORY_NAME": TERRITORY, "YEAR": YEAR,
                         **_collection_pct_series(base)})
    # COLL_C & COLL_D untuk SEG_A saja
    for ind in ["COLL_C", "COLL_D"]:
        rows.append({"INDICATOR": ind, "SEGMENT": "SEG_A",
                     "TERRITORY_NAME": TERRITORY, "YEAR": YEAR,
                     **_collection_pct_series(COLL_BASE[ind]["SEG_A"])})

    df = pd.DataFrame(rows, columns=["INDICATOR", "SEGMENT", "TERRITORY_NAME",
                                     "YEAR"] + MONTHS)
    path = INPUT_DIR / "data_collection.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="collection", index=False)  # header di row 0
        if "Sheet" in xw.book.sheetnames:
            del xw.book["Sheet"]
    print(f"  ✓ {path.name}")


# ============================================================================
# 5-9. Sales CSV (tab-separated, UTF-16, header di baris 0)
# ============================================================================
def _periods():
    return [YEAR * 100 + m for m in range(1, 13)]  # 202601..202612


def build_sales_csv():
    periods = _periods()

    # 5. PRODA FILE.csv (source A): TERRITORY, PRODUCT, PERIOD, SALES_QTY
    #    PROD_A = PRODUCT_A_1..A_4 ; PROD_C = PRODUCT_C_1..C_3 ; semua SEG_A
    rows = []
    proda_products = ["PRODUCT_A_1", "PRODUCT_A_2", "PRODUCT_A_3", "PRODUCT_A_4"]
    prodc_products = ["PRODUCT_C_1", "PRODUCT_C_2", "PRODUCT_C_3"]
    for per in periods:
        for prod in proda_products:
            rows.append({"TERRITORY": TERRITORY, "PRODUCT": prod,
                         "PERIOD": per, "SALES_QTY": 50})   # 4 x 50 = 200 / bulan
        for prod in prodc_products:
            rows.append({"TERRITORY": TERRITORY, "PRODUCT": prod,
                         "PERIOD": per, "SALES_QTY": 30})   # 3 x 30 = 90 / bulan
    pd.DataFrame(rows, columns=["TERRITORY", "PRODUCT", "PERIOD", "SALES_QTY"]) \
        .to_csv(INPUT_DIR / "PRODA FILE.csv", sep="\t", encoding="utf-16", index=False)
    print("  ✓ PRODA FILE.csv")

    # 6. PRODB FILE.csv (source A): TERRITORY_NEW, PERIOD, SALES_SPEED  (SEG_A)
    rows = [{"TERRITORY_NEW": TERRITORY, "PERIOD": per, "SALES_SPEED": 1500}
            for per in periods]
    pd.DataFrame(rows, columns=["TERRITORY_NEW", "PERIOD", "SALES_SPEED"]) \
        .to_csv(INPUT_DIR / "PRODB FILE.csv", sep="\t", encoding="utf-16", index=False)
    print("  ✓ PRODB FILE.csv")

    # division → nilai per bulan untuk source B
    DIV_PRODA = {"DIV_B": 120, "DIV_C": 80, "DIV_D": 40}   # SEG_B/C/D
    DIV_PRODC = {"DIV_B": 60,  "DIV_C": 40, "DIV_D": 20}
    DIV_PRODB = {"DIV_B": 1200, "DIV_C": 800, "DIV_D": 400}

    # 7. PRODA FILE 2.csv (source B): territory_new, PRODUCT_NAME, division_new, period, sales_qty
    rows = []
    for per in periods:
        for div, qty in DIV_PRODA.items():
            rows.append({"territory_new": TERRITORY, "PRODUCT_NAME": "PROD_A",
                         "division_new": div, "period": per, "sales_qty": qty})
    pd.DataFrame(rows, columns=["territory_new", "PRODUCT_NAME", "division_new",
                                "period", "sales_qty"]) \
        .to_csv(INPUT_DIR / "PRODA FILE 2.csv", sep="\t", encoding="utf-16", index=False)
    print("  ✓ PRODA FILE 2.csv")

    # 8. PRODB FILE 2.csv (source B): territory_new, division_new, period, sales_speed
    rows = []
    for per in periods:
        for div, spd in DIV_PRODB.items():
            rows.append({"territory_new": TERRITORY, "division_new": div,
                         "period": per, "sales_speed": spd})
    pd.DataFrame(rows, columns=["territory_new", "division_new", "period", "sales_speed"]) \
        .to_csv(INPUT_DIR / "PRODB FILE 2.csv", sep="\t", encoding="utf-16", index=False)
    print("  ✓ PRODB FILE 2.csv")

    # 9. PRODC FILE.csv (source B): territory_new, PRODUCT_NAME, division_new, period, sales_qty
    rows = []
    for per in periods:
        for div, qty in DIV_PRODC.items():
            rows.append({"territory_new": TERRITORY, "PRODUCT_NAME": "PROD_C",
                         "division_new": div, "period": per, "sales_qty": qty})
    pd.DataFrame(rows, columns=["territory_new", "PRODUCT_NAME", "division_new",
                                "period", "sales_qty"]) \
        .to_csv(INPUT_DIR / "PRODC FILE.csv", sep="\t", encoding="utf-16", index=False)
    print("  ✓ PRODC FILE.csv")


def main():
    print(f"Generate dummy data → {INPUT_DIR}")
    build_revenue_file()
    build_target_historical("data target.xlsx",     factor=0.95)  # target ~95% actual
    build_target_historical("data historical.xlsx", factor=0.88)  # YoY base ~88%
    build_collection_data()
    build_sales_csv()
    print("Selesai. 9 file dummy siap di data/input/.")


if __name__ == "__main__":
    main()
