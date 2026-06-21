# src/processors/input_handler.py
# Menangani semua interactive input dari user.

from config import MONTH_COLS_MTD, MONTH_NUMBER_MAP

VALID_MONTHS = set(MONTH_COLS_MTD)

SEGMENT_CONFIG = {
    "SEG_A": {"code": "CODE_A", "scaling_label": "SCALING"},
    "SEG_B": {"code": "CODE_B", "scaling_label": "SCALING"},
    "SEG_C": {"code": "CODE_C", "scaling_label": "SCALING"},
    "SEG_D": {"code": "CODE_D", "scaling_label": "SCALING"},
}


# ── Input primitives ──────────────────────────────────────────────────────────

def _month_idx(month: str) -> int:
    return {v: k for k, v in MONTH_NUMBER_MAP.items()}[month.upper()]

def _ask_month(prompt: str) -> str:
    while True:
        raw = input(prompt).strip().upper()
        if raw in VALID_MONTHS:
            return raw
        print("  ⚠️  Format tidak valid. Gunakan: JAN/FEB/MAR/APR/MAY/JUN/JUL/AUG/SEP/OCT/NOV/DEC")

def _ask_scaling(prompt: str) -> float:
    while True:
        raw = input(prompt).strip()
        if raw == "": return 0.0
        try:
            return float(raw.replace(".", "").replace(",", ""))
        except ValueError:
            print("  ⚠️  Input tidak valid. Masukkan angka (contoh: 150000000)")

def _ask_territory() -> str:
    while True:
        raw = input("  Territory? (contoh: AREA_01) : ").strip().upper()
        if raw: return raw
        print("  ⚠️  Territory tidak boleh kosong.")

def _ask_yes_no(prompt: str) -> bool:
    while True:
        raw = input(prompt).strip().lower()
        if raw in ["y", "yes"]: return True
        if raw in ["n", "no"]:  return False
        print("  ⚠️  Ketik 'y' untuk ya atau 'n' untuk tidak.")


# ── Revenue segment confirmation ──────────────────────────────────────────────

def _confirm_revenue_segment(segment: str, report_month: str) -> dict:
    """
    Konfirmasi cut off + scaling per bulan untuk satu segment Revenue.
    Loop sampai user konfirmasi benar.
    """
    cfg          = SEGMENT_CONFIG[segment]
    code         = cfg["code"]
    scl_label    = cfg["scaling_label"]
    report_month = report_month.upper()
    report_idx   = _month_idx(report_month)

    while True:
        print(f"\n  {'─'*56}")
        print(f"  Konfirmasi Revenue Segment {segment} ({code})")
        print(f"  {'─'*56}")

        valid_until = _ask_month(
            f"  Revenue {segment} valid sampai bulan apa?\n"
            f"  (contoh: APR) : "
        )
        valid_idx = _month_idx(valid_until)

        # Deteksi bulan outlook otomatis
        outlook_months = [
            MONTH_NUMBER_MAP[idx]
            for idx in range(valid_idx + 1, report_idx)
        ]

        # Tanya scaling per bulan outlook
        scaling_map = {}
        if outlook_months:
            print(f"\n  Perlu hitung outlook untuk: {', '.join(outlook_months)}")
            for month in outlook_months:
                scaling_map[month] = _ask_scaling(
                    f"  {scl_label} {segment} untuk {month} (Rp)\n"
                    f"  (Enter jika 0) : "
                )
        else:
            print(f"\n  Data sudah valid — tidak perlu outlook.")

        # Ringkasan
        print(f"\n  {'─'*40}")
        print(f"  Ringkasan Revenue {segment}:")
        print(f"  {'─'*40}")
        print(f"  Valid s/d : {valid_until}")
        if outlook_months:
            for month in outlook_months:
                print(f"  {scl_label} {month:<4}: Rp {scaling_map[month]:>20,.0f}")
        else:
            print(f"  Outlook   : Tidak perlu")
        print(f"  {'─'*40}")

        if _ask_yes_no("\n  Sudah benar? (y/n) : "):
            return {"valid_until": valid_until, "scaling_map": scaling_map}

        print(f"\n  ↩️  Input diulang untuk Revenue {segment}...\n")


# ── GROWTH confirmation ────────────────────────────────────────────────────────

def _confirm_ngtma(report_month: str) -> dict:
    """
    Konfirmasi cut off GROWTH — satu nilai berlaku untuk semua segment.
    Loop sampai user konfirmasi benar.

    Returns:
        { "valid_until": str }
    """
    report_month = report_month.upper()

    while True:
        print(f"\n  {'─'*56}")
        print(f"  Konfirmasi GROWTH (berlaku semua segment)")
        print(f"  {'─'*56}")

        valid_until = _ask_month(
            f"  GROWTH valid sampai bulan apa?\n"
            f"  (contoh: APR) : "
        )

        # Ringkasan
        print(f"\n  {'─'*40}")
        print(f"  Ringkasan GROWTH:")
        print(f"  {'─'*40}")
        print(f"  Valid s/d      : {valid_until}")
        print(f"  Berlaku untuk  : SEG_A, SEG_B, SEG_C, SEG_D")
        print(f"  Outlook        : Diambil dari Excel langsung")
        print(f"  {'─'*40}")

        if _ask_yes_no("\n  Sudah benar? (y/n) : "):
            return {"valid_until": valid_until}

        print(f"\n  ↩️  Input diulang untuk GROWTH...\n")


# ── Sales Connectivity confirmation ────────────────────────────────────────────────────────

def _confirm_sales_connectivity(report_month: str) -> dict:
    """
    Konfirmasi cut off Sales Connectivity — satu nilai berlaku untuk semua segment.
    """
    report_month = report_month.upper()

    while True:
        print(f"\n  {'─'*56}")
        print(f"  Konfirmasi Sales Connectivity (berlaku semua segment)")
        print(f"  {'─'*56}")

        valid_until = _ask_month(
            f"  Sales Connectivity valid sampai bulan apa?\n"
            f"  (contoh: APR) : "
        )

        print(f"\n  {'─'*40}")
        print(f"  Ringkasan Sales Connectivity:")
        print(f"  {'─'*40}")
        print(f"  Valid s/d      : {valid_until}")
        print(f"  Sub-indicators : PROD_A, PROD_B, PROD_C")
        print(f"  Berlaku untuk  : SEG_A, SEG_B, SEG_C, SEG_D")
        print(f"  {'─'*40}")

        if _ask_yes_no("\n  Sudah benar? (y/n) : "):
            return {"valid_until": valid_until}

        print(f"\n  ↩️  Input diulang untuk Sales Connectivity...\n")


# ── Collection confirmation ────────────────────────────────────────────────────────
def _confirm_collection(report_month: str) -> dict:
    report_month = report_month.upper()

    while True:
        print(f"\n  {'─'*56}")
        print(f"  Konfirmasi Collection Performance (berlaku semua segment)")
        print(f"  {'─'*56}")

        valid_until = _ask_month(
            f"  Collection valid sampai bulan apa?\n"
            f"  (contoh: APR) : "
        )

        print(f"\n  {'─'*40}")
        print(f"  Ringkasan Collection Performance:")
        print(f"  {'─'*40}")
        print(f"  Valid s/d      : {valid_until}")
        print(f"  Indikator      : COLL_A, COLL_B, COLL_C, COLL_D")
        print(f"  Berlaku untuk  : SEG_A, SEG_B, SEG_C, SEG_D, TOTAL")
        print(f"  {'─'*40}")

        if _ask_yes_no("\n  Sudah benar? (y/n) : "):
            return {"valid_until": valid_until}

        print(f"\n  ↩️  Input diulang untuk Collection...\n")


# ── Main entry point ──────────────────────────────────────────────────────────

def get_report_inputs(report_month: str) -> dict:
    """
    Kumpulkan semua input untuk generate laporan.

    Returns:
        {
          "territory": str,
          "revenue": {
              "SEG_A": {"valid_until": str, "scaling_map": dict},
              "SEG_B": {...},
              "SEG_C": {...},
              "SEG_D": {...},
          },
          "growth": {
              "valid_until": str,  ← berlaku untuk semua segment
          }
        }
    """
    print(f"\n{'─'*60}")
    print(f"  [Konfigurasi Laporan]")
    print(f"{'─'*60}")

    territory = _ask_territory()

    # ── Revenue per segment ───────────────────────────────────────────────────
    print(f"\n{'─'*60}")
    print(f"  [Revenue — Konfirmasi per Segment]")
    print(f"{'─'*60}")

    revenue_inputs = {}
    for segment in ["SEG_A", "SEG_B", "SEG_C", "SEG_D"]:
        revenue_inputs[segment] = _confirm_revenue_segment(segment, report_month)

    # ── GROWTH ─────────────────────────────────────────────────────────────────
    print(f"\n{'─'*60}")
    print(f"  [GROWTH — Konfirmasi Cut Off]")
    print(f"{'─'*60}")

    ngtma_input = _confirm_ngtma(report_month)

    # ── Sales Connectivity ─────────────────────────────────────────────────────────────────
    print(f"\n{'─'*60}")
    print(f"  [Sales Connectivity — Konfirmasi Cut Off]")
    print(f"{'─'*60}")

    sc_input = _confirm_sales_connectivity(report_month)

    # ── Collection ─────────────────────────────────────────────────────────────────
    print(f"\n{'─'*60}")
    print(f"  [Collection — Konfirmasi Cut Off]")
    print(f"{'─'*60}")

    collection_input = _confirm_collection(report_month)

  # ── Konfirmasi Akhir ─────────────────────────────────────────────────────────────────

    print(f"\n{'='*60}")
    print(f"  ✓ Semua input confirmed")
    print(f"{'='*60}\n")

    return {
    "territory"          : territory,
    "revenue"            : revenue_inputs,
    "growth"             : ngtma_input,
    "sales_connectivity" : sc_input, # ← tambah ini
    "collection"         : collection_input  # ← tambah ini
    }
