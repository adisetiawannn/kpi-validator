# main.py
# Entry point KPI Validator — [nama perusahaan/divisi/unit etc]
# Jalankan: python main.py --month 6 --year 2026

import argparse
from pathlib import Path
from config import MONTH_NUMBER_MAP, SOURCE_REVENUE, SOURCE_TARGET, SOURCE_HISTORICAL

from src.processors.segments.sales_processor import calculate_sales_connectivity
from src.processors.aggregator import aggregate_revenue, aggregate_ngtma, aggregate_sales_connectivity, print_aggregated_result
from src.processors.segments.collection_processor import build_collection_results

def parse_args():
    parser = argparse.ArgumentParser(description="KPI Validator — Automated KPI Report Generator")
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--year",  type=int, required=True)
    parser.add_argument("--revenue",    type=str, default=f"data/input/{SOURCE_REVENUE}")
    parser.add_argument("--target",     type=str, default=f"data/input/{SOURCE_TARGET}")
    parser.add_argument("--historical", type=str, default=f"data/input/{SOURCE_HISTORICAL}")
    return parser.parse_args()


def _ask_yes_no(prompt):
    while True:
        raw = input(prompt).strip().lower()
        if raw in ["y","yes"]: return True
        if raw in ["n","no"]:  return False
        print("  ⚠️  Ketik 'y' atau 'n'.")


def main():
    args         = parse_args()
    report_month = MONTH_NUMBER_MAP[args.month]

    print(f"\n{'='*60}")
    print(f"  KPI Validator — {report_month} {args.year}")
    print(f"{'='*60}")

    # ── Step 1: Validasi file ─────────────────────────────────────────────────
    print(f"\n[Step 1/5] Validasi file source...")
    missing = [p for p in [args.revenue, args.target, args.historical]
               if not Path(p).exists()]
    if missing:
        print("  ✗ File tidak ditemukan:")
        for f in missing: print(f"    - {f}")
        return
    print(f"  ✓ Semua file ditemukan")

    # ── Step 2: Konfirmasi input ──────────────────────────────────────────────
    print(f"\n[Step 2/5] Konfirmasi data...")
    from src.processors.input_handler import get_report_inputs
    inputs    = get_report_inputs(report_month)
    territory = inputs["territory"]

    # ── Step 3: Proses Revenue ────────────────────────────────────────────────
    print(f"\n[Step 3/5] Memproses data...\n")

    from src.processors.segments.bs_processor  import calculate_bs,  print_bs_result
    from src.processors.segments.gs_processor  import calculate_gs,  print_gs_result
    from src.processors.segments.dss_processor import calculate_dss, print_dss_result
    from src.processors.segments.dps_processor import calculate_dps, print_dps_result
    from src.processors.segments.ngtma_processor import calculate_ngtma, print_ngtma_result
    from src.processors.aggregator import aggregate_revenue, aggregate_ngtma, print_aggregated_result

    BASE = dict(
        revenue_file    = args.revenue,
        target_file     = args.target,
        historical_file = args.historical,
        territory       = territory,
        year            = args.year,
        report_month    = report_month,
    )

    # Revenue
    rev = inputs["revenue"]
    r_bs  = calculate_bs (**BASE, valid_until=rev["SEG_A"]["valid_until"],
                          scaling_map=rev["SEG_A"]["scaling_map"])
    r_gs  = calculate_gs (**BASE, valid_until=rev["SEG_B"]["valid_until"],
                          scaling_map=rev["SEG_B"]["scaling_map"])
    r_dss = calculate_dss(**BASE, valid_until=rev["SEG_C"]["valid_until"],
                          scaling_map=rev["SEG_C"]["scaling_map"])
    r_dps = calculate_dps(**BASE, valid_until=rev["SEG_D"]["valid_until"],
                          scaling_map=rev["SEG_D"]["scaling_map"])

    agg_rev = aggregate_revenue(
        results_bs=r_bs, results_gs=r_gs,
        results_dss=r_dss, results_dps=r_dps,
        target_file=args.target, historical_file=args.historical,
        territory=territory, report_month=report_month,
    )

    # GROWTH
    ngtma_valid = inputs["growth"]["valid_until"]
    n_bs  = calculate_ngtma(**BASE, segment="SEG_A", valid_until=ngtma_valid)
    n_gs  = calculate_ngtma(**BASE, segment="SEG_B", valid_until=ngtma_valid)
    n_dss = calculate_ngtma(**BASE, segment="SEG_C", valid_until=ngtma_valid)
    n_dps = calculate_ngtma(**BASE, segment="SEG_D", valid_until=ngtma_valid)

    agg_ngtma = aggregate_ngtma(
        results_ngtma_bs=n_bs, results_ngtma_gs=n_gs,
        results_ngtma_dss=n_dss, results_ngtma_dps=n_dps,
        target_file=args.target, historical_file=args.historical,
        territory=territory, report_month=report_month,
    )

    # Sales Connectivity
    sc_valid = inputs["sales_connectivity"]["valid_until"]
    sc_bs  = calculate_sales_connectivity(**dict(
        territory=territory, period=args.year * 100 + args.month, year=args.year,
        report_month=report_month, valid_until=sc_valid, segment="SEG_A",
        path_target=args.target, path_historical=args.historical
        ,))

    sc_gs  = calculate_sales_connectivity(**dict(
        territory=territory, period=args.year * 100 + args.month, year=args.year,
        report_month=report_month, valid_until=sc_valid, segment="SEG_B",
        path_target=args.target, path_historical=args.historical
        ,))

    sc_dss = calculate_sales_connectivity(**dict(
        territory=territory, period=args.year * 100 + args.month, year=args.year,
        report_month=report_month, valid_until=sc_valid, segment="SEG_C",
        path_target=args.target, path_historical=args.historical
        ,))

    sc_dps = calculate_sales_connectivity(**dict(
        territory=territory, period=args.year * 100 + args.month, year=args.year,
        report_month=report_month, valid_until=sc_valid, segment="SEG_D",
        path_target=args.target, path_historical=args.historical
        ,))

    agg_sc = aggregate_sales_connectivity(
        results_bs=sc_bs, results_gs=sc_gs,
        results_dss=sc_dss, results_dps=sc_dps,
        target_file=args.target, historical_file=args.historical,
        territory=territory, report_month=report_month,
    )

    # Collection
    collection_results = build_collection_results(
        territory    = territory,
        period       = args.year * 100 + args.month,
        report_month = args.year * 100 + args.month,
        filepath     = args.target,
        )


    # ── Step 4: Preview terminal ──────────────────────────────────────────────
    print(f"\n[Step 4/5] Preview — {report_month} {args.year} | {territory}\n")

    print("── REVENUE ──────────────────────────────────────────────────")
    print_bs_result(r_bs)
    print_gs_result(r_gs)
    print_dss_result(r_dss)
    print_dps_result(r_dps)
    print_aggregated_result(agg_rev["bsgs"])
    print_aggregated_result(agg_rev["total"])

    print("── GROWTH ────────────────────────────────────────────────────")
    print_ngtma_result(n_bs)
    print_ngtma_result(n_gs)
    print_ngtma_result(n_dss)
    print_ngtma_result(n_dps)
    print_aggregated_result(agg_ngtma["bsgs"])
    print_aggregated_result(agg_ngtma["total"])

    print("── SALES CONNECTIVITY ───────────────────────────────────")
    # print per segment dan agregasi


    # ── Step 5: Export Excel ──────────────────────────────────────────────────
    print(f"\n[Step 5/5] Export Excel")
    print(f"  Output: data/output/kpi_summary_{territory}_{report_month}_{args.year}.xlsx")

    if not _ask_yes_no("\n  Generate Excel sekarang? (y/n) : "):
        print("\n  Export dibatalkan.")
        print(f"{'='*60}\n")
        return


    # EXPORT DATA TO EXCEL
    from src.exporters.excel_exporter import export_to_excel
    filepath = export_to_excel(
        r_bs=r_bs, r_gs=r_gs, r_dss=r_dss, r_dps=r_dps,
        agg_bsgs=agg_rev["bsgs"], agg_total=agg_rev["total"],
        n_bs=n_bs, n_gs=n_gs, n_dss=n_dss, n_dps=n_dps,
        agg_ngtma_bsgs=agg_ngtma["bsgs"], agg_ngtma_total=agg_ngtma["total"],
        sc_bs=sc_bs, sc_gs=sc_gs, sc_dss=sc_dss, sc_dps=sc_dps,  # ← tambah
        agg_sc=agg_sc,                                              # ← tambah
        report_month=report_month, year=args.year,
        territory=territory, historical_file=args.historical,
        collection_results=collection_results,
        output_dir="data/output",
        )

    print(f"\n{'='*60}")
    print(f"  ✓ Selesai — {report_month} {args.year} | {territory}")
    print(f"  ✓ File: {filepath}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
