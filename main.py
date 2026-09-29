#!/usr/bin/env python3
"""
main.py - Unified Entry Point for NSE & BSE Scrapers
Scrape corporate actions and dividend records from NSE, BSE, or both concurrently.
"""

import os
import sys
import time
import argparse
from concurrent.futures import ThreadPoolExecutor
from scraper_core import (
    NSEScraper,
    BSEScraper,
    read_companies_from_file,
    save_merged_csv,
    get_base_path,
)


def run_nse(input_file: str, output_file: str, workers: int):
    print("\n--- 🚀 Running NSE Scraper ---")
    companies = read_companies_from_file(input_file)
    if not companies:
        print(f"❌ No NSE companies found in '{input_file}'.")
        return None
    scraper = NSEScraper()
    df = scraper.scrape_all(companies, max_workers=workers, show_progress=True)
    if not df.empty:
        save_merged_csv(df, output_file)
    return df


def run_bse(input_file: str, output_file: str, workers: int):
    print("\n--- 🚀 Running BSE Scraper ---")
    companies = read_companies_from_file(input_file)
    if not companies:
        print(f"❌ No BSE companies found in '{input_file}'.")
        return None
    scraper = BSEScraper()
    df = scraper.scrape_all(companies, max_workers=workers, show_progress=True)
    if not df.empty:
        save_merged_csv(df, output_file)
    return df


def main():
    parser = argparse.ArgumentParser(
        description="Unified NSE & BSE Dividend Data Downloader",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument(
        "-e", "--exchange",
        choices=["all", "nse", "bse"],
        default="all",
        help="Which exchange to scrape ('nse', 'bse', or 'all')"
    )
    parser.add_argument(
        "--all",
        action="store_const",
        const="all",
        dest="exchange",
        help="Convenience alias to scrape both exchanges"
    )
    parser.add_argument(
        "--nse-input",
        default="input/companies_nse.txt",
        help="Input company list for NSE (default: input/companies_nse.txt)"
    )
    parser.add_argument(
        "--bse-input",
        default="input/companies_bse.txt",
        help="Input company list for BSE (default: input/companies_bse.txt)"
    )
    parser.add_argument(
        "--nse-output",
        default="output/merged_nse_dividend_data.csv",
        help="Output CSV file for NSE (default: output/merged_nse_dividend_data.csv)"
    )
    parser.add_argument(
        "--bse-output",
        default="output/merged_bse_dividend_data.csv",
        help="Output CSV file for BSE (default: output/merged_bse_dividend_data.csv)"
    )
    parser.add_argument(
        "-w", "--workers",
        type=int,
        default=4,
        help="Number of worker threads per exchange"
    )
    parser.add_argument(
        "--export-icici-tata",
        action="store_true",
        help="Pull latest NSE/BSE lists, verify links, and export ICICI & Tata Excel"
    )
    parser.add_argument(
        "--export-dividends",
        action="store_true",
        help="Scrape all dividend & corporate action records for Tata & ICICI companies into an Excel workbook"
    )
    parser.add_argument(
        "--export-unclaimed-shareholders",
        action="store_true",
        help="Crawl and export shareholder-level unclaimed dividend records with names for Tata & ICICI companies"
    )
    parser.add_argument(
        "--export-all",
        action="store_true",
        help="Run all 3 extraction pipelines (Verified Companies, Corporate Dividends, Shareholder Unclaimed Dividends)"
    )

    args = parser.parse_args()

    if args.export_all:
        from pull_and_verify_companies import run_pipeline
        from download_dividend_data_xlsx import main as run_dividend_export
        from scrape_shareholder_unclaimed_dividends import main as run_shareholder_export

        print("\n🚀 [1/3] Running Output Type 1: Verified Companies Pipeline...")
        run_pipeline(workers=args.workers if args.workers > 4 else 8)
        print("\n🚀 [2/3] Running Output Type 2: Corporate Dividend Declarations Pipeline...")
        run_dividend_export()
        print("\n🚀 [3/3] Running Output Type 3: Shareholder Unclaimed Dividends Pipeline...")
        run_shareholder_export()
        return

    if args.export_icici_tata:
        from pull_and_verify_companies import run_pipeline
        run_pipeline(workers=args.workers if args.workers > 4 else 8)
        return

    if args.export_dividends:
        from download_dividend_data_xlsx import main as run_dividend_export
        run_dividend_export()
        return

    if args.export_unclaimed_shareholders:
        from scrape_shareholder_unclaimed_dividends import main as run_shareholder_export
        run_shareholder_export()
        return
    base_path = get_base_path()

    nse_in = args.nse_input if os.path.isabs(args.nse_input) else os.path.join(base_path, args.nse_input)
    bse_in = args.bse_input if os.path.isabs(args.bse_input) else os.path.join(base_path, args.bse_input)
    nse_out = args.nse_output if os.path.isabs(args.nse_output) else os.path.join(base_path, args.nse_output)
    bse_out = args.bse_output if os.path.isabs(args.bse_output) else os.path.join(base_path, args.bse_output)

    print("=" * 65)
    print("📈 Ultra-Fast Indian Exchanges Scraper (NSE & BSE)")
    print(f"Target: {args.exchange.upper()} | Workers: {args.workers}")
    print("=" * 65)

    overall_start = time.time()
    nse_records = 0
    bse_records = 0

    if args.exchange == "all":
        # Run both exchanges concurrently in background threads
        with ThreadPoolExecutor(max_workers=2) as executor:
            fut_nse = executor.submit(run_nse, nse_in, nse_out, args.workers)
            fut_bse = executor.submit(run_bse, bse_in, bse_out, args.workers)

            df_nse = fut_nse.result()
            df_bse = fut_bse.result()

            nse_records = len(df_nse) if df_nse is not None and not df_nse.empty else 0
            bse_records = len(df_bse) if df_bse is not None and not df_bse.empty else 0

    elif args.exchange == "nse":
        df_nse = run_nse(nse_in, nse_out, args.workers)
        nse_records = len(df_nse) if df_nse is not None and not df_nse.empty else 0

    elif args.exchange == "bse":
        df_bse = run_bse(bse_in, bse_out, args.workers)
        bse_records = len(df_bse) if df_bse is not None and not df_bse.empty else 0

    total_time = time.time() - overall_start
    print("\n" + "=" * 65)
    print(f"✨ Finished in {total_time:.2f} seconds!")
    if args.exchange in ("all", "nse"):
        print(f"   • NSE records extracted: {nse_records}")
    if args.exchange in ("all", "bse"):
        print(f"   • BSE records extracted: {bse_records}")
    print("=" * 65)


if __name__ == "__main__":
    main()
