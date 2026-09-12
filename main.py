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
        default="companies_nse.txt",
        help="Input company list for NSE"
    )
    parser.add_argument(
        "--bse-input",
        default="companies_bse.txt",
        help="Input company list for BSE"
    )
    parser.add_argument(
        "--nse-output",
        default="merged_nse_dividend_data.csv",
        help="Output CSV file for NSE"
    )
    parser.add_argument(
        "--bse-output",
        default="merged_bse_dividend_data.csv",
        help="Output CSV file for BSE"
    )
    parser.add_argument(
        "-w", "--workers",
        type=int,
        default=4,
        help="Number of worker threads per exchange"
    )

    args = parser.parse_args()
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
