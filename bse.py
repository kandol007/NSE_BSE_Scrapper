#!/usr/bin/env python3
"""
bse.py - High-Performance BSE Dividend & Corporate Actions Scraper
Fetches corporate actions and dividend records from BSE India in seconds.
"""

import os
import sys
import time
import argparse
from scraper_core import BSEScraper, read_companies_from_file, save_merged_csv, get_base_path


def main():
    parser = argparse.ArgumentParser(description="High-Speed BSE Dividend & Corporate Actions Scraper")
    parser.add_argument(
        "-i", "--input",
        default="companies_bse.txt",
        help="Path to input file containing company name and scrip code (default: companies_bse.txt)"
    )
    parser.add_argument(
        "-o", "--output",
        default="merged_bse_dividend_data.csv",
        help="Path to output merged CSV file (default: merged_bse_dividend_data.csv)"
    )
    parser.add_argument(
        "-w", "--workers",
        type=int,
        default=6,
        help="Number of concurrent worker threads (default: 6)"
    )
    parser.add_argument(
        "-c", "--code",
        help="Scrape a single BSE scrip code directly (e.g. 500325)"
    )
    args = parser.parse_args()

    base_path = get_base_path()
    input_file = args.input if os.path.isabs(args.input) else os.path.join(base_path, args.input)
    output_file = args.output if os.path.isabs(args.output) else os.path.join(base_path, args.output)

    print("=" * 60)
    print("⚡ High-Speed BSE Dividend Scraper (API Engine)")
    print("=" * 60)

    if args.code:
        companies = {f"Scrip_{args.code}": str(args.code)}
    else:
        companies = read_companies_from_file(input_file)
        if not companies:
            print(f"❌ No companies found in '{input_file}'. Exiting.")
            sys.exit(1)

    print(f"📋 Loaded {len(companies)} company/companies to process.")
    start_time = time.time()

    scraper = BSEScraper()
    df = scraper.scrape_all(companies, max_workers=args.workers, show_progress=True)

    elapsed = time.time() - start_time
    print("-" * 60)
    if not df.empty:
        save_merged_csv(df, output_file)
        print(f"✅ Scraping completed in {elapsed:.2f}s! Total records: {len(df)}")
    else:
        print(f"⚠️ Scraping finished in {elapsed:.2f}s, but no records were found.")
    print("=" * 60)


if __name__ == "__main__":
    main()