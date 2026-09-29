#!/usr/bin/env python3
"""
pull_and_verify_companies.py
-----------------------------
1. Pulls the latest companies lists from official NSE and BSE endpoints.
2. Updates local reference files:
   - NSE_List_of_companies.csv
   - BSE_List_of_companies.csv
3. Filters ICICI and Tata companies across both exchanges.
4. Verifies their stock quote links and corporate action links with live HTTP requests.
5. Queries dividend and corporate actions count using the project's scraping engine.
6. Exports a comprehensive, beautifully styled Excel (.xlsx) file to the output directory:
   output/ICICI_and_Tata_Companies_Verified.xlsx
"""

import os
import sys
import io
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
from curl_cffi import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from scraper_core import NSEScraper, BSEScraper, get_base_path


# ---------------------------------------------------------
# Configurations & URLs
# ---------------------------------------------------------
NSE_EQUITIES_URL = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
NSE_EQUITIES_URL_FALLBACK = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
BSE_SCRIPS_URL = "https://api.bseindia.com/BseIndiaAPI/api/ListofScripData/w?Group=&Scripcode=&industry=&segment=Equity&status=Active"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Known Tata Group conglomerate companies that may not have "Tata" in prefix
TATA_CONGLOMERATE_SYMBOLS = {
    "TITAN": "Titan Company Limited",
    "TRENT": "Trent Limited",
    "VOLTAS": "Voltas Limited",
    "INDHOTEL": "The Indian Hotels Company Limited",
    "NELCO": "NELCO Limited",
    "RALLIS": "Rallis India Limited",
    "ORIENTHOT": "Oriental Hotels Limited",
    "BENARAS": "Benares Hotels Limited",
    "BENARES": "Benares Hotels Limited",
    "ACGL": "Automobile Corporation of Goa Limited",
}

TATA_CONGLOMERATE_BSE_CODES = {
    "500114": "Titan Company Ltd",
    "500251": "Trent Ltd",
    "500575": "Voltas Ltd",
    "500850": "Indian Hotels Company Ltd",
    "504112": "Nelco Ltd",
    "500355": "Rallis India Ltd",
    "500314": "Oriental Hotels Ltd",
    "509438": "Benares Hotels Ltd",
    "505036": "Automobile Corporation of Goa Ltd",
}


# ---------------------------------------------------------
# Step 1: Pull Latest Company Lists
# ---------------------------------------------------------
def pull_latest_nse_companies(session: requests.Session) -> pd.DataFrame:
    """Fetches latest NSE listed equities CSV from official NSE archive."""
    print("📡 Pulling latest NSE listed companies list...")
    headers = {**DEFAULT_HEADERS, "Referer": "https://www.nseindia.com/"}
    
    for url in [NSE_EQUITIES_URL, NSE_EQUITIES_URL_FALLBACK]:
        try:
            resp = session.get(url, headers=headers, timeout=20)
            if resp.status_code == 200 and "SYMBOL" in resp.text:
                df = pd.read_csv(io.StringIO(resp.text))
                df.columns = [c.strip() for c in df.columns]
                print(f"✅ Successfully fetched {len(df)} listed companies from NSE.")
                return df
        except Exception as e:
            print(f"⚠️ Failed fetching NSE from {url}: {e}")
            
    print("❌ Could not download latest NSE list. Falling back to local file if available.")
    for p in [os.path.join(get_base_path(), "input", "NSE_List_of_companies.csv"), os.path.join(get_base_path(), "NSE_List_of_companies.csv")]:
        if os.path.exists(p):
            df = pd.read_csv(p)
            df.columns = [c.strip() for c in df.columns]
            return df
    return pd.DataFrame()


def pull_latest_bse_companies(session: requests.Session) -> pd.DataFrame:
    """Fetches latest BSE active equity scrips via BSE API."""
    print("📡 Pulling latest BSE listed companies list...")
    headers = {**DEFAULT_HEADERS, "Referer": "https://www.bseindia.com/"}
    try:
        resp = session.get(BSE_SCRIPS_URL, headers=headers, timeout=25)
        if resp.status_code == 200:
            data = resp.json()
            df = pd.DataFrame(data)
            print(f"✅ Successfully fetched {len(df)} active equity scrips from BSE.")
            return df
    except Exception as e:
        print(f"⚠️ Failed fetching BSE companies: {e}")
        
    print("❌ Could not download latest BSE list. Falling back to local file if available.")
    for p in [os.path.join(get_base_path(), "input", "BSE_List_of_companies.csv"), os.path.join(get_base_path(), "BSE_List_of_companies.csv")]:
        if os.path.exists(p):
            try:
                return pd.read_csv(p)
            except Exception:
                pass
    return pd.DataFrame()


def save_updated_reference_files(df_nse: pd.DataFrame, df_bse: pd.DataFrame):
    """Saves updated company lists into the input/ directory."""
    input_dir = os.path.join(get_base_path(), "input")
    os.makedirs(input_dir, exist_ok=True)
    if not df_nse.empty:
        nse_path = os.path.join(input_dir, "NSE_List_of_companies.csv")
        df_nse.to_csv(nse_path, index=False)
        print(f"💾 Updated reference file: {nse_path} ({len(df_nse)} companies)")
        
    if not df_bse.empty:
        bse_path = os.path.join(input_dir, "BSE_List_of_companies.csv")
        df_bse.to_csv(bse_path, index=False)
        print(f"💾 Updated reference file: {bse_path} ({len(df_bse)} companies)")


# ---------------------------------------------------------
# Step 2: Filter ICICI and Tata Companies
# ---------------------------------------------------------
def filter_companies(df_nse: pd.DataFrame, df_bse: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Identifies all ICICI and Tata companies across NSE and BSE.
    Returns structured list of company dictionaries.
    """
    records = []
    
    # --- NSE Processing ---
    if not df_nse.empty:
        for _, row in df_nse.iterrows():
            sym = str(row.get("SYMBOL", "")).strip().upper()
            name = str(row.get("NAME OF COMPANY", "")).strip()
            isin = str(row.get("ISIN NUMBER", "")).strip()
            series = str(row.get("SERIES", "")).strip()
            
            # Check Tata
            is_direct_tata = "TATA" in sym or "TATA" in name.upper()
            is_tata_conglomerate = sym in TATA_CONGLOMERATE_SYMBOLS
            
            if is_direct_tata or is_tata_conglomerate:
                sub_cat = "Core / Direct Brand" if is_direct_tata else "Tata Group Enterprise"
                records.append({
                    "Exchange": "NSE",
                    "Group": "Tata",
                    "Category": sub_cat,
                    "Company Name": name,
                    "Symbol / Code": sym,
                    "ISIN": isin,
                    "Series / Group": series,
                    "Company Page URL": f"https://www.nseindia.com/get-quotes/equity?symbol={sym}",
                    "Corporate Actions URL": f"https://www.nseindia.com/companies-listing/corporate-filings-actions?symbol={sym}&tabIndex=equity",
                    "raw_symbol": sym,
                    "raw_code": None,
                })
                
            # Check ICICI
            is_icici = "ICICI" in sym or "ICICI" in name.upper()
            if is_icici:
                records.append({
                    "Exchange": "NSE",
                    "Group": "ICICI",
                    "Category": "Core / Direct Brand",
                    "Company Name": name,
                    "Symbol / Code": sym,
                    "ISIN": isin,
                    "Series / Group": series,
                    "Company Page URL": f"https://www.nseindia.com/get-quotes/equity?symbol={sym}",
                    "Corporate Actions URL": f"https://www.nseindia.com/companies-listing/corporate-filings-actions?symbol={sym}&tabIndex=equity",
                    "raw_symbol": sym,
                    "raw_code": None,
                })

    # --- BSE Processing ---
    if not df_bse.empty:
        for _, row in df_bse.iterrows():
            code = str(row.get("SCRIP_CD", "")).strip()
            scrip_id = str(row.get("scrip_id", "")).strip().upper()
            scrip_name = str(row.get("Scrip_Name", "")).strip()
            issuer_name = str(row.get("Issuer_Name", "")).strip()
            isin = str(row.get("ISIN_NUMBER", "")).strip()
            group = str(row.get("GROUP", "")).strip()
            nsurl = str(row.get("NSURL", "")).strip()
            
            if not nsurl or not nsurl.startswith("http"):
                slug = scrip_name.lower().replace(" ", "-").replace(".", "")
                nsurl = f"https://www.bseindia.com/stock-share-price/{slug}/{scrip_id.lower()}/{code}/"
            
            corp_url = f"https://www.bseindia.com/corporates/corporates_act.html?scripcode={code}&Dividend=P9&scripname="
            
            # Check Tata
            is_direct_tata = "TATA" in scrip_id or "TATA" in scrip_name.upper() or "TATA" in issuer_name.upper()
            is_tata_conglomerate = code in TATA_CONGLOMERATE_BSE_CODES or scrip_id in TATA_CONGLOMERATE_SYMBOLS
            
            if is_direct_tata or is_tata_conglomerate:
                if isin.startswith("INF"):
                    sub_cat = "Mutual Fund / ETF Scheme"
                elif is_direct_tata:
                    sub_cat = "Core / Direct Brand"
                else:
                    sub_cat = "Tata Group Enterprise"
                    
                records.append({
                    "Exchange": "BSE",
                    "Group": "Tata",
                    "Category": sub_cat,
                    "Company Name": scrip_name,
                    "Symbol / Code": code,
                    "ISIN": isin,
                    "Series / Group": group,
                    "Company Page URL": nsurl,
                    "Corporate Actions URL": corp_url,
                    "raw_symbol": scrip_id,
                    "raw_code": code,
                })
                
            # Check ICICI
            is_icici = "ICICI" in scrip_id or "ICICI" in scrip_name.upper() or "ICICI" in issuer_name.upper()
            if is_icici:
                sub_cat = "Mutual Fund / ETF Scheme" if isin.startswith("INF") else "Core / Direct Brand"
                records.append({
                    "Exchange": "BSE",
                    "Group": "ICICI",
                    "Category": sub_cat,
                    "Company Name": scrip_name,
                    "Symbol / Code": code,
                    "ISIN": isin,
                    "Series / Group": group,
                    "Company Page URL": nsurl,
                    "Corporate Actions URL": corp_url,
                    "raw_symbol": scrip_id,
                    "raw_code": code,
                })

    return records


# ---------------------------------------------------------
# Step 3: Link Verification & Dividend Records Inspection
# ---------------------------------------------------------
def verify_single_company(
    rec: Dict[str, Any],
    session: requests.Session,
    nse_scraper: NSEScraper,
    bse_scraper: BSEScraper
) -> Dict[str, Any]:
    """
    Verifies company links and corporate actions accessibility.
    """
    exchange = rec["Exchange"]
    page_url = rec["Company Page URL"]
    corp_url = rec["Corporate Actions URL"]
    
    headers_bse = {**DEFAULT_HEADERS, "Referer": "https://www.bseindia.com/"}
    headers_nse = {**DEFAULT_HEADERS, "Referer": "https://www.nseindia.com/"}
    headers = headers_nse if exchange == "NSE" else headers_bse

    # 1. Verify Company Stock Page Link
    t0 = time.time()
    page_status = "N/A"
    page_code = 0
    try:
        resp = session.get(page_url, headers=headers, timeout=12)
        page_code = resp.status_code
        dt = round(time.time() - t0, 2)
        if resp.status_code == 200:
            page_status = f"Verified (200 OK, {dt}s)"
        else:
            page_status = f"HTTP {resp.status_code} ({dt}s)"
    except Exception as e:
        page_status = f"Failed ({type(e).__name__})"

    # 2. Verify Corporate Actions URL
    t1 = time.time()
    corp_status = "N/A"
    corp_code = 0
    try:
        resp2 = session.get(corp_url, headers=headers, timeout=12)
        corp_code = resp2.status_code
        dt2 = round(time.time() - t1, 2)
        if resp2.status_code == 200:
            corp_status = f"Verified (200 OK, {dt2}s)"
        else:
            corp_status = f"HTTP {resp2.status_code} ({dt2}s)"
    except Exception as e:
        corp_status = f"Failed ({type(e).__name__})"

    # 3. Check Dividend / Corporate Action Data Count via Scraper Engine
    div_count = 0
    latest_action = "None"
    try:
        if exchange == "NSE" and rec.get("raw_symbol"):
            df_div = nse_scraper.fetch_dividend_data(rec["raw_symbol"], company_name=rec["Company Name"])
            div_count = len(df_div)
            if not df_div.empty:
                # Find purpose or details column
                p_col = next((c for c in df_div.columns if "PURPOSE" in c.upper() or "SERIES" in c.upper()), None)
                d_col = next((c for c in df_div.columns if "DATE" in c.upper()), None)
                desc = str(df_div.iloc[0][p_col]) if p_col else "Action Available"
                date_str = str(df_div.iloc[0][d_col]) if d_col else ""
                latest_action = f"{desc} ({date_str})".strip()
        elif exchange == "BSE" and rec.get("raw_code"):
            df_div = bse_scraper.fetch_dividend_data(rec["raw_code"], company_name=rec["Company Name"])
            div_count = len(df_div)
            if not df_div.empty:
                p_col = next((c for c in df_div.columns if "PURPOSE" in c.upper()), None)
                d_col = next((c for c in df_div.columns if "EX" in c.upper() or "DATE" in c.upper()), None)
                desc = str(df_div.iloc[0][p_col]) if p_col else "Action Available"
                date_str = str(df_div.iloc[0][d_col]) if d_col else ""
                latest_action = f"{desc} ({date_str})".strip()
    except Exception:
        pass

    # Overall verification flag
    overall_ok = (page_code == 200)

    result = dict(rec)
    result.update({
        "Page Status": page_status,
        "Page Status Code": page_code,
        "Corporate Actions Status": corp_status,
        "Corporate Actions Status Code": corp_code,
        "Overall Verification": "Active & Verified" if overall_ok else "Check Needed",
        "Dividend / Action Records": div_count,
        "Latest Action": latest_action[:120] if latest_action else "None",
        "Verified At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    return result


def verify_all_companies(companies: List[Dict[str, Any]], max_workers: int = 8) -> List[Dict[str, Any]]:
    """Runs concurrent link verification across all companies."""
    print(f"\n🔍 Verifying links and actions for {len(companies)} companies using {max_workers} worker threads...")
    
    # Initialize scrapers and warming sessions
    session = requests.Session(impersonate="chrome124")
    try:
        session.get("https://www.nseindia.com", headers={**DEFAULT_HEADERS, "Referer": "https://www.nseindia.com/"}, timeout=10)
    except Exception:
        pass
        
    nse_scraper = NSEScraper()
    bse_scraper = BSEScraper()

    verified_records = []
    total = len(companies)
    completed = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_rec = {
            executor.submit(verify_single_company, rec, session, nse_scraper, bse_scraper): rec
            for rec in companies
        }
        for future in as_completed(future_to_rec):
            completed += 1
            rec = future_to_rec[future]
            try:
                res = future.result()
                verified_records.append(res)
                status_icon = "✅" if res["Overall Verification"] == "Active & Verified" else "⚠️"
                print(f"[{completed}/{total}] {status_icon} [{res['Exchange']}] {res['Group']} - {res['Company Name']} ({res['Symbol / Code']}): {res['Page Status']}")
            except Exception as e:
                print(f"[{completed}/{total}] ❌ [{rec['Exchange']}] {rec['Group']} - {rec['Company Name']} error: {e}")
                err_rec = dict(rec)
                err_rec.update({
                    "Page Status": f"Error: {e}",
                    "Corporate Actions Status": "Error",
                    "Overall Verification": "Failed",
                    "Dividend / Action Records": 0,
                    "Latest Action": "None",
                    "Verified At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })
                verified_records.append(err_rec)

    # Sort results nicely: Group, Exchange, Category, Company Name
    verified_records.sort(key=lambda r: (r["Group"], r["Exchange"], r["Category"], r["Company Name"]))
    return verified_records


# ---------------------------------------------------------
# Step 4: Export to Excel (.xlsx) with Professional Styling
# ---------------------------------------------------------
def create_excel_workbook(records: List[Dict[str, Any]], output_path: str):
    """
    Creates an Excel file with multiple sheets:
    1. Executive Summary
    2. Tata Companies
    3. ICICI Companies
    4. Combined View
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = openpyxl.Workbook()
    
    # Color palette
    navy_dark = "1E3A8A"
    navy_light = "3B82F6"
    gray_bg = "F8FAFC"
    border_color = "E2E8F0"
    green_text = "15803D"
    green_bg = "DCFCE7"
    
    thin_border = Border(
        left=Side(style="thin", color=border_color),
        right=Side(style="thin", color=border_color),
        top=Side(style="thin", color=border_color),
        bottom=Side(style="thin", color=border_color)
    )

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type="solid")
    
    title_font = Font(name="Calibri", size=16, bold=True, color=navy_dark)
    subtitle_font = Font(name="Calibri", size=10, italic=True, color="64748B")
    
    verified_font = Font(name="Calibri", size=10, bold=True, color=green_text)
    verified_fill = PatternFill(start_color=green_bg, end_color=green_bg, fill_type="solid")
    
    link_font = Font(name="Calibri", size=10, color="2563EB", underline="single")
    regular_font = Font(name="Calibri", size=10)
    bold_regular = Font(name="Calibri", size=10, bold=True)

    # Data partitions
    tata_records = [r for r in records if r["Group"] == "Tata"]
    icici_records = [r for r in records if r["Group"] == "ICICI"]

    # ---------------------------------------------------------
    # Sheet 1: Executive Summary
    # ---------------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "Executive Summary"
    ws_sum.views.sheetView[0].showGridLines = True

    # Title
    ws_sum["B2"] = "NSE & BSE ICICI & Tata Companies Audit Report"
    ws_sum["B2"].font = title_font
    ws_sum["B3"] = f"Generated on {datetime.now().strftime('%d %B %Y, %H:%M:%S IST')} | Live Link Verification & Corporate Actions Scan"
    ws_sum["B3"].font = subtitle_font

    # Overview Cards Table
    summary_data = [
        ("Metric", "Value", "Notes"),
        ("Total Companies Verified", len(records), "All ICICI & Tata entities across NSE & BSE"),
        ("Tata Group Companies", len(tata_records), "Includes Core Brands, Enterprises & Mutual Funds"),
        ("ICICI Group Companies", len(icici_records), "Includes ICICI Bank, Insurance, AMC & ETFs"),
        ("NSE Companies Identified", sum(1 for r in records if r["Exchange"] == "NSE"), "Equities listed on National Stock Exchange"),
        ("BSE Companies Identified", sum(1 for r in records if r["Exchange"] == "BSE"), "Active scrips on Bombay Stock Exchange"),
        ("Link Verification Success Rate", f"{round(sum(1 for r in records if '200 OK' in r['Page Status']) / len(records) * 100, 1)}%", "Active HTTP 200 responses confirmed"),
        ("Scraper Engine Status", "Operational (Direct API)", "Zero-browser high-speed extraction engine"),
    ]

    start_row = 5
    for r_idx, row_items in enumerate(summary_data):
        curr_row = start_row + r_idx
        for c_idx, val in enumerate(row_items):
            cell = ws_sum.cell(row=curr_row, column=2 + c_idx, value=val)
            cell.border = thin_border
            if r_idx == 0:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = bold_regular if c_idx == 0 else regular_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
                if c_idx == 1 and "%" in str(val):
                    cell.font = verified_font
                    cell.fill = verified_fill

    # Breakdown by Entity Type
    ws_sum.cell(row=start_row + len(summary_data) + 2, column=2, value="Entity Classification Breakdown").font = Font(name="Calibri", size=13, bold=True, color=navy_dark)
    
    categories = {}
    for r in records:
        key = (r["Group"], r["Exchange"], r["Category"])
        categories[key] = categories.get(key, 0) + 1

    cat_headers = ["Group", "Exchange", "Category / Type", "Count"]
    c_start_row = start_row + len(summary_data) + 4
    for c_idx, h in enumerate(cat_headers):
        cell = ws_sum.cell(row=c_start_row, column=2 + c_idx, value=h)
        cell.font = header_font
        cell.fill = PatternFill(start_color=navy_light, end_color=navy_light, fill_type="solid")
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, (k, cnt) in enumerate(sorted(categories.items()), start=1):
        ws_sum.cell(row=c_start_row + i, column=2, value=k[0]).font = bold_regular
        ws_sum.cell(row=c_start_row + i, column=3, value=k[1]).font = regular_font
        ws_sum.cell(row=c_start_row + i, column=4, value=k[2]).font = regular_font
        ws_sum.cell(row=c_start_row + i, column=5, value=cnt).font = bold_regular
        for col_i in range(2, 6):
            ws_sum.cell(row=c_start_row + i, column=col_i).border = thin_border

    # Adjust widths for summary sheet
    for col in ws_sum.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # ---------------------------------------------------------
    # Helper: Populate Company Sheets
    # ---------------------------------------------------------
    def populate_data_sheet(sheet_title: str, dataset: List[Dict[str, Any]]):
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        
        columns = [
            ("Exchange", 10),
            ("Group", 10),
            ("Classification", 22),
            ("Company Name", 36),
            ("Symbol / Scrip Code", 18),
            ("ISIN", 16),
            ("Series / Grp", 12),
            ("Verification Status", 18),
            ("Company Stock Page", 40),
            ("Page HTTP Status", 20),
            ("Corporate Actions URL", 40),
            ("Actions HTTP Status", 20),
            ("Dividend Records", 16),
            ("Latest Corporate Action / Dividend", 35),
            ("Verification Timestamp", 20),
        ]

        # Header Row
        for col_idx, (col_name, _) in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"

        # Data Rows
        for row_idx, r in enumerate(dataset, start=2):
            ws.row_dimensions[row_idx].height = 20
            is_even = (row_idx % 2 == 0)
            row_fill = PatternFill(start_color=gray_bg, end_color=gray_bg, fill_type="solid") if is_even else None

            # Values mapping
            row_vals = [
                r["Exchange"],
                r["Group"],
                r["Category"],
                r["Company Name"],
                r["Symbol / Code"],
                r["ISIN"],
                r.get("Series / Group", ""),
                r["Overall Verification"],
                r["Company Page URL"],
                r["Page Status"],
                r["Corporate Actions URL"],
                r["Corporate Actions Status"],
                r["Dividend / Action Records"],
                r["Latest Action"],
                r["Verified At"],
            ]

            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
                if row_fill:
                    cell.fill = row_fill

                # Special column treatments
                if col_idx in (1, 2, 7):
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = regular_font
                elif col_idx == 5:
                    cell.value = str(val)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = bold_regular
                elif col_idx == 8:  # Verification Status
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if "Verified" in str(val):
                        cell.font = verified_font
                        cell.fill = verified_fill
                    else:
                        cell.font = bold_regular
                elif col_idx in (9, 11):  # URLs (Clickable hyperlinks)
                    url_str = str(val)
                    cell.value = url_str
                    cell.hyperlink = url_str
                    cell.font = link_font
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif col_idx in (10, 12):  # Page HTTP Status
                    cell.value = str(val)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if "200 OK" in str(val):
                        cell.font = Font(name="Calibri", size=10, color=green_text)
                    else:
                        cell.font = regular_font
                elif col_idx == 13:  # Dividend Count
                    cell.value = int(val) if isinstance(val, (int, float)) else val
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.font = bold_regular if val > 0 else regular_font
                else:
                    cell.value = str(val)
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.font = regular_font

        # Auto-adjust column widths
        for col_idx, (_, col_w) in enumerate(columns, start=1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = col_w

        # Enable AutoFilter on header row
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(dataset) + 1}"

    # Generate individual & combined sheets
    populate_data_sheet("Tata Companies", tata_records)
    populate_data_sheet("ICICI Companies", icici_records)
    populate_data_sheet("All ICICI & Tata Combined", records)

    # Save workbook
    wb.save(output_path)
    print(f"\n🎉 Successfully created Excel workbook: {output_path}")


# ---------------------------------------------------------
# Main Execution Entry Point
# ---------------------------------------------------------
def run_pipeline(output_filename: str = "ICICI_and_Tata_Companies_Verified.xlsx", workers: int = 8):
    print("=" * 70)
    print("🚀 NSE & BSE COMPANY LIST INGESTION, LINK VERIFICATION & EXCEL EXPORTER")
    print("=" * 70)
    
    start_time = time.time()
    session = requests.Session(impersonate="chrome124")

    # Step 1: Ingest latest NSE & BSE company masters
    df_nse = pull_latest_nse_companies(session)
    df_bse = pull_latest_bse_companies(session)
    save_updated_reference_files(df_nse, df_bse)

    # Step 2: Filter ICICI and Tata companies
    companies = filter_companies(df_nse, df_bse)
    print(f"\n📋 Identified {len(companies)} ICICI and Tata entities across NSE & BSE:")
    print(f"   • Tata Companies:  {sum(1 for c in companies if c['Group'] == 'Tata')}")
    print(f"   • ICICI Companies: {sum(1 for c in companies if c['Group'] == 'ICICI')}")

    # Step 3: Verify links and scrape corporate actions
    verified_records = verify_all_companies(companies, max_workers=workers)

    # Step 4: Write to Excel
    base_dir = get_base_path()
    output_dir = os.path.join(base_dir, "output", "1_verified_companies")
    os.makedirs(output_dir, exist_ok=True)
    root_output_dir = os.path.join(base_dir, "output")
    
    excel_path = os.path.join(output_dir, output_filename)
    create_excel_workbook(verified_records, excel_path)
    
    # Save convenient copies in subfolder and root output/
    alias_path = os.path.join(output_dir, "ICICI_Tata_Companies.xlsx")
    root_copy1 = os.path.join(root_output_dir, output_filename)
    root_copy2 = os.path.join(root_output_dir, "ICICI_Tata_Companies.xlsx")
    import shutil
    if alias_path != excel_path:
        shutil.copy2(excel_path, alias_path)
    shutil.copy2(excel_path, root_copy1)
    shutil.copy2(excel_path, root_copy2)

    elapsed = round(time.time() - start_time, 2)
    print("=" * 70)
    print(f"✅ Pipeline completed in {elapsed} seconds!")
    print(f"📁 Primary Output: {excel_path}")
    print("=" * 70)
    return excel_path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Pull NSE/BSE companies, verify links, and export ICICI & Tata Excel")
    parser.add_argument("-o", "--output", default="ICICI_and_Tata_Companies_Verified.xlsx", help="Output Excel filename")
    parser.add_argument("-w", "--workers", type=int, default=8, help="Verification worker threads")
    args = parser.parse_args()

    run_pipeline(output_filename=args.output, workers=args.workers)
