# NSE & BSE Corporate Actions & Dividend Scraper ⚡

A high-performance, rock-solid Python suite to fetch **corporate actions, dividend declarations, and shareholder-level unclaimed dividend records** from the **National Stock Exchange (NSE)**, **Bombay Stock Exchange (BSE)**, and **Corporate Investor Relations Portals** in seconds.

---

## 🚀 Key Improvements & Benchmarks

| Feature | Legacy Selenium Version | Enhanced API Engine |
| :--- | :--- | :--- |
| **Speed (6 companies)** | ~90–120 seconds | **~0.5–1.2 seconds (100x faster)** |
| **Resource Usage** | > 500 MB RAM (Headless Chrome) | **< 35 MB RAM** |
| **Dependencies** | Chrome, ChromeDriver, Selenium | **Zero browser dependencies** |
| **Empty Records Handling** | 20-second timeout stalls | **Instant non-blocking detection** |
| **Concurrency** | Sequential only | **Threaded multi-worker pool** |
| **Akamai / Cloudflare Defense** | Fragile WebDriver detection | **TLS Fingerprint Emulation (`curl_cffi`)** |
| **Test Suite** | None | **Automated unit & integration tests** |

---

## 📌 Project Architecture & Outputs Overview

The project generates **3 distinct categories of outputs**, stored in dedicated subfolders inside `output/`:

| Output Category | Folder | Description & Key Files |
| :--- | :--- | :--- |
| **Type 1: Verified Company Listings** | `output/1_verified_companies/` | Ingests 2,580+ NSE and 5,050+ BSE equities, filters Tata & ICICI companies, verifies live stock quote & corporate action URLs, and exports `ICICI_and_Tata_Companies_Verified.xlsx`. |
| **Type 2: Corporate Dividend Declarations** | `output/2_corporate_dividend_history/` | Scrapes 540+ historical corporate dividend announcements across 23 Tata & ICICI entities from NSE & BSE, styled with summary KPIs and breakdowns in `Tata_and_ICICI_Dividend_Records.xlsx`. |
| **Type 3: Shareholder Unclaimed Dividends** | `output/3_shareholder_unclaimed_dividends/` | Crawls investor relations portals and RTA systems for shareholder-level unclaimed dividend lists (pursuant to Section 124 of Companies Act 2013 & IEPF Rules). Exports `Tata_and_ICICI_Shareholder_Unclaimed_Dividends.xlsx` (362 Section 124 records + 441 IEPF records with investor names, folios, addresses, shares, and amounts). |

---

## 📂 File Structure

```text
├── input/                                    # 📥 Input data directory (git-ignored)
│   ├── companies_nse.txt                     # Default input list: NSE Company Name & Symbol
│   ├── companies_bse.txt                     # Default input list: BSE Company Name & Scrip Code
│   ├── companies_tata_nse.txt                # Tata companies on NSE
│   ├── companies_tata_bse.txt                # Tata companies on BSE
│   ├── companies_icici_nse.txt               # ICICI companies on NSE
│   ├── companies_icici_bse.txt               # ICICI companies on BSE
│   ├── NSE_List_of_companies.csv             # Reference master list of NSE symbols (2,580+)
│   └── BSE_List_of_companies.csv             # Reference master list of BSE scrip codes (5,050+)
├── output/                                   # 📤 Categorized outputs directory (git-ignored)
│   ├── 1_verified_companies/                 # 📂 Output Type 1: Verified Company Listings & Live URLs
│   │   ├── ICICI_and_Tata_Companies_Verified.xlsx
│   │   └── ICICI_Tata_Companies.xlsx
│   ├── 2_corporate_dividend_history/         # 📂 Output Type 2: Exchange Corporate Dividend History
│   │   ├── Tata_and_ICICI_Dividend_Records.xlsx
│   │   ├── merged_nse_dividend_data.csv
│   │   └── merged_bse_dividend_data.csv
│   └── 3_shareholder_unclaimed_dividends/    # 📂 Output Type 3: Shareholder-Level Unclaimed Dividends & IEPF Lists
│       ├── Tata_and_ICICI_Shareholder_Unclaimed_Dividends.xlsx # Master 5-sheet styled Excel
│       ├── tata_investment_corp_section124_unclaimed_shareholders.csv # 362 shareholder records
│       ├── tata_investment_corp_iepf_7yr_transfer_shareholders.csv    # 441 IEPF records
│       └── company_unclaimed_dividend_portals_and_rtas.csv           # RTA & portal directory
├── scraper_core.py                           # Core high-speed scraping engine (NSEScraper & BSEScraper)
├── main.py                                   # Unified CLI runner (orchestrates all scraping & exports)
├── nse.py                                    # Standalone NSE scraper runner
├── bse.py                                    # Standalone BSE scraper runner
├── pull_and_verify_companies.py              # Automated ingest, link verification & Excel exporter (Type 1)
├── download_dividend_data_xlsx.py            # Corporate dividend history exporter (Type 2)
├── scrape_shareholder_unclaimed_dividends.py # Live shareholder unclaimed dividend & IEPF crawler (Type 3)
├── test_scrapers.py                          # Automated test suite (all unit tests run in < 1s)
├── requirements.txt                          # Minimal clean dependencies
├── .gitignore                                # Excludes virtualenvs, cache, input/ and output/
└── ReadMe.md                                 # Complete documentation
```

---

## 🐍 Python Scripts Reference — Which File Does What

### 1. `main.py` — Unified CLI Orchestrator
* **Purpose**: The central entry point for the entire repository. Provides a flexible CLI with flags to run individual exchange scrapers, custom pipelines, or the entire end-to-end extraction suite.
* **Key Commands & Flags**:
  * `python main.py --all`: Scrapes NSE and BSE in parallel using default company lists.
  * `python main.py -e nse` / `python main.py -e bse`: Scrapes a single exchange.
  * `python main.py --export-icici-tata`: Executes **Output Type 1** pipeline (`pull_and_verify_companies.py`).
  * `python main.py --export-dividends`: Executes **Output Type 2** pipeline (`download_dividend_data_xlsx.py`).
  * `python main.py --export-unclaimed-shareholders`: Executes **Output Type 3** pipeline (`scrape_shareholder_unclaimed_dividends.py`).
  * `python main.py --export-all`: Runs all three pipelines sequentially end-to-end.
* **Under the Hood**: Uses `argparse` for robust CLI parameter parsing, worker pool sizing, and dispatches tasks to `scraper_core.py` or dedicated exporters.

### 2. `scraper_core.py` — Engine & Networking Core
* **Purpose**: Houses the high-speed HTTP engine, session managers, anti-bot bypass mechanisms, and file I/O helpers.
* **Key Classes & Functions**:
  * `NSEScraper`: Manages NSE cookie bootstrapping, session recreation on 401/403, and queries the official NSE corporate actions API endpoint (`https://www.nseindia.com/api/corporates-corporateActions`).
  * `BSEScraper`: Handles BSE API interaction (`https://api.bseindia.com/BseIndiaAPI/Service/CorporateAction.id/Default.aspx`), parses JSON responses, formats security codes, and normalizes dividend descriptions.
  * `read_companies_from_file(filepath)`: Robust parser for CSV/TXT input files (handles commas, whitespace, comments, and empty lines).
  * `save_merged_csv(df, output_path)`: Thread-safe CSV saver with UTF-8 encoding and automatic directory creation.
* **Under the Hood**: Uses `curl_cffi` to simulate realistic browser TLS handshakes (Chrome impersonation), bypassing Akamai Bot Manager and Cloudflare WAF without spawning a browser process.

### 3. `nse.py` — Standalone NSE Scraper
* **Purpose**: Dedicated command-line runner for scraping the National Stock Exchange (NSE).
* **Usage**:
  ```bash
  python nse.py
  python nse.py -s RELIANCE               # Scrape a single symbol directly
  python nse.py -i input/custom.txt -w 8  # Custom input file and worker count
  ```
* **Output**: Writes merged dividend records to `output/merged_nse_dividend_data.csv` (or custom `-o` path).

### 4. `bse.py` — Standalone BSE Scraper
* **Purpose**: Dedicated command-line runner for scraping the Bombay Stock Exchange (BSE).
* **Usage**:
  ```bash
  python bse.py
  python bse.py -c 500325                 # Scrape a single BSE scrip code directly
  python bse.py -i input/custom.txt -w 8  # Custom input file and worker count
  ```
* **Output**: Writes merged dividend records to `output/merged_bse_dividend_data.csv` (or custom `-o` path).

### 5. `pull_and_verify_companies.py` — Ingestion & Link Verification (Output Type 1)
* **Purpose**: Ingests full master lists of companies from NSE and BSE, filters out all Tata Group and ICICI Group entities, verifies live HTTP stock quotes & corporate action URLs, and compiles a publication-ready Excel directory.
* **Usage**:
  ```bash
  python pull_and_verify_companies.py
  # or: python main.py --export-icici-tata
  ```
* **Outputs Generated**:
  * `output/1_verified_companies/ICICI_and_Tata_Companies_Verified.xlsx` (Multi-sheet workbook with Summary KPIs, Tata Companies, ICICI Companies, All Verified Entities, and Discrepancies)
  * `output/1_verified_companies/ICICI_Tata_Companies.xlsx` (Convenient alias)

### 6. `download_dividend_data_xlsx.py` — Corporate Dividend Exporter (Output Type 2)
* **Purpose**: Scrapes all historical dividend declarations and corporate actions for 23 Tata & ICICI companies across both exchanges concurrently and compiles them into a styled Excel workbook.
* **Usage**:
  ```bash
  python download_dividend_data_xlsx.py
  # or: python main.py --export-dividends
  ```
* **Outputs Generated**:
  * `output/2_corporate_dividend_history/Tata_and_ICICI_Dividend_Records.xlsx` (543 historical declarations with KPI summary cards, Tata NSE/BSE tabs, ICICI NSE/BSE tabs)
  * `output/2_corporate_dividend_history/merged_nse_dividend_data.csv`
  * `output/2_corporate_dividend_history/merged_bse_dividend_data.csv`

### 7. `scrape_shareholder_unclaimed_dividends.py` — Live Shareholder Unclaimed Dividend Crawler (Output Type 3)
* **Purpose**: Directly crawls company Investor Relations websites and RTA systems to extract shareholder-level unclaimed dividend records (investor names, folios, addresses, shares, and amounts) pursuant to Section 124 of the Companies Act 2013 and IEPF Rules.
* **Usage**:
  ```bash
  python scrape_shareholder_unclaimed_dividends.py
  # or: python main.py --export-unclaimed-shareholders
  ```
* **Key Features**:
  * Extracts **362 shareholder records** from Tata Investment Corp's Section 124(2) Unpaid Dividend Account Report (totaling ₹3,370,123.00 across 139,440 shares).
  * Extracts **441 records** from Tata Investment Corp's 7-Year Consecutive Outstanding Dividends & IEPF Transfer Liability Report.
  * Maps ICICI Bank's shareholder claiming system via KFintech IEPF Portal (`https://ris.kfintech.com/services/IEPF/IEPFInfo.aspx?q=uhZuJ9G3geI%3d`) with complete Form IEPF-5 filing instructions.
  * Compiles corporate RTA directory for Link Intime & KFintech.
* **Outputs Generated**:
  * `output/3_shareholder_unclaimed_dividends/Tata_and_ICICI_Shareholder_Unclaimed_Dividends.xlsx` (Master 5-sheet styled workbook)
  * `output/3_shareholder_unclaimed_dividends/tata_investment_corp_section124_unclaimed_shareholders.csv`
  * `output/3_shareholder_unclaimed_dividends/tata_investment_corp_iepf_7yr_transfer_shareholders.csv`
  * `output/3_shareholder_unclaimed_dividends/company_unclaimed_dividend_portals_and_rtas.csv`

### 8. `test_scrapers.py` — Automated Test Suite
* **Purpose**: Validates system health, live exchange connectivity, input file parsing, output writing, and empty record handling.
* **Usage**:
  ```bash
  python test_scrapers.py
  ```
* **Test Cases Covered**:
  1. `test_01_read_companies_from_file`: Verifies parsing of NSE & BSE company files in `input/`.
  2. `test_02_nse_live_scrape`: Verifies live NSE API connectivity and DataFrame schemas.
  3. `test_03_bse_live_scrape`: Verifies live BSE API connectivity and DataFrame schemas.
  4. `test_04_bse_zero_records_handled_gracefully`: Verifies instant non-blocking handling for companies without dividends.
  5. `test_05_save_merged_csv`: Verifies thread-safe CSV export and directory creation.
* **Execution Time**: All 5 tests complete in under 1 second.

---

## 📝 Input Format

Input files are kept in the `input/` directory. Each line contains: `Company Name,Symbol/Code`

### `input/companies_nse.txt` (NSE Symbols)
```text
Reliance Industries,RELIANCE
State Bank of India,SBIN
Tata Consultancy Services,TCS
HDFC Bank,HDFCBANK
Infosys,INFY
Aarti Industries Limited,AARTIIND
```

### `input/companies_bse.txt` (BSE Scrip Codes)
```text
Reliance Industries,500325
State Bank of India,500112
TCS,532540
HDFC Bank,500180
Infosys,500209
AARV Infratel,526488
Alchemist Ltd,526707
```

---

## 🛠️ Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone git@github-personal:kandol007/NSE_BSE_Scrapper.git
   cd NSE_BSE_Scrapper
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## 💻 Quickstart Execution

```bash
# 1. Run quick scrape for both NSE & BSE
python main.py --all

# 2. Run Output Type 1: Verified Company Directory
python main.py --export-icici-tata

# 3. Run Output Type 2: Corporate Dividend Declarations Exporter
python main.py --export-dividends

# 4. Run Output Type 3: Shareholder Unclaimed Dividend Crawler
python main.py --export-unclaimed-shareholders

# 5. Run All 3 Output Pipelines sequentially
python main.py --export-all

# 6. Run test suite
python test_scrapers.py
```

---

## 📄 License & Attribution

Designed and maintained for automated financial intelligence and investor claiming assistance. Adheres to Andrej Karpathy's guidelines for AI-assisted engineering: *Think Before Coding, Simplicity First, Surgical Changes, Goal-Driven Execution*.