# nse_scraper.py (Modified and ready for EXE conversion)

import os
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

def get_base_path():
    """ Get absolute path to the base directory, works for dev and for PyInstaller """
    # MODIFIED: This function finds the EXE's directory when frozen.
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def read_companies_from_file(filename="companies_nse.txt"):
    """ Reads a list of companies from a text file in the same directory. """
    # MODIFIED: Uses the base path to reliably find the input file.
    file_path = os.path.join(get_base_path(), filename)
    companies_dict = {}
    try:
        with open(file_path, 'r') as f:
            for line in f:
                if not line.strip():
                    continue
                parts = line.strip().split(',')
                if len(parts) == 2:
                    name = parts[0].strip()
                    symbol = parts[1].strip()
                    companies_dict[name] = symbol
    except FileNotFoundError:
        print(f"❌ File '{filename}' not found in the application folder.")
        return None
    return companies_dict

# --- Main Script ---

# MODIFIED: Set the base path for all file operations.
BASE_PATH = get_base_path()
DOWNLOAD_DIR = os.path.join(BASE_PATH, "downloads_nse")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

companies = read_companies_from_file()

if not companies:
    print("No companies to process. Exiting in 10 seconds.")
    time.sleep(10)
else:
    # --- Chrome options are fine, no changes needed here ---
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/115.0.0.0 Safari/537.36")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True
    })

    print("Setting up WebDriver...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    print("Initializing session with NSE India...")
    driver.get("https://www.nseindia.com")
    time.sleep(3)

    final_data = []

    for name, symbol in companies.items():
        print(f"🔍 Downloading for {name} ({symbol})")
        try:
            url = f"https://www.nseindia.com/companies-listing/corporate-filings-actions?symbol={symbol}&tabIndex=equity"
            driver.get(url)

            download_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "CFcorpactionsEquity-download"))
            )

            driver.execute_script("arguments[0].scrollIntoView(true);", download_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", download_btn)

            # A more reliable wait for the download to complete
            WebDriverWait(driver, 20, 1).until(
                lambda d: any(f.endswith('.csv') for f in os.listdir(DOWNLOAD_DIR))
            )

            files = os.listdir(DOWNLOAD_DIR)
            csv_files = [os.path.join(DOWNLOAD_DIR, f) for f in files if f.endswith(".csv")]
            latest_file = max(csv_files, key=os.path.getctime)

            df = pd.read_csv(latest_file, encoding="utf-8", engine="python")
            df["Company"] = name
            final_data.append(df)
            os.remove(latest_file)

        except TimeoutException:
            print(f"❌ Timeout for {name}: Download button not found or page did not load.")
        except Exception as e:
            print(f"❌ Error for {name}: {e}")

    driver.quit()

    if final_data:
        all_df = pd.concat(final_data, ignore_index=True)
        # MODIFIED: Save the final output file next to the EXE.
        output_path = os.path.join(BASE_PATH, "merged_nse_dividend_data.csv")
        all_df.to_csv(output_path, index=False)
        print(f"✅ Saved merged data to {output_path}")
    else:
        print("⚠️ No data was downloaded.")

    # MODIFIED: Pause at the end so the user can read the output.
    print("\nScript finished. This window will close in 15 seconds.")
    time.sleep(15)


# python nse.py