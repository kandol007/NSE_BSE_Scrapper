# bse_scraper.py (Your Modified Script)

import os
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service # MODIFIED: Import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def get_base_path():
    """ Get absolute path to the base directory, works for dev and for PyInstaller """
    # MODIFIED: This function helps find files when run as an .exe
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # sets the sys.frozen attribute and stores the path in sys.executable.
        return os.path.dirname(sys.executable)
    else:
        # If run as a normal .py script, the base path is the script's directory.
        return os.path.dirname(os.path.abspath(__file__))

def read_companies_from_file(filename="companies_bse.txt"):
    """
    Reads a list of companies from a text file located in the same directory as the script/exe.
    """
    # MODIFIED: Use the base path to find the companies file
    file_path = os.path.join(get_base_path(), filename)
    
    companies_dict = {}
    try:
        with open(file_path, 'r') as f:
            for line in f:
                if not line.strip():
                    continue
                parts = line.strip().split(',')
                if len(parts) == 2:
                    companies_dict[parts[0].strip()] = parts[1].strip()
                else:
                    print(f"⚠️ Skipping malformed line: {line.strip()}")
    except FileNotFoundError:
        print(f"❌ Error: The file '{filename}' was not found in the same folder as the application.")
        print(f"   Please make sure '{filename}' is next to the .exe file.")
        return None
    return companies_dict

# --- Main Script ---

# MODIFIED: Use the base path for downloads
BASE_PATH = get_base_path()
DOWNLOAD_DIR = os.path.join(BASE_PATH, "downloads_bse")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

companies = read_companies_from_file()

if not companies:
    print("No companies to process. Exiting in 10 seconds.")
    time.sleep(10)
else:
    # Setup Chrome options
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True
    })
    chrome_options.add_argument("--headless=new") 
    
    # MODIFIED: This automatically manages chromedriver.exe for you!
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)

    final_data = []

    for name, code in companies.items():
        print(f"🔍 Downloading dividend data for {name} ({code})")
        try:
            url = f"https://www.bseindia.com/corporates/corporates_act.html?scripcode={code}&Dividend=P9&scripname="
            driver.get(url)

            download_icon = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'fa-download')]"))
            ) 
            driver.execute_script("arguments[0].click();", download_icon)
            
            # This is a more reliable wait than time.sleep()
            # It waits up to 20 seconds for a new .csv file to appear.
            WebDriverWait(driver, 20, 1).until(
                lambda d: any(f.endswith('.csv') for f in os.listdir(DOWNLOAD_DIR))
            )

            files = os.listdir(DOWNLOAD_DIR)
            csv_files = [os.path.join(DOWNLOAD_DIR, f) for f in files if f.endswith(".csv")]
            latest_file = max(csv_files, key=os.path.getctime)

            df = pd.read_csv(latest_file, encoding="utf-8", engine='python')
            df["Company"] = name
            final_data.append(df)
            
            os.remove(latest_file)

        except TimeoutException:
            print(f"❌ Failed for {name}: Could not find the download button or file did not download in time.")
        except Exception as e:
            print(f"❌ An unexpected error occurred for {name} ({code}): {e}")

    driver.quit()

    if final_data:
        all_df = pd.concat(final_data, ignore_index=True)
        # MODIFIED: Save the final CSV in the base path
        final_csv_path = os.path.join(BASE_PATH, "merged_bse_dividend_data.csv")
        all_df.to_csv(final_csv_path, index=False)
        print(f"\n✅ Saved merged data to '{final_csv_path}'")
    else:
        print("\n⚠️ No data was downloaded to save.")

    print("\nScript finished. This window will close in 15 seconds.")
    time.sleep(15)


# python bse.py