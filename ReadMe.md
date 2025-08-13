# NSE and BSE Dividend Data Downloader (Selenium)

This Python script automates the process of downloading **corporate actions/dividend CSV files** from the **NSE India** and **BSE India** website for a list of companies, merges them, and saves them into a single CSV file.

---

## 📌 Features
- Reads a list of companies and their stock symbols from `companies_nse.txt` and `companies_bse.txt`.
- Uses **Selenium** with Chrome in headless mode to visit each company's NSE and BSE corporate filings page.
- Automatically clicks the **"Download (.csv)"** button for equity corporate actions.
- Waits for each CSV to download, processes the data, and merges all into one file:  
  `merged_nse_dividend_data.csv` and `merged_bse_dividend_data.csv`
- Cleans up downloaded CSVs after processing.

---

## 📂 File Structure
├── companies_nse.txt # Input: Company name and symbol list for NSE
├── companies_bse.txt # Input: Company name and symbol list for BSE
├── downloads_bse/ # Temporary CSV downloads for BSE script
├── downloads_nse/ # Temporary CSV downloads for NSE script
├── nse.py # Main script for NSE
├── bse.py # Main script for BSE
├── merged_nse_dividend_data.csv # Final merged data
├── merged_nse_dividend_data.csv # Final merged data
├── Requirements.txt
└── README.md # Documentation


---

## 📝 companies_nse.txt Format
Each line should contain:
Company Name,Symbol

Example:  For NSE 
Reliance Industries,RELIANCE
State Bank of India,SBIN
Tata Consultancy Services,TCS
HDFC Bank,HDFCBANK
Infosys,INFY

Example:  For BSE 
Reliance Industries,500325
State Bank of India,500112
TCS,532540
HDFC Bank,500180
Infosys,500209


---

## 🚀 How to Run

1️⃣ Install Dependencies
Make sure you have Python 3 installed, then run:
bash
pip install selenium webdriver-manager pandas


2️⃣ Prepare the Company List
Edit companies_nse.txt and companies_bse.txt with your desired companies and their NSE and BSE symbols.

3️⃣ Run the Script
python nse.py
python bse.py


## There is list for both NSE(NSE_List_of_companies.csv) and BSE(BSE_List_of_companies.csv) in which it contains the list of companies with company codes to use in companies_nse.txt and companies_bse.txt so you can add the companies with symbol for both scripts according to their setups