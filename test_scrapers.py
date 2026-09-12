"""
test_scrapers.py - Automated Unit & Integration Tests
Validates company parsing, live API connectivity, data schemas, and error handling.
"""

import os
import unittest
import pandas as pd
from scraper_core import (
    NSEScraper,
    BSEScraper,
    read_companies_from_file,
    save_merged_csv,
    get_base_path
)


class TestScrapers(unittest.TestCase):

    def setUp(self):
        self.base_path = get_base_path()

    def test_01_read_companies_from_file(self):
        """Test parsing of companies files."""
        nse_file = os.path.join(self.base_path, "companies_nse.txt")
        companies = read_companies_from_file(nse_file)
        self.assertIsInstance(companies, dict)
        self.assertGreater(len(companies), 0)
        self.assertIn("Reliance Industries", companies)
        self.assertEqual(companies["Reliance Industries"], "RELIANCE")

        bse_file = os.path.join(self.base_path, "companies_bse.txt")
        bse_companies = read_companies_from_file(bse_file)
        self.assertIsInstance(bse_companies, dict)
        self.assertGreater(len(bse_companies), 0)
        self.assertIn("Reliance Industries", bse_companies)
        self.assertEqual(bse_companies["Reliance Industries"], "500325")

    def test_02_nse_live_scrape(self):
        """Test live extraction of dividend data from NSE."""
        scraper = NSEScraper()
        df = scraper.fetch_dividend_data("RELIANCE", company_name="Reliance Industries")
        self.assertIsInstance(df, pd.DataFrame)
        self.assertFalse(df.empty, "NSE RELIANCE returned no records")
        self.assertIn("SYMBOL", df.columns)
        self.assertIn("PURPOSE", df.columns)
        self.assertIn("Company", df.columns)
        self.assertEqual(df["Company"].iloc[0], "Reliance Industries")

    def test_03_bse_live_scrape(self):
        """Test live extraction of dividend data from BSE."""
        scraper = BSEScraper()
        df = scraper.fetch_dividend_data("500325", company_name="Reliance Industries")
        self.assertIsInstance(df, pd.DataFrame)
        self.assertFalse(df.empty, "BSE 500325 returned no records")
        self.assertIn("Security Code", df.columns)
        self.assertIn("Company", df.columns)
        self.assertEqual(df["Company"].iloc[0], "Reliance Industries")

    def test_04_bse_zero_records_handled_gracefully(self):
        """Test that scrips without dividend history return empty DataFrame instantly."""
        scraper = BSEScraper()
        # 526488 has no dividend actions
        df = scraper.fetch_dividend_data("526488", company_name="AARV Infratel")
        self.assertIsInstance(df, pd.DataFrame)
        self.assertTrue(df.empty)

    def test_05_save_merged_csv(self):
        """Test saving merged dataframe to disk."""
        test_df = pd.DataFrame([
            {"Company": "Test Corp", "Symbol": "TEST", "Dividend": "10.0"}
        ])
        test_out = os.path.join(self.base_path, "test_output.csv")
        try:
            saved = save_merged_csv(test_df, test_out)
            self.assertTrue(saved)
            self.assertTrue(os.path.exists(test_out))
            loaded = pd.read_csv(test_out)
            self.assertEqual(len(loaded), 1)
            self.assertEqual(loaded["Company"].iloc[0], "Test Corp")
        finally:
            if os.path.exists(test_out):
                os.remove(test_out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
