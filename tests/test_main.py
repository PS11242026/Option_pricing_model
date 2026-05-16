import io
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

import pandas as pd

import main


class MainDatasetTests(unittest.TestCase):
    def test_build_full_dataset_announces_fetch_and_processing_progress(self):
        companies = pd.DataFrame([{"Ticker": "AAA", "MarketCap": 100.0}])

        with (
            patch("main.get_top_companies", return_value=companies) as get_top_companies,
            patch("main.fetch_active_option_rows", return_value=[]),
        ):
            stream = io.StringIO()
            with redirect_stdout(stream):
                dataset = main.build_full_dataset()

        output = stream.getvalue()
        self.assertIn("Fetching live market data....", output)
        self.assertIn("Fetched market data successfully", output)
        self.assertIn("Processing market data", output)
        self.assertIn("[--------------------------------]   0% (0/1)", output)
        self.assertIn("[################################] 100% (1/1) AAA skipped", output)
        get_top_companies.assert_called_once_with(limit=main.DEFAULT_COMPANY_LIMIT)
        self.assertTrue(dataset.empty)

    def test_format_progress_bar_uses_completed_work_fraction(self):
        self.assertEqual(
            main.format_progress_bar(2, 5, width=10),
            "[####------]  40% (2/5)",
        )

    def test_run_engine_cli_includes_pricing_engine_tracker(self):
        dataset = pd.DataFrame(
            [
                {
                    "Ticker": "AAA",
                    "Type": "call",
                    "MarketOptionType": "call",
                    "ContractSymbol": "AAA_CALL",
                    "LastTradeDate": "2026-05-15T15:59:00+00:00",
                    "Spot": 100.0,
                    "Strike": 105.0,
                    "Expiry": "2026-06-19",
                    "t(years)": 0.1,
                    "Volatility": 0.25,
                    "DividendYield": 0.0,
                    "Market": 4.0,
                    "BSM": 6.5,
                    "AbsMis": 2.5,
                    "MarketCap": 100.0,
                }
            ]
        )

        with (
            patch("main.print_header"),
            patch("main.build_full_dataset", return_value=dataset),
            patch("main.print_original_style_results"),
            patch("main.save_workbook"),
        ):
            stream = io.StringIO()
            with redirect_stdout(stream):
                main.run_engine_cli()

        self.assertIn("Portfolio Tracker: Top Mispriced Options", stream.getvalue())


if __name__ == "__main__":
    unittest.main()
