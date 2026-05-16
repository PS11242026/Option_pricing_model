import io
import unittest
from contextlib import redirect_stdout

import pandas as pd

from pricing_engine import (
    DEFAULT_TOP_N,
    arbitrage_from_spread,
    build_tracker_table,
    display_tracker_table,
    moneyness_from_contract,
    print_tracker_table,
    signal_from_spread,
)


class PortfolioTrackerTests(unittest.TestCase):
    def test_default_top_n_is_fifty_contracts(self):
        self.assertEqual(DEFAULT_TOP_N, 50)

    def test_signal_from_spread_maps_direction(self):
        self.assertEqual(signal_from_spread(2.5), "BUY")
        self.assertEqual(signal_from_spread(-1.0), "SELL")
        self.assertEqual(signal_from_spread(0.0), "HOLD")

    def test_arbitrage_from_spread_describes_edge(self):
        self.assertEqual(arbitrage_from_spread("call", 1.5), "Long CALL; market below BSM")
        self.assertEqual(arbitrage_from_spread("put", -0.5), "Short PUT; market above BSM")
        self.assertEqual(arbitrage_from_spread("put", 0.0), "No edge in PUT")

    def test_moneyness_from_contract_classifies_options(self):
        self.assertEqual(
            moneyness_from_contract("call", 110.0, 100.0),
            "IN-THE-MONEY",
        )
        self.assertEqual(
            moneyness_from_contract("put", 110.0, 100.0),
            "OUT-OF-THE-MONEY",
        )
        self.assertEqual(
            moneyness_from_contract("call", 100.0, 100.0),
            "AT-THE-MONEY",
        )

    def test_build_tracker_table_sorts_by_absolute_mispricing(self):
        dataset = pd.DataFrame(
            [
                {
                    "Ticker": "AAA",
                    "Type": "call",
                    "ContractSymbol": "AAA_CALL",
                    "Spot": 100.0,
                    "Strike": 105.0,
                    "Expiry": "2026-06-19",
                    "Market": 4.0,
                    "BSM": 6.5,
                },
                {
                    "Ticker": "BBB",
                    "Type": "put",
                    "ContractSymbol": "BBB_PUT",
                    "Spot": 95.0,
                    "Strike": 90.0,
                    "Expiry": "2026-07-17",
                    "Market": 8.0,
                    "BSM": 3.0,
                },
                {
                    "Ticker": "CCC",
                    "Type": "call",
                    "ContractSymbol": "CCC_CALL",
                    "Spot": 120.0,
                    "Strike": 125.0,
                    "Expiry": "2026-08-21",
                    "Market": 5.0,
                    "BSM": 5.5,
                },
            ]
        )

        tracker = build_tracker_table(dataset, top_n=2)

        self.assertEqual(list(tracker["Ticker"]), ["BBB", "AAA"])
        self.assertEqual(list(tracker["Signal"]), ["SELL", "BUY"])
        self.assertEqual(
            list(tracker["Arbitrage"]),
            ["Short PUT; market above BSM", "Long CALL; market below BSM"],
        )
        self.assertEqual(
            list(tracker["Moneyness"]),
            ["OUT-OF-THE-MONEY", "OUT-OF-THE-MONEY"],
        )

    def test_display_tracker_table_formats_ranked_contract_view(self):
        tracker = pd.DataFrame(
            [
                {
                    "Ticker": "BBB",
                    "Type": "put",
                    "ContractSymbol": "BBB_PUT",
                    "Spot": 95.0,
                    "Strike": 90.0,
                    "Expiry": "2026-07-17",
                    "Market": 8.0,
                    "BSM": 3.0,
                    "Spread": -5.0,
                    "AbsMis": 5.0,
                    "Signal": "SELL",
                    "Arbitrage": "Short PUT; market above BSM",
                    "Moneyness": "OUT-OF-THE-MONEY",
                }
            ]
        )

        display = display_tracker_table(tracker)

        self.assertEqual(
            list(display.columns),
            [
                "Rank",
                "Ticker",
                "Option",
                "Symbol",
                "Spot",
                "Strike",
                "Expiry",
                "DTE",
                "Market",
                "BSM",
                "Edge",
                "Abs Mis",
                "Signal",
                "Moneyness",
                "Trade",
            ],
        )
        self.assertEqual(display.iloc[0]["Rank"], 1)
        self.assertEqual(display.iloc[0]["Edge"], "-5.00")
        self.assertEqual(display.iloc[0]["Abs Mis"], "$5.00")
        self.assertNotIn("Edge %", display.columns)

    def test_print_tracker_table_includes_summary_and_table_output(self):
        tracker = pd.DataFrame(
            [
                {
                    "Ticker": "BBB",
                    "Type": "put",
                    "ContractSymbol": "BBB_PUT",
                    "Spot": 95.0,
                    "Strike": 90.0,
                    "Expiry": "2026-07-17",
                    "Market": 8.0,
                    "BSM": 3.0,
                    "Spread": -5.0,
                    "AbsMis": 5.0,
                    "Signal": "SELL",
                    "Arbitrage": "Short PUT; market above BSM",
                    "Moneyness": "OUT-OF-THE-MONEY",
                },
                {
                    "Ticker": "AAA",
                    "Type": "call",
                    "ContractSymbol": "AAA_CALL",
                    "Spot": 100.0,
                    "Strike": 105.0,
                    "Expiry": "2026-06-19",
                    "Market": 4.0,
                    "BSM": 6.5,
                    "Spread": 2.5,
                    "AbsMis": 2.5,
                    "Signal": "BUY",
                    "Arbitrage": "Long CALL; market below BSM",
                    "Moneyness": "OUT-OF-THE-MONEY",
                },
            ]
        )

        stream = io.StringIO()
        with redirect_stdout(stream):
            print_tracker_table(tracker, top_n=2, risk_free_rate=0.037)

        output = stream.getvalue()

        self.assertIn("Signal mix: BUY 1 | SELL 1 | HOLD 0", output)
        self.assertIn("Largest edge: BBB PUT", output)
        self.assertIn("Moneyness mix: ITM 0 | OTM 2 | ATM 0", output)
        self.assertIn("Rank Ticker Option", output)
        self.assertIn("BBB_PUT", output)
        self.assertIn("Short PUT; market above BSM", output)
        self.assertNotIn("1. BBB PUT", output)


if __name__ == "__main__":
    unittest.main()
