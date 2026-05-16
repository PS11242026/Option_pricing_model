import unittest
from types import SimpleNamespace

import pandas as pd

from data_fetch import (
    DEFAULT_COMPANY_LIMIT,
    parse_market_cap,
    parse_reuters_market_cap_table,
    option_market_price,
    prepare_active_options,
    reuters_symbol_to_yfinance,
    select_last_active_option,
    select_last_active_options_by_type,
    ticker_from_reuters_stock_link,
)


class OptionSelectionTests(unittest.TestCase):
    """Regression tests for Yahoo option-chain selection helpers."""

    def test_default_company_limit_is_one_hundred(self):
        self.assertEqual(DEFAULT_COMPANY_LIMIT, 100)

    def test_reuters_symbol_to_yfinance_handles_exchange_and_share_classes(self):
        self.assertEqual(reuters_symbol_to_yfinance("NVDA.OQ"), "NVDA")
        self.assertEqual(reuters_symbol_to_yfinance("BRKb.N"), "BRK-B")

    def test_ticker_from_reuters_stock_link_extracts_company_symbol(self):
        stock = "[Berkshire Hathaway](https://www.reuters.com/companies/BRKb.N)"

        self.assertEqual(ticker_from_reuters_stock_link(stock), "BRK-B")

    def test_parse_market_cap_accepts_raw_numbers_and_display_values(self):
        self.assertEqual(parse_market_cap("5.5 trillion"), 5.5e12)
        self.assertEqual(parse_market_cap("946.4 billion"), 946.4e9)
        self.assertEqual(parse_market_cap(5457369000000), 5457369000000.0)

    def test_parse_reuters_market_cap_table_returns_top_100_rows(self):
        rows = [
            "Stock,Market Value, ",
            "[NVIDIA](https://www.reuters.com/companies/NVDA.OQ),5.5 trillion,5457369000000",
            "[Alphabet](https://www.reuters.com/companies/GOOGL.OQ),4.8 trillion,4788509000000",
            "[Berkshire Hathaway](https://www.reuters.com/companies/BRKb.N),1.0 trillion,1040998000000",
        ]
        for index in range(4, 111):
            rows.append(
                f"[Company {index}](https://www.reuters.com/companies/T{index}.N),"
                f"{index}.0 billion,{index * 1_000_000_000}"
            )
        csv_text = "\n".join(rows)

        companies = parse_reuters_market_cap_table(csv_text, limit=DEFAULT_COMPANY_LIMIT)

        self.assertEqual(len(companies), 100)
        self.assertEqual(list(companies.columns), ["Ticker", "MarketCap"])
        self.assertEqual(companies.iloc[0]["Ticker"], "NVDA")
        self.assertEqual(companies.iloc[2]["Ticker"], "BRK-B")
        self.assertEqual(companies.iloc[0]["MarketCap"], 5457369000000.0)

    def test_market_price_uses_last_price_not_midpoint(self):
        option = pd.Series({"lastPrice": 3.25, "bid": 2.00, "ask": 2.20})

        self.assertEqual(option_market_price(option), 3.25)

    def test_prepare_active_options_keeps_usable_implied_volatility(self):
        options = pd.DataFrame(
            [
                {
                    "contractSymbol": "XYZ260117C00100000",
                    "strike": 100,
                    "lastPrice": 1.25,
                    "impliedVolatility": 0.30,
                    "lastTradeDate": "2026-01-10T15:59:00Z",
                },
                {
                    "contractSymbol": "XYZ260117C00105000",
                    "strike": 105,
                    "lastPrice": 1.75,
                    "impliedVolatility": 0.01,
                    "lastTradeDate": "2026-01-10T15:58:00Z",
                },
            ]
        )

        active = prepare_active_options(options, "2026-01-17", "call")

        self.assertEqual(len(active), 1)
        self.assertEqual(active.iloc[0]["contractSymbol"], "XYZ260117C00100000")
        self.assertEqual(active.iloc[0]["OptionType"], "call")
        self.assertEqual(active.iloc[0]["Expiry"], "2026-01-17")

    def test_select_last_active_option_uses_most_recent_trade_across_chains(self):
        older_calls = pd.DataFrame(
            [
                {
                    "contractSymbol": "XYZ260117C00100000",
                    "strike": 100,
                    "lastPrice": 5.00,
                    "impliedVolatility": 0.25,
                    "lastTradeDate": "2026-01-10T15:59:00Z",
                    "volume": 100,
                    "openInterest": 200,
                }
            ]
        )
        newer_puts = pd.DataFrame(
            [
                {
                    "contractSymbol": "XYZ260124P00095000",
                    "strike": 95,
                    "lastPrice": 4.00,
                    "impliedVolatility": 0.35,
                    "lastTradeDate": "2026-01-11T15:59:00Z",
                    "volume": 10,
                    "openInterest": 20,
                }
            ]
        )
        empty = pd.DataFrame()
        ticker = SimpleNamespace(
            option_chain=lambda expiry: SimpleNamespace(
                calls=older_calls if expiry == "2026-01-17" else empty,
                puts=newer_puts if expiry == "2026-01-24" else empty,
            )
        )

        selected = select_last_active_option(ticker, ["2026-01-17", "2026-01-24"])

        self.assertIsNotNone(selected)
        self.assertEqual(selected["contractSymbol"], "XYZ260124P00095000")
        self.assertEqual(selected["OptionType"], "put")

    def test_select_last_active_options_by_type_keeps_call_and_put_markets_separate(self):
        calls = pd.DataFrame(
            [
                {
                    "contractSymbol": "XYZ260117C00105000",
                    "strike": 105,
                    "lastPrice": 1.50,
                    "impliedVolatility": 0.20,
                    "lastTradeDate": "2026-01-10T15:59:00Z",
                    "volume": 100,
                    "openInterest": 200,
                }
            ]
        )
        puts = pd.DataFrame(
            [
                {
                    "contractSymbol": "XYZ260117P00095000",
                    "strike": 95,
                    "lastPrice": 2.75,
                    "impliedVolatility": 0.30,
                    "lastTradeDate": "2026-01-10T15:58:00Z",
                    "volume": 50,
                    "openInterest": 100,
                }
            ]
        )
        ticker = SimpleNamespace(
            option_chain=lambda expiry: SimpleNamespace(
                calls=calls,
                puts=puts,
            )
        )

        selected = select_last_active_options_by_type(ticker, ["2026-01-17"])

        self.assertEqual(selected["call"]["contractSymbol"], "XYZ260117C00105000")
        self.assertEqual(selected["put"]["contractSymbol"], "XYZ260117P00095000")
        self.assertEqual(selected["call"]["Market"], 1.50)
        self.assertEqual(selected["put"]["Market"], 2.75)


if __name__ == "__main__":
    unittest.main()
