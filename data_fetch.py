"""Market data retrieval for the Top 25 option-pricing engine."""

from __future__ import annotations

from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import yfinance as yf


# Settings used while filtering market data and estimating volatility.
TRADING_DAYS = 252
MIN_OPTION_PRICE = 0.50
MIN_VOLATILITY = 0.05
MIN_TIME_TO_MATURITY = 1 / 365


def normalize_yfinance_ticker(symbol: str) -> str:
    """Normalize ticker formats for yfinance compatibility."""
    # Yahoo Finance expects share classes like BRK-B instead of BRK.B.
    return symbol.strip().upper().replace(".", "-")


def get_top25_companies() -> pd.DataFrame:
    """Return fixed Top 25 companies by market cap."""
    # Keeping this list here makes the project easier to run without extra CSV files.
    data = [
        ("NVDA", 4.90e12),
        ("GOOGL", 4.10e12),
        ("AAPL", 4.00e12),
        ("MSFT", 3.10e12),
        ("AMZN", 2.70e12),
        ("AVGO", 1.90e12),
        ("META", 1.70e12),
        ("TSLA", 1.50e12),
        ("BRK-B", 1.00e12),
        ("WMT", 1.00e12),
        ("LLY", 8.75e11),
        ("JPM", 8.32e11),
        ("XOM", 6.08e11),
        ("V", 6.04e11),
        ("JNJ", 5.64e11),
        ("MU", 5.13e11),
        ("ORCL", 5.03e11),
        ("MA", 4.64e11),
        ("AMD", 4.53e11),
        ("COST", 4.43e11),
        ("NFLX", 4.09e11),
        ("BAC", 3.84e11),
        ("CAT", 3.69e11),
        ("ABBV", 3.68e11),
        ("CVX", 3.66e11),
    ]
    return pd.DataFrame(data, columns=["Ticker", "MarketCap"])


def get_spot_price(ticker: yf.Ticker) -> float:
    """Fetch the latest available spot price with close fallback."""
    # Try the quick quote first because it is the fastest way to get the current price.
    try:
        last_price = ticker.fast_info.get("last_price")
        if last_price and last_price > 0:
            return float(last_price)
    except Exception:
        pass

    # If the quick quote fails, use the most recent daily close instead.
    history = ticker.history(period="1d", auto_adjust=False)
    if history.empty:
        raise ValueError("No one-day price data returned.")

    # Do not continue with missing prices because BSM needs real numbers.
    close = history["Close"].dropna()
    if close.empty:
        raise ValueError("One-day price data has no close values.")
    return float(close.iloc[-1])


def historical_volatility(symbol: str, window: int = 60) -> float | None:
    """Compute annualized historical volatility from recent daily log returns."""
    # Return None when Yahoo data is missing so the ticker can be skipped cleanly.
    try:
        data = yf.Ticker(symbol).history(period="1y", auto_adjust=True)
        if data.empty or "Close" not in data:
            return None

        # Use daily log returns as the base for historical volatility.
        close = data["Close"].dropna()
        returns = np.log(close / close.shift(1)).dropna()
        if len(returns) < window:
            return None

        # Convert the daily volatility into an annualized number.
        volatility = returns.rolling(window).std().iloc[-1] * np.sqrt(TRADING_DAYS)
        if pd.isna(volatility):
            return None
        return float(volatility)
    except Exception:
        return None


def years_to_expiry(expiry: str) -> float:
    """Convert an expiry date string into years with a small positive floor."""
    # Treat the option as expiring at market close on the listed date.
    expiry_close = datetime.strptime(expiry, "%Y-%m-%d") + timedelta(hours=16)
    seconds = (expiry_close - datetime.now()).total_seconds()

    # Same-day expiries still need a small positive time value for the formula.
    return max(seconds / (365 * 24 * 60 * 60), MIN_TIME_TO_MATURITY)


def option_market_price(option: pd.Series) -> float | None:
    """Prefer bid/ask midpoint, then fall back to last traded price."""
    # Yahoo sometimes leaves fields blank, so default missing quotes to zero.
    bid = float(option.get("bid", 0.0) or 0.0)
    ask = float(option.get("ask", 0.0) or 0.0)
    last = float(option.get("lastPrice", 0.0) or 0.0)

    # The bid/ask midpoint is usually better than an old last-traded price.
    if bid > 0 and ask > 0 and ask >= bid:
        return (bid + ask) / 2

    # If bid/ask is not usable, last trade is still better than no price.
    if last > 0:
        return last
    return None


def select_atm_option(options: pd.DataFrame, spot: float) -> pd.Series | None:
    """Select the nearest ATM option with enough market price to be useful."""
    # Skip chains that are empty or missing the basic price column.
    if options.empty or "lastPrice" not in options:
        return None

    # Add one clean market price column and filter out tiny quotes.
    tradable = options.copy()
    tradable["Market"] = tradable.apply(option_market_price, axis=1)
    tradable = tradable[tradable["Market"] > MIN_OPTION_PRICE]
    if tradable.empty:
        return None

    # The ATM contract is the strike closest to the current stock price.
    tradable["MoneynessDistance"] = (tradable["strike"] - spot).abs() / spot
    return tradable.sort_values("MoneynessDistance").iloc[0]


def fetch_atm_option_rows(symbol: str, market_cap: float) -> list[dict[str, float | str]]:
    """Fetch ATM call and put rows for one ticker."""
    # Normalize once, then reuse the same Yahoo Finance object for this ticker.
    yf_symbol = normalize_yfinance_ticker(symbol)
    ticker = yf.Ticker(yf_symbol)

    # Pull together the inputs needed to price the nearest expiry.
    try:
        spot = get_spot_price(ticker)
        expiries = ticker.options
        if not expiries:
            return []

        expiry = expiries[0]
        chain = ticker.option_chain(expiry)
        volatility = historical_volatility(yf_symbol)
        if volatility is None or volatility < MIN_VOLATILITY:
            return []

        time_to_maturity = years_to_expiry(expiry)
    except Exception:
        return []

    # Build one row for the best call and one row for the best put.
    rows: list[dict[str, float | str]] = []
    for options, option_type in ((chain.calls, "call"), (chain.puts, "put")):
        option = select_atm_option(options, spot)
        if option is None:
            continue

        # Store the values that main.py will price and write to Excel.
        strike = float(option["strike"])
        rows.append(
            {
                "Ticker": yf_symbol,
                "MarketCap": market_cap,
                "Type": option_type,
                "Spot": spot,
                "Strike": strike,
                "Expiry": expiry,
                "t(years)": time_to_maturity,
                "Volatility": volatility,
                "Market": float(option["Market"]),
            }
        )

    return rows
