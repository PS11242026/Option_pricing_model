"""Market data retrieval for the option-pricing engine."""

from __future__ import annotations

import re
from io import StringIO
from datetime import datetime, timedelta
from pathlib import Path
from urllib.request import Request, urlopen

import pandas as pd
import yfinance as yf


# Settings used while filtering market data and choosing liquid contracts.
MIN_OPTION_PRICE = 0.50
MIN_VOLATILITY = 0.05
MIN_TIME_TO_MATURITY = 1 / (365 * 24)
YFINANCE_CACHE_DIR = Path(__file__).resolve().parent / ".yfinance_cache"
REUTERS_MARKET_CAP_CSV_URL = (
    "https://graphics.thomsonreuters.com/automated-charts/production/"
    "market-cap-ranking-table/data.csv"
)
REUTERS_USER_AGENT = "Mozilla/5.0 option-pricing-model/0.1"

YFINANCE_CACHE_DIR.mkdir(exist_ok=True)
yf.set_tz_cache_location(str(YFINANCE_CACHE_DIR))


def normalize_yfinance_ticker(symbol: str) -> str:
    """Normalize ticker formats for yfinance compatibility."""
    # Yahoo Finance expects share classes like BRK-B instead of BRK.B.
    return symbol.strip().upper().replace(".", "-")


def reuters_symbol_to_yfinance(symbol: str) -> str:
    """Convert a Reuters company symbol into the yfinance ticker format."""
    base_symbol = symbol.strip().split(".", maxsplit=1)[0]
    share_class = re.fullmatch(r"([A-Z]+)([a-z])", base_symbol)
    if share_class:
        base_symbol = f"{share_class.group(1)}-{share_class.group(2).upper()}"
    return normalize_yfinance_ticker(base_symbol)


def ticker_from_reuters_stock_link(stock: str) -> str | None:
    """Extract the ticker from a Reuters company markdown link."""
    match = re.search(r"/companies/([A-Za-z0-9.]+)", str(stock))
    if not match:
        return None
    return reuters_symbol_to_yfinance(match.group(1))


def parse_market_cap(value: object) -> float | None:
    """Parse Reuters market-cap values into raw dollars."""
    if value is None or pd.isna(value):
        return None

    if isinstance(value, int | float):
        market_cap = float(value)
        return market_cap if market_cap > 0 else None

    text = str(value).strip().replace(",", "")
    if not text:
        return None

    try:
        market_cap = float(text)
        return market_cap if market_cap > 0 else None
    except ValueError:
        pass

    match = re.fullmatch(r"\$?([0-9]+(?:\.[0-9]+)?)\s*(trillion|billion|million)", text, re.I)
    if not match:
        return None

    multiplier = {
        "trillion": 1e12,
        "billion": 1e9,
        "million": 1e6,
    }[match.group(2).lower()]
    return float(match.group(1)) * multiplier


DEFAULT_COMPANY_LIMIT = 100


def parse_reuters_market_cap_table(
    csv_text: str, limit: int = DEFAULT_COMPANY_LIMIT
) -> pd.DataFrame:
    """Parse Reuters' live market-cap ranking CSV into ticker and market-cap rows."""
    ranking = pd.read_csv(StringIO(csv_text))
    if "Stock" not in ranking or "Market Value" not in ranking:
        raise ValueError("Reuters market-cap CSV is missing expected columns.")

    raw_market_cap_column = next(
        (column for column in ranking.columns if not str(column).strip()),
        "Market Value",
    )
    parsed = pd.DataFrame(
        {
            "Ticker": ranking["Stock"].apply(ticker_from_reuters_stock_link),
            "MarketCap": ranking[raw_market_cap_column].apply(parse_market_cap),
        }
    )
    parsed = parsed.dropna(subset=["Ticker", "MarketCap"]).copy()
    parsed["MarketCap"] = parsed["MarketCap"].astype(float)
    return parsed.head(limit).reset_index(drop=True)


def fetch_reuters_market_cap_csv() -> str:
    """Fetch Reuters' live S&P 500 market-cap ranking CSV."""
    request = Request(REUTERS_MARKET_CAP_CSV_URL, headers={"User-Agent": REUTERS_USER_AGENT})
    with urlopen(request, timeout=30) as response:
        return response.read().decode("utf-8-sig")


def get_top_companies(limit: int = DEFAULT_COMPANY_LIMIT) -> pd.DataFrame:
    """Return current S&P 500 companies by Reuters market-cap ranking."""
    return parse_reuters_market_cap_table(fetch_reuters_market_cap_csv(), limit=limit)


def get_top25_companies() -> pd.DataFrame:
    """Backward-compatible wrapper for the current top-company universe."""
    return get_top_companies()


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


def get_dividend_yield(ticker: yf.Ticker) -> float:
    """Fetch the current annual dividend yield as a decimal, defaulting to zero."""
    try:
        info = ticker.info
    except Exception:
        return 0.0

    dividend_yield = info.get("trailingAnnualDividendYield")
    if dividend_yield is None or pd.isna(dividend_yield):
        dividend_yield = info.get("dividendYield")
        if dividend_yield is None or pd.isna(dividend_yield):
            return 0.0

        # Yahoo's dividendYield field is commonly in percentage-point units:
        # 0.38 means 0.38%, not 38%.
        dividend_yield = float(dividend_yield) / 100

    dividend_yield = float(dividend_yield)
    if dividend_yield <= 0:
        return 0.0
    return dividend_yield


def years_to_expiry(expiry: str) -> float:
    """Convert an expiry date string into years with a small positive floor."""
    # Treat the option as expiring at market close on the listed date.
    expiry_close = datetime.strptime(expiry, "%Y-%m-%d") + timedelta(hours=16)
    seconds = (expiry_close - datetime.now()).total_seconds()

    # Same-day expiries still need a small positive time value for the formula.
    return max(seconds / (365 * 24 * 60 * 60), MIN_TIME_TO_MATURITY)


def option_market_price(option: pd.Series) -> float | None:
    """Return Yahoo's last traded option price when it is usable."""
    last = float(option.get("lastPrice", 0.0) or 0.0)

    if last > 0:
        return last
    return None


def option_implied_volatility(option: pd.Series) -> float | None:
    """Return a usable implied volatility from a Yahoo option row."""
    volatility = option.get("impliedVolatility")
    if volatility is None or pd.isna(volatility):
        return None

    volatility = float(volatility)
    if volatility < MIN_VOLATILITY:
        return None
    return volatility


def option_numeric_column(options: pd.DataFrame, column: str) -> pd.Series:
    """Return a numeric option-chain column, defaulting missing columns to zero."""
    if column in options:
        return pd.to_numeric(options[column], errors="coerce").fillna(0)
    return pd.Series(0, index=options.index, dtype="float64")


def prepare_active_options(options: pd.DataFrame, expiry: str, option_type: str) -> pd.DataFrame:
    """Return usable option rows annotated for active-contract selection."""
    if options.empty or "lastPrice" not in options or "lastTradeDate" not in options:
        return pd.DataFrame()

    tradable = options.copy()
    tradable["Market"] = tradable.apply(option_market_price, axis=1)
    tradable["Volatility"] = tradable.apply(option_implied_volatility, axis=1)
    tradable["LastTradeDate"] = pd.to_datetime(tradable["lastTradeDate"], errors="coerce", utc=True)
    tradable["Volume"] = option_numeric_column(tradable, "volume")
    tradable["OpenInterest"] = option_numeric_column(tradable, "openInterest")
    tradable["Market"] = pd.to_numeric(tradable["Market"], errors="coerce")
    tradable["Volatility"] = pd.to_numeric(tradable["Volatility"], errors="coerce")
    tradable = tradable[
        (tradable["Market"] > MIN_OPTION_PRICE)
        & tradable["Volatility"].notna()
        & tradable["LastTradeDate"].notna()
    ].copy()
    if tradable.empty:
        return pd.DataFrame()

    tradable["Expiry"] = expiry
    tradable["OptionType"] = option_type
    return tradable


def select_last_active_option(
    ticker: yf.Ticker, expiries: tuple[str, ...] | list[str]
) -> pd.Series | None:
    """Select the most recently traded Yahoo option contract for one ticker."""
    candidates: list[pd.DataFrame] = []

    for expiry in expiries:
        try:
            chain = ticker.option_chain(expiry)
        except Exception:
            continue

        for options, option_type in ((chain.calls, "call"), (chain.puts, "put")):
            active = prepare_active_options(options, expiry, option_type)
            if not active.empty:
                candidates.append(active)

    if not candidates:
        return None

    tradable = pd.concat(candidates, ignore_index=True)
    return tradable.sort_values(
        ["LastTradeDate", "Volume", "OpenInterest", "Market"],
        ascending=[False, False, False, False],
    ).iloc[0]


def select_last_active_options_by_type(
    ticker: yf.Ticker, expiries: tuple[str, ...] | list[str]
) -> dict[str, pd.Series]:
    """Select the most recently traded Yahoo call and put contracts for one ticker."""
    candidates: dict[str, list[pd.DataFrame]] = {"call": [], "put": []}

    for expiry in expiries:
        try:
            chain = ticker.option_chain(expiry)
        except Exception:
            continue

        for options, option_type in ((chain.calls, "call"), (chain.puts, "put")):
            active = prepare_active_options(options, expiry, option_type)
            if not active.empty:
                candidates[option_type].append(active)

    selected: dict[str, pd.Series] = {}
    for option_type, option_frames in candidates.items():
        if not option_frames:
            continue

        tradable = pd.concat(option_frames, ignore_index=True)
        selected[option_type] = tradable.sort_values(
            ["LastTradeDate", "Volume", "OpenInterest", "Market"],
            ascending=[False, False, False, False],
        ).iloc[0]

    return selected


def fetch_active_option_rows(symbol: str, market_cap: float) -> list[dict[str, float | str]]:
    """Fetch rows priced from the last actively traded call and put for one ticker."""
    # Normalize once, then reuse the same Yahoo Finance object for this ticker.
    yf_symbol = normalize_yfinance_ticker(symbol)
    ticker = yf.Ticker(yf_symbol)

    # Pull together the inputs needed to price the nearest expiry.
    try:
        spot = get_spot_price(ticker)
        dividend_yield = get_dividend_yield(ticker)
        expiries = ticker.options
        if not expiries:
            return []

        selected_options = select_last_active_options_by_type(ticker, expiries)
        if not selected_options:
            return []
    except Exception:
        return []

    # Price each BSM side against the market price and IV of the matching option type.
    rows: list[dict[str, float | str]] = []
    for option_type in ("call", "put"):
        option = selected_options.get(option_type)
        if option is None:
            continue

        expiry = str(option["Expiry"])
        strike = float(option["strike"])
        contract_symbol = str(option.get("contractSymbol", ""))
        last_trade_date = option["LastTradeDate"].isoformat()
        rows.append(
            {
                "Ticker": yf_symbol,
                "MarketCap": market_cap,
                "Type": option_type,
                "MarketOptionType": str(option["OptionType"]),
                "ContractSymbol": contract_symbol,
                "LastTradeDate": last_trade_date,
                "Spot": spot,
                "Strike": strike,
                "Expiry": expiry,
                "t(years)": years_to_expiry(expiry),
                "Volatility": float(option["Volatility"]),
                "DividendYield": dividend_yield,
                "Market": float(option["Market"]),
            }
        )

    return rows


def fetch_atm_option_rows(symbol: str, market_cap: float) -> list[dict[str, float | str]]:
    """Backward-compatible wrapper for callers using the old ATM function name."""
    return fetch_active_option_rows(symbol, market_cap)
