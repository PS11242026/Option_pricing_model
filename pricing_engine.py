"""Command-line tracker for the most mispriced options in the dataset."""

from __future__ import annotations

import argparse
from datetime import date

import pandas as pd

from main import RISK_FREE_RATE, run_engine


DEFAULT_TOP_N = 50


def mispricing_spread(row: pd.Series) -> float:
    """Return theoretical minus market price for one option row."""
    return float(row["BSM"]) - float(row["Market"])


def signal_from_spread(spread: float) -> str:
    """Map BSM-vs-market spread into a directional signal."""
    if spread > 0:
        return "BUY"
    if spread < 0:
        return "SELL"
    return "HOLD"


def arbitrage_from_spread(option_type: str, spread: float) -> str:
    """Describe the model-driven opportunity for one contract."""
    option_label = option_type.upper()
    if spread > 0:
        return f"Long {option_label}; market below BSM"
    if spread < 0:
        return f"Short {option_label}; market above BSM"
    return f"No edge in {option_label}"


def moneyness_from_contract(option_type: str, spot: float, strike: float) -> str:
    """Classify an option as in/out/at the money from spot vs strike."""
    option_name = str(option_type).lower()
    if spot == strike:
        return "AT-THE-MONEY"
    if option_name == "call":
        return "IN-THE-MONEY" if spot > strike else "OUT-OF-THE-MONEY"
    if option_name == "put":
        return "IN-THE-MONEY" if spot < strike else "OUT-OF-THE-MONEY"
    return "UNKNOWN"


def build_tracker_table(df: pd.DataFrame, top_n: int = DEFAULT_TOP_N) -> pd.DataFrame:
    """Return the top mispriced contracts ranked by absolute mispricing."""
    if df.empty:
        return pd.DataFrame()

    tracker_df = df.copy()
    if "AbsMis" not in tracker_df.columns:
        tracker_df["AbsMis"] = (tracker_df["BSM"] - tracker_df["Market"]).abs()

    tracker_df["Spread"] = tracker_df.apply(mispricing_spread, axis=1)
    tracker_df["Signal"] = tracker_df["Spread"].apply(signal_from_spread)
    tracker_df["Arbitrage"] = tracker_df.apply(
        lambda row: arbitrage_from_spread(str(row["Type"]), float(row["Spread"])),
        axis=1,
    )
    tracker_df["Moneyness"] = tracker_df.apply(
        lambda row: moneyness_from_contract(
            str(row["Type"]),
            float(row["Spot"]),
            float(row["Strike"]),
        ),
        axis=1,
    )

    ranked = tracker_df.sort_values(
        ["AbsMis", "Ticker", "Type"],
        ascending=[False, True, True],
    ).head(top_n)

    columns = [
        "Ticker",
        "Type",
        "ContractSymbol",
        "Spot",
        "Strike",
        "Expiry",
        "Market",
        "BSM",
        "Spread",
        "AbsMis",
        "Signal",
        "Arbitrage",
        "Moneyness",
    ]
    return ranked[columns].reset_index(drop=True)


def money(value: float) -> str:
    return f"${value:,.2f}"


def signed_money(value: float) -> str:
    return f"{value:+,.2f}"


def percent(value: float) -> str:
    return f"{value:.1%}"


def format_expiry(expiry: object) -> str:
    parsed = pd.to_datetime(expiry, errors="coerce")
    if pd.isna(parsed):
        return str(expiry)
    return parsed.strftime("%b %d, %Y")


def days_to_expiry(expiry: object) -> int | None:
    parsed = pd.to_datetime(expiry, errors="coerce")
    if pd.isna(parsed):
        return None
    return max((parsed.date() - date.today()).days, 0)


def contract_label(row: pd.Series) -> str:
    dte = days_to_expiry(row["Expiry"])
    expiry = format_expiry(row["Expiry"])
    dte_text = f"{dte}d" if dte is not None else "n/a"
    return (
        f"{row['Ticker']} {str(row['Type']).upper()} {money(float(row['Strike']))} "
        f"exp {expiry} ({dte_text})"
    )


def mispricing_percent(spread: float, market_price: float) -> float:
    if market_price == 0:
        return 0.0
    return abs(spread) / market_price


def display_tracker_table(tracker_df: pd.DataFrame) -> pd.DataFrame:
    """Return a string-friendly view of the ranked contracts for CLI display."""
    if tracker_df.empty:
        return pd.DataFrame()

    display_df = tracker_df.copy().reset_index(drop=True)
    display_df.insert(0, "Rank", display_df.index + 1)
    display_df["DTE"] = display_df["Expiry"].apply(days_to_expiry)
    display_df["Expiry"] = display_df["Expiry"].apply(format_expiry)
    display_df["Spot"] = display_df["Spot"].apply(lambda value: money(float(value)))
    display_df["Strike"] = display_df["Strike"].apply(lambda value: money(float(value)))
    display_df["Market"] = display_df["Market"].apply(lambda value: money(float(value)))
    display_df["BSM"] = display_df["BSM"].apply(lambda value: money(float(value)))
    display_df["Spread"] = display_df["Spread"].apply(lambda value: signed_money(float(value)))
    display_df["AbsMis"] = display_df["AbsMis"].apply(lambda value: money(float(value)))
    display_df["DTE"] = display_df["DTE"].apply(
        lambda value: str(int(value)) if value is not None else "n/a"
    )

    return display_df[
        [
            "Rank",
            "Ticker",
            "Type",
            "ContractSymbol",
            "Spot",
            "Strike",
            "Expiry",
            "DTE",
            "Market",
            "BSM",
            "Spread",
            "AbsMis",
            "Signal",
            "Moneyness",
            "Arbitrage",
        ]
    ].rename(
        columns={
            "Type": "Option",
            "ContractSymbol": "Symbol",
            "Spread": "Edge",
            "AbsMis": "Abs Mis",
            "Arbitrage": "Trade",
        }
    )


def print_tracker_table(tracker_df: pd.DataFrame, top_n: int, risk_free_rate: float) -> None:
    """Render top option opportunities as a summary plus table."""
    print("Portfolio Tracker: Top Mispriced Options")
    print(f"Risk-free rate: {risk_free_rate:.2%}")
    print(f"Contracts shown: {min(top_n, len(tracker_df))}")

    if tracker_df.empty:
        print("\nNo option data available.")
        return

    summary = {
        "BUY": int((tracker_df["Signal"] == "BUY").sum()),
        "SELL": int((tracker_df["Signal"] == "SELL").sum()),
        "HOLD": int((tracker_df["Signal"] == "HOLD").sum()),
    }
    moneyness_summary = {
        "IN-THE-MONEY": int((tracker_df["Moneyness"] == "IN-THE-MONEY").sum()),
        "OUT-OF-THE-MONEY": int((tracker_df["Moneyness"] == "OUT-OF-THE-MONEY").sum()),
        "AT-THE-MONEY": int((tracker_df["Moneyness"] == "AT-THE-MONEY").sum()),
    }
    strongest = tracker_df.iloc[0]

    print()
    print(f"Signal mix: BUY {summary['BUY']} | SELL {summary['SELL']} | HOLD {summary['HOLD']}")
    print(
        "Largest edge: "
        f"{contract_label(strongest)} | "
        f"{strongest['Signal']} {money(float(strongest['AbsMis']))} "
        f"({percent(mispricing_percent(float(strongest['Spread']), float(strongest['Market'])))} of market)"
    )
    print(
        "Moneyness mix: "
        f"ITM {moneyness_summary['IN-THE-MONEY']} | "
        f"OTM {moneyness_summary['OUT-OF-THE-MONEY']} | "
        f"ATM {moneyness_summary['AT-THE-MONEY']}"
    )
    print()
    print(display_tracker_table(tracker_df).to_string(index=False))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Track the top mispriced options using the existing Yahoo Finance "
            "and Black-Scholes-Merton pricing pipeline."
        )
    )
    parser.add_argument(
        "--top",
        type=int,
        default=DEFAULT_TOP_N,
        help="Number of mispriced contracts to show (default: 50).",
    )
    parser.add_argument(
        "--rate",
        type=float,
        default=RISK_FREE_RATE,
        help="Risk-free rate used for BSM pricing as a decimal (default: 0.037).",
    )
    return parser.parse_args()


def run_tracker() -> None:
    args = parse_args()
    top_n = max(1, args.top)
    dataset = run_engine(risk_free_rate=float(args.rate), include_tracker=False)
    tracker_df = build_tracker_table(dataset, top_n=top_n)
    print_tracker_table(tracker_df, top_n=top_n, risk_free_rate=float(args.rate))


if __name__ == "__main__":
    run_tracker()
