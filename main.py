"""Top 25 option-pricing and mispricing detection engine."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from bsm import option_price
from data_fetch import fetch_active_option_rows, get_top25_companies


# Main settings for the pricing run and the Excel output.
RISK_FREE_RATE = 0.037
WORKBOOK_XLSX = "option_pricing_analysis.xlsx"
OUTPUT_COLUMNS = [
    "Ticker",
    "Type",
    "MarketOptionType",
    "ContractSymbol",
    "LastTradeDate",
    "Spot",
    "Strike",
    "Expiry",
    "t(years)",
    "Volatility",
    "DividendYield",
    "Market",
    "BSM",
    "AbsMis",
]


def print_header() -> None:
    """Print the model overview and formula definitions shown at startup."""
    # Print the assumptions up front so the terminal output explains itself.
    print(
        "Option Mispricing Detection Model for the 25 largest S&P 500 companies "
        "by market cap using Black-Scholes-Merton call and put pricing with "
        "each side's last actively traded option implied volatility"
    )
    print()
    print("Variable definitions:")
    print("C = call option price")
    print("P = put option price")
    print("S = spot price of the underlying stock")
    print("X = strike price")
    print("r = risk-free rate")
    print("T = time to expiration in years")
    print("sigma = option-chain implied volatility")
    print("q = continuous dividend yield")
    print("N(x) = cumulative distribution function of the standard normal distribution")
    print("e = Euler's number")
    print("d1 and d2 = Black-Scholes-Merton distribution inputs")
    print()
    print("C = S e^(-qT) N(d1) - X e^(-rT) N(d2)")
    print("P = X e^(-rT) N(-d2) - S e^(-qT) N(-d1)")
    print("d1 = [ln(S/X) + (r - q + sigma^2 / 2)T] / [sigma sqrt(T)]")
    print("d2 = d1 - sigma sqrt(T)")
    print()
    print("Outputs: Excel workbook with dataset and absolute mispricing chart")


def build_full_dataset() -> pd.DataFrame:
    """Fetch Top 25 active options and compute BSM mispricing metrics."""
    # Start with the fixed ticker list and collect the option rows we can price.
    companies = get_top25_companies()
    rows: list[dict[str, float | str]] = []

    # Handle each ticker on its own so one bad Yahoo response does not stop the run.
    for _, company in companies.iterrows():
        ticker = str(company["Ticker"])
        market_cap = float(company["MarketCap"])
        option_rows = fetch_active_option_rows(ticker, market_cap)
        if not option_rows:
            print(f"{ticker}: skipped, no usable active option data")
            continue

        # Add the BSM price and the gap between model price and market price.
        for row in option_rows:
            row["BSM"] = option_price(
                float(row["Spot"]),
                float(row["Strike"]),
                float(row["t(years)"]),
                RISK_FREE_RATE,
                float(row["Volatility"]),
                str(row["Type"]),
                float(row["DividendYield"]),
            )
            row["AbsMis"] = abs(float(row["BSM"]) - float(row["Market"]))
            rows.append(row)

    # No rows means Yahoo did not return enough usable data for this run.
    if not rows:
        return pd.DataFrame()

    # Keep the biggest companies first, with each ticker's call and put together.
    df = pd.DataFrame(rows)
    return df.sort_values(["MarketCap", "Ticker", "Type"], ascending=[False, True, True]).reset_index(drop=True)


def cleaned_output_df(df: pd.DataFrame) -> pd.DataFrame:
    """Return the workbook dataset without market cap."""
    # Drop the internal market-cap field and give the mispricing column a clearer name.
    output_df = df[OUTPUT_COLUMNS].copy()
    return output_df.rename(columns={"AbsMis": "Absolute Mispricing"})


def save_workbook(df: pd.DataFrame, path: str = WORKBOOK_XLSX) -> None:
    """Save Excel workbook with dataset and absolute mispricing chart."""
    # Choose a safe output path in case the workbook is already open in Excel.
    output_path = writable_workbook_path(path)

    # Prepare the data that appears in Excel and the labels used by the chart.
    output_df = cleaned_output_df(df)
    chart_df = output_df[["Ticker", "Spot", "Absolute Mispricing"]].copy()
    chart_df.insert(
        0,
        "Label",
        output_df["Ticker"]
        + " "
        + output_df["Type"].str.upper()
        + " $"
        + chart_df["Spot"].map(lambda value: f"{value:,.2f}"),
    )

    # Write the dataset, chart helper data, and chart into separate sheets.
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, sheet_name="Full Dataset", index=False)
        chart_df.to_excel(writer, sheet_name="Chart Data", index=False)

        # Grab the workbook and sheets so we can format them after writing.
        workbook = writer.book
        dataset_sheet = writer.sheets["Full Dataset"]
        chart_data_sheet = writer.sheets["Chart Data"]
        charts_sheet = workbook.add_worksheet("Charts")

        # Reuse these formats anywhere the workbook shows money, percent, or decimals.
        money_format = workbook.add_format({"num_format": "$#,##0.00"})
        percent_format = workbook.add_format({"num_format": "0.00%"})
        number_format = workbook.add_format({"num_format": "0.0000"})

        # Freeze headers and size columns so the workbook opens cleanly.
        dataset_sheet.freeze_panes(1, 0)
        chart_data_sheet.freeze_panes(1, 0)
        dataset_sheet.set_column("A:B", 12)
        dataset_sheet.set_column("C:C", 16)
        dataset_sheet.set_column("D:D", 22)
        dataset_sheet.set_column("E:E", 26)
        dataset_sheet.set_column("F:G", 12, money_format)
        dataset_sheet.set_column("H:H", 14)
        dataset_sheet.set_column("I:I", 12, number_format)
        dataset_sheet.set_column("J:K", 14, percent_format)
        dataset_sheet.set_column("L:N", 12, money_format)

        rows = len(chart_df) + 1
        abs_mis_col = chart_df.columns.get_loc("Absolute Mispricing")
        label_col = chart_df.columns.get_loc("Label")

        # Add a simple bar chart showing where the largest pricing gaps are.
        bar = workbook.add_chart({"type": "column"})
        bar.add_series(
            {
                "name": "Absolute Mispricing",
                "categories": ["Chart Data", 1, label_col, rows - 1, label_col],
                "values": ["Chart Data", 1, abs_mis_col, rows - 1, abs_mis_col],
            }
        )
        bar.set_title({"name": "Absolute Mispricing by Option"})
        bar.set_x_axis({"name": "Ticker / Spot Price"})
        bar.set_y_axis({"name": "Absolute Mispricing"})
        bar.set_legend({"none": True})
        charts_sheet.insert_chart("B2", bar, {"x_scale": 1.8, "y_scale": 1.35})

    print(f"\nSaved {output_path}")


def writable_workbook_path(path: str) -> str:
    """Return a writable workbook path, falling back when the target is open."""
    # Opening in append mode tells us whether Excel has locked the file.
    target = Path(path)
    try:
        with target.open("a+b"):
            pass
        return str(target)
    except PermissionError:
        # If it is locked, save a timestamped copy instead of crashing.
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = target.with_name(f"{target.stem}_{timestamp}{target.suffix}")
        print(f"\n{target} is locked, likely open in Excel. Saving to {fallback} instead.")
        return str(fallback)


def money(value: float) -> str:
    """Format a number as a two-decimal dollar amount without the dollar sign."""
    # Keep all printed dollar values in the same style.
    return f"{value:,.2f}"


def print_original_style_results(df: pd.DataFrame) -> None:
    """Print one original-style call/put summary per ticker."""
    # Print one combined call/put summary for each ticker.
    for ticker, group in df.groupby("Ticker", sort=False):
        call = group[group["Type"] == "call"]
        put = group[group["Type"] == "put"]

        # Pull the shared ticker fields and allow for a missing call or put.
        first = group.iloc[0]
        call_row = call.iloc[0] if not call.empty else None
        put_row = put.iloc[0] if not put.empty else None

        # Format each side only if that option type was found.
        market_call = money(float(call_row["Market"])) if call_row is not None else "n/a"
        market_put = money(float(put_row["Market"])) if put_row is not None else "n/a"
        bsm_call = money(float(call_row["BSM"])) if call_row is not None else "n/a"
        bsm_put = money(float(put_row["BSM"])) if put_row is not None else "n/a"
        call_mis = money(float(call_row["BSM"] - call_row["Market"])) if call_row is not None else "n/a"
        put_mis = money(float(put_row["BSM"] - put_row["Market"])) if put_row is not None else "n/a"

        # Show a compact summary for this ticker in the terminal.
        print(f"\n{ticker}")
        print(f"  Market cap: ${money(float(first['MarketCap']))}")
        active_call = (
            f"{call_row['ContractSymbol']} (last traded {call_row['LastTradeDate']})"
            if call_row is not None
            else "n/a"
        )
        active_put = (
            f"{put_row['ContractSymbol']} (last traded {put_row['LastTradeDate']})"
            if put_row is not None
            else "n/a"
        )
        print(f"  Active Yahoo call / put: {active_call} / {active_put}")
        print(f"  Spot price: ${money(float(first['Spot']))}")
        call_strike = money(float(call_row["Strike"])) if call_row is not None else "n/a"
        put_strike = money(float(put_row["Strike"])) if put_row is not None else "n/a"
        print(f"  Strike call / put: ${call_strike} / ${put_strike}")
        call_expiry = (
            f"{call_row['Expiry']} ({float(call_row['t(years)']):.4f} years)"
            if call_row is not None
            else "n/a"
        )
        put_expiry = (
            f"{put_row['Expiry']} ({float(put_row['t(years)']):.4f} years)"
            if put_row is not None
            else "n/a"
        )
        print(f"  Expiry call / put: {call_expiry} / {put_expiry}")
        call_volatility = f"{float(call_row['Volatility']):.2%}" if call_row is not None else "n/a"
        put_volatility = f"{float(put_row['Volatility']):.2%}" if put_row is not None else "n/a"
        print(f"  Implied volatility call / put: {call_volatility} / {put_volatility}")
        print(f"  Dividend yield: {float(first['DividendYield']):.2%}")
        print(f"  Market call / put: ${market_call} / ${market_put}")
        print(f"  BSM call / put: ${bsm_call} / ${bsm_put}")
        print(f"  Mispricing call / put: ${call_mis} / ${put_mis}")


def run_engine() -> None:
    """Run the complete command-line workflow."""
    # Show the setup first, then fetch live market data.
    print_header()
    print(f"Risk-free rate assumption: {RISK_FREE_RATE:.2%}")

    # Build results, print them, and save the workbook when data is available.
    df = build_full_dataset()
    if df.empty:
        print("\nNo data found")
        return

    print_original_style_results(df)
    save_workbook(df)


if __name__ == "__main__":
    run_engine()
