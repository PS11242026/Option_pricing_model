"""Web front end for the option pricing engine (Streamlit).

Run locally:  streamlit run streamlit_app.py

Tab 1 prices any European option with Black-Scholes-Merton.
Tab 2 scans live Yahoo Finance option chains for the chosen tickers and ranks
BSM-vs-market mispricing with the same engine the CLI uses. If live data is
unavailable (Yahoo throttles some cloud hosts), it falls back to the bundled
snapshot in sample_data/.
"""

from __future__ import annotations

import io
import math
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

from bsm import call_price, put_price
from main import RISK_FREE_RATE
from pricing_engine import build_tracker_table

SNAPSHOT = Path(__file__).resolve().parent / "sample_data" / "option_snapshot.csv"
DEFAULT_TICKERS = ["AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "META", "TSLA", "JPM"]
MAX_TICKERS = 15

st.set_page_config(page_title="Option Mispricing Model", page_icon="📈",
                   layout="wide")
st.title("Option Mispricing Detection Model")
st.caption("Black-Scholes-Merton pricing vs live market prices. Educational "
           "project: model signals are not investment advice.")

calc_tab, scan_tab, about_tab = st.tabs(
    ["🧮 Pricing calculator", "📡 Live mispricing scan", "ℹ️ How it works"])


# --------------------------------------------------------------- calculator
def bsm_both(S, K, T, r, sigma, q):
    return call_price(S, K, T, r, sigma, q), put_price(S, K, T, r, sigma, q)


with calc_tab:
    c1, c2, c3 = st.columns(3)
    S = c1.number_input("Spot price (S)", min_value=0.01, value=100.0, step=1.0)
    K = c1.number_input("Strike price (K)", min_value=0.01, value=100.0, step=1.0)
    days = c2.number_input("Days to expiry", min_value=1, value=30, step=1)
    sigma = c2.number_input("Volatility σ (%)", min_value=0.1, value=25.0,
                            step=0.5) / 100
    r = c3.number_input("Risk-free rate r (%)", value=RISK_FREE_RATE * 100,
                        step=0.1) / 100
    q = c3.number_input("Dividend yield q (%)", min_value=0.0, value=0.0,
                        step=0.1) / 100
    T = days / 365

    call, put = bsm_both(S, K, T, r, sigma, q)
    parity_gap = (call - put) - (S * math.exp(-q * T) - K * math.exp(-r * T))
    m1, m2, m3 = st.columns(3)
    m1.metric("Call value", f"${call:,.4f}")
    m2.metric("Put value", f"${put:,.4f}")
    m3.metric("Put-call parity gap", f"{parity_gap:.2e}",
              help="C − P − (S·e^(−qT) − K·e^(−rT)); ≈ 0 when the model is "
                   "internally consistent.")

    spots = np.linspace(max(0.01, S * 0.5), S * 1.5, 121)
    curve = pd.DataFrame(
        [bsm_both(s, K, T, r, sigma, q) for s in spots],
        columns=["Call", "Put"], index=pd.Index(spots.round(2), name="Spot"))
    st.markdown("**Option value across spot prices** (strike, expiry, σ, r, "
                "q held fixed)")
    st.line_chart(curve, height=320)


# --------------------------------------------------------------- live scan
@st.cache_data(ttl=900, show_spinner=False)
def scan_live(tickers: tuple[str, ...], r: float) -> pd.DataFrame:
    """Same steps as main.build_full_dataset, for a chosen ticker list."""
    from bsm import option_price
    from data_fetch import fetch_active_option_rows

    rows = []
    for ticker in tickers:
        for row in fetch_active_option_rows(ticker, market_cap=0.0):
            row["BSM"] = option_price(
                float(row["Spot"]), float(row["Strike"]),
                float(row["t(years)"]), r, float(row["Volatility"]),
                str(row["Type"]), float(row["DividendYield"]))
            row["AbsMis"] = abs(row["BSM"] - float(row["Market"]))
            rows.append(row)
    return pd.DataFrame(rows)


def load_snapshot() -> pd.DataFrame:
    return pd.read_csv(SNAPSHOT)


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name="Mispricing")
    return buf.getvalue()


with scan_tab:
    st.markdown("Pulls the **last actively traded call and put** at the "
                "nearest expiry for each ticker, prices them with BSM using "
                "the option's own implied volatility, and ranks contracts by "
                "absolute mispricing.")
    tickers = st.multiselect(
        "Tickers", options=sorted(set(DEFAULT_TICKERS + [
            "BRK-B", "V", "UNH", "XOM", "JNJ", "WMT", "MA", "PG", "HD", "COST",
            "NFLX", "AMD", "ORCL", "KO", "PEP", "BAC", "AVGO", "LLY"])),
        default=DEFAULT_TICKERS, max_selections=MAX_TICKERS)
    custom = st.text_input("…or add symbols (comma separated)", "")
    tickers = list(dict.fromkeys(
        tickers + [t.strip().upper() for t in custom.split(",") if t.strip()]
    ))[:MAX_TICKERS]
    top_n = st.slider("Contracts to show", 5, 50, 20)
    run = st.button("Run live scan", type="primary", disabled=not tickers)

    if run:
        with st.spinner(f"Fetching option chains for {len(tickers)} tickers…"):
            try:
                data = scan_live(tuple(tickers), RISK_FREE_RATE)
            except Exception as exc:  # noqa: BLE001 - any live-data failure -> snapshot
                data = pd.DataFrame()
                st.caption(f"Live fetch error: {type(exc).__name__}")
        source = "live"
        if data.empty:
            data = load_snapshot()
            source = "snapshot"
        st.session_state["scan"] = (data, source)

    if "scan" in st.session_state:
        data, source = st.session_state["scan"]
        if source == "snapshot":
            st.warning("Live market data was unavailable from this host, so "
                       "this shows the bundled snapshot of the top-100 S&P 500 "
                       f"names (last trade {data['LastTradeDate'].max()[:10]}).")
        else:
            st.success(f"Live data for {data['Ticker'].nunique()} tickers "
                       "(cached for 15 minutes).")
        tracker = build_tracker_table(data, top_n=top_n)
        st.dataframe(
            tracker, hide_index=True, width="stretch",
            column_config={
                c: st.column_config.NumberColumn(format="$%.2f")
                for c in ("Spot", "Strike", "Market", "BSM", "Spread", "AbsMis")})
        chart = tracker.assign(Contract=tracker["Ticker"] + " " +
                               tracker["Type"].str.upper())
        st.markdown("**BSM minus market price** (positive = market below model)")
        st.bar_chart(chart.set_index("Contract")["Spread"], height=320)
        st.download_button("Download full dataset (Excel)", to_excel(data),
                           file_name="option_mispricing.xlsx")


# --------------------------------------------------------------- about
with about_tab:
    st.markdown(
        """
**Model.** European Black-Scholes-Merton with continuous dividend yield:
`C = S·e^(−qT)·N(d1) − K·e^(−rT)·N(d2)`, `P = K·e^(−rT)·N(−d2) − S·e^(−qT)·N(−d1)`.

**Data.** Yahoo Finance spot, dividend yield and option chains (`yfinance`);
the CLI ranks the top-100 S&P 500 companies by live Reuters market cap.

**Signal.** Spread = BSM − market. Positive → market below model (long);
negative → market above model (short). Contracts are filtered for a minimum
price, minimum implied volatility and a recent last trade.

**Limits.** European formula applied to American-style US equity options;
implied vol from the quote itself makes the spread a consistency check rather
than a forecast. Not investment advice.

Source code and the CLI version: see the repository README.
"""
    )
