"""Web front end: calculator matches bsm.py, and the scan falls back to the
bundled snapshot when live data is unavailable (no network used)."""

from pathlib import Path

import pandas as pd
import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest

from bsm import call_price, put_price

APP = str(Path(__file__).resolve().parents[1] / "streamlit_app.py")


def test_calculator_matches_bsm():
    at = AppTest.from_file(APP, default_timeout=60).run()
    assert not at.exception
    values = {m.label: m.value for m in at.metric}
    call = call_price(100, 100, 30 / 365, 0.037, 0.25, 0.0)
    put = put_price(100, 100, 30 / 365, 0.037, 0.25, 0.0)
    assert values["Call value"] == f"${call:,.4f}"
    assert values["Put value"] == f"${put:,.4f}"


def test_snapshot_is_valid_engine_input():
    snap = pd.read_csv(Path(APP).parent / "sample_data" / "option_snapshot.csv")
    for col in ("Ticker", "Type", "Spot", "Strike", "Market", "BSM", "AbsMis"):
        assert col in snap.columns
    assert len(snap) > 50


def test_scan_falls_back_to_snapshot(monkeypatch):
    import data_fetch
    monkeypatch.setattr(data_fetch, "fetch_active_option_rows",
                        lambda *a, **k: [])
    at = AppTest.from_file(APP, default_timeout=60).run()
    run = next(b for b in at.button if b.label == "Run live scan")
    at = run.click().run()
    assert not at.exception
    assert any("snapshot" in w.value for w in at.warning)
    assert len(at.dataframe) == 1


def test_main_py_serves_web_app_under_streamlit():
    """A hosted app whose entry point is main.py must show the web app,
    not start the long-running CLI scan."""
    main_py = str(Path(APP).parent / "main.py")
    at = AppTest.from_file(main_py, default_timeout=60).run()
    assert not at.exception
    assert at.title[0].value == "Option Mispricing Detection Model"
    assert {m.label for m in at.metric} >= {"Call value", "Put value"}
