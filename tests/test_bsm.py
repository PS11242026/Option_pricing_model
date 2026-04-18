import unittest

from bsm import call_price, option_price, put_price


class BlackScholesMertonTests(unittest.TestCase):
    """Regression tests for the Black-Scholes-Merton formula helpers."""

    def test_known_call_and_put_prices(self):
        # These benchmark values make sure the core formulas stay unchanged.
        self.assertAlmostEqual(call_price(100, 100, 1, 0.05, 0.2), 10.4506, places=4)
        self.assertAlmostEqual(put_price(100, 100, 1, 0.05, 0.2), 5.5735, places=4)

    def test_known_call_and_put_prices_with_dividend_yield(self):
        # BSM with continuous dividends should match known benchmark values.
        self.assertAlmostEqual(call_price(100, 100, 1, 0.05, 0.2, 0.02), 9.2270, places=4)
        self.assertAlmostEqual(put_price(100, 100, 1, 0.05, 0.2, 0.02), 6.3301, places=4)

    def test_option_type_dispatch(self):
        # The wrapper should match the direct call and put helpers.
        self.assertAlmostEqual(
            option_price(100, 100, 1, 0.05, 0.2, "call"),
            call_price(100, 100, 1, 0.05, 0.2),
        )
        self.assertAlmostEqual(
            option_price(100, 100, 1, 0.05, 0.2, "put"),
            put_price(100, 100, 1, 0.05, 0.2),
        )

    def test_invalid_inputs_raise(self):
        # Bad inputs should raise clear errors instead of returning nonsense.
        with self.assertRaises(ValueError):
            call_price(0, 100, 1, 0.05, 0.2)
        with self.assertRaises(ValueError):
            option_price(100, 100, 1, 0.05, 0.2, "straddle")


if __name__ == "__main__":
    unittest.main()
