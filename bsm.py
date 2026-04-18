"""Black Scholes Merton pricing formulas for European options."""

from __future__ import annotations

import math

from scipy.stats import norm


def _d1_d2(S: float, K: float, T: float, r: float, sigma: float) -> tuple[float, float]:
    """Check the inputs and calculate the two BSM helper values."""
    # These checks catch values that would break the math below.
    if S <= 0:
        raise ValueError("Spot price must be positive.")
    if K <= 0:
        raise ValueError("Strike price must be positive.")
    if T <= 0:
        raise ValueError("Time to maturity must be positive.")
    if sigma <= 0:
        raise ValueError("Volatility must be positive.")

    # d1 and d2 are the values used inside the standard normal CDF.
    sqrt_t = math.sqrt(T)
    d1 = (math.log(S / K) + (r + 0.5 * sigma**2) * T) / (sigma * sqrt_t)
    d2 = d1 - sigma * sqrt_t
    return d1, d2


def call_price(S: float, K: float, T: float, r: float, sigma: float) -> float:
    """Calculate a European call option price using Black Scholes Merton."""
    # Use the same d1/d2 calculation for every BSM formula.
    d1, d2 = _d1_d2(S, K, T, r, sigma)

    # A call gives the holder the right to buy at the strike price.
    return S * norm.cdf(d1) - K * math.exp(-r * T) * norm.cdf(d2)


def put_price(S: float, K: float, T: float, r: float, sigma: float) -> float:
    """Calculate a European put option price using Black-Scholes-Merton."""
    # Use the same d1/d2 calculation for every BSM formula.
    d1, d2 = _d1_d2(S, K, T, r, sigma)

    # A put gives the holder the right to sell at the strike price.
    return K * math.exp(-r * T) * norm.cdf(-d2) - S * norm.cdf(-d1)


def option_price(S: float, K: float, T: float, r: float, sigma: float, option_type: str) -> float:
    """Calculate a European option price for ``call`` or ``put``."""
    # Let inputs like "Call", "CALL", and "call" behave the same way.
    option_type = option_type.lower()

    # Pick the matching formula based on the option type.
    if option_type == "call":
        return call_price(S, K, T, r, sigma)
    if option_type == "put":
        return put_price(S, K, T, r, sigma)

    # Anything else is probably a typo, so stop instead of giving a bad result.
    raise ValueError("option_type must be 'call' or 'put'.")
