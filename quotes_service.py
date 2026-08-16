"""Live quotes for the header watchlist strip, via yfinance. Optional
dependency — if yfinance isn't installed, every quote comes back tagged
with an error instead of crashing the bridge."""
import time

_CACHE = {}       # ticker -> (fetched_at, quote_dict)
_CACHE_TTL = 20   # seconds — cheap protection against a hot refresh loop


def fetch_quotes(tickers):
    """Returns one {"ticker", "price", "change", "changePct", "error"} dict per input
    ticker, same order as given. Cached briefly so rapid reloads don't hammer Yahoo."""
    if not tickers:
        return []

    now = time.time()
    stale = [t for t in tickers if t not in _CACHE or now - _CACHE[t][0] > _CACHE_TTL]

    if stale:
        try:
            import yfinance as yf
        except ImportError:
            for t in stale:
                _CACHE[t] = (now, _error(t, "yfinance isn't installed on the server"))
        else:
            for t in stale:
                _CACHE[t] = (now, _fetch_one(yf, t))

    return [_CACHE[t][1] for t in tickers]


def _fetch_one(yf, ticker):
    try:
        fi = yf.Ticker(ticker).fast_info
        price = fi.last_price
        prev = fi.previous_close
    except Exception as exc:                                      # noqa: BLE001
        return _error(ticker, str(exc) or "Couldn't reach the quote server.")

    if price is None or prev is None:
        return _error(ticker, "No quote found for this ticker.")

    change = price - prev
    change_pct = (change / prev * 100) if prev else 0.0
    return {"ticker": ticker, "price": round(price, 2), "change": round(change, 2),
            "changePct": round(change_pct, 2), "error": None}


def _error(ticker, message):
    return {"ticker": ticker, "price": None, "change": None, "changePct": None, "error": message}


def fetch_history(ticker, period="3mo"):
    """Returns (points, error) — points is a list of {"date": "YYYY-MM-DD", "close": float},
    oldest first. Not cached: only fetched when a chart is actually opened."""
    try:
        import yfinance as yf
    except ImportError:
        return None, "yfinance isn't installed on the server"

    try:
        hist = yf.Ticker(ticker).history(period=period)
    except Exception as exc:                                        # noqa: BLE001
        return None, str(exc) or "Couldn't reach the quote server."

    if hist.empty:
        return None, "No price history found for this ticker."

    points = [
        {"date": idx.strftime("%Y-%m-%d"), "close": round(float(row["Close"]), 2)}
        for idx, row in hist.iterrows()
    ]
    return points, None
