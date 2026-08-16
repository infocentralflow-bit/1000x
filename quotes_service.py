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


# range key -> (yfinance period, interval). Short ranges need an intraday
# interval or "1d"/"5d" would come back as one or five points, not a chart.
HISTORY_RANGES = {
    "1d":  ("1d", "5m"),
    "5d":  ("5d", "15m"),
    "1mo": ("1mo", "1d"),
    "3mo": ("3mo", "1d"),
    "6mo": ("6mo", "1d"),
    "1y":  ("1y", "1d"),
    "ytd": ("ytd", "1d"),
}


def fetch_history(ticker, range_key="3mo"):
    """Returns (points, error) — points is a list of {"date", "close"}, oldest first.
    "date" is "YYYY-MM-DD" for daily+ intervals, "YYYY-MM-DD HH:MM" for intraday ones.
    Not cached: only fetched when a chart is actually opened."""
    if range_key not in HISTORY_RANGES:
        return None, "Invalid range."
    period, interval = HISTORY_RANGES[range_key]

    try:
        import yfinance as yf
    except ImportError:
        return None, "yfinance isn't installed on the server"

    try:
        hist = yf.Ticker(ticker).history(period=period, interval=interval)
    except Exception as exc:                                        # noqa: BLE001
        return None, str(exc) or "Couldn't reach the quote server."

    if hist.empty:
        return None, "No price history found for this ticker."

    date_fmt = "%Y-%m-%d %H:%M" if interval.endswith(("m", "h")) else "%Y-%m-%d"
    points = [
        {"date": idx.strftime(date_fmt), "close": round(float(row["Close"]), 2)}
        for idx, row in hist.iterrows()
    ]
    return points, None
