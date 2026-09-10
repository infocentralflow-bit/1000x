"""Live quotes for the header watchlist strip, via yfinance. Optional
dependency — if yfinance isn't installed, every quote comes back tagged
with an error instead of crashing the bridge."""
import concurrent.futures
import math
import time

_CACHE = {}       # ticker -> (fetched_at, quote_dict)
_CACHE_TTL = 20   # seconds — cheap protection against a hot refresh loop
_FETCH_TIMEOUT = 8   # seconds — a stalled/rate-limited symbol shouldn't block the rest of the batch


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
            # Fetched concurrently, each with its own timeout — yfinance has no
            # built-in per-request deadline, so a single stalled/rate-limited
            # symbol (fetched serially) could otherwise hang the whole batch
            # response, silently blocking every other ticker in it too.
            # Not a context manager on purpose: `with ... as pool` calls
            # shutdown(wait=True) on exit, which blocks until every submitted
            # thread finishes — including the one we just gave up waiting on —
            # defeating the timeout entirely. shutdown(wait=False) lets this
            # function return promptly; the stuck thread is orphaned (Python
            # can't force-kill a thread) but no longer blocks the response.
            pool = concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(stale)))
            futures = {pool.submit(_fetch_one, yf, t): t for t in stale}
            for fut, t in futures.items():
                try:
                    _CACHE[t] = (now, fut.result(timeout=_FETCH_TIMEOUT))
                except concurrent.futures.TimeoutError:
                    _CACHE[t] = (now, _error(t, "Quote request timed out."))
            pool.shutdown(wait=False)

    return [_CACHE[t][1] for t in tickers]


def _fetch_one(yf, ticker):
    try:
        fi = yf.Ticker(ticker).fast_info
        price = fi.last_price
        prev = fi.previous_close
    except Exception as exc:                                      # noqa: BLE001
        return _error(ticker, str(exc) or "Couldn't reach the quote server.")

    if price is None or prev is None or not math.isfinite(price) or not math.isfinite(prev):
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
    Not cached: only fetched when a chart is actually opened. `ticker` isn't
    restricted to the saved watchlist here — the bridge validates user-supplied
    tickers before calling this; the benchmark overlay passes BENCHMARK_TICKER
    (an index symbol, "^GSPC") straight through, since that never comes from
    the client."""
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
    # A gap day (holiday, thin intraday trading, etc.) can leave Close as NaN.
    # json.dumps happily emits a bare `NaN` token for that, which isn't valid
    # JSON and breaks JSON.parse() on the frontend — so drop those points
    # rather than ship a value that can't even be parsed.
    points = [
        {"date": idx.strftime(date_fmt), "close": round(float(row["Close"]), 2)}
        for idx, row in hist.iterrows()
        if math.isfinite(row["Close"])
    ]
    if not points:
        return None, "No price history found for this ticker."
    return points, None


BENCHMARK_TICKER = "^GSPC"    # S&P 500 — the only benchmark the chart overlay supports

_SNAPSHOT_CACHE = {}          # ticker -> (fetched_at, snapshot_dict)
_SNAPSHOT_CACHE_TTL = 3600    # seconds — market cap/P-E/sector move far slower than price


def _safe_num(v):
    # NaN is a valid float but not valid JSON — json.dumps emits a bare `NaN`
    # token for it, which breaks JSON.parse() on the frontend (the same class
    # of bug fetch_history() guards against above).
    return v if isinstance(v, (int, float)) and math.isfinite(v) else None


def _empty_snapshot(ticker, error=None):
    return {"ticker": ticker, "marketCap": None, "trailingPE": None, "forwardPE": None, "dividendYield": None,
            "fiftyTwoWeekLow": None, "fiftyTwoWeekHigh": None, "sector": None, "error": error}


def fetch_snapshot(ticker):
    """Returns one {"ticker", "marketCap", "trailingPE", "forwardPE", "dividendYield",
    "fiftyTwoWeekLow", "fiftyTwoWeekHigh", "sector", "error"} dict. Every field
    degrades independently to None — one missing field never blanks the rest.
    "error" is only set when nothing at all came back. Cached for an hour:
    these move far slower than price, so there's no reason to refetch on
    every chart open."""
    now = time.time()
    cached = _SNAPSHOT_CACHE.get(ticker)
    if cached and now - cached[0] <= _SNAPSHOT_CACHE_TTL:
        return cached[1]

    try:
        import yfinance as yf
    except ImportError:
        snap = _empty_snapshot(ticker, "yfinance isn't installed on the server")
        _SNAPSHOT_CACHE[ticker] = (now, snap)
        return snap

    t = yf.Ticker(ticker)
    snap = _empty_snapshot(ticker)

    # fast_info can raise per-attribute for fields a given ticker type doesn't
    # have (e.g. an ETF) — check each independently so one miss doesn't blank
    # the others.
    try:
        snap["marketCap"] = _safe_num(t.fast_info.market_cap)
    except Exception:                                              # noqa: BLE001
        pass
    try:
        snap["fiftyTwoWeekLow"] = _safe_num(t.fast_info.year_low)
    except Exception:                                              # noqa: BLE001
        pass
    try:
        snap["fiftyTwoWeekHigh"] = _safe_num(t.fast_info.year_high)
    except Exception:                                              # noqa: BLE001
        pass

    # .info is one much heavier network call (a full quoteSummary scrape) —
    # if it fails, trailingPE/dividendYield/sector fail together, which is
    # correct since they share this one fetch.
    try:
        info = t.info or {}
    except Exception:                                              # noqa: BLE001
        info = {}
    snap["trailingPE"] = _safe_num(info.get("trailingPE"))
    snap["forwardPE"] = _safe_num(info.get("forwardPE"))
    snap["sector"] = info.get("sector") or None
    dy = _safe_num(info.get("dividendYield"))
    if dy is not None:
        # Yahoo's raw field has flip-flopped between a fraction (0.006) and a
        # whole percent (0.6) across API changes — normalize by magnitude so
        # this always comes out as a percent number (e.g. 2.53 meaning 2.53%).
        snap["dividendYield"] = dy * 100 if dy <= 1 else dy

    if all(snap[k] is None for k in
           ("marketCap", "trailingPE", "forwardPE", "dividendYield", "fiftyTwoWeekLow", "fiftyTwoWeekHigh", "sector")):
        snap["error"] = "No snapshot data found for this ticker."

    _SNAPSHOT_CACHE[ticker] = (now, snap)
    return snap


# ── Options: expirations, chain, and per-contract quotes/history ───────────
# fetch_history() above already works unchanged for an option contract
# symbol — Yahoo's chart endpoint is symbol-agnostic, it just happens to
# also carry real historical bars for currently-active option contracts.
# That's the one piece of this section that needed no new code at all.

_OPTION_EXP_CACHE = {}        # ticker -> (fetched_at, (expirations, error))
_OPTION_EXP_CACHE_TTL = 300   # seconds — expiration dates don't change intraday

_OPTION_CHAIN_CACHE = {}      # (ticker, expiration) -> (fetched_at, (chain, error))
_OPTION_CHAIN_CACHE_TTL = 600   # seconds — longer than _CACHE_TTL: options
                                 # endpoints are more prone to Yahoo rate-limiting
                                 # than the plain quote endpoint, and an expiration's
                                 # chain doesn't need second-by-second freshness.


def fetch_option_expirations(ticker):
    """Returns (expirations, error) — expirations is a list of "YYYY-MM-DD"
    strings, nearest first, or None on error."""
    now = time.time()
    cached = _OPTION_EXP_CACHE.get(ticker)
    if cached and now - cached[0] <= _OPTION_EXP_CACHE_TTL:
        return cached[1]

    try:
        import yfinance as yf
    except ImportError:
        result = (None, "yfinance isn't installed on the server")
        _OPTION_EXP_CACHE[ticker] = (now, result)
        return result

    try:
        expirations = list(yf.Ticker(ticker).options)
    except Exception as exc:                                          # noqa: BLE001
        result = (None, str(exc) or "Couldn't reach the options server.")
        _OPTION_EXP_CACHE[ticker] = (now, result)
        return result

    result = (expirations, None) if expirations else (None, "No options found for this ticker.")
    _OPTION_EXP_CACHE[ticker] = (now, result)
    return result


# Columns pulled from yfinance's option_chain() calls/puts DataFrames.
_CHAIN_NUMERIC_FIELDS = ("strike", "lastPrice", "bid", "ask", "impliedVolatility", "openInterest", "volume")


def _chain_rows(df):
    rows = []
    for _, row in df.iterrows():
        rec = {f: _safe_num(row.get(f)) for f in _CHAIN_NUMERIC_FIELDS}
        rec["contractSymbol"] = str(row.get("contractSymbol") or "")
        rec["inTheMoney"] = bool(row.get("inTheMoney"))
        rows.append(rec)
    return rows


def fetch_option_chain(ticker, expiration):
    """Returns ({"calls": [...], "puts": [...]}, error) for one expiration.
    Each row has contractSymbol/strike/lastPrice/bid/ask/impliedVolatility/
    openInterest/volume/inTheMoney — NaN-valued numeric fields come back as
    None (same reasoning as fetch_history's NaN guard above: a bare NaN
    breaks JSON.parse() on the frontend)."""
    key = (ticker, expiration)
    now = time.time()
    cached = _OPTION_CHAIN_CACHE.get(key)
    if cached and now - cached[0] <= _OPTION_CHAIN_CACHE_TTL:
        return cached[1]

    try:
        import yfinance as yf
    except ImportError:
        result = (None, "yfinance isn't installed on the server")
        _OPTION_CHAIN_CACHE[key] = (now, result)
        return result

    try:
        chain = yf.Ticker(ticker).option_chain(expiration)
    except Exception as exc:                                          # noqa: BLE001
        result = (None, str(exc) or "Couldn't reach the options server.")
        _OPTION_CHAIN_CACHE[key] = (now, result)
        return result

    calls, puts = _chain_rows(chain.calls), _chain_rows(chain.puts)
    result = ({"calls": calls, "puts": puts}, None) if (calls or puts) \
        else (None, "No options found for this expiration.")
    _OPTION_CHAIN_CACHE[key] = (now, result)
    return result


_OPTION_QUOTE_CACHE = {}   # contractSymbol -> (fetched_at, quote_dict) — kept
                            # separate from stock _CACHE so the two ticker
                            # spaces (plain tickers vs. OCC contract symbols)
                            # never share a slot.


def fetch_option_quotes(contracts):
    """Same shape/contract as fetch_quotes(): one {"ticker","price","change",
    "changePct","error"} dict per input contract symbol, same order, briefly
    cached. Whether fast_info behaves for an OCC option symbol the way it
    does for a stock ticker is undocumented on Yahoo's side, so this tries
    fast_info first and falls back to deriving price/previous-close from the
    last two daily bars of fetch_history() — the one path already confirmed
    to return real data for an active contract — if fast_info comes back
    empty or raises."""
    if not contracts:
        return []

    now = time.time()
    stale = [c for c in contracts if c not in _OPTION_QUOTE_CACHE or now - _OPTION_QUOTE_CACHE[c][0] > _CACHE_TTL]
    if stale:
        try:
            import yfinance as yf
        except ImportError:
            for c in stale:
                _OPTION_QUOTE_CACHE[c] = (now, _error(c, "yfinance isn't installed on the server"))
        else:
            pool = concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(stale)))
            futures = {pool.submit(_fetch_one_option, yf, c): c for c in stale}
            for fut, c in futures.items():
                try:
                    _OPTION_QUOTE_CACHE[c] = (now, fut.result(timeout=_FETCH_TIMEOUT))
                except concurrent.futures.TimeoutError:
                    _OPTION_QUOTE_CACHE[c] = (now, _error(c, "Quote request timed out."))
            pool.shutdown(wait=False)

    return [_OPTION_QUOTE_CACHE[c][1] for c in contracts]


def _fetch_one_option(yf_mod, symbol):
    pair = None
    try:
        fi = yf_mod.Ticker(symbol).fast_info
        price, prev = fi.last_price, fi.previous_close
        if price is not None and prev is not None and math.isfinite(price) and math.isfinite(prev):
            pair = (price, prev)
    except Exception:                                                  # noqa: BLE001
        pass

    if pair is None:
        points, _err = fetch_history(symbol, "5d")
        if points:
            pair = (points[-1]["close"], points[-2]["close"] if len(points) > 1 else points[-1]["close"])

    if pair is None:
        return _error(symbol, "No quote found for this contract.")
    price, prev = pair
    change = price - prev
    change_pct = (change / prev * 100) if prev else 0.0
    return {"ticker": symbol, "price": round(price, 2), "change": round(change, 2),
            "changePct": round(change_pct, 2), "error": None}
