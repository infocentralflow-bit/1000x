"""
Plain-English explanation for the Option Insight view.

The dashboard computes every number in the browser (OptionMath in
dashboard.html) and posts one "facts" object here. Claude is asked only to
explain those numbers in words — never to calculate, predict, or advise. The
page always shows its own template summary, so this is an optional extra:
with no ANTHROPIC_API_KEY set, the endpoint reports "not configured" and the
button stays hidden.

Cost guards: one call per button click, identical facts are served from a
10-minute cache, and calls are spaced at least 5 seconds apart.
"""

import hashlib
import json
import os
import threading
import time

MODEL = "claude-opus-5-5"
MAX_FACTS_BYTES = 32 * 1024
CACHE_TTL_S = 600
MIN_INTERVAL_S = 5.0
REQUIRED_SECTIONS = ("contract", "overview", "categories")

SYSTEM_PROMPT = """You explain a single stock option to a retail investor in plain English.

You receive a JSON object of numbers that were already calculated by the app \
(Black-Scholes-Merton estimates). Rules:
- Use ONLY the numbers in the JSON. Quote them exactly as given. Do not calculate, \
estimate, round differently, or introduce any new number.
- Do not give buy, sell, or hold advice, and do not predict where the stock will go.
- Explain how the outcome depends on WHERE the stock goes, WHEN it gets there, and \
what implied volatility does — use the timeSensitivity, timeDecay and ivRisk facts.
- Mention the break-even requirement and that the full premium (100%) can be lost.
- If thesis.directionMisaligned is true, point it out.
- 120 to 180 words, short paragraphs, no headings, no bullet lists, no markdown."""

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()

_lock = threading.Lock()
_cache = {}            # sha256 -> (expires_at, text)
_last_call = 0.0
_client = None


def configured():
    return bool(ANTHROPIC_API_KEY)


def validate_facts(facts):
    """Returns (canonical_json, error). Canonical JSON doubles as the cache key."""
    if not isinstance(facts, dict):
        return None, "facts must be an object."
    missing = [k for k in REQUIRED_SECTIONS if not isinstance(facts.get(k), dict)]
    if missing:
        return None, "facts is missing: " + ", ".join(missing)
    canonical = json.dumps(facts, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    if len(canonical.encode("utf-8")) > MAX_FACTS_BYTES:
        return None, "facts is too large."
    return canonical, None


def _get_client():
    global _client
    if _client is None:
        import anthropic  # lazy: the bridge runs fine without the SDK installed
        _client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    return _client


def _call_claude(canonical):
    """Returns (text, error)."""
    try:
        import anthropic
    except ImportError:
        return None, "The anthropic package isn't installed on the server."
    try:
        resp = _get_client().beta.messages.create(
            model=MODEL,
            max_tokens=4096,
            betas=["server-side-fallback-2026-07-01"],
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": "Explain this option using only these facts:\n" + canonical}],
            extra_body={"output_config": {"effort": "low"}, "fallbacks": "default"},
        )
    except anthropic.AuthenticationError:
        return None, "The server's ANTHROPIC_API_KEY was rejected."
    except anthropic.RateLimitError:
        return None, "The AI service is busy right now — try again in a minute."
    except anthropic.APIStatusError as e:
        return None, f"The AI service returned an error ({e.status_code})."
    except anthropic.APIConnectionError:
        return None, "Couldn't reach the AI service."
    if getattr(resp, "stop_reason", None) == "refusal":
        return None, "The AI declined to explain this one."
    text = "".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", "") == "text").strip()
    if not text:
        return None, "The AI returned an empty explanation."
    return text, None


def explain(facts):
    """Returns (text, error, http_status)."""
    global _last_call
    if not configured():
        return None, "AI explanations aren't configured on the server.", 400
    canonical, err = validate_facts(facts)
    if err:
        return None, err, 400
    key = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    now = time.time()
    with _lock:
        hit = _cache.get(key)
        if hit and hit[0] > now:
            return hit[1], None, 200
        if now - _last_call < MIN_INTERVAL_S:
            return None, "Please wait a few seconds between explanations.", 429
        _last_call = now
    text, err = _call_claude(canonical)
    if err:
        return None, err, 502
    with _lock:
        for k in [k for k, (exp, _) in _cache.items() if exp <= now]:
            del _cache[k]
        _cache[key] = (now + CACHE_TTL_S, text)
    return text, None, 200
