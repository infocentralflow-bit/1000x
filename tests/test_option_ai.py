"""Tests for option_ai.py — no network: the Anthropic client is stubbed.
Run: python3 -m unittest tests/test_option_ai.py"""
import os
import sys
import types
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import option_ai  # noqa: E402

FACTS = {"contract": {"ticker": "XYZ", "strike": 170}, "overview": {"breakEven": 181.5}, "categories": {"timeDecay": "Low"}}


class _Block:
    def __init__(self, text):
        self.type, self.text = "text", text


class _Resp:
    def __init__(self, text, stop_reason="end_turn"):
        self.content, self.stop_reason = [_Block(text)], stop_reason


class _StubClient:
    def __init__(self, resp):
        self.calls = []
        outer = self

        class _Messages:
            def create(self, **kw):
                outer.calls.append(kw)
                return resp
        self.beta = types.SimpleNamespace(messages=_Messages())


class OptionAiTest(unittest.TestCase):
    def setUp(self):
        option_ai._cache.clear()
        option_ai._last_call = 0.0
        option_ai.ANTHROPIC_API_KEY = "test-key"
        sys.modules.setdefault("anthropic", types.SimpleNamespace(
            AuthenticationError=type("AuthenticationError", (Exception,), {}),
            RateLimitError=type("RateLimitError", (Exception,), {}),
            APIStatusError=type("APIStatusError", (Exception,), {}),
            APIConnectionError=type("APIConnectionError", (Exception,), {}),
        ))

    def tearDown(self):
        option_ai._client = None

    def test_not_configured(self):
        option_ai.ANTHROPIC_API_KEY = ""
        text, err, status = option_ai.explain(FACTS)
        self.assertIsNone(text)
        self.assertEqual(status, 400)
        self.assertFalse(option_ai.configured())

    def test_validation(self):
        self.assertEqual(option_ai.explain("nope")[2], 400)
        self.assertIn("missing", option_ai.explain({"contract": {}})[1])
        big = dict(FACTS, pad={"x": "y" * (option_ai.MAX_FACTS_BYTES + 1)})
        self.assertIn("too large", option_ai.explain(big)[1])

    def test_calls_claude_with_facts_and_caches(self):
        stub = _StubClient(_Resp("Plain explanation."))
        option_ai._client = stub
        text, err, status = option_ai.explain(FACTS)
        self.assertEqual((text, err, status), ("Plain explanation.", None, 200))
        kw = stub.calls[0]
        self.assertEqual(kw["model"], option_ai.MODEL)
        self.assertIn('"breakEven":181.5', kw["messages"][0]["content"])
        self.assertIn("Do not calculate", kw["system"])
        # Same facts again: served from cache, no second call, no rate limit.
        self.assertEqual(option_ai.explain(dict(FACTS))[0], "Plain explanation.")
        self.assertEqual(len(stub.calls), 1)

    def test_rate_limit_between_new_calls(self):
        option_ai._client = _StubClient(_Resp("ok"))
        option_ai.explain(FACTS)
        other = dict(FACTS, overview={"breakEven": 190})
        self.assertEqual(option_ai.explain(other)[2], 429)

    def test_refusal(self):
        option_ai._client = _StubClient(_Resp("", stop_reason="refusal"))
        text, err, status = option_ai.explain(FACTS)
        self.assertIsNone(text)
        self.assertEqual(status, 502)


if __name__ == "__main__":
    unittest.main()
