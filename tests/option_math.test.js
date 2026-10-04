// Unit tests for the OptionMath block embedded in dashboard.html.
// Run: node tests/option_math.test.js
// The block is extracted between its start/end marker comments and evaluated
// in isolation, so these tests exercise exactly the code the page runs.
const fs = require("fs");
const path = require("path");
const assert = require("assert");

const html = fs.readFileSync(process.env.OPTION_MATH_SRC || path.join(__dirname, "..", "dashboard.html"), "utf8");
const start = html.indexOf("/* ── OptionMath:start");
const end = html.indexOf("/* ── OptionMath:end ── */");
if (start < 0 || end < 0) throw new Error("OptionMath markers not found in dashboard.html");
const OptionMath = new Function(html.slice(start, end) + "\nreturn OptionMath;")();
const M = OptionMath;

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; }
  catch (e) { console.error("FAIL:", name, "\n ", e.message); process.exitCode = 1; }
}
const near = (a, b, tol, msg) => assert.ok(Math.abs(a - b) <= tol, `${msg || ""} expected ${b}, got ${a} (tol ${tol})`);

test("normCdf reference values", () => {
  near(M.normCdf(0), 0.5, 1e-15);
  near(M.normCdf(1.96), 0.9750021048517795, 1e-12);
  near(M.normCdf(-1), 0.15865525393145707, 1e-12);
  near(M.normCdf(-8), 6.22096057427178e-16, 1e-20);
});

test("BSM textbook prices (S=K=100, T=1, r=5%, σ=20%)", () => {
  near(M.bsm(100, 100, 1, 0.05, 0, 0.2, "call"), 10.4506, 1e-4);
  near(M.bsm(100, 100, 1, 0.05, 0, 0.2, "put"), 5.5735, 1e-4);
});

test("BSM Hull example (S=42, K=40, r=10%, σ=20%, T=0.5)", () => {
  near(M.bsm(42, 40, 0.5, 0.1, 0, 0.2, "call"), 4.76, 5e-3);
  near(M.bsm(42, 40, 0.5, 0.1, 0, 0.2, "put"), 0.81, 5e-3);
});

test("put-call parity with dividend yield", () => {
  for (const [S, K, T, r, q, s] of [[150, 170, 0.7, 0.04, 0.01, 0.55], [90, 80, 2, 0.03, 0.025, 0.3]]) {
    const c = M.bsm(S, K, T, r, q, s, "call"), p = M.bsm(S, K, T, r, q, s, "put");
    near(c - p, S * Math.exp(-q * T) - K * Math.exp(-r * T), 1e-9);
  }
});

test("Greeks match central finite differences", () => {
  const S = 150, K = 170, T = 257 / 365, r = 0.04, q = 0.01, s = 0.55;
  for (const type of ["call", "put"]) {
    const g = M.greeks(S, K, T, r, q, s, type);
    const f = (S_, T_, r_, s_) => M.bsm(S_, K, T_, r_, q, s_, type);
    const h = 1e-3;
    near(g.delta, (f(S + h, T, r, s) - f(S - h, T, r, s)) / (2 * h), 1e-6, type + " delta");
    near(g.gamma, (f(S + h, T, r, s) - 2 * f(S, T, r, s) + f(S - h, T, r, s)) / (h * h), 1e-4, type + " gamma");
    near(g.vega, (f(S, T, r, s + 1e-4) - f(S, T, r, s - 1e-4)) / 2e-4 / 100, 1e-6, type + " vega");
    near(g.rho, (f(S, T, r + 1e-5, s) - f(S, T, r - 1e-5, s)) / 2e-5 / 100, 1e-6, type + " rho");
    const dt = 1 / 365;
    near(g.theta, (f(S, T - dt, r, s) - f(S, T + dt, r, s)) / 2, 1e-5, type + " theta/day");
  }
  assert.strictEqual(M.greeks(150, 170, 0, 0.04, 0, 0.5, "call"), null);
});

test("implied vol round-trips across a grid; out-of-bounds premiums return null", () => {
  for (const type of ["call", "put"])
    for (const K of [80, 100, 130])
      for (const T of [0.05, 0.5, 2])
        for (const s of [0.1, 0.45, 1.2]) {
          const p = M.bsm(100, K, T, 0.04, 0.01, s, type);
          if (p < 1e-6) continue;
          const iv = M.impliedVol(p, 100, K, T, 0.04, 0.01, type);
          // Always: the solved vol reprices the premium. Only where vega is
          // material is sigma itself identifiable (deep ITM short-dated
          // options price the same across a wide band of vols).
          near(M.bsm(100, K, T, 0.04, 0.01, iv, type), p, 1e-6, `${type} K${K} T${T} reprice`);
          if (M.greeks(100, K, T, 0.04, 0.01, s, type).vega > 1e-3)
            near(iv, s, 1e-5, `${type} K${K} T${T} sigma`);
        }
  assert.strictEqual(M.impliedVol(5, 150, 100, 0.5, 0.04, 0, "call"), null, "below intrinsic");
  assert.strictEqual(M.impliedVol(500, 150, 100, 0.5, 0.04, 0, "call"), null, "above max");
});

test("value at expiration is exactly intrinsic", () => {
  assert.strictEqual(M.value(190, 170, 0, 0.04, 0, 0.5, "call"), 20);
  assert.strictEqual(M.value(150, 170, 0, 0.04, 0, 0.5, "call"), 0);
  assert.strictEqual(M.value(150, 170, 0, 0.04, 0, 0.5, "put"), 20);
});

test("value never drops below intrinsic (deep ITM put)", () => {
  const v = M.value(50, 150, 1, 0.08, 0, 0.1, "put");
  assert.ok(v >= 100 - 1e-12, `got ${v}`);
});

const userExample = { type: "call", S: 150, K: 170, premium: 11.5, dte: 257, iv: 0.55, r: 0.04, q: 0 };

test("overview matches the user's worked example", () => {
  const o = M.overview(userExample);
  near(o.breakEven, 181.5, 1e-9);
  near(o.requiredMove, 0.21, 1e-9);
  assert.strictEqual(o.moneyness, "OTM");
  near(o.moneynessPct, 20 / 150, 1e-12);
  assert.strictEqual(o.intrinsic, 0);
  near(o.extrinsic, 11.5, 1e-12);
  near(o.cost, 1150, 1e-9);
  near(o.maxLoss, 1150, 1e-9);
});

test("put overview is direction-aware", () => {
  const o = M.overview({ type: "put", S: 150, K: 140, premium: 6, dte: 90, iv: 0.4, r: 0.04, q: 0 });
  near(o.breakEven, 134, 1e-12);
  near(o.requiredMove, -16 / 150, 1e-12);
  assert.strictEqual(o.moneyness, "OTM");
});

test("P&L and return at expiration follow the payoff formula", () => {
  const v = M.valueAt(userExample, 190, userExample.dte, 0.55);
  near(v, 20, 1e-12);
  near(M.pnl(userExample, v), 850, 1e-9);
  near(M.ret(userExample, v), 8.5 / 11.5, 1e-12);
});

test("reaching the same price earlier is worth more (time sensitivity)", () => {
  const early = M.valueAt(userExample, 190, 30, 0.55);
  const late = M.valueAt(userExample, 190, 180, 0.55);
  const atExp = M.valueAt(userExample, 190, userExample.dte, 0.55);
  assert.ok(early > late && late > atExp, `${early} > ${late} > ${atExp}`);
});

test("lower future IV lowers the projected value (IV crush)", () => {
  assert.ok(M.valueAt(userExample, 190, 30, 0.40) < M.valueAt(userExample, 190, 30, 0.55));
});

test("decay curve is non-increasing for an OTM call and ends at intrinsic", () => {
  const d = M.decay(userExample);
  for (let i = 1; i < d.length; i++) assert.ok(d[i].value <= d[i - 1].value + 1e-9, `day ${d[i].day}`);
  assert.strictEqual(d[d.length - 1].day, 257);
  assert.strictEqual(d[d.length - 1].value, 0);
  near(d[0].remaining + d[0].lost, 1, 1e-12);
});

test("overTime grid matches valueAt at the same points", () => {
  const rows = M.overTime(userExample, [130, 190], 0.5, [0, 60, 257]);
  near(rows[1].points[1].value, M.valueAt(userExample, 190, 60, 0.5), 1e-12);
  assert.strictEqual(rows[0].points[2].value, 0);
});

test("daySteps covers today through expiration with no duplicates", () => {
  const s = M.daySteps(257);
  assert.strictEqual(s[0], 0);
  assert.strictEqual(s[s.length - 1], 257);
  assert.strictEqual(new Set(s).size, s.length);
  assert.deepStrictEqual(M.daySteps(0), [0]);
  assert.deepStrictEqual(M.daySteps(3), [0, 1, 2, 3]);
});

test("calendar-day helpers", () => {
  assert.strictEqual(M.daysBetween("2026-10-04", "2027-06-18"), 257);
  assert.strictEqual(M.isoDate(M.addDays("2026-10-04", 30)), "2026-11-03");
});

test("default chart prices span bear..bull with round steps", () => {
  const p = M.defaultPrices(150, 135, 230);
  assert.strictEqual(p.length, 6);
  assert.ok(p[0] <= 127.5 && p[p.length - 1] >= 210, JSON.stringify(p));
  const step = p[1] - p[0];
  for (let i = 2; i < p.length; i++) near(p[i] - p[i - 1], step, 1e-9);
});

console.log(`${passed} OptionMath tests passed${process.exitCode ? " (with failures above)" : ""}`);
