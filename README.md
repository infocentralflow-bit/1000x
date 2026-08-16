# Projection Model — Scenario Terminal

A dashboard front-end for `Financial_Projection_Model.xlsx`. Dark theme, gold accents,
KPI cards, an interactive projection chart, and the Base / Bear / Bull tables with every
assumption editable live.

## Run it

**As a desktop app (for now)** — double-click the **"Projection Model"** shortcut on your
Desktop. It silently starts the bridge in the background (no console window, no login
prompt — this path never leaves your PC) and opens the dashboard in its own app-style
window: no browser tabs, no address bar, its own taskbar icon. Closing the window doesn't
stop the bridge; if you want it fully stopped, end the `pythonw.exe` process, or just leave
it running — it's cheap and only listens on localhost.

If you ever need to recreate that shortcut, it points `pythonw.exe` at
`launch_app.py` with this folder as the working directory — see that file's docstring.

**As a real installed PWA (the more "refined" option)** — with the bridge running, open
`http://127.0.0.1:8765/` in Edge or Chrome and use the browser's own **install icon** in the
address bar ("Install this site as an app" / a ⊕-in-a-box icon). This registers a proper app
with Windows — Start Menu entry, its own icon (pulled straight from `manifest.json`, same
gold "FP" mark), and it works offline once loaded (a service worker caches the shell; your
live data — `/api/*` — is never cached, so it always hits the network for real). This
supersedes the desktop shortcut's `--app=` window trick with a browser-native equivalent;
either one is fine to use, the PWA install is just the more standard path going forward.

**As a server you start manually** — double-click **`run_dashboard.bat`**, or:

```bash
python excel_bridge.py
```

This one asks for a login (see [Viewing it on your phone](#viewing-it-on-your-phone) below
for why) and opens a normal browser tab at <http://127.0.0.1:8765/>. `Ctrl+C` stops it.

You can also just open `dashboard.html` directly — it works standalone as a calculator,
but pulling values from the workbook needs the bridge (a browser page can't reach into Excel).

## Where Company Inputs data comes from

There's no manual refresh button in the dashboard — on load, it silently asks the bridge
for the workbook's numbers once and fills in Company Inputs if it finds any. The bridge
tries, in order:

| # | Path | Needs | Result |
|---|------|-------|--------|
| 1 | Runs Excel's `RefreshAll` (Power Query), saves, reads `Co_*` back | Excel + `pywin32` | Live from workbook |
| 2 | Calls the same two Alpha Vantage endpoints the M code uses, then writes the row to `Data_API` | An API key in the workbook's **API Key** cell | Live from Alpha Vantage |
| 3 | Returns whatever values the workbook last saved | — | Live from workbook |

If your workbook reports `Workbook.Queries.Count == 0` (Power Query never saved into the
file), path 1 won't find anything — leave an Alpha Vantage key in the workbook's **API Key**
cell so path 2 can fill the Company Inputs cells instead. Path 1 starts working by itself the
moment you finish connecting the query in Excel. Any value can still be typed over by hand
in the dashboard regardless of where it came from.

## Formula parity

Every number is computed with the workbook's own formulas — verified line by line against
the sheet:

| Row | Formula |
|-----|---------|
| Revenue | `C25 = Co_Revenue`, then `Dt = C(t-1) * (1 + growth_t)` |
| Net Income | `C27 = Co_NetIncome`, then `Dt = C(t-1) * (1 + NI growth_t)` |
| Net Income Margin | `= NI / Revenue` (0 when revenue is 0) |
| Shares Outstanding | `= Co_Shares`, held flat across all five years |
| EPS | `C31 = Co_EPS` (reported); `Dt = NI_t / shares_t` |
| Share Price Low / High | `= EPS_t × P/E_t` |
| CAGR | `= (final / initial) ^ (1/4) − 1` |

Zero denominators resolve to 0, matching the sheet's `IFERROR(..., 0)` wrappers.

**One deliberate difference:** the workbook leaves 2028–2030 growth and P/E blank as
*ENTER VALUES* placeholders, which makes those years compute to 0. The dashboard carries
the prior year's assumption forward so the projection is never dead. Every carried cell is
editable — type over it and everything downstream recalculates instantly.

## Saving named projections

Top bar: a **New**, **Save**, and **Open** icon, plus the project name button.

- **New** (also `Ctrl`+`Alt`+`N`) — blanks everything: Company Inputs, all three scenarios'
  assumptions back to their 10%/3%/18% defaults, unlocks 2026, clears the active project. If
  there's anything unsaved worth losing, it asks first; sample data or an already-saved
  project never prompts unnecessarily.
- **Save** (also `Ctrl`/`Cmd`+`S`) — first time, prompts for a name; after that, saves over
  the same slot. A gold dot appears next to the name whenever you have unsaved changes, so
  you always know if what's on screen matches what's saved.
- Click the **project name** itself to rename it.
- **Open** shows every saved projection — ticker, revenue, and "updated 3 hr ago" at a
  glance — with **Load**, **Rename**, **Duplicate**, and **Delete** on each row. Duplicate
  is the easy way to branch a scenario ("Base Case" → "Base Case (copy)" → tweak the copy)
  without touching the original.
- **Export all** downloads every saved projection as one `.json` file; **Import** loads one
  back in — the manual way to move projections to a different browser or device (e.g. your
  phone's copy of `dashboard.html`, which has no bridge to sync through).
- **Automatic sync between machines running the bridge**: every Save also writes the
  projection as a file under `ProjectionApp/saves/` — inside this OneDrive-synced folder —
  so any other Windows machine signed into the same OneDrive account picks it up on its own,
  no setup required. Opening **Open** (or just launching the app) checks for anything new
  from another machine and merges it in; the newer copy always wins if the same projection
  was edited in two places. This only covers machines running `excel_bridge.py` — your
  phone's standalone `dashboard.html` still needs Export/Import.
- Closing the tab/window with unsaved changes prompts you first, same as any document editor.

## Entering 2026 by hand, then locking it

The four Company Inputs cards (Revenue, Net Income, EPS, Shares Outstanding) are **editable
number fields** — click the big number and type. Net Margin stays calculated.

Press **Lock 2026** (padlock, top right of that section) once the base year is right. Locking:

- turns the four fields read-only and gold, with a padlock on each card;
- **dims the entire 2026 column** in all three scenario tables and puts a padlock in its
  header, so your eye goes straight to 2027–2030;
- makes the 2026 P/E cells read-only too, while every forward year stays editable;
- **protects the base year from the startup load** — the automatic fetch won't overwrite
  locked values, it just tells you it didn't. Unlock first if you want it to win.

Click it again to unlock. Everything you type — values, lock state, and all forward-year
assumptions — is saved in the browser on that device, so it survives a reload or a phone
being closed. **Reset** clears it and returns to defaults.

## Viewing it on your phone

**Away from home / on mobile data — copy the single file.** `dashboard.html` is entirely
self-contained (no server, no internet, no fonts to download) and it's already inside your
OneDrive folder. Open the OneDrive app on your phone → `Documents/FWDBOTCODE/ProjectionApp/`
→ `dashboard.html` → open it in Safari or Chrome. Editing, locking, the chart and all three
scenarios work fully offline; only the automatic company-data load on startup needs the
bridge. Tip: if OneDrive insists on
downloading rather than opening it, email the file to yourself and open the attachment, or
use "Share → Copy to Files" first. Add it to your home screen and it behaves like an app.

**At home on the same Wi-Fi — use the bridge.** Start it with the `--lan` flag:

```bash
python excel_bridge.py --lan
```

It prints the exact phone URL. Windows Firewall will likely ask to allow Python the first
time — allow it on **Private** networks. Without `--lan` the bridge stays on loopback and is
unreachable from any other device, which is the safe default.

**On cellular data, away from home — tunnel it with ngrok.** This makes the bridge reachable
from anywhere, so it's now protected with a username/password (see below) — a bare
`--lan` bridge has no such protection and should never be tunneled directly.

1. In one window, start the bridge normally (no `--lan` needed — ngrok talks to
   `127.0.0.1` directly, so there's no reason to also expose it to your LAN):
   ```bash
   python excel_bridge.py
   ```
   It prints a login line like `owner / rPofnyQA3PMi` — that's your password, generated
   once and cached in `.bridge_credentials.json` (delete that file for a new one).

2. **One-time only:** sign up free at <https://dashboard.ngrok.com/signup>, then copy your
   authtoken from <https://dashboard.ngrok.com/get-started/your-authtoken> and run it
   yourself in a terminal (this is your own account credential, so I won't handle it for
   you):
   ```bash
   ngrok config add-authtoken YOUR_TOKEN_HERE
   ```

3. Double-click **`start_tunnel.bat`** (or run `ngrok http 8765`). It prints a public
   `https://something.ngrok-free.app` URL — **use the `https://` one**, not `http://`, so
   your password isn't sent in the clear.

4. Open that URL on your phone. Your browser will show its normal login prompt — enter
   the username/password from step 1. After that it behaves like the LAN version, startup
   data load included, from anywhere with signal.

The free ngrok URL changes every time you restart the tunnel, and the tunnel only exists
while `start_tunnel.bat` is running — closing it takes the app off the public internet
again immediately.

## Reading the dashboard

- **Gold values** — calculated, same as the workbook's gold cells.
- **Green values** — linked from Company Inputs (the 2026 column).
- **Dashed cells** — editable assumptions, so you always know the numbers aren't required
  to match the workbook.
- Chart: hover for a crosshair and tooltip, or focus it and use **← →**. Series are
  labelled directly at their endpoints as well as in the legend, so colour is never the
  only cue.

## Chart colours

The three series colours are a validated categorical palette, not a guess — checked on all
pairs against the dark card surface (`#16161a`):

| Series | Hex | |
|---|---|---|
| Base | `#3987e5` | contrast 4.96:1 |
| Bear | `#d43232` | contrast 3.70:1 |
| Bull | `#06b07a` | contrast 6.44:1 |

Worst-pair colour-blind separation ΔE **10.6** (target ≥ 8), worst normal-vision ΔE **22.7**
(floor ≥ 15). If you re-colour these, re-run the check — red/green pairs fail easily, and
most vivid red/green combinations in this lightness band do.

## Files

| File | |
|---|---|
| `dashboard.html` | The whole app — no build step, no dependencies, works offline |
| `excel_bridge.py` | Local server: reads the workbook, drives Excel, calls the API, hosts saved-projection files |
| `notion_service.py` | Notion API client used by the Notion Research feature |
| `quotes_service.py` | yfinance wrapper used by the watchlist strip |
| `run_dashboard.bat` | One-click launcher (asks for login) |
| `launch_app.py` / desktop shortcut | Silent launcher for the no-login local app window |
| `start_tunnel.bat` | ngrok tunnel for cellular access |
| `manifest.json`, `sw.js`, `icons/` | PWA install support (see "Viewing it on your phone" / desktop app section) |
| `saves/` | One `.json` file per saved projection — created on first Save, lives here so OneDrive syncs it |
| `app_icon.ico` | Desktop shortcut icon |

## Dependencies

- `openpyxl` — required
- `pywin32` — optional, enables the real Excel Power Query refresh (path 1)
- `yfinance` — optional, powers the watchlist strip; without it, watchlist tickers just show
  "no quote" instead of a price

```bash
pip install openpyxl pywin32 yfinance
```

## Running it always-on (no PC required)

ngrok/`start_tunnel.bat` only work while your PC is on — that's what's actually running the
bridge. To be reachable with your computer off, the app has to move to a cloud host, which
can't run Excel: **the workbook-refresh path is unavailable there — Alpha Vantage still
works normally.** The code already detects this environment (`$PORT` set) and adjusts —
binds `0.0.0.0`, skips opening a local browser, skips the Excel attempt without complaint.

This folder is already a git repo (`git log` to see the initial commit) with everything a
host needs: `requirements.txt`, `render.yaml`. What's left needs your own accounts, so I
can't do it from here:

1. **Create a GitHub repo** (github.com → New repository, empty, no README/license). Then,
   from this folder:
   ```bash
   git remote add origin https://github.com/<you>/<repo-name>.git
   git branch -M main
   git push -u origin main
   ```
2. **Sign up at [render.com](https://render.com)** (free, no card required for this tier) →
   **New → Blueprint** → connect the GitHub repo you just pushed. Render reads `render.yaml`
   automatically.
3. Render will ask for two environment variables it can't guess — set them in its dashboard:
   - `BRIDGE_USERNAME` — pick anything, e.g. `owner`
   - `BRIDGE_PASSWORD` — a real password; this gates the whole site once it's public
4. Deploy. Render gives you a permanent `https://<something>.onrender.com` URL — that one
   doesn't change on restart, unlike ngrok's.

Two things worth knowing about the free tier: it **spins down after 15 minutes idle** and
takes ~30–60 seconds to wake back up on the next visit (a paid plan removes this); and saved
projections written via the bridge (`saves/`) live on that ephemeral filesystem, wiped on
every redeploy — the dashboard's own `localStorage` is unaffected, so nothing you've saved
in-browser is at risk, but treat the cloud copy as not persistent server-side. Export/Import
(in the Open modal) is the reliable way to move projections to or from the cloud instance.

## Notion Research (optional)

Push a saved projection to a Notion page — revenue/net income/EPS/implied valuation per
scenario, plus a notes field that syncs back to Notion automatically as you type. Uses a
Notion **internal integration** (one token, set by you), not OAuth — there's no "each user
connects their own workspace" step because this app doesn't have separate user accounts.

1. Create an integration at [notion.so/my-integrations](https://www.notion.so/my-integrations)
   → **New integration** → give it a name → copy its **Internal Integration Secret**.
2. In Notion, open the database you want stock research pages written into → **•••** menu →
   **Connections** → add the integration you just created. (Without this step the app can see
   the database exists but can't read or write to it.)
3. Set `NOTION_TOKEN` to that secret — in Render's environment variables if it's deployed
   there, or as a local env var (`set NOTION_TOKEN=secret_...` / `export NOTION_TOKEN=...`)
   before running `excel_bridge.py` locally.
4. In the app: save a projection (Notion sync needs a name to attach to), then use the
   **Database** button in the Notion Research section to pick which database to sync into,
   then **Sync to Notion**.

The token never reaches the browser — every Notion API call happens in `excel_bridge.py`
(`notion_service.py`), same as the Alpha Vantage calls already do. Which Notion page belongs
to which saved projection is stored the same way saved projections themselves are (Postgres
if `NEON_DATABASE_URL` is set, a local file otherwise).

## Watchlist

A short-list strip under the top bar with a live price and day change per ticker — no API
key needed, no separate setup. Needs the bridge (same as everything else that talks to a
server), so it stays hidden if you open `dashboard.html` standalone.

- Tap **+** to add a ticker, **×** on a chip to remove one.
- Click a chip's ticker/price to open a price chart for it, with a range picker —
  **1D, 5D, 1M, 3M, 6M, 1Y, YTD** (1D/5D use intraday bars, everything else daily closes).
  Fetched on demand per range, not cached or pre-fetched — only requested when you actually
  open a chart or switch ranges. The price/change header updates to match the selected
  range (e.g. the 6M button shows the 6-month change, not the day's).
- Hover (or arrow-key through, once the chart has focus) any point on the chart for a
  crosshair and a tooltip with that point's exact date and close.
- Quotes come from `yfinance` (Yahoo Finance), fetched server-side in `quotes_service.py` and
  cached for 20 seconds so a fast reload doesn't refetch every ticker; the strip itself
  polls every 60 seconds while the tab is visible.
- The list is shared across your saved projections (one bridge, one watchlist) and stored the
  same way everything else here is — Postgres if `NEON_DATABASE_URL` is set, a local
  `watchlist.json` file otherwise. Capped at 12 tickers.

## Notes

- The bridge binds to `127.0.0.1` only — nothing is exposed to your network.
- Your API key is never written anywhere except the workbook cell you choose to put it in.
- Alpha Vantage's free tier allows 25 requests/day and 5/minute; each startup load uses 2.
