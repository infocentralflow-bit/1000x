"""
Excel bridge for the Projection Model dashboard.

Serves dashboard.html on http://127.0.0.1:8765 and exposes three endpoints the
dashboard calls:

    GET  /api/health   -> {"ok": true}                     (used to detect the bridge)
    GET  /api/model    -> Company Inputs read from the .xlsx (cached values)
    POST /api/refresh  -> writes the ticker into the workbook, triggers Excel's
                          Power Query refresh (RefreshAll) via COM, waits for the
                          queries to settle, saves, and returns the fresh values.

Refresh has a documented fallback chain, so it degrades instead of dying:

    1. Excel COM (pywin32 + Excel installed)  -> real Power Query refresh
    2. Alpha Vantage called directly here     -> same two endpoints the M code uses
    3. Cached values already in the workbook  -> whatever the last refresh left

Only openpyxl is required. pywin32 is optional (step 1); without it the bridge
still works via step 2.

Run:  python excel_bridge.py            (or double-click run_dashboard.bat)
"""

import base64
import hmac
import json
import os
import re
import secrets
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.normpath(os.path.join(HERE, "..", "Financial_Projection_Model.xlsx"))
SHEET = "Projection Model"
# Cloud hosts (Render, etc.) assign a port via $PORT and expect the app to
# bind it — local runs never set this, so the local default is untouched.
PORT = int(os.environ.get("PORT", 8765))
IS_CLOUD = "PORT" in os.environ
CREDENTIALS_FILE = os.path.join(HERE, ".bridge_credentials.json")


# ── auth ─────────────────────────────────────────────────────────────────────
# Basic Auth in front of every route. This matters most once the bridge is
# reachable off this machine (LAN or a public tunnel like ngrok) — without it,
# anyone with the URL could read the model and trigger an Excel refresh.
# Credentials are generated once and cached locally so they survive restarts;
# delete .bridge_credentials.json to force a new password.
def load_or_create_credentials():
    # Cloud deploys set these explicitly (Render's dashboard env vars) so the
    # password is known up front instead of buried in a log line, and so it
    # survives a redeploy on a host with an ephemeral filesystem.
    env_user, env_pass = os.environ.get("BRIDGE_USERNAME"), os.environ.get("BRIDGE_PASSWORD")
    if env_user and env_pass:
        return env_user, env_pass
    if os.path.exists(CREDENTIALS_FILE):
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if data.get("username") and data.get("password"):
                return data["username"], data["password"]
        except Exception:                                        # noqa: BLE001
            pass
    creds = {"username": "owner", "password": secrets.token_urlsafe(9)}
    try:
        with open(CREDENTIALS_FILE, "w", encoding="utf-8") as fh:
            json.dump(creds, fh)
    except Exception:                                             # noqa: BLE001
        pass
    return creds["username"], creds["password"]


AUTH_USER, AUTH_PASS = load_or_create_credentials()
AUTH_TOKEN = base64.b64encode(f"{AUTH_USER}:{AUTH_PASS}".encode()).decode()
AUTH_ENABLED = "--no-auth" not in sys.argv[1:]

# Cell addresses of the Company Inputs block. These mirror the workbook's
# defined names (Ticker / ApiKey / Co_Revenue / Co_NetIncome / Co_EPS /
# Co_Shares); we resolve by name where possible and fall back to these.
FALLBACK_CELLS = {
    "ticker": "C3",
    "apiKey": "C4",
    "revenue": "C5",
    "netIncome": "C6",
    "eps": "C7",
    "shares": "C8",
}
NAME_MAP = {
    "ticker": "Ticker",
    "apiKey": "ApiKey",
    "revenue": "Co_Revenue",
    "netIncome": "Co_NetIncome",
    "eps": "Co_EPS",
    "shares": "Co_Shares",
}


# ── workbook reading ────────────────────────────────────────────────────────
def resolve_cells(wb):
    """Map logical field -> A1 address, preferring the workbook's defined names."""
    out = dict(FALLBACK_CELLS)
    for field, name in NAME_MAP.items():
        dn = wb.defined_names.get(name)
        if dn is None:
            continue
        try:
            # attr_text looks like: 'Projection Model'!$C$5
            ref = dn.attr_text.split("!")[-1].replace("$", "")
            if ref:
                out[field] = ref
        except Exception:
            pass
    return out


# ── saved projections: file-backed, so OneDrive syncs them across machines ──
# This folder already lives inside a OneDrive-synced directory, so writing
# each saved projection as its own JSON file here is enough to get real
# cross-device sync for free — no server, no accounts, just files OneDrive
# was already going to sync. The dashboard's localStorage save stays the
# fast local copy; this is the layer that lets a second PC see it too.
SAVES_DIR = os.path.join(HERE, "saves")
PROJECT_ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,80}$")


def _project_path(project_id):
    if not PROJECT_ID_RE.match(project_id or ""):
        return None
    return os.path.join(SAVES_DIR, project_id + ".json")


def list_saved_projects():
    """Every saved projection, keyed by id — same shape as the client's map."""
    out = {}
    if not os.path.isdir(SAVES_DIR):
        return out
    for name in os.listdir(SAVES_DIR):
        if not name.endswith(".json"):
            continue
        try:
            with open(os.path.join(SAVES_DIR, name), "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict) and data.get("id"):
                out[data["id"]] = data
        except (OSError, json.JSONDecodeError):
            continue  # skip a corrupt file rather than fail the whole list
    return out


def write_saved_project(data):
    """Create or overwrite one project's file. Returns (ok, error_message)."""
    if not isinstance(data, dict) or not data.get("id") or not data.get("name"):
        return False, "Project must include at least 'id' and 'name'."
    path = _project_path(data["id"])
    if path is None:
        return False, "Invalid project id."
    os.makedirs(SAVES_DIR, exist_ok=True)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
    except OSError as exc:
        return False, f"Could not write project file: {exc}"
    return True, None


def delete_saved_project(project_id):
    path = _project_path(project_id)
    if path is None or not os.path.isfile(path):
        return False
    try:
        os.remove(path)
        return True
    except OSError:
        return False


def read_company():
    """Read the Company Inputs block from the workbook's cached values."""
    if not os.path.exists(WORKBOOK):
        return {"error": f"Workbook not found: {WORKBOOK}"}
    try:
        wb_f = openpyxl.load_workbook(WORKBOOK)                  # for defined names
        cells = resolve_cells(wb_f)
        wb_f.close()
        wb = openpyxl.load_workbook(WORKBOOK, data_only=True)    # cached values
        ws = wb[SHEET]

        def num(addr):
            v = ws[addr].value
            try:
                return float(v)
            except (TypeError, ValueError):
                return 0.0

        data = {
            "ticker": str(ws[cells["ticker"]].value or "").strip().upper(),
            "revenue": num(cells["revenue"]),
            "netIncome": num(cells["netIncome"]),
            "eps": num(cells["eps"]),
            "shares": num(cells["shares"]),
            "workbook": WORKBOOK,
        }
        wb.close()
        return data
    except Exception as exc:                                     # noqa: BLE001
        return {"error": f"Could not read workbook: {exc}"}


def read_api_key():
    """Pull the Alpha Vantage key out of the workbook, if the user stored it there."""
    try:
        wb_f = openpyxl.load_workbook(WORKBOOK)
        cells = resolve_cells(wb_f)
        wb_f.close()
        wb = openpyxl.load_workbook(WORKBOOK, data_only=True)
        val = wb[SHEET][cells["apiKey"]].value
        wb.close()
        key = str(val or "").strip()
        return "" if key.upper().startswith("YOUR_API_KEY") else key
    except Exception:                                            # noqa: BLE001
        return ""


# ── path 1: real Excel Power Query refresh via COM ──────────────────────────
def refresh_via_excel(ticker, api_key):
    """Drive Excel itself: set the ticker, RefreshAll (Power Query), save, read back.

    Returns a dict on success, or raises with a readable reason.
    """
    try:
        import pythoncom  # noqa: F401  (pywin32)
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("pywin32 not installed") from exc

    import pythoncom
    pythoncom.CoInitialize()
    excel = None
    opened_here = False
    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except Exception:                                        # noqa: BLE001
            excel = win32.DispatchEx("Excel.Application")
            opened_here = True
        excel.DisplayAlerts = False

        target = os.path.normcase(os.path.abspath(WORKBOOK))
        wb = None
        for i in range(1, excel.Workbooks.Count + 1):
            cand = excel.Workbooks.Item(i)
            try:
                if os.path.normcase(os.path.abspath(cand.FullName)) == target:
                    wb = cand
                    break
            except Exception:                                    # noqa: BLE001
                continue
        if wb is None:
            wb = excel.Workbooks.Open(os.path.abspath(WORKBOOK))
            opened_here = True

        ws = wb.Worksheets(SHEET)

        # Write the ticker (and key, if supplied) so Power Query picks them up.
        def set_named(name, fallback_addr, value):
            if value in (None, ""):
                return
            try:
                wb.Names(name).RefersToRange.Value = value
            except Exception:                                    # noqa: BLE001
                ws.Range(fallback_addr).Value = value

        set_named("Ticker", FALLBACK_CELLS["ticker"], ticker)
        if api_key:
            set_named("ApiKey", FALLBACK_CELLS["apiKey"], api_key)

        # The actual Power Query refresh.
        wb.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:                                        # noqa: BLE001
            time.sleep(4)
        excel.Calculate()

        def num(name, addr):
            try:
                v = wb.Names(name).RefersToRange.Value
            except Exception:                                    # noqa: BLE001
                v = ws.Range(addr).Value
            try:
                return float(v)
            except (TypeError, ValueError):
                return 0.0

        data = {
            "ticker": str(ws.Range(FALLBACK_CELLS["ticker"]).Value or ticker).strip().upper(),
            "revenue": num("Co_Revenue", FALLBACK_CELLS["revenue"]),
            "netIncome": num("Co_NetIncome", FALLBACK_CELLS["netIncome"]),
            "eps": num("Co_EPS", FALLBACK_CELLS["eps"]),
            "shares": num("Co_Shares", FALLBACK_CELLS["shares"]),
            "source": "excel",
            "message": "Excel Power Query refreshed.",
        }

        try:
            wb.Save()
        except Exception:                                        # noqa: BLE001
            pass
        if opened_here:
            try:
                wb.Close(SaveChanges=True)
                excel.Quit()
            except Exception:                                    # noqa: BLE001
                pass
        return data
    finally:
        pythoncom.CoUninitialize()


# ── path 2: call Alpha Vantage directly (same endpoints as the M code) ──────
def fetch_alpha_vantage(ticker, api_key):
    """Server-side twin of the workbook's M code. No CORS problem from here."""
    if not api_key:
        raise RuntimeError("No Alpha Vantage API key. Enter one in the dashboard "
                           "or in the workbook's API Key cell.")

    def get(fn):
        url = ("https://www.alphavantage.co/query?function=" + fn
               + "&symbol=" + urllib.parse.quote(ticker)
               + "&apikey=" + urllib.parse.quote(api_key))
        req = urllib.request.Request(url, headers={"User-Agent": "ProjectionModel/1.0"})
        with urllib.request.urlopen(req, timeout=25) as resp:
            return json.loads(resp.read().decode("utf-8"))

    overview = get("OVERVIEW")
    income = get("INCOME_STATEMENT")

    # Alpha Vantage signals throttling / bad keys with these keys instead of data.
    for payload in (overview, income):
        for flag in ("Note", "Information", "Error Message"):
            if isinstance(payload, dict) and flag in payload:
                raise RuntimeError(str(payload[flag])[:300])

    reports = (income or {}).get("annualReports") or []
    if not reports:
        raise RuntimeError(f"No annual income statement returned for {ticker}.")
    latest = reports[0]

    def num(value, divisor=1.0):
        try:
            return float(value) / divisor
        except (TypeError, ValueError):
            return 0.0

    return {
        "ticker": ticker,
        "revenue": num(latest.get("totalRevenue"), 1e6),
        "netIncome": num(latest.get("netIncome"), 1e6),
        "eps": num(overview.get("EPS")),
        "shares": num(overview.get("SharesOutstanding"), 1e6),
        "source": "alphavantage",
        "message": "Fetched from Alpha Vantage (Excel not driven).",
    }


def write_into_workbook(data):
    """Best-effort: park fetched values on Data_API so Excel shows them too.

    Silently skips when the workbook is open in Excel (file locked) — the
    dashboard already has the numbers either way.
    """
    try:
        wb = openpyxl.load_workbook(WORKBOOK)
        if "Data_API" not in wb.sheetnames:
            wb.close()
            return False
        ws = wb["Data_API"]
        target = None
        for r in range(2, max(3, ws.max_row + 2)):
            cell = ws.cell(row=r, column=1).value
            if cell is None or str(cell).strip().upper() == data["ticker"]:
                target = r
                break
        if target is None:
            target = ws.max_row + 1
        ws.cell(row=target, column=1, value=data["ticker"])
        ws.cell(row=target, column=2, value=data["revenue"])
        ws.cell(row=target, column=3, value=data["netIncome"])
        ws.cell(row=target, column=4, value=data["eps"])
        ws.cell(row=target, column=5, value=data["shares"])
        wb.save(WORKBOOK)
        wb.close()
        return True
    except Exception:                                            # noqa: BLE001
        return False


# ── refresh orchestration ───────────────────────────────────────────────────
def refresh_chain(ticker, api_key):
    """Try each refresh path in order. Returns (payload, http_status)."""
    notes = []

    # 1. real Power Query refresh, driven through Excel itself
    try:
        data = refresh_via_excel(ticker, api_key)
        if any(data[k] for k in ("revenue", "netIncome", "eps", "shares")):
            return data, 200
        notes.append("Excel refreshed but Company Inputs are still empty "
                     "(the Power Query is not connected in the saved workbook).")
    except Exception as exc:                                     # noqa: BLE001
        notes.append(f"Excel COM path unavailable: {exc}")

    # 2. call the same endpoints the M code uses, from here
    try:
        data = fetch_alpha_vantage(ticker, api_key)
        wrote = write_into_workbook(data)
        data["message"] = ("Fetched from Alpha Vantage"
                           + (" and written to Data_API." if wrote
                              else " (workbook is open, so it was not written)."))
        if notes:
            data["note"] = " ".join(notes)
        return data, 200
    except Exception as exc:                                     # noqa: BLE001
        notes.append(f"Alpha Vantage path failed: {exc}")

    # 3. fall back to whatever the workbook already holds
    cached = read_company()
    if "error" not in cached and any(
        cached[k] for k in ("revenue", "netIncome", "eps", "shares")
    ):
        cached["source"] = "workbook-cache"
        cached["message"] = "Showing the workbook's last saved values."
        cached["note"] = " ".join(notes)
        return cached, 200

    return {"error": " ".join(notes) or "Refresh failed."}, 502


# ── HTTP layer ──────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    server_version = "ProjectionBridge/1.0"

    def log_message(self, fmt, *args):
        sys.stderr.write("  %s\n" % (fmt % args))

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")

    def _json(self, obj, status=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):                                        # noqa: N802
        # Preflight requests never carry credentials — don't gate them, or
        # browsers can't even ask permission to send the real request.
        self.send_response(204)
        self._cors()
        self.end_headers()

    def _authorized(self):
        if not AUTH_ENABLED:
            return True
        header = self.headers.get("Authorization", "")
        if not header.startswith("Basic "):
            return False
        # constant-time compare — this endpoint is reachable from the public
        # internet once tunneled, so a timing side-channel is a real risk.
        return hmac.compare_digest(header[6:], AUTH_TOKEN)

    def _require_auth(self):
        body = b"Authentication required."
        self.send_response(401)
        self.send_header("WWW-Authenticate", 'Basic realm="Projection Model"')
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):                                            # noqa: N802
        if not self._authorized():
            self._require_auth()
            return
        path = urllib.parse.urlparse(self.path).path

        if path in ("/", "/index.html", "/dashboard.html"):
            fp = os.path.join(HERE, "dashboard.html")
            try:
                with open(fp, "rb") as fh:
                    body = fh.read()
            except OSError:
                self._json({"error": "dashboard.html missing"}, 404)
                return
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)
            return

        # PWA assets — manifest, service worker, icons. Static, cacheable, and
        # what makes the browser offer a real "Install app" prompt.
        STATIC_FILES = {
            "/manifest.json": ("manifest.json", "application/manifest+json"),
            "/sw.js": ("sw.js", "application/javascript"),
        }
        if path in STATIC_FILES:
            rel, ctype = STATIC_FILES[path]
            self._serve_static(os.path.join(HERE, rel), ctype)
            return

        if path.startswith("/icons/"):
            # Path-traversal guard: resolve, then require it stay under icons/.
            icons_dir = os.path.realpath(os.path.join(HERE, "icons"))
            requested = os.path.realpath(os.path.join(HERE, path.lstrip("/")))
            if os.path.commonpath([icons_dir, requested]) == icons_dir and os.path.isfile(requested):
                self._serve_static(requested, "image/png")
            else:
                self._json({"error": "not found"}, 404)
            return

        if path == "/api/health":
            self._json({"ok": True, "workbook": WORKBOOK,
                        "workbookExists": os.path.exists(WORKBOOK)})
            return

        if path == "/api/model":
            data = read_company()
            self._json(data, 500 if "error" in data else 200)
            return

        if path == "/api/projects":
            self._json(list_saved_projects())
            return

        self._json({"error": "not found"}, 404)

    def _serve_static(self, filepath, content_type):
        try:
            with open(filepath, "rb") as fh:
                body = fh.read()
        except OSError:
            self._json({"error": "not found"}, 404)
            return
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "public, max-age=3600")
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        try:
            return json.loads(self.rfile.read(length) or b"{}")
        except json.JSONDecodeError:
            return {}

    def do_POST(self):                                           # noqa: N802
        if not self._authorized():
            self._require_auth()
            return
        path = urllib.parse.urlparse(self.path).path

        if path == "/api/refresh":
            payload = self._read_json_body()
            ticker = str(payload.get("ticker") or "").strip().upper()
            api_key = str(payload.get("apiKey") or "").strip() or read_api_key()
            if not ticker:
                ticker = (read_company().get("ticker") or "AAPL")
            data, status = refresh_chain(ticker, api_key)
            self._json(data, status)
            return

        if path == "/api/projects":
            payload = self._read_json_body()
            ok, err = write_saved_project(payload)
            if not ok:
                self._json({"error": err}, 400)
                return
            self._json({"ok": True})
            return

        self._json({"error": "not found"}, 404)

    def do_DELETE(self):                                         # noqa: N802
        if not self._authorized():
            self._require_auth()
            return
        path = urllib.parse.urlparse(self.path).path
        prefix = "/api/projects/"
        if not path.startswith(prefix):
            self._json({"error": "not found"}, 404)
            return
        project_id = path[len(prefix):]
        ok = delete_saved_project(project_id)
        self._json({"ok": ok})


def lan_ip():
    """Best-effort local network address, for opening the app on a phone."""
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))       # no packets sent; just picks the route
        return s.getsockname()[0]
    except Exception:                                            # noqa: BLE001
        return None
    finally:
        s.close()


def main():
    # --lan (or --host 0.0.0.0) exposes the app to your local network so a
    # phone on the same Wi-Fi can open it. Off by default: loopback only.
    # On a cloud host (Render etc.) $PORT is set and there's no LAN concept —
    # bind 0.0.0.0 automatically, since that's the only way it's reachable.
    args = sys.argv[1:]
    host = "0.0.0.0" if IS_CLOUD else "127.0.0.1"
    if "--lan" in args:
        host = "0.0.0.0"
    if "--host" in args:
        try:
            host = args[args.index("--host") + 1]
        except IndexError:
            pass

    print("=" * 68)
    print("  Projection Model — Excel bridge")
    print("=" * 68)
    if IS_CLOUD:
        print("  mode     : cloud (no Excel available — Alpha Vantage refresh only)")
    print(f"  workbook : {WORKBOOK}")
    print(f"  exists   : {os.path.exists(WORKBOOK)}")
    if AUTH_ENABLED:
        print(f"  login    : {AUTH_USER} / {AUTH_PASS}")
        if not (os.environ.get("BRIDGE_USERNAME") and os.environ.get("BRIDGE_PASSWORD")):
            print(f"             (saved in .bridge_credentials.json — delete it for a new password)")
    else:
        print("  login    : DISABLED (--no-auth) — do not expose this beyond localhost")
    try:
        import win32com.client  # noqa: F401
        print("  pywin32  : available (Excel Power Query refresh enabled)")
    except ImportError:
        print("  pywin32  : NOT installed — refresh will use Alpha Vantage directly.")
        if not IS_CLOUD:
            print("             enable the Excel path with:  pip install pywin32")

    url = f"http://127.0.0.1:{PORT}/"
    print(f"  serving  : {url}" if not IS_CLOUD else f"  serving  : 0.0.0.0:{PORT}")
    if IS_CLOUD:
        pass
    elif host == "0.0.0.0":
        ip = lan_ip()
        if ip:
            print(f"  on phone : http://{ip}:{PORT}/   (same Wi-Fi)")
            print("             Windows Firewall may ask to allow Python — say yes,")
            print("             and make sure it's allowed on Private networks.")
        else:
            print("  on phone : could not determine this machine's LAN address")
    else:
        print(f"  phone    : restart with  --lan  to allow access from your phone")
    print("  press Ctrl+C to stop")
    print("=" * 68)

    httpd = ThreadingHTTPServer((host, PORT), Handler)
    if "--no-open" not in args and not IS_CLOUD:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.")
        httpd.server_close()


if __name__ == "__main__":
    main()
