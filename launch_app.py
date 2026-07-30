"""
Desktop-app launcher for the Projection Model dashboard.

Double-clicking the desktop shortcut runs this with pythonw.exe (no console
window). It:

  1. Starts excel_bridge.py in the background if it isn't already running
     (--no-open so it doesn't also pop a normal browser tab, --no-auth
     because this path never leaves 127.0.0.1 — nothing to protect against).
  2. Waits for it to come up.
  3. Opens it in an Edge/Chrome "app window" — no tabs, no address bar, its
     own taskbar icon — so it feels like a real app rather than a browser tab.
     Falls back to a normal browser tab if neither browser is found.

This is the quick desktop-app path. Nothing here is exposed off this
machine; use excel_bridge.py --lan or start_tunnel.bat for that.
"""

import os
import socket
import subprocess
import sys
import time
import webbrowser

HERE = os.path.dirname(os.path.abspath(__file__))
BRIDGE = os.path.join(HERE, "excel_bridge.py")
PORT = 8765
URL = f"http://127.0.0.1:{PORT}/"

BROWSER_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]


def port_open(host="127.0.0.1", port=PORT, timeout=0.4):
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def start_bridge():
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(
        [sys.executable, BRIDGE, "--no-open", "--no-auth"],
        cwd=HERE,
        creationflags=creationflags,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def wait_for_bridge(tries=40, interval=0.25):
    for _ in range(tries):
        if port_open():
            return True
        time.sleep(interval)
    return False


def open_app_window():
    browser = next((p for p in BROWSER_CANDIDATES if os.path.exists(p)), None)
    if browser:
        subprocess.Popen([browser, f"--app={URL}", "--window-size=1440,980"])
    else:
        webbrowser.open(URL)


def main():
    if not port_open():
        start_bridge()
        wait_for_bridge()
    open_app_window()


if __name__ == "__main__":
    main()
