"""Opening the app window, and stopping when it closes."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

def _app_browser_candidates() -> list:
    c = []
    if sys.platform == "win32":
        for env in ("PROGRAMFILES(X86)", "PROGRAMFILES", "LOCALAPPDATA"):
            base = os.environ.get(env)
            if base:
                c.append(Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
                c.append(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe")
    elif sys.platform == "darwin":
        for app in ("Microsoft Edge", "Google Chrome", "Chromium", "Brave Browser"):
            c.append(Path("/Applications") / f"{app}.app" / "Contents" / "MacOS" / app)
            c.append(Path.home() / "Applications" / f"{app}.app" / "Contents" / "MacOS" / app)
    else:
        for name in ("microsoft-edge", "google-chrome", "chromium", "chromium-browser"):
            p = shutil.which(name)
            if p:
                c.append(Path(p))
    return [p for p in c if p.exists()]


def open_window(url: str, plain_browser: bool = False) -> None:
    if not plain_browser:
        for exe in _app_browser_candidates():
            try:
                subprocess.Popen([str(exe), f"--app={url}", "--window-size=1360,880"],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
                return
            except OSError:
                continue
    webbrowser.open(url)


def _existing_instance(port: int) -> bool:
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # never send localhost via a proxy
    try:
        with opener.open(f"http://127.0.0.1:{port}/health", timeout=2) as resp:
            return resp.status == 200 and json.loads(resp.read()).get("app") == "planner"
    except Exception:
        return False


def _watchdog(server: PlannerServer):
    """Quit when the window has been closed (no heartbeat), unless --keep-running."""
    while True:
        time.sleep(2)
        if server.keep_running:
            continue
        now = time.time()
        if server.bye_at and now - server.bye_at > 8:
            break
        if server.seen_page and now - server.last_seen > 180:
            break
    print("Window closed, so Planner has stopped.")
    server.shutdown()
