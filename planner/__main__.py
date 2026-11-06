"""Command line: py -m planner [--demo] [--port N] [--browser] [--no-window] [--keep-running]"""
from __future__ import annotations

import argparse
import secrets
import sys
import threading

from . import APP_NAME, DEFAULT_PORT, VERSION
from .app import PlannerApp
from .config import Config, config_path
from .server import PlannerServer
from .window import _existing_instance, _watchdog, open_window

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Outlook calendar + Todoist tasks in one week view.")
    ap.add_argument("--demo", action="store_true", help="use sample data instead of your accounts")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"local port (default {DEFAULT_PORT})")
    ap.add_argument("--browser", action="store_true", help="open in your normal browser instead of an app window")
    ap.add_argument("--no-window", action="store_true", help="don't open a window (just print the address)")
    ap.add_argument("--keep-running", action="store_true", help="keep running after the window is closed")
    args = ap.parse_args(argv)

    cfg = Config(config_path(), memory_only=args.demo)
    app = PlannerApp(cfg, demo=args.demo)
    token = secrets.token_urlsafe(24)

    server = None
    try:
        server = PlannerServer(("127.0.0.1", args.port), app, token, args.keep_running or args.no_window)
    except OSError:
        if not args.demo and _existing_instance(args.port):
            url = f"http://127.0.0.1:{args.port}/"
            print(f"Planner is already running at {url}. Opening it.")
            if not args.no_window:
                open_window(url, args.browser)
            return 0
        server = PlannerServer(("127.0.0.1", 0), app, token, args.keep_running or args.no_window)

    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"{APP_NAME} {VERSION}{' (demo data)' if args.demo else ''} is running at {url}")
    print(f"Settings file: {cfg.path}" if not args.demo else "Demo mode: nothing is saved.")
    print("Close the Planner window (or press Ctrl+C here) to stop it.")
    if not args.no_window:
        threading.Timer(0.4, open_window, args=(url, args.browser)).start()
    threading.Thread(target=_watchdog, args=(server,), daemon=True).start()
    try:
        server.serve_forever(poll_interval=0.5)
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
