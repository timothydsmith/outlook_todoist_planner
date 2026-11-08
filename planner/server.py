"""The local web server: serves the page and the JSON API to it."""
from __future__ import annotations

import datetime as dt
import json
import os
import secrets
import sys
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from . import APP_NAME, VERSION
from .app import PlannerApp
from .errors import SourceError


UI_DIR = Path(__file__).with_name("ui")
# The only files the page may load besides index.html. Nothing else on disk is reachable.
UI_ASSETS = {"/app.js": ("app.js", "text/javascript; charset=utf-8"),
             "/styles.css": ("styles.css", "text/css; charset=utf-8")}

class PlannerServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = sys.platform != "win32"  # on Windows this would let two copies share a port

    def __init__(self, addr, app: PlannerApp, token: str, keep_running: bool):
        super().__init__(addr, PlannerHandler)
        self.app = app
        self.token = token
        self.keep_running = keep_running
        port = self.server_address[1]
        self.allowed_hosts = {f"127.0.0.1:{port}", f"localhost:{port}"}
        self.last_seen = time.time()
        self.seen_page = False
        self.bye_at = None


class PlannerHandler(BaseHTTPRequestHandler):
    server_version = f"{APP_NAME}/{VERSION}"
    server: PlannerServer

    def log_message(self, fmt, *args):
        if os.environ.get("PLANNER_DEBUG"):
            sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _send(self, code: int, obj=None, body: bytes | None = None, ctype="application/json; charset=utf-8"):
        if body is None:
            body = json.dumps(obj if obj is not None else {}).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()
        self.wfile.write(body)

    def _host_ok(self) -> bool:
        return self.headers.get("Host", "") in self.server.allowed_hosts

    def _token_ok(self, query: dict) -> bool:
        supplied = self.headers.get("X-Planner-Token") or (query.get("t") or [""])[0]
        return secrets.compare_digest(supplied, self.server.token)

    def _alive(self):
        self.server.last_seen = time.time()
        self.server.seen_page = True
        self.server.bye_at = None

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, {"error": "Forbidden"})
        u = urllib.parse.urlsplit(self.path)
        q = urllib.parse.parse_qs(u.query)
        if u.path == "/health":
            return self._send(200, {"app": "planner", "version": VERSION})
        if u.path in ("/", "/index.html"):
            self._alive()
            html = (UI_DIR / "index.html").read_text(encoding="utf-8").replace("__PLANNER_TOKEN__", self.server.token)
            return self._send(200, body=html.encode("utf-8"), ctype="text/html; charset=utf-8")
        if u.path in UI_ASSETS:
            name, ctype = UI_ASSETS[u.path]
            return self._send(200, body=(UI_DIR / name).read_bytes(), ctype=ctype)
        if not u.path.startswith("/api/") or not self._token_ok(q):
            return self._send(404, {"error": "Not found"})
        app = self.server.app
        if u.path == "/api/settings":
            s = app.cfg.public()
            s["demo"] = app.demo
            s["version"] = VERSION
            return self._send(200, s)
        if u.path == "/api/data":
            try:
                start = dt.date.fromisoformat(q.get("start", [""])[0])
                end = dt.date.fromisoformat(q.get("end", [""])[0])
            except ValueError:
                return self._send(400, {"error": "start and end must be YYYY-MM-DD"})
            if not (start < end <= start + dt.timedelta(days=62)):
                return self._send(400, {"error": "Date range must be 1–62 days"})
            force = q.get("force", ["0"])[0] == "1"
            self._alive()
            return self._send(200, app.data(start, end, force))
        return self._send(404, {"error": "Not found"})

    def do_POST(self):
        if not self._host_ok():
            return self._send(403, {"error": "Forbidden"})
        u = urllib.parse.urlsplit(self.path)
        q = urllib.parse.parse_qs(u.query)
        if not self._token_ok(q):
            return self._send(403, {"error": "Forbidden"})
        length = min(int(self.headers.get("Content-Length") or 0), 1_000_000)
        raw = self.rfile.read(length) if length else b""
        try:
            body = json.loads(raw) if raw.strip() else {}
        except ValueError:
            body = {}
        if not isinstance(body, dict):
            body = {}
        app = self.server.app
        try:
            if u.path == "/api/ping":
                self._alive()
                return self._send(200, {"ok": True})
            if u.path == "/api/bye":
                if not self.server.keep_running:
                    self.server.bye_at = time.time()
                return self._send(200, {"ok": True})
            if u.path == "/api/settings":
                if app.demo:
                    return self._send(400, {"error": "Settings can't be changed in demo mode. "
                                                     "Start the app without --demo to use your own data."})
                app.cfg.update(body)
                app.settings_changed()
                return self._send(200, app.cfg.public())
            if u.path == "/api/task/close":
                app.close_task(str(body["id"]))
                return self._send(200, {"ok": True})
            if u.path == "/api/task/reopen":
                app.reopen_task(str(body["id"]))
                return self._send(200, {"ok": True})
            if u.path == "/api/task/schedule":
                minutes = body.get("duration")
                app.schedule_task(str(body["id"]), str(body["date"]), body.get("time") or None,
                                  int(minutes) if minutes else None)
                return self._send(200, {"ok": True})
            if u.path == "/api/task/create":
                minutes = body.get("duration")
                app.create_task(str(body.get("content") or ""), body.get("date") or None,
                                body.get("time") or None, int(minutes) if minutes else None)
                return self._send(200, {"ok": True})
        except SourceError as exc:
            return self._send(502, {"error": str(exc)})
        except (KeyError, ValueError, TypeError) as exc:
            return self._send(400, {"error": f"Bad request: {exc}"})
        return self._send(404, {"error": "Not found"})
