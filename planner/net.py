"""HTTP helper (honours system proxy settings; uses the OS certificate store)."""
from __future__ import annotations

import ssl
import urllib.error
import urllib.request

from . import APP_NAME, VERSION


def _ssl_context() -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    try:  # python.org builds on macOS ship without CA certs; certifi fills the gap if present
        import certifi  # type: ignore
        ctx.load_verify_locations(certifi.where())
    except Exception:
        pass
    return ctx


SSL_CTX = _ssl_context()


def http_request(method: str, url: str, headers: dict | None = None, body: bytes | None = None,
                 timeout: float = 25.0):
    req = urllib.request.Request(url, data=body, method=method, headers=dict(headers or {}))
    req.add_header("User-Agent", f"{APP_NAME}/{VERSION} (personal desktop app)")
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=SSL_CTX) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        try:
            return exc.code, exc.read() or b""
        finally:
            exc.close()
