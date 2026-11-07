"""Runs helper.ps1 once and talks to it, one JSON line per request.

The helper is started on the first request, not when the app starts, so nothing touches Outlook
unless the "Outlook on this PC" calendar is in use. If the helper dies or stops answering, it's
restarted once; if that fails too, the request raises SourceError with a plain explanation.
"""
from __future__ import annotations

import collections
import json
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

from ..errors import SourceError

HELPER = Path(__file__).with_name("helper.ps1")

NEEDS_CLASSIC = ("This needs classic Outlook for Windows to be installed. You don't need to open it: Planner "
                 "runs it in the background. (The new Outlook can't be read this way.)")
TOO_SLOW = ("Classic Outlook took too long to answer. If it's showing a dialog box (for example asking you "
            "to sign in), deal with that and then refresh.")


def default_command() -> list:
    return ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(HELPER)]


class _HelperGone(Exception):
    """The helper exited or stopped answering. Worth one restart."""


class HelperError(SourceError):
    """The helper answered, but the request failed. `detail` is the helper's own message."""

    def __init__(self, message: str, detail: str):
        super().__init__(message)
        self.detail = detail


class OutlookBridge:
    def __init__(self, command: list | None = None, first_timeout: float = 60.0, timeout: float = 30.0):
        self.command = command            # None = the real PowerShell helper
        self.first_timeout = first_timeout
        self.timeout = timeout
        self._proc: subprocess.Popen | None = None
        self._lines: queue.Queue | None = None
        self._stderr: collections.deque = collections.deque(maxlen=20)
        self._lock = threading.Lock()
        self._seq = 0
        self._answered = False            # has the current helper answered anything yet?

    @property
    def running(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    def request(self, op: str, **args) -> dict:
        """Send one request and return the helper's reply (a dict with ok == True)."""
        with self._lock:
            try:
                return self._exchange(op, args)
            except _HelperGone:
                self._stop()
            try:
                return self._exchange(op, args)
            except _HelperGone as exc:
                self._stop()
                raise SourceError(str(exc)) from None

    def close(self) -> None:
        """Ask the helper to quit (it closes Outlook if it started it), then make sure it's gone."""
        with self._lock:
            if self.running:
                try:
                    self._exchange("quit", {}, timeout=15)
                except (_HelperGone, SourceError):
                    pass
                try:
                    self._proc.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    pass
            self._stop()

    # -- internals -------------------------------------------------------------
    def _start(self) -> None:
        cmd = self.command
        if cmd is None:
            if sys.platform != "win32":
                raise SourceError("Reading from the Outlook desktop app only works on Windows. "
                                  "Use the published calendar (ICS) link instead.")
            cmd = default_command()
        try:
            proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except FileNotFoundError:
            raise SourceError("Couldn't find PowerShell on this PC.") from None
        lines: queue.Queue = queue.Queue()
        self._stderr.clear()
        threading.Thread(target=self._pump, args=(proc.stdout, lines), daemon=True).start()
        threading.Thread(target=self._drain, args=(proc.stderr,), daemon=True).start()
        self._proc, self._lines, self._answered = proc, lines, False

    @staticmethod
    def _pump(stream, lines: queue.Queue) -> None:
        for raw in iter(stream.readline, b""):
            lines.put(raw)
        lines.put(None)  # end of output: the helper has exited

    def _drain(self, stream) -> None:
        for raw in iter(stream.readline, b""):
            text = raw.decode("utf-8", "replace").strip()
            if text:
                self._stderr.append(text)

    def _stop(self) -> None:
        proc, self._proc = self._proc, None
        if proc is None:
            return
        if proc.poll() is None:
            proc.kill()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            pass
        for stream in (proc.stdin, proc.stdout, proc.stderr):
            try:
                stream.close()
            except OSError:
                pass

    def _gone(self) -> str:
        detail = next((ln for ln in reversed(self._stderr) if ln), "")
        return ("The Outlook helper stopped unexpectedly. " + NEEDS_CLASSIC
                + (f" Details: {detail[:240]}" if detail else ""))

    def _exchange(self, op: str, args: dict, timeout: float | None = None) -> dict:
        if not self.running:
            self._start()
        self._seq += 1
        seq = self._seq
        try:
            self._proc.stdin.write((json.dumps({"id": seq, "op": op, **args}) + "\n").encode("utf-8"))
            self._proc.stdin.flush()
        except OSError:
            raise _HelperGone(self._gone()) from None
        wait = timeout if timeout is not None else (self.timeout if self._answered else self.first_timeout)
        deadline = time.monotonic() + wait
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise _HelperGone(TOO_SLOW)
            try:
                raw = self._lines.get(timeout=remaining)
            except queue.Empty:
                raise _HelperGone(TOO_SLOW) from None
            if raw is None:
                raise _HelperGone(self._gone())
            try:
                resp = json.loads(raw.decode("utf-8", "replace").strip().lstrip("﻿"))
            except ValueError:
                continue  # stray output from PowerShell
            if isinstance(resp, dict) and resp.get("id") == seq:
                break     # anything else answers an earlier request we gave up on
        self._answered = True
        if not resp.get("ok"):
            detail = str(resp.get("error") or "no details")[:240]
            raise HelperError(f"Couldn't read the Outlook desktop app. {NEEDS_CLASSIC} Details: {detail}", detail)
        return resp
