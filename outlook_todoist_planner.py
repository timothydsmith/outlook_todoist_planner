#!/usr/bin/env python3
"""
Planner: your Outlook calendar and Todoist tasks in one week view.

It runs a small web app on your own computer (http://127.0.0.1:8765) and opens it
in its own window. Nothing gets installed and no admin consent is needed: it uses
only Python 3.9+ and its standard library.

    python outlook_todoist_planner.py           start the app
    python outlook_todoist_planner.py --demo    try it with sample data
    python outlook_todoist_planner.py --help    other options

Where the data comes from
  Todoist   your personal API token (Todoist > Settings > Integrations > Developer)
  Outlook   either the ICS link from "Publish a calendar" in Outlook on the web,
            or (Windows only) the classic Outlook desktop app already on your PC

Settings are saved in a small JSON file in your user profile (see config_path()).
Your Todoist token is stored there in plain text, readable only by your account.
"""
from __future__ import annotations

import argparse
import calendar as _calendar
import datetime as dt
import json
import os
import re
import secrets
import shutil
import socket
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:  # pragma: no cover
    ZoneInfo = None

APP_NAME = "Planner"
VERSION = "1.0"
DEFAULT_PORT = 8765
TODOIST_API = os.environ.get("PLANNER_TODOIST_API", "https://api.todoist.com/api/v1")
TODOIST_TASK_URL = "https://app.todoist.com/app/task/{id}"


# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------

DEFAULTS = {
    "todoist_token": "",
    "todoist_filter": "",
    "calendar_source": "ics",          # "ics" | "outlook_desktop" | "none"
    "ics_url": "",
    "week_start": 1,                   # 1 = Monday, 0 = Sunday
    "show_weekends": True,
    "day_start_hour": 7,
    "default_task_minutes": 30,
    "refresh_minutes": 5,
}


def config_path() -> Path:
    override = os.environ.get("PLANNER_CONFIG")
    if override:
        return Path(override)
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config"))
    return base / "OutlookTodoistPlanner" / "settings.json"


class Config:
    def __init__(self, path: Path, memory_only: bool = False):
        self.path = path
        self.memory_only = memory_only
        self.lock = threading.Lock()
        self.data = dict(DEFAULTS)
        if not memory_only and path.exists():
            try:
                loaded = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    for k in DEFAULTS:
                        if k in loaded:
                            self.data[k] = loaded[k]
            except (OSError, ValueError) as exc:
                print(f"Couldn't read settings from {path}: {exc}. Starting with defaults.")

    def get(self, key):
        with self.lock:
            return self.data.get(key, DEFAULTS.get(key))

    def snapshot(self) -> dict:
        with self.lock:
            return dict(self.data)

    def public(self) -> dict:
        """Settings for the UI. The token itself never goes back to the page."""
        d = self.snapshot()
        tok = d.pop("todoist_token", "") or ""
        d["has_token"] = bool(tok)
        d["token_hint"] = ("…" + tok[-4:]) if len(tok) >= 8 else ""
        d["platform"] = sys.platform
        return d

    def update(self, changes: dict) -> None:
        clean = {}
        for k, v in changes.items():
            if k not in DEFAULTS:
                continue
            if k == "todoist_token":
                v = str(v or "").strip()
                if not v:
                    continue  # blank means "keep the current token"
                if v == "__clear__":
                    v = ""
            elif k in ("todoist_filter", "ics_url"):
                v = str(v or "").strip()
            elif k == "calendar_source":
                v = v if v in ("ics", "outlook_desktop", "none") else "ics"
            elif k == "show_weekends":
                v = bool(v)
            elif k == "week_start":
                v = 0 if int(v) == 0 else 1
            elif k == "day_start_hour":
                v = min(12, max(0, int(v)))
            elif k == "default_task_minutes":
                v = min(480, max(5, int(v)))
            elif k == "refresh_minutes":
                v = min(60, max(1, int(v)))
            clean[k] = v
        with self.lock:
            self.data.update(clean)
            data = dict(self.data)
        if self.memory_only:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
        try:
            os.chmod(tmp, 0o600)
        except OSError:
            pass
        os.replace(tmp, self.path)


# --------------------------------------------------------------------------
# HTTP helper (honours system proxy settings; uses the OS certificate store)
# --------------------------------------------------------------------------

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


# --------------------------------------------------------------------------
# Todoist
# --------------------------------------------------------------------------

TODOIST_COLORS = {
    "berry_red": "#b8256f", "red": "#db4035", "orange": "#ff9933", "yellow": "#fad000",
    "olive_green": "#afb83b", "lime_green": "#7ecc49", "green": "#299438",
    "mint_green": "#6accbc", "teal": "#158fad", "sky_blue": "#14aaf5",
    "light_blue": "#96c3eb", "blue": "#4073ff", "grape": "#884dff", "violet": "#af38eb",
    "lavender": "#eb96eb", "magenta": "#e05194", "salmon": "#ff8d85",
    "charcoal": "#808080", "grey": "#b8b8b8", "taupe": "#ccac93",
}


class SourceError(Exception):
    """A problem with Todoist or the calendar, worded for the person using the app."""


class TodoistClient:
    def __init__(self, token: str, base: str = TODOIST_API):
        self.token = token
        self.base = base.rstrip("/")

    def _req(self, method: str, path: str, params: dict | None = None, payload: dict | None = None):
        url = self.base + path
        if params:
            url += "?" + urllib.parse.urlencode(params)
        headers = {"Authorization": f"Bearer {self.token}"}
        body = None
        if payload is not None:
            body = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        try:
            status, data = http_request(method, url, headers, body)
        except (urllib.error.URLError, socket.timeout, OSError) as exc:
            raise SourceError(f"Couldn't reach Todoist ({_reason(exc)}). Check your internet connection.")
        if status in (401, 403):
            raise SourceError("Todoist didn't accept the API token. Copy it again from "
                              "Todoist › Settings › Integrations › Developer and paste it in Settings.")
        if status == 404:
            raise SourceError("Todoist couldn't find that task. It may have been completed or deleted elsewhere.")
        if status == 429:
            raise SourceError("Todoist is rate-limiting requests. Wait a minute, then refresh.")
        if status >= 400:
            detail = data[:300].decode("utf-8", "replace").strip()
            raise SourceError(f"Todoist returned an error (HTTP {status}): {detail}")
        if not data or not data.strip():
            return None
        try:
            return json.loads(data)
        except ValueError:
            raise SourceError("Todoist sent a response this app couldn't read.")

    def _paged(self, path: str, params: dict | None = None) -> list:
        out: list = []
        cursor = None
        for _ in range(200):
            p = dict(params or {})
            p["limit"] = 200
            if cursor:
                p["cursor"] = cursor
            data = self._req("GET", path, p)
            if isinstance(data, list):  # tolerate an unpaged response
                out.extend(data)
                break
            data = data or {}
            out.extend(data.get("results") or [])
            cursor = data.get("next_cursor")
            if not cursor:
                break
        return out

    def tasks(self, filter_query: str = "") -> list:
        if filter_query:
            return self._paged("/tasks/filter", {"query": filter_query})
        return self._paged("/tasks")

    def projects(self) -> list:
        return self._paged("/projects")

    def close(self, task_id: str):
        self._req("POST", f"/tasks/{urllib.parse.quote(task_id)}/close")

    def reopen(self, task_id: str):
        self._req("POST", f"/tasks/{urllib.parse.quote(task_id)}/reopen")

    def update(self, task_id: str, payload: dict):
        return self._req("POST", f"/tasks/{urllib.parse.quote(task_id)}", payload=payload)

    def create(self, payload: dict):
        return self._req("POST", "/tasks", payload=payload)


_ISO_RE = re.compile(
    r"^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?\s*(Z|[+-]\d{2}:?\d{2})?$")


def parse_iso(raw: str):
    """'2026-11-06' -> date; '2026-11-06T09:00:00' -> naive; with Z/offset -> aware."""
    m = _ISO_RE.match((raw or "").strip())
    if not m:
        raise ValueError(f"not an ISO date: {raw!r}")
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if m.group(4) is None:
        return dt.date(y, mo, d)
    val = dt.datetime(y, mo, d, int(m.group(4)), int(m.group(5)), int(m.group(6) or 0))
    tz = m.group(7)
    if tz:
        if tz == "Z":
            val = val.replace(tzinfo=dt.timezone.utc)
        else:
            sign = -1 if tz[0] == "-" else 1
            digits = tz[1:].replace(":", "")
            off = dt.timedelta(hours=int(digits[:2]), minutes=int(digits[2:4]))
            val = val.replace(tzinfo=dt.timezone(sign * off))
    return val


def to_local_naive(value: dt.datetime) -> dt.datetime:
    """Aware -> this computer's local wall time (naive). Naive values are already local."""
    if value.tzinfo is None:
        return value
    return value.astimezone().replace(tzinfo=None)


def normalize_task(t: dict, projects: dict, unknown_zones: set | None = None) -> dict:
    """unknown_zones, if given, collects time-zone names that couldn't be resolved."""
    due = t.get("due") or None
    nd = None
    if due:
        raw = due.get("datetime") or due.get("date") or ""
        try:
            parsed = parse_iso(raw)
        except ValueError:
            parsed = None
        if isinstance(parsed, dt.datetime):
            if parsed.tzinfo is None and due.get("timezone"):
                tz = zone_by_name(due["timezone"])
                if tz is not None:
                    parsed = parsed.replace(tzinfo=tz)
                elif unknown_zones is not None:
                    unknown_zones.add(str(due["timezone"]))
            local = to_local_naive(parsed)
            nd = {"date": local.strftime("%Y-%m-%d"), "time": local.strftime("%H:%M")}
        elif isinstance(parsed, dt.date):
            nd = {"date": parsed.isoformat(), "time": None}
        if nd is not None:
            nd["recurring"] = bool(due.get("is_recurring"))
            nd["string"] = due.get("string") or ""
    minutes = None
    dur = t.get("duration")
    if isinstance(dur, dict) and dur.get("amount"):
        try:
            minutes = int(dur["amount"]) * (1440 if dur.get("unit") == "day" else 1)
        except (TypeError, ValueError):
            minutes = None
    proj = projects.get(str(t.get("project_id")), {})
    tid = str(t.get("id"))
    return {
        "id": tid,
        "content": t.get("content") or "",
        "description": (t.get("description") or "")[:1500],
        "priority": int(t.get("priority") or 1),
        "project": proj.get("name", ""),
        "projectColor": TODOIST_COLORS.get(proj.get("color"), "#9aa3ad"),
        "labels": list(t.get("labels") or []),
        "due": nd,
        "duration": minutes,
        "isSubtask": bool(t.get("parent_id")),
        "url": TODOIST_TASK_URL.format(id=urllib.parse.quote(tid)),
    }


def schedule_payload(date_str: str, time_str: str | None, minutes: int | None) -> dict:
    """Turn a local date/time chosen in the UI into a Todoist update payload."""
    day = dt.date.fromisoformat(date_str)
    if not time_str:
        return {"due_date": day.isoformat()}
    hh, mm = (int(x) for x in time_str.split(":")[:2])
    local = dt.datetime(day.year, day.month, day.day, hh, mm)
    utc = local.astimezone(dt.timezone.utc)  # naive -> treated as this computer's local time
    payload = {"due_datetime": utc.strftime("%Y-%m-%dT%H:%M:%SZ")}
    if minutes:
        payload["duration"] = int(minutes)
        payload["duration_unit"] = "minute"
    return payload


# --------------------------------------------------------------------------
# iCalendar (.ics) parsing with time zones and recurrence
# --------------------------------------------------------------------------

WEEKDAY_CODES = {"MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4, "SA": 5, "SU": 6}

# Fallback only: Outlook feeds normally carry their own VTIMEZONE definitions.
WINDOWS_TZ = {
    "AUS Eastern Standard Time": "Australia/Sydney", "E. Australia Standard Time": "Australia/Brisbane",
    "Cen. Australia Standard Time": "Australia/Adelaide", "AUS Central Standard Time": "Australia/Darwin",
    "W. Australia Standard Time": "Australia/Perth", "Tasmania Standard Time": "Australia/Hobart",
    "New Zealand Standard Time": "Pacific/Auckland", "UTC": "UTC", "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik", "W. Europe Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris", "Central Europe Standard Time": "Europe/Budapest",
    "Eastern Standard Time": "America/New_York", "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver", "Pacific Standard Time": "America/Los_Angeles",
    "Singapore Standard Time": "Asia/Singapore", "China Standard Time": "Asia/Shanghai",
    "Tokyo Standard Time": "Asia/Tokyo", "India Standard Time": "Asia/Kolkata",
    "Korea Standard Time": "Asia/Seoul", "SE Asia Standard Time": "Asia/Bangkok",
}


def unfold_lines(text: str) -> list:
    out: list = []
    for line in re.split(r"\r\n|\n|\r", text):
        if line[:1] in (" ", "\t") and out:
            out[-1] += line[1:]
        elif line:
            out.append(line)
    return out


def _split_unquoted(s: str, sep: str) -> list:
    parts, buf, in_q = [], [], False
    for ch in s:
        if ch == '"':
            in_q = not in_q
            buf.append(ch)
        elif ch == sep and not in_q:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf))
    return parts


def parse_content_line(line: str):
    in_q = False
    colon = -1
    for i, ch in enumerate(line):
        if ch == '"':
            in_q = not in_q
        elif ch == ":" and not in_q:
            colon = i
            break
    if colon <= 0:
        return None
    head, value = line[:colon], line[colon + 1:]
    parts = _split_unquoted(head, ";")
    name = parts[0].strip().upper()
    if "." in name:  # strip vCard-style group prefixes
        name = name.rsplit(".", 1)[1]
    params = {}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            params[k.strip().upper()] = v.strip().strip('"')
    return name, params, value


class Component:
    __slots__ = ("name", "props", "children")

    def __init__(self, name: str):
        self.name = name
        self.props: list = []
        self.children: list = []

    def first(self, key: str):
        for n, p, v in self.props:
            if n == key:
                return p, v
        return None

    def every(self, key: str) -> list:
        return [(p, v) for n, p, v in self.props if n == key]

    def walk(self, name: str):
        for c in self.children:
            if c.name == name:
                yield c
            yield from c.walk(name)


def parse_components(text: str) -> Component:
    root = Component("ROOT")
    stack = [root]
    for line in unfold_lines(text):
        cl = parse_content_line(line)
        if not cl:
            continue
        n, p, v = cl
        if n == "BEGIN":
            c = Component(v.strip().upper())
            stack[-1].children.append(c)
            stack.append(c)
        elif n == "END":
            if len(stack) > 1:
                stack.pop()
        else:
            stack[-1].props.append(cl)
    return root


def unescape_text(v: str) -> str:
    return re.sub(r"\\([\\;,nN])", lambda m: "\n" if m.group(1) in "nN" else m.group(1), v)


_DT_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$")


def parse_dt_value(value: str, params: dict, tzres=None):
    """ICS DATE / DATE-TIME -> date | aware datetime | naive (floating) datetime."""
    m = _DT_RE.match(value.strip())
    if not m:
        raise ValueError(f"bad ICS date {value!r}")
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if m.group(4) is None:
        return dt.date(y, mo, d)
    naive = dt.datetime(y, mo, d, int(m.group(4)), int(m.group(5)), int(m.group(6) or 0))
    if m.group(7):
        return naive.replace(tzinfo=dt.timezone.utc)
    tzid = params.get("TZID")
    if tzid and tzres is not None:
        tz = tzres(tzid)
        if tz is not None:
            return naive.replace(tzinfo=tz)
    return naive


def parse_offset(s: str) -> dt.timedelta:
    s = s.strip()
    sign = -1 if s.startswith("-") else 1
    digits = s.lstrip("+-")
    h, m = int(digits[0:2]), int(digits[2:4] or 0)
    sec = int(digits[4:6] or 0) if len(digits) >= 6 else 0
    return sign * dt.timedelta(hours=h, minutes=m, seconds=sec)


def parse_duration(s: str) -> dt.timedelta:
    m = re.match(r"^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$", s.strip())
    if not m:
        raise ValueError(f"bad duration {s!r}")
    sign = -1 if m.group(1) == "-" else 1
    w, d, h, mi, se = (int(x or 0) for x in m.groups()[1:])
    return sign * dt.timedelta(weeks=w, days=d, hours=h, minutes=mi, seconds=se)


# ---- recurrence rules ------------------------------------------------------

def parse_rrule(s: str) -> dict:
    r = {"FREQ": None, "INTERVAL": 1, "COUNT": None, "UNTIL": None, "BYDAY": [], "BYMONTHDAY": [],
         "BYMONTH": [], "BYSETPOS": [], "BYYEARDAY": [], "WKST": 0}
    for part in s.split(";"):
        if "=" not in part:
            continue
        k, v = part.split("=", 1)
        k, v = k.strip().upper(), v.strip()
        try:
            if k == "FREQ":
                r["FREQ"] = v.upper()
            elif k == "INTERVAL":
                r["INTERVAL"] = max(1, int(v))
            elif k == "COUNT":
                r["COUNT"] = max(0, int(v))
            elif k == "UNTIL":
                r["UNTIL"] = v
            elif k == "BYDAY":
                for x in v.split(","):
                    mm = re.match(r"^([+-]?\d{1,2})?(MO|TU|WE|TH|FR|SA|SU)$", x.strip().upper())
                    if mm:
                        r["BYDAY"].append((int(mm.group(1)) if mm.group(1) else None, WEEKDAY_CODES[mm.group(2)]))
            elif k in ("BYMONTHDAY", "BYMONTH", "BYSETPOS", "BYYEARDAY"):
                r[k] = [int(x) for x in v.split(",") if x.strip()]
            elif k == "WKST":
                r["WKST"] = WEEKDAY_CODES.get(v.upper(), 0)
        except ValueError:
            continue
    return r


def _add_months(d: dt.date, n: int) -> dt.date:
    idx = d.year * 12 + (d.month - 1) + n
    return dt.date(idx // 12, idx % 12 + 1, 1)


def _weekdays_in(first: dt.date, last: dt.date, wd: int) -> list:
    d = first + dt.timedelta(days=(wd - first.weekday()) % 7)
    out = []
    while d <= last:
        out.append(d)
        d += dt.timedelta(days=7)
    return out


def _pick(days: list, n: int | None) -> list:
    if n is None:
        return days
    if n > 0:
        return [days[n - 1]] if n <= len(days) else []
    return [days[n]] if -n <= len(days) else []


def _monthday_dates(y: int, m: int, mdays: list) -> list:
    last = _calendar.monthrange(y, m)[1]
    out = []
    for md in mdays:
        day = md if md > 0 else last + md + 1
        if 1 <= day <= last:
            out.append(dt.date(y, m, day))
    return out


def _period_candidates(rule: dict, pstart: dt.date, first: dt.date) -> list:
    freq = rule["FREQ"]
    bm, bmd, bd, byd = rule["BYMONTH"], rule["BYMONTHDAY"], rule["BYDAY"], rule["BYYEARDAY"]
    wdset = {w for _, w in bd}
    days: list = []
    if freq == "DAILY":
        d = pstart
        if bm and d.month not in bm:
            return []
        if bmd and d not in _monthday_dates(d.year, d.month, bmd):
            return []
        if bd and d.weekday() not in wdset:
            return []
        return [d]
    if freq == "WEEKLY":
        if bd:
            days = sorted({pstart + dt.timedelta(days=(w - pstart.weekday()) % 7) for w in wdset})
        else:
            days = [pstart + dt.timedelta(days=(first.weekday() - pstart.weekday()) % 7)]
        if bm:
            days = [d for d in days if d.month in bm]
        return days
    if freq == "MONTHLY":
        y, m = pstart.year, pstart.month
        if bm and m not in bm:
            return []
        if bmd:
            days = _monthday_dates(y, m, bmd)
            if bd:
                days = [d for d in days if d.weekday() in wdset]
        elif bd:
            last = dt.date(y, m, _calendar.monthrange(y, m)[1])
            s = set()
            for n, w in bd:
                s.update(_pick(_weekdays_in(dt.date(y, m, 1), last, w), n))
            days = sorted(s)
        else:
            days = _monthday_dates(y, m, [first.day]) if first.day <= _calendar.monthrange(y, m)[1] else []
        return sorted(set(days))
    if freq == "YEARLY":
        y = pstart.year
        ylen = 366 if _calendar.isleap(y) else 365
        if byd:
            for yd in byd:
                n = yd if yd > 0 else ylen + yd + 1
                if 1 <= n <= ylen:
                    days.append(dt.date(y, 1, 1) + dt.timedelta(days=n - 1))
            if bm:
                days = [d for d in days if d.month in bm]
            if bmd:
                days = [d for d in days if d in _monthday_dates(d.year, d.month, bmd)]
            if bd:
                days = [d for d in days if d.weekday() in wdset]
        elif bmd:
            for m in (bm or range(1, 13)):
                days.extend(_monthday_dates(y, m, bmd))
            if bd:
                days = [d for d in days if d.weekday() in wdset]
        elif bd:
            s = set()
            if bm:
                for m in bm:
                    last = dt.date(y, m, _calendar.monthrange(y, m)[1])
                    for n, w in bd:
                        s.update(_pick(_weekdays_in(dt.date(y, m, 1), last, w), n))
            else:
                for n, w in bd:
                    s.update(_pick(_weekdays_in(dt.date(y, 1, 1), dt.date(y, 12, 31), w), n))
            days = list(s)
        else:
            for m in (bm or [first.month]):
                if first.day <= _calendar.monthrange(y, m)[1]:
                    days.append(dt.date(y, m, first.day))
        return sorted(set(days))
    return []


def _apply_setpos(days: list, setpos: list) -> list:
    if not setpos:
        return days
    out = set()
    for p in setpos:
        if p > 0 and p <= len(days):
            out.add(days[p - 1])
        elif p < 0 and -p <= len(days):
            out.add(days[p])
    return sorted(out)


def expand_rrule(dtstart: dt.datetime, rule: dict, win_start: dt.datetime, win_end: dt.datetime,
                 until: dt.datetime | None = None, limit: int = 5000) -> list:
    """Occurrences (naive wall-clock datetimes) of a rule that fall in [win_start, win_end).

    DTSTART always counts as the first occurrence, as RFC 5545 requires.
    """
    freq = rule.get("FREQ")
    out: list = []
    if dtstart < win_end and dtstart >= win_start and (until is None or dtstart <= until):
        out.append(dtstart)
    if freq not in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
        return out
    interval = rule["INTERVAL"]
    count = rule["COUNT"]
    if count is not None and count <= 1:
        return out
    first = dtstart.date()
    tpart = dtstart.time()
    if freq == "DAILY":
        p0 = first
    elif freq == "WEEKLY":
        p0 = first - dt.timedelta(days=(first.weekday() - rule["WKST"]) % 7)
    elif freq == "MONTHLY":
        p0 = dt.date(first.year, first.month, 1)
    else:
        p0 = dt.date(first.year, 1, 1)

    k = 0
    if count is None:  # jump straight to the window; COUNT rules must be walked from the start
        target = win_start.date()
        if target > first:
            if freq == "DAILY":
                n = (target - p0).days // interval
            elif freq == "WEEKLY":
                n = (target - p0).days // (7 * interval)
            elif freq == "MONTHLY":
                n = ((target.year - p0.year) * 12 + target.month - p0.month) // interval
            else:
                n = (target.year - p0.year) // interval
            k = max(0, n - 1)

    produced = 1  # DTSTART
    end_date = win_end.date()
    for _ in range(200000):
        if freq == "DAILY":
            ps = p0 + dt.timedelta(days=k * interval)
        elif freq == "WEEKLY":
            ps = p0 + dt.timedelta(days=7 * k * interval)
        elif freq == "MONTHLY":
            if p0.year * 12 + k * interval > 9998 * 12:
                break
            ps = _add_months(p0, k * interval)
        else:
            if p0.year + k * interval > 9998:
                break
            ps = dt.date(p0.year + k * interval, 1, 1)
        k += 1
        if ps > end_date:
            break
        if until is not None and ps > until.date():
            break
        days = _apply_setpos(_period_candidates(rule, ps, first), rule["BYSETPOS"])
        for d in days:
            c = dt.datetime.combine(d, tpart)
            if c <= dtstart:
                continue
            if until is not None and c > until:
                return out
            produced += 1
            if count is not None and produced > count:
                return out
            if c >= win_end:
                return out
            if c >= win_start:
                out.append(c)
                if len(out) >= limit:
                    return out
    return out


# ---- time zones defined inside the feed -----------------------------------

class VTimezone(dt.tzinfo):
    """tzinfo built from an ICS VTIMEZONE block (what Outlook feeds use)."""

    def __init__(self, tzid: str, comp: Component):
        self.tzid = tzid
        self.observances = []
        for c in comp.children:
            if c.name not in ("STANDARD", "DAYLIGHT"):
                continue
            start, to = c.first("DTSTART"), c.first("TZOFFSETTO")
            if not start or not to:
                continue
            try:
                onset = parse_dt_value(start[1], {})
                if isinstance(onset, dt.date) and not isinstance(onset, dt.datetime):
                    onset = dt.datetime.combine(onset, dt.time())
                onset = onset.replace(tzinfo=None)
                off_to = parse_offset(to[1])
                frm = c.first("TZOFFSETFROM")
                off_from = parse_offset(frm[1]) if frm else off_to
            except ValueError:
                continue
            rr = c.first("RRULE")
            rule = parse_rrule(rr[1]) if rr else None
            rdates = []
            for p, v in c.every("RDATE"):
                for item in v.split(","):
                    try:
                        rd = parse_dt_value(item, {})
                        if isinstance(rd, dt.datetime):
                            rdates.append(rd.replace(tzinfo=None))
                    except ValueError:
                        pass
            self.observances.append((c.name, onset, off_from, off_to, rule, rdates))
        self._cache: dict = {}

    def _onsets(self, year: int) -> list:
        hit = self._cache.get(year)
        if hit is not None:
            return hit
        ws, we = dt.datetime(year - 1, 1, 1), dt.datetime(year + 1, 1, 1)
        lst = []
        for kind, onset, off_from, off_to, rule, rdates in self.observances:
            if rule:
                until = None
                if rule.get("UNTIL"):
                    try:
                        u = parse_dt_value(rule["UNTIL"], {})
                        until = (u if isinstance(u, dt.datetime) else dt.datetime.combine(u, dt.time(23, 59, 59)))
                        until = until.replace(tzinfo=None) + off_from
                    except ValueError:
                        until = None
                for occ in expand_rrule(onset, rule, ws, we, until):
                    lst.append((occ, off_to, off_from, kind))
            elif onset < we:
                lst.append((onset, off_to, off_from, kind))
            for rd in rdates:
                if rd < we:
                    lst.append((rd, off_to, off_from, kind))
        lst.sort(key=lambda x: x[0])
        if len(self._cache) > 64:
            self._cache.clear()
        self._cache[year] = lst
        return lst

    def _lookup(self, naive: dt.datetime):
        lst = self._onsets(naive.year)
        cur = None
        for onset, off_to, off_from, kind in lst:
            if onset <= naive:
                cur = (off_to, kind)
            else:
                break
        if cur is None:
            if lst:
                return lst[0][2], "STANDARD"
            if self.observances:
                return self.observances[0][3], self.observances[0][0]
            return dt.timedelta(0), "STANDARD"
        return cur

    def utcoffset(self, d):
        return self._lookup(d.replace(tzinfo=None))[0] if d is not None else None

    def dst(self, d):
        return dt.timedelta(0)

    def tzname(self, d):
        return self.tzid

    def fromutc(self, d):
        naive = d.replace(tzinfo=None)
        guess = naive + self._lookup(naive)[0]
        return (naive + self._lookup(guess)[0]).replace(tzinfo=self)

    def __repr__(self):
        return f"VTimezone({self.tzid!r})"


def zone_by_name(name: str):
    """A named time zone ('Australia/Melbourne'), or None if unknown.

    Python on Windows has no time-zone database unless the 'tzdata' package is installed,
    so there every name is unknown.
    """
    if ZoneInfo is None or not name:
        return None
    try:
        return ZoneInfo(name)
    except Exception:
        return None


def has_tz_database() -> bool:
    return zone_by_name("UTC") is not None


def timezone_message(names: list) -> str:
    shown = ", ".join(names[:3]) + (" and others" if len(names) > 3 else "")
    if has_tz_database():
        return (f"Some items use time zones Planner doesn't recognise ({shown}), "
                "so they're shown as if they were in your own time zone.")
    cmd = "py -m pip install tzdata" if sys.platform == "win32" else "python3 -m pip install tzdata"
    return (f"Some items use time zones this computer's Python doesn't know ({shown}), so they're shown "
            f"as if they were in your own time zone. To fix it, close Planner, run {cmd}, then start it again.")


def make_tz_resolver(vtimezones: dict, unknown: set | None = None):
    """unknown, if given, collects the TZIDs that couldn't be resolved."""
    cache: dict = {}

    def resolve(tzid: str):
        key = (tzid or "").strip().strip('"')
        if key in cache:
            return cache[key]
        tz = vtimezones.get(key)
        if tz is None and key.upper() in ("UTC", "GMT", "Z", "ETC/UTC", "COORDINATED UNIVERSAL TIME"):
            tz = dt.timezone.utc
        if tz is None:
            tz = zone_by_name(key) or zone_by_name(WINDOWS_TZ.get(key, ""))
        if tz is None and key and unknown is not None:
            unknown.add(key)
        cache[key] = tz  # None -> treat as local time
        return tz

    return resolve


# ---- events ----------------------------------------------------------------

_MEETING_URL = re.compile(
    r"https://(?:[\w-]+\.)?(?:teams\.microsoft\.com|teams\.live\.com|zoom\.us|[\w-]+\.zoom\.us|meet\.google\.com|"
    r"[\w-]+\.webex\.com)/[^\s<>\"')\]]+", re.I)

_BUSY_MAP = {"FREE": "free", "TENTATIVE": "tentative", "BUSY": "busy", "OOF": "oof",
             "WORKINGELSEWHERE": "elsewhere"}


def _instance_key(value, all_day: bool) -> str:
    if isinstance(value, dt.datetime):
        if all_day:
            return value.date().isoformat()
        if value.tzinfo is None:
            value = value.astimezone()  # floating = this computer's time
        return value.astimezone(dt.timezone.utc).strftime("%Y%m%dT%H%M%S")
    return value.isoformat()


class ICalendar:
    def __init__(self, text: str):
        root = parse_components(text)
        vtz = {}
        for c in root.walk("VTIMEZONE"):
            tzid = c.first("TZID")
            if tzid:
                vtz[tzid[1].strip().strip('"')] = VTimezone(tzid[1].strip(), c)
        self.unknown_zones: set = set()
        self.tzres = make_tz_resolver(vtz, self.unknown_zones)
        self.masters: list = []
        self.singles: list = []
        self.overrides: dict = {}
        self.skipped = 0
        for c in root.walk("VEVENT"):
            try:
                ev = self._read_event(c)
            except (ValueError, TypeError, OverflowError):
                self.skipped += 1
                continue
            if ev is None:
                continue
            if ev["rid"] is not None:
                self.overrides.setdefault(ev["uid"], {})[ev["rid"]] = ev
            elif ev["rule"] or ev["rdates"]:
                self.masters.append(ev)
            else:
                self.singles.append(ev)

    def _read_event(self, c: Component):
        s = c.first("DTSTART")
        if not s:
            return None
        start = parse_dt_value(s[1], s[0], self.tzres)
        all_day = not isinstance(start, dt.datetime)
        e = c.first("DTEND")
        if e:
            end = parse_dt_value(e[1], e[0], self.tzres)
            if all_day and isinstance(end, dt.datetime):
                end = end.date()
            if not all_day and not isinstance(end, dt.datetime):
                end = dt.datetime.combine(end, dt.time()).replace(tzinfo=start.tzinfo)
            if all_day:
                duration = end - start
            elif (start.tzinfo is None) != (end.tzinfo is None):
                duration = dt.timedelta(0)
            elif start.tzinfo is not None and end.tzinfo is not start.tzinfo:
                duration = end.astimezone(dt.timezone.utc) - start.astimezone(dt.timezone.utc)
            else:
                duration = end - start  # same zone: wall-clock length, right for recurring events
        else:
            du = c.first("DURATION")
            duration = parse_duration(du[1]) if du else (dt.timedelta(days=1) if all_day else dt.timedelta(0))
        if duration < dt.timedelta(0):
            duration = dt.timedelta(0)
        if all_day and duration < dt.timedelta(days=1):
            duration = dt.timedelta(days=1)

        def text(key):
            v = c.first(key)
            return unescape_text(v[1]).strip() if v else ""

        status = text("STATUS").upper()
        summary = text("SUMMARY") or "(No title)"
        cancelled = status == "CANCELLED" or summary.lower().startswith(("canceled:", "cancelled:"))
        busy = _BUSY_MAP.get(text("X-MICROSOFT-CDO-BUSYSTATUS").upper())
        if busy is None:
            busy = "free" if text("TRANSP").upper() == "TRANSPARENT" else ("tentative" if status == "TENTATIVE" else "busy")
        location = text("LOCATION")
        url = ""
        for key in ("X-MICROSOFT-SKYPETEAMSMEETINGURL", "X-MICROSOFT-ONLINEMEETINGEXTERNALLINK",
                    "X-GOOGLE-CONFERENCE", "URL"):
            val = text(key)
            if val.startswith("https://"):
                url = val
                break
        if not url:
            m = _MEETING_URL.search(location + "\n" + text("DESCRIPTION"))
            if m:
                url = m.group(0)

        rule = None
        rr = c.first("RRULE")
        if rr:
            rule = parse_rrule(rr[1])
            if rule.get("FREQ") not in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
                rule = None
        exdates = set()
        for p, v in c.every("EXDATE"):
            for item in v.split(","):
                try:
                    exdates.add(_instance_key(parse_dt_value(item, p, self.tzres), all_day))
                except ValueError:
                    pass
        rdates = []
        for p, v in c.every("RDATE"):
            if p.get("VALUE", "").upper() == "PERIOD":
                continue
            for item in v.split(","):
                try:
                    rdates.append(parse_dt_value(item, p, self.tzres))
                except ValueError:
                    pass
        rid = None
        r = c.first("RECURRENCE-ID")
        if r:
            rid = _instance_key(parse_dt_value(r[1], r[0], self.tzres), all_day)
        uid = text("UID") or f"nouid-{id(c)}"
        return {"uid": uid, "start": start, "duration": duration, "all_day": all_day, "title": summary,
                "location": location, "busy": busy, "cancelled": cancelled, "status": status, "url": url,
                "rule": rule, "exdates": exdates, "rdates": rdates, "rid": rid}

    # -- expansion ------------------------------------------------------------
    def _occurrences(self, ev: dict, ws: dt.datetime, we: dt.datetime) -> list:
        """Start values (same type as ev['start']) of a master event near [ws, we)."""
        start, dur = ev["start"], ev["duration"]
        margin = dur + dt.timedelta(days=1)
        if ev["all_day"]:
            s0 = dt.datetime.combine(start, dt.time())
            lo = dt.datetime.combine(ws.date(), dt.time()) - margin
            hi = dt.datetime.combine(we.date(), dt.time()) + dt.timedelta(days=1)
            until = self._until(ev, None, all_day=True)
            occ = [o.date() for o in expand_rrule(s0, ev["rule"], lo, hi, until)] if ev["rule"] else []
            occ += [r if not isinstance(r, dt.datetime) else r.date() for r in ev["rdates"]]
            return occ
        tz = start.tzinfo
        frame = (lambda x: x.astimezone(tz).replace(tzinfo=None)) if tz else (lambda x: to_local_naive(x))
        lo, hi = frame(ws) - margin, frame(we) + dt.timedelta(days=1)
        until = self._until(ev, tz)
        occ = expand_rrule(start.replace(tzinfo=None), ev["rule"], lo, hi, until) if ev["rule"] else []
        out = [o.replace(tzinfo=tz) if tz else o for o in occ]
        for r in ev["rdates"]:
            if isinstance(r, dt.datetime):
                out.append(r.astimezone(tz) if (tz and r.tzinfo) else r)
        return out

    def _until(self, ev: dict, tz, all_day: bool = False):
        rule = ev["rule"]
        if not rule or not rule.get("UNTIL"):
            return None
        try:
            u = parse_dt_value(rule["UNTIL"], {}, self.tzres)
        except ValueError:
            return None
        if all_day:
            d = u.date() if isinstance(u, dt.datetime) else u
            return dt.datetime.combine(d, dt.time(23, 59, 59))
        if not isinstance(u, dt.datetime):
            return dt.datetime.combine(u, dt.time(23, 59, 59))
        if u.tzinfo is not None:
            return u.astimezone(tz).replace(tzinfo=None) if tz else to_local_naive(u)
        return u

    def events_between(self, start_date: dt.date, end_date: dt.date) -> list:
        """Event instances overlapping [start_date, end_date) in this computer's local time."""
        ws = dt.datetime.combine(start_date, dt.time()).astimezone()
        we = dt.datetime.combine(end_date, dt.time()).astimezone()
        out = []
        for ev in self.singles:
            item = self._emit(ev, ev["start"], start_date, end_date, ws, we)
            if item:
                out.append(item)
        for ev in self.masters:
            overrides = self.overrides.get(ev["uid"], {})
            for occ in self._occurrences(ev, ws, we):
                key = _instance_key(occ, ev["all_day"])
                if key in ev["exdates"] or key in overrides:
                    continue
                item = self._emit(ev, occ, start_date, end_date, ws, we, key)
                if item:
                    out.append(item)
        for uid, by_rid in self.overrides.items():
            for rid, ev in by_rid.items():
                item = self._emit(ev, ev["start"], start_date, end_date, ws, we, rid)
                if item:
                    out.append(item)
        out.sort(key=lambda x: (x["start"], x["title"]))
        return out

    def _emit(self, ev, occ_start, start_date, end_date, ws, we, key=None):
        if ev["status"] == "CANCELLED":
            return None
        if ev["all_day"]:
            s = occ_start
            e = s + dt.timedelta(days=max(1, ev["duration"].days))
            if not (s < end_date and e > start_date):
                return None
            return {"id": f"{ev['uid']}|{key or s.isoformat()}", "title": ev["title"], "allDay": True,
                    "start": s.isoformat(), "end": e.isoformat(), "location": ev["location"],
                    "busy": ev["busy"], "cancelled": ev["cancelled"], "url": ev["url"]}
        s = occ_start
        if s.tzinfo is not None:
            e = (s.replace(tzinfo=None) + ev["duration"]).replace(tzinfo=s.tzinfo)
            s_loc, e_loc = to_local_naive(s), to_local_naive(e)
        else:
            s_loc, e_loc = s, s + ev["duration"]
        lo, hi = to_local_naive(ws), to_local_naive(we)
        if not (s_loc < hi and (e_loc > lo or (e_loc == s_loc and s_loc >= lo))):
            return None
        return {"id": f"{ev['uid']}|{key or _instance_key(s, False)}", "title": ev["title"], "allDay": False,
                "start": s_loc.strftime("%Y-%m-%dT%H:%M"), "end": e_loc.strftime("%Y-%m-%dT%H:%M"),
                "location": ev["location"], "busy": ev["busy"], "cancelled": ev["cancelled"], "url": ev["url"]}


# --------------------------------------------------------------------------
# Outlook desktop (Windows, classic Outlook) via PowerShell + COM. Read-only.
# --------------------------------------------------------------------------

_PS_SCRIPT = (
    "$ErrorActionPreference='Stop';"
    "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
    "$ic=[Globalization.CultureInfo]::InvariantCulture;"
    "$s=[datetime]::ParseExact('@START@','yyyy-MM-dd',$ic);"
    "$e=[datetime]::ParseExact('@END@','yyyy-MM-dd',$ic);"
    "$ol=New-Object -ComObject Outlook.Application;"
    "$items=$ol.GetNamespace('MAPI').GetDefaultFolder(9).Items;"
    "$items.Sort('[Start]');"
    "$items.IncludeRecurrences=$true;"
    "$f='[Start] < '''+$e.ToString('g')+''' AND [End] > '''+$s.ToString('g')+'''';"
    "try{$res=$items.Restrict($f)}catch{$res=$items};"
    "$out=New-Object System.Collections.ArrayList;"
    "foreach($a in $res){"
    "if($a.Start -ge $e){break};"
    "if($a.End -le $s){continue};"
    "[void]$out.Add([pscustomobject]@{id=[string]$a.EntryID+'|'+$a.Start.ToString('s');"
    "title=[string]$a.Subject;start=$a.Start.ToString('s');end=$a.End.ToString('s');"
    "allDay=[bool]$a.AllDayEvent;location=[string]$a.Location;busy=[int]$a.BusyStatus});"
    "if($out.Count -ge 3000){break}"
    "};"
    "ConvertTo-Json -InputObject @($out) -Compress"
)

_OUTLOOK_BUSY = {0: "free", 1: "tentative", 2: "busy", 3: "oof", 4: "elsewhere"}


def outlook_desktop_events(start_date: dt.date, end_date: dt.date) -> list:
    if sys.platform != "win32":
        raise SourceError("Reading from the Outlook desktop app only works on Windows. "
                          "Use the published calendar (ICS) link instead.")
    script = _PS_SCRIPT.replace("@START@", start_date.isoformat()).replace("@END@", end_date.isoformat())
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, timeout=120, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except FileNotFoundError:
        raise SourceError("Couldn't find PowerShell on this PC.")
    except subprocess.TimeoutExpired:
        raise SourceError("Outlook took too long to answer. Make sure Outlook is open and not showing a dialog.")
    text = proc.stdout.decode("utf-8", "replace").lstrip("﻿").strip()
    if proc.returncode != 0 or not text.startswith("["):
        err = proc.stderr.decode("utf-8", "replace").strip().splitlines()
        detail = next((ln for ln in err if ln.strip()), "no details")
        raise SourceError("Couldn't read the Outlook desktop app. This needs classic Outlook for Windows "
                          f"(the new Outlook can't be read this way). Details: {detail[:240]}")
    rows = json.loads(text or "[]")
    out = []
    for r in rows:
        try:
            s = dt.datetime.fromisoformat(r["start"])
            e = dt.datetime.fromisoformat(r["end"])
        except (KeyError, ValueError, TypeError):
            continue
        busy = _OUTLOOK_BUSY.get(r.get("busy"), "busy")
        title = r.get("title") or "(No title)"
        loc = r.get("location") or ""
        m = _MEETING_URL.search(loc)
        base = {"id": r.get("id") or f"{title}|{r['start']}", "title": title, "location": loc, "busy": busy,
                "cancelled": title.lower().startswith(("canceled:", "cancelled:")), "url": m.group(0) if m else ""}
        if r.get("allDay"):
            sd, ed = s.date(), e.date()
            if ed <= sd:
                ed = sd + dt.timedelta(days=1)
            base.update(allDay=True, start=sd.isoformat(), end=ed.isoformat())
        else:
            base.update(allDay=False, start=s.strftime("%Y-%m-%dT%H:%M"), end=e.strftime("%Y-%m-%dT%H:%M"))
        out.append(base)
    return out


# --------------------------------------------------------------------------
# Demo data (used with --demo; also handy for trying the app before setup)
# --------------------------------------------------------------------------

def demo_ics(today: dt.date) -> str:
    monday = today - dt.timedelta(days=today.weekday())

    def stamp(day_offset, hh, mm):
        d = monday + dt.timedelta(days=day_offset)
        return f"{d:%Y%m%d}T{hh:02d}{mm:02d}00"

    def day(day_offset):
        return f"{monday + dt.timedelta(days=day_offset):%Y%m%d}"

    ev = []

    def add(lines):
        ev.append("BEGIN:VEVENT\r\n" + "\r\n".join(lines) + "\r\nEND:VEVENT")

    add(["UID:demo-standup", f"DTSTART:{stamp(-14, 9, 0)}", f"DTEND:{stamp(-14, 9, 15)}",
         "RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR", f"EXDATE:{stamp(3, 9, 0)}",
         "SUMMARY:Team check-in", "LOCATION:Microsoft Teams Meeting",
         "DESCRIPTION:Join: https://teams.microsoft.com/l/meetup-join/demo"])
    add(["UID:demo-standup", f"RECURRENCE-ID:{stamp(1, 9, 0)}", f"DTSTART:{stamp(1, 9, 30)}",
         f"DTEND:{stamp(1, 9, 45)}", "SUMMARY:Team check-in (moved)", "LOCATION:Microsoft Teams Meeting"])
    add(["UID:demo-ops", f"DTSTART:{stamp(-7, 10, 0)}", f"DTEND:{stamp(-7, 11, 0)}",
         "RRULE:FREQ=WEEKLY;INTERVAL=1;BYDAY=MO", "SUMMARY:Operations weekly", "LOCATION:Room 3.20"])
    add(["UID:demo-abc", f"DTSTART:{stamp(1, 13, 30)}", f"DTEND:{stamp(1, 15, 30)}",
         "SUMMARY:Leadership meeting", "LOCATION:Boardroom"])
    add(["UID:demo-tent", f"DTSTART:{stamp(1, 14, 0)}", f"DTEND:{stamp(1, 15, 0)}",
         "SUMMARY:Process review working group", "X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE"])
    add(["UID:demo-1on1", f"DTSTART:{stamp(2, 11, 0)}", f"DTEND:{stamp(2, 11, 30)}",
         "SUMMARY:1:1 with Priya", "LOCATION:Microsoft Teams Meeting"])
    add(["UID:demo-lunch", f"DTSTART:{stamp(2, 12, 30)}", f"DTEND:{stamp(2, 13, 30)}",
         "SUMMARY:Lunch", "TRANSP:TRANSPARENT", "X-MICROSOFT-CDO-BUSYSTATUS:FREE"])
    add(["UID:demo-plan", f"DTSTART:{stamp(3, 9, 30)}", f"DTEND:{stamp(3, 12, 0)}",
         "SUMMARY:Planning workshop: 2027 roadmap", "LOCATION:Level 2 meeting room"])
    add(["UID:demo-leave", f"DTSTART;VALUE=DATE:{day(4)}", f"DTEND;VALUE=DATE:{day(5)}",
         "SUMMARY:Working from the other office", "X-MICROSOFT-CDO-BUSYSTATUS:WORKINGELSEWHERE"])
    add(["UID:demo-exec", f"DTSTART:{stamp(4, 15, 0)}", f"DTEND:{stamp(4, 16, 0)}",
         "SUMMARY:Executive briefing prep", "LOCATION:Level 4 office"])
    add(["UID:demo-conf", f"DTSTART;VALUE=DATE:{day(2)}", f"DTEND;VALUE=DATE:{day(4)}",
         "SUMMARY:Industry conference (Wed–Thu)"])
    add(["UID:demo-cancel", f"DTSTART:{stamp(0, 15, 0)}", f"DTEND:{stamp(0, 15, 30)}",
         "SUMMARY:Canceled: Vendor demo"])
    return "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Planner demo//EN\r\n" + "\r\n".join(ev) + "\r\nEND:VCALENDAR\r\n"


def demo_tasks(today: dt.date) -> list:
    monday = today - dt.timedelta(days=today.weekday())

    def d(off):
        return (monday + dt.timedelta(days=off)).isoformat()

    def task(i, content, project, priority=1, due=None, duration=None, recurring=False, description=""):
        t = {"id": f"demo{i}", "content": content, "project_id": project, "priority": priority,
             "description": description, "labels": [], "parent_id": None}
        if due:
            t["due"] = {"date": due, "string": "every weekday" if recurring else "", "is_recurring": recurring}
        if duration:
            t["duration"] = {"amount": duration, "unit": "minute"}
        return t

    yesterday = (today - dt.timedelta(days=1)).isoformat()
    return [
        task(1, "Send leadership meeting papers to the committee", "w", 4, yesterday),
        task(2, "Reply to the team about the schedule change", "w", 3, (today - dt.timedelta(days=3)).isoformat()),
        task(3, "Draft briefing note on the service review", "w", 4, d(0) + "T13:00:00", 90,
             description="Two pages. Lead with the demand trend, then the three options."),
        task(4, "Review steering committee minutes", "w", 2, d(1) + "T11:00:00", 45),
        task(5, "Inbox to zero", "w", 1, d(2) + "T16:30:00", 30, recurring=True),
        task(6, "Book travel for the industry conference", "p", 2, d(1)),
        task(7, "Prep questions for 1:1 with Priya", "w", 1, d(2) + "T10:30:00", 30),
        task(8, "Professional development plan: write the goals section", "p", 3, d(4) + "T08:00:00", 60),
        task(9, "Pick up dry cleaning", "p", 1, d(5)),
        task(10, "Check in on the 2027 roadmap survey responses", "w", 2, d(3)),
        task(11, "Read the industry report on remote work", "w", 1),
        task(12, "Update the risk register", "w", 3),
        task(13, "Book a session with the writing group", "p", 1),
    ]


DEMO_PROJECTS = [{"id": "w", "name": "Work", "color": "blue"}, {"id": "p", "name": "Personal", "color": "green"}]


# --------------------------------------------------------------------------
# The app: fetching, caching and the local web server
# --------------------------------------------------------------------------

def _reason(exc: Exception) -> str:
    r = getattr(exc, "reason", None) or exc
    return str(r)[:160]


class PlannerApp:
    def __init__(self, cfg: Config, demo: bool = False):
        self.cfg = cfg
        self.demo = demo
        self.lock = threading.Lock()
        self.fetch_lock = threading.Lock()
        self._ics = {"url": None, "fetched": 0.0, "cal": None}
        self._tasks = {"key": None, "fetched": 0.0, "tasks": None, "projects": {}}
        self._outlook: dict = {}
        if demo:
            today = dt.date.today()
            self._demo_cal = ICalendar(demo_ics(today))
            self._demo_tasks = {t["id"]: t for t in demo_tasks(today)}
            self._demo_seq = 100

    # -- calendar --------------------------------------------------------------
    def events(self, start: dt.date, end: dt.date, force: bool) -> list:
        if self.demo:
            return self._demo_cal.events_between(start, end)
        source = self.cfg.get("calendar_source")
        ttl = 60.0 * float(self.cfg.get("refresh_minutes") or 5)
        if source == "none":
            return []
        if source == "outlook_desktop":
            key = (start, end)
            hit = self._outlook.get(key)
            if hit and not force and time.time() - hit[0] < ttl:
                return hit[1]
            evs = outlook_desktop_events(start, end)
            if len(self._outlook) > 20:
                self._outlook.clear()
            self._outlook[key] = (time.time(), evs)
            return evs
        url = (self.cfg.get("ics_url") or "").strip()
        if not url:
            return []
        if url.lower().startswith("webcal://"):
            url = "https://" + url[len("webcal://"):]
        age = time.time() - self._ics["fetched"]
        stale = self._ics["url"] != url or self._ics["cal"] is None or age > ttl or (force and age > 15)
        if stale:
            with self.fetch_lock:
                try:
                    status, body = http_request("GET", url, {"Accept": "text/calendar, */*"}, timeout=45)
                except (urllib.error.URLError, socket.timeout, OSError) as exc:
                    raise SourceError(f"Couldn't download your Outlook calendar ({_reason(exc)}).")
                except ValueError:
                    raise SourceError("The calendar link doesn't look like a web address. "
                                      "Paste the ICS link from Outlook into Settings.")
                if status >= 400:
                    raise SourceError(f"Outlook returned HTTP {status} for the calendar link. If you recently "
                                      "unpublished or re-published the calendar, copy the new ICS link into Settings.")
                text = body.decode("utf-8", "replace")
                if "BEGIN:VCALENDAR" not in text[:4000].upper():
                    raise SourceError("The calendar link didn't return a calendar. Make sure you copied the ICS link "
                                      "(ending in .ics), not the HTML one.")
                self._ics = {"url": url, "fetched": time.time(), "cal": ICalendar(text)}
        return self._ics["cal"].events_between(start, end)

    # -- tasks -----------------------------------------------------------------
    def tasks(self, force: bool) -> list:
        if self.demo:
            projects = {p["id"]: p for p in DEMO_PROJECTS}
            return [normalize_task(t, projects) for t in self._demo_tasks.values()]
        token = self.cfg.get("todoist_token")
        if not token:
            return []
        filt = self.cfg.get("todoist_filter") or ""
        key = (token, filt)
        c = self._tasks
        if not force and c["key"] == key and c["tasks"] is not None and time.time() - c["fetched"] < 45:
            return c["tasks"]
        client = TodoistClient(token)
        raw = client.tasks(filt)
        projects = c["projects"] if (c["key"] and c["key"][0] == token and c["projects"] and not force) else None
        if projects is None:
            try:
                projects = {str(p.get("id")): p for p in client.projects()}
            except SourceError:
                projects = {}
        zones: set = set()
        tasks = [normalize_task(t, projects, zones) for t in raw if not t.get("checked") and not t.get("is_deleted")]
        self._tasks = {"key": key, "fetched": time.time(), "tasks": tasks, "projects": projects, "zones": zones}
        return tasks

    def _invalidate_tasks(self):
        self._tasks["fetched"] = 0.0

    def data(self, start: dt.date, end: dt.date, force: bool) -> dict:
        errors = []
        events, tasks = [], []
        try:
            events = self.events(start, end, force)
        except SourceError as exc:
            errors.append({"source": "calendar", "message": str(exc)})
        except Exception as exc:  # keep the app usable whatever the feed contains
            errors.append({"source": "calendar", "message": f"Couldn't read the calendar: {exc}"})
        try:
            tasks = self.tasks(force)
        except SourceError as exc:
            errors.append({"source": "todoist", "message": str(exc)})
        except Exception as exc:
            errors.append({"source": "todoist", "message": f"Couldn't read Todoist: {exc}"})
        zones = self.unknown_zones()
        if zones:
            errors.append({"source": "timezones", "message": timezone_message(zones)})
        return {"events": events, "tasks": tasks, "errors": errors, "demo": self.demo,
                "updated": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")}

    def unknown_zones(self) -> list:
        """Time zones used by the current calendar feed or tasks that couldn't be resolved."""
        if self.demo:
            return []
        found = set(self._tasks.get("zones") or ()) if self.cfg.get("todoist_token") else set()
        cal = self._ics["cal"]
        if cal is not None and self.cfg.get("calendar_source") == "ics":
            found |= cal.unknown_zones
        return sorted(found)

    # -- task actions ----------------------------------------------------------
    def _client(self) -> TodoistClient:
        token = self.cfg.get("todoist_token")
        if not token:
            raise SourceError("Add your Todoist API token in Settings first.")
        return TodoistClient(token)

    def close_task(self, task_id: str):
        if self.demo:
            self._demo_closed = getattr(self, "_demo_closed", {})
            t = self._demo_tasks.pop(task_id, None)
            if t:
                self._demo_closed[task_id] = t
            return
        self._client().close(task_id)
        self._invalidate_tasks()

    def reopen_task(self, task_id: str):
        if self.demo:
            t = getattr(self, "_demo_closed", {}).pop(task_id, None)
            if t:
                self._demo_tasks[task_id] = t
            return
        self._client().reopen(task_id)
        self._invalidate_tasks()

    def schedule_task(self, task_id: str, date_str: str, time_str: str | None, minutes: int | None):
        payload = schedule_payload(date_str, time_str, minutes)
        if self.demo:
            t = self._demo_tasks.get(task_id)
            if t is None:
                raise SourceError("That task no longer exists.")
            due = dict(t.get("due") or {})
            due["date"] = date_str + (f"T{time_str}:00" if time_str else "")
            t["due"] = due
            if minutes and time_str:
                t["duration"] = {"amount": int(minutes), "unit": "minute"}
            return
        self._client().update(task_id, payload)
        self._invalidate_tasks()

    def create_task(self, content: str, date_str: str | None, time_str: str | None, minutes: int | None):
        content = content.strip()
        if not content:
            raise SourceError("Type a task name first.")
        payload = {"content": content[:500]}
        if date_str:
            payload.update(schedule_payload(date_str, time_str, minutes))
        if self.demo:
            self._demo_seq += 1
            tid = f"demo{self._demo_seq}"
            t = {"id": tid, "content": content, "project_id": "w", "priority": 1, "labels": []}
            if date_str:
                t["due"] = {"date": date_str + (f"T{time_str}:00" if time_str else ""), "is_recurring": False}
                if minutes and time_str:
                    t["duration"] = {"amount": int(minutes), "unit": "minute"}
            self._demo_tasks[tid] = t
            return
        self._client().create(payload)
        self._invalidate_tasks()


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
            html = UI_HTML.replace("__PLANNER_TOKEN__", self.server.token)
            return self._send(200, body=html.encode("utf-8"), ctype="text/html; charset=utf-8")
        if not u.path.startswith("/api/") or not self._token_ok(q):
            return self._send(404, {"error": "Not found"})
        app = self.server.app
        if u.path == "/api/settings":
            s = app.cfg.public()
            s["demo"] = app.demo
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
                app._invalidate_tasks()
                app._ics["fetched"] = 0.0
                app._outlook.clear()
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


# --------------------------------------------------------------------------
# Window + lifecycle
# --------------------------------------------------------------------------

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


UI_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Planner</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect x='3' y='5' width='26' height='24' rx='4' fill='%2324405F'/%3E%3Crect x='3' y='5' width='26' height='7' rx='3' fill='%23F0B23A'/%3E%3Cpath d='M10 20l4 4 8-9' stroke='%23fff' stroke-width='3' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
<style>
:root{
  --paper:#F4F5F1;
  --sheet:#FFFFFF;
  --ink:#1E2C3D;
  --ink-2:#56626F;
  --ink-3:#8B95A0;
  --rule:#DADFE2;
  --rule-soft:#ECEEEF;
  --wknd:#EEF0EB;
  --event:#24405F;
  --event-ink:#F4F7FB;
  --marker:#F0B23A;
  --marker-soft:#FBEBC6;
  --p1:#D1453B; --p2:#E3861B; --p3:#2F6FDB; --p4:#A7AFB8;
  --hour:48px;
  --gutter:58px;
  --serif:"Iowan Old Style","Palatino Linotype","Book Antiqua",Palatino,Georgia,serif;
  --sans:"Segoe UI Variable Text","Segoe UI",-apple-system,BlinkMacSystemFont,system-ui,"Helvetica Neue",Arial,sans-serif;
}
*{box-sizing:border-box}
html,body{height:100%;margin:0}
body{font:13px/1.35 var(--sans);color:var(--ink);background:var(--paper);display:flex;flex-direction:column;overflow:hidden;-webkit-font-smoothing:antialiased}
button{font:inherit;color:inherit}
a{color:var(--p3)}
:focus-visible{outline:2px solid var(--p3);outline-offset:2px}
[hidden]{display:none!important}

/* ---------- top bar ---------- */
.bar{display:flex;align-items:center;gap:14px;padding:10px 16px;background:var(--sheet);border-bottom:1px solid var(--rule);min-height:56px}
.bar h1{font:400 23px/1.1 var(--serif);margin:0 6px 0 4px;letter-spacing:-.005em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1 1 auto;min-width:0}
.nav,.tools{display:flex;align-items:center;gap:6px;flex:0 0 auto}
.btn,.icon{border:1px solid var(--rule);background:var(--sheet);border-radius:7px;height:32px;cursor:pointer}
.btn{padding:0 12px;font-weight:600}
.icon{width:32px;display:inline-grid;place-items:center;padding:0}
.btn:hover,.icon:hover{background:#F2F4F6;border-color:#C9CFD4}
.icon svg{width:16px;height:16px}
.icon.spin svg{animation:spin .9s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.seg{display:inline-flex;border:1px solid var(--rule);border-radius:7px;overflow:hidden;height:32px}
.seg button{border:0;background:var(--sheet);padding:0 12px;cursor:pointer;color:var(--ink-2)}
.seg button+button{border-left:1px solid var(--rule)}
.seg button[aria-pressed="true"]{background:var(--ink);color:#fff}
.status{color:var(--ink-3);font-size:12px;min-width:92px;text-align:right;white-space:nowrap}
.only-narrow{display:none}

/* ---------- banner ---------- */
.banner{background:#FFF7E3;border-bottom:1px solid #EED9A4;padding:8px 16px;display:flex;flex-direction:column;gap:4px}
.banner p{margin:0;display:flex;gap:10px;align-items:baseline}
.banner .src{font-weight:600;flex:0 0 auto}
.banner code{white-space:nowrap;font-size:12px}
.banner button{border:0;background:none;color:var(--p3);cursor:pointer;padding:0;font-weight:600;text-decoration:underline;text-underline-offset:2px}
.banner.demo{background:#EEF3F8;border-color:#D3DEE9}

/* ---------- layout ---------- */
.main{flex:1 1 auto;display:flex;min-height:0;position:relative}
.side{width:280px;flex:0 0 280px;background:var(--sheet);border-right:1px solid var(--rule);overflow:auto;padding:14px 12px 24px}
.cal{flex:1 1 auto;display:flex;flex-direction:column;min-width:0}
.cal-head{flex:0 0 auto;background:var(--sheet);border-bottom:1px solid var(--rule);overflow:hidden;scrollbar-gutter:stable}
.scroller{flex:1 1 auto;overflow-y:auto;overflow-x:hidden;scrollbar-gutter:stable;position:relative}
.row,.grid{display:grid;grid-template-columns:var(--gutter) repeat(var(--ncols,7),minmax(0,1fr))}

/* ---------- day headers: the desk-diary numerals ---------- */
.dayhead{border:0;background:none;text-align:left;padding:8px 10px 6px;display:flex;align-items:center;gap:8px;cursor:pointer;border-left:1px solid var(--rule-soft);min-width:0}
.dayhead:hover .wd{color:var(--ink)}
.dayhead .num{font:400 28px/1 var(--serif);font-variant-numeric:lining-nums;min-width:38px;height:38px;display:inline-grid;place-items:center;border-radius:50%;margin-left:-6px}
.dayhead .wd{color:var(--ink-2);font-size:12px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.dayhead .wd b{display:block;font-weight:600;color:inherit}
.dayhead.today .num{background:var(--marker);color:var(--ink)}
.dayhead.past .num{color:var(--ink-3)}

/* ---------- all-day row ---------- */
.allday .gut{font-size:11px;color:var(--ink-3);padding:6px 6px 0 8px;text-align:right}
.ad-cell{border-left:1px solid var(--rule-soft);padding:3px 4px 5px;min-height:30px;display:flex;flex-direction:column;gap:3px;min-width:0}
.ad-cell.wknd{background:var(--wknd)}
.ad-cell.drop{background:var(--marker-soft)}
.more{border:0;background:none;color:var(--ink-2);font-size:11.5px;text-align:left;padding:1px 4px;cursor:pointer}
.more:hover{color:var(--ink);text-decoration:underline}

.chip{display:flex;align-items:center;gap:5px;border-radius:4px;padding:2px 6px;font-size:12px;line-height:1.3;cursor:pointer;min-width:0;white-space:nowrap}
.chip .t{overflow:hidden;text-overflow:ellipsis;min-width:0}
.chip.ev{background:var(--event);color:var(--event-ink)}
.chip.ev.busy-free,.chip.ev.busy-elsewhere{background:#E6ECF2;color:var(--ink)}
.chip.ev.busy-oof{background:#5B4A78}
.chip.task{background:var(--sheet);border:1px solid #C7CED4;border-left:3px solid var(--pc);padding-left:4px}

/* ---------- time grid ---------- */
.grid{position:relative;height:calc(var(--hour) * 24)}
.hours{position:relative}
.hours span{position:absolute;right:8px;transform:translateY(-50%);font-size:11px;color:var(--ink-3);font-variant-numeric:tabular-nums;white-space:nowrap}
.col{position:relative;border-left:1px solid var(--rule);
  background-image:linear-gradient(var(--rule) 1px,transparent 1px),linear-gradient(var(--rule-soft) 1px,transparent 1px);
  background-size:100% var(--hour),100% calc(var(--hour) / 2)}
.col.wknd{background-color:var(--wknd)}
.col.today{background-color:#FDFBF4}
.now{position:absolute;left:-1px;right:0;height:2px;background:var(--marker);z-index:6;pointer-events:none}
.now::before{content:"";position:absolute;left:-5px;top:-4px;width:10px;height:10px;border-radius:50%;background:var(--marker)}

.blk{position:absolute;border-radius:5px;padding:3px 6px 2px;overflow:hidden;font-size:12px;line-height:1.28;cursor:pointer;z-index:2;min-height:18px}
.blk .t{font-weight:600;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.blk .m{display:block;opacity:.82;font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-variant-numeric:tabular-nums}
.blk.tall .t{white-space:normal;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical}
.blk.short{padding-top:1px;padding-bottom:0;display:flex;gap:6px;align-items:baseline}
.blk.short .m{order:-1;flex:0 0 auto}
.blk.short .t{flex:1 1 auto;min-width:0}
.blk.ev{background:var(--event);color:var(--event-ink);box-shadow:0 0 0 1px var(--paper)}
.blk.ev.busy-tentative{background:#E3E9F0;color:var(--ink);border:1px solid var(--event);
  background-image:repeating-linear-gradient(135deg,rgba(36,64,95,.13) 0 5px,transparent 5px 10px)}
.blk.ev.busy-free{background:#EEF2F6;color:var(--ink-2);border:1px solid #C6D1DC}
.blk.ev.busy-oof{background:#5B4A78}
.blk.ev.busy-elsewhere{background:#E9EEF3;color:var(--ink);border:1px dashed var(--event)}
.blk.ev.cut-top{border-top-left-radius:0;border-top-right-radius:0}
.blk.ev.cut-bottom{border-bottom-left-radius:0;border-bottom-right-radius:0}
.cancelled .t{text-decoration:line-through}
.cancelled{opacity:.6}

.blk.task{background:var(--sheet);color:var(--ink);border:1px solid #C3CAD1;border-left:3px solid var(--pc);padding-left:22px;z-index:3;touch-action:none;
  box-shadow:0 1px 0 rgba(30,44,61,.06)}
.blk.task:hover{border-color:#9FA9B3;border-left-color:var(--pc)}
.blk.task .check{position:absolute;left:4px;top:4px}
.blk.task.short .check{top:2px}
.blk .resize{position:absolute;left:0;right:0;bottom:0;height:7px;cursor:ns-resize}
.blk.lifted{opacity:.35}
.recur{display:inline-block;width:11px;height:11px;vertical-align:-1px;margin-left:3px;opacity:.7}

.check{width:14px;height:14px;border-radius:50%;border:1.5px solid var(--pc);background:var(--sheet);padding:0;cursor:pointer;flex:0 0 auto;display:inline-grid;place-items:center}
.check:hover{background:color-mix(in srgb,var(--pc) 18%,white)}
.check svg{width:9px;height:9px;opacity:0}
.check:hover svg{opacity:.8}
.done .check{background:var(--pc)}
.done .check svg{opacity:1;color:#fff}
.done .t{text-decoration:line-through;color:var(--ink-3)}
.done{transition:opacity .35s ease .15s;opacity:.25}

.pc1{--pc:var(--p1)} .pc2{--pc:var(--p2)} .pc3{--pc:var(--p3)} .pc4{--pc:var(--p4)}

.preview{position:absolute;left:3px;right:3px;border:2px dashed var(--ink);border-radius:5px;background:rgba(240,178,58,.18);z-index:8;pointer-events:none;padding:2px 6px;font-size:11.5px;font-weight:600;font-variant-numeric:tabular-nums}
.ghost{position:fixed;z-index:50;pointer-events:none;background:var(--sheet);border:1px solid #9FA9B3;border-left:3px solid var(--pc);border-radius:5px;padding:4px 8px;font-size:12px;max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;box-shadow:0 6px 18px rgba(30,44,61,.18)}
body.dragging,body.dragging *{cursor:grabbing!important;user-select:none}
.new-task{position:absolute;left:3px;right:3px;z-index:9}
.new-task input{width:100%;font:inherit;padding:5px 7px;border:2px solid var(--ink);border-radius:5px;background:var(--sheet);box-shadow:0 4px 12px rgba(30,44,61,.15)}
.ad-cell .new-task{position:static}

/* ---------- sidebar ---------- */
.side h2{font:600 13px/1.3 var(--sans);margin:4px 4px 8px;display:flex;align-items:center;gap:8px}
.side h2 .count{font-weight:400;color:var(--ink-3)}
.side section+section{margin-top:20px}
.side ul{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:2px}
.item{display:grid;grid-template-columns:auto 1fr;gap:2px 9px;align-items:start;padding:6px 6px;border-radius:6px;cursor:grab;touch-action:none}
.item:hover{background:#F2F4F5}
.item .check{margin-top:2px}
.item .t{overflow-wrap:anywhere}
.item .meta{grid-column:2;color:var(--ink-3);font-size:11.5px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.item .meta .late{color:var(--p1)}
.dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:4px;vertical-align:1px}
.disclose{border:0;background:none;padding:0;cursor:pointer;font:inherit;display:flex;align-items:center;gap:6px}
.disclose svg{width:12px;height:12px;transition:transform .15s}
.disclose[aria-expanded="true"] svg{transform:rotate(90deg)}
.filter{width:100%;margin:0 0 8px;font:inherit;padding:6px 8px;border:1px solid var(--rule);border-radius:6px}
.empty{color:var(--ink-3);margin:2px 6px}
.hint{color:var(--ink-3);font-size:12px;margin:22px 6px 0;line-height:1.45}

/* ---------- popover ---------- */
.pop{position:fixed;z-index:40;width:330px;max-width:calc(100vw - 16px);background:var(--sheet);border:1px solid var(--rule);border-radius:10px;box-shadow:0 12px 32px rgba(30,44,61,.18);padding:14px 16px 12px}
.pop h3{font:600 15px/1.3 var(--sans);margin:0 0 6px;padding-right:18px;overflow-wrap:anywhere}
.pop p{margin:0 0 4px;color:var(--ink-2)}
.pop .desc{white-space:pre-wrap;color:var(--ink);margin-top:8px;max-height:140px;overflow:auto;overflow-wrap:anywhere}
.pop .close{position:absolute;top:8px;right:8px;border:0;background:none;width:24px;height:24px;cursor:pointer;color:var(--ink-3);font-size:18px;line-height:1}
.pop .actions{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}
.pop .actions .btn{height:30px;display:inline-flex;align-items:center;text-decoration:none;color:var(--ink)}
.pop .actions .btn.primary{background:var(--ink);color:#fff;border-color:var(--ink)}
.pop .note{font-size:12px;color:var(--ink-3);margin-top:8px}

/* ---------- toast ---------- */
.toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);background:var(--ink);color:#fff;border-radius:8px;padding:10px 14px;z-index:60;display:flex;gap:14px;align-items:center;box-shadow:0 8px 24px rgba(30,44,61,.25);max-width:min(560px,calc(100vw - 32px))}
.toast button{border:0;background:none;color:var(--marker);font-weight:700;cursor:pointer;padding:0}
.toast.err{background:#7A2620}

/* ---------- settings ---------- */
dialog{border:0;border-radius:12px;padding:0;width:min(560px,calc(100vw - 24px));max-height:calc(100vh - 32px);box-shadow:0 24px 60px rgba(30,44,61,.3);color:var(--ink)}
dialog::backdrop{background:rgba(30,44,61,.35)}
dialog form{padding:20px 22px 18px;overflow:auto;max-height:calc(100vh - 32px)}
dialog h2{font:400 24px/1.2 var(--serif);margin:0 0 4px}
dialog .intro{color:var(--ink-2);margin:0 0 14px;line-height:1.5}
fieldset{border:0;padding:0;margin:0 0 18px}
legend{font-weight:700;margin-bottom:8px;padding:0}
label.f{display:block;margin:0 0 10px}
label.f>span{display:block;font-weight:600;margin-bottom:4px}
.f input[type=text],.f input[type=password],.f input[type=url],.f select{width:100%;font:inherit;padding:7px 9px;border:1px solid #C7CED4;border-radius:6px;background:#fff}
.help{color:var(--ink-2);font-size:12px;margin:4px 0 0;line-height:1.45}
.radio{display:flex;gap:8px;align-items:flex-start;margin:0 0 8px;cursor:pointer}
.radio input{margin-top:3px}
.radio small{display:block;color:var(--ink-2)}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:0 14px}
.check-row{display:flex;gap:8px;align-items:center;margin:4px 0 10px}
.dialog-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:6px}
.dialog-actions .primary{background:var(--ink);color:#fff;border-color:var(--ink)}
.form-error{color:#9B2C24;margin:0 0 10px}

@media (max-width:820px){
  .side{display:none}
  .side.open{display:block;position:absolute;inset:0 auto 0 0;z-index:30;box-shadow:8px 0 24px rgba(30,44,61,.15)}
  .only-narrow{display:inline-grid}
  .status{display:none}
  .bar{gap:8px;padding:8px 10px}
  .bar h1{display:none}
  .tools{margin-left:auto}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
</style>
</head>
<body>
<header class="bar">
  <div class="nav">
    <button class="icon only-narrow" data-act="side" aria-label="Show task lists" title="Task lists">
      <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M3 4h10M3 8h10M3 12h7"/></svg></button>
    <button class="icon" data-act="prev" aria-label="Previous" title="Previous (←)">
      <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10 3L5 8l5 5"/></svg></button>
    <button class="icon" data-act="next" aria-label="Next" title="Next (→)">
      <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3l5 5-5 5"/></svg></button>
    <button class="btn" data-act="today" title="Today (T)">Today</button>
  </div>
  <h1 id="range" aria-live="polite"></h1>
  <div class="tools">
    <div class="seg" role="group" aria-label="View">
      <button data-view="day" aria-pressed="false" title="Day view (D)">Day</button><button data-view="week" aria-pressed="true" title="Week view (W)">Week</button>
    </div>
    <span id="status" class="status"></span>
    <button class="icon" data-act="refresh" aria-label="Refresh" title="Refresh (R)">
      <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M13.5 8a5.5 5.5 0 1 1-1.6-3.9"/><path d="M13.5 2.5v3h-3"/></svg></button>
    <button class="icon" data-act="settings" aria-label="Settings" title="Settings">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg></button>
  </div>
</header>
<div id="banner" class="banner" hidden></div>
<div class="main">
  <aside class="side" id="side" aria-label="Task lists"></aside>
  <section class="cal" id="cal">
    <div class="cal-head" id="calHead"></div>
    <div class="scroller" id="scroller"><div class="grid" id="grid"></div></div>
  </section>
</div>
<div id="pop" class="pop" role="dialog" hidden></div>
<div id="toast" class="toast" role="status" hidden></div>

<dialog id="settings" aria-labelledby="settingsTitle">
  <form method="dialog" id="settingsForm" novalidate>
    <h2 id="settingsTitle">Settings</h2>
    <p class="intro" id="settingsIntro" hidden>Connect Todoist and your Outlook calendar to get started. Everything stays on this computer, apart from the requests the app makes to Todoist and Outlook.</p>
    <p class="form-error" id="settingsError" hidden></p>
    <fieldset>
      <legend>Todoist</legend>
      <label class="f"><span>API token</span>
        <input type="password" name="todoist_token" autocomplete="off" spellcheck="false" placeholder="Paste your token">
      </label>
      <p class="help" id="tokenHelp">In Todoist, open Settings › Integrations › Developer and copy the API token.</p>
      <label class="f" style="margin-top:10px"><span>Only show tasks matching a filter (optional)</span>
        <input type="text" name="todoist_filter" spellcheck="false" placeholder="For example: #Work | @work">
      </label>
      <p class="help">Uses Todoist's own filter syntax. Leave blank to show every task.</p>
    </fieldset>
    <fieldset>
      <legend>Outlook calendar</legend>
      <label class="radio"><input type="radio" name="calendar_source" value="ics">
        <span>Published calendar link<small>Works on Windows and Mac. In Outlook on the web, go to Settings › Calendar › Shared calendars › Publish a calendar, choose your calendar and a detail level, then copy the ICS link.</small></span></label>
      <label class="radio" id="desktopOption"><input type="radio" name="calendar_source" value="outlook_desktop">
        <span>Outlook desktop app on this PC<small>Windows with classic Outlook only. Reads your calendar straight from Outlook, so nothing is published.</small></span></label>
      <label class="radio"><input type="radio" name="calendar_source" value="none">
        <span>No calendar<small>Show Todoist tasks only.</small></span></label>
      <label class="f" id="icsField" style="margin-top:8px"><span>ICS link</span>
        <input type="url" name="ics_url" spellcheck="false" placeholder="https://outlook.office365.com/owa/calendar/…/calendar.ics">
      </label>
      <p class="help" id="icsHelp">Anyone who has this link can see what you published, so keep it to yourself. "Can view titles and locations" is enough for this app.</p>
    </fieldset>
    <fieldset>
      <legend>View</legend>
      <div class="cols">
        <label class="f"><span>Week starts on</span>
          <select name="week_start"><option value="1">Monday</option><option value="0">Sunday</option></select></label>
        <label class="f"><span>Scroll to</span>
          <select name="day_start_hour"></select></label>
        <label class="f"><span>Default task length</span>
          <select name="default_task_minutes"><option value="15">15 minutes</option><option value="30">30 minutes</option><option value="45">45 minutes</option><option value="60">1 hour</option><option value="90">1.5 hours</option></select></label>
        <label class="f"><span>Refresh every</span>
          <select name="refresh_minutes"><option value="2">2 minutes</option><option value="5">5 minutes</option><option value="10">10 minutes</option><option value="15">15 minutes</option><option value="30">30 minutes</option></select></label>
      </div>
      <label class="check-row"><input type="checkbox" name="show_weekends"> Show weekends</label>
    </fieldset>
    <div class="dialog-actions">
      <button class="btn" value="cancel" formnovalidate>Cancel</button>
      <button class="btn primary" id="saveSettings" value="save">Save</button>
    </div>
  </form>
</dialog>

<script>
"use strict";
const TOKEN = "__PLANNER_TOKEN__";
const SNAP = 15;
const HOUR_PX = 48;

const state = {
  view: "week", anchor: startOfDay(new Date()), events: [], tasks: [], errors: [], settings: null,
  updated: null, loading: false, demo: false, nodateOpen: false, nodateFilter: "", adExpanded: false,
  firstPaint: true, dragging: false, pendingRender: false,
};
let loadSeq = 0;
const $ = (s, el = document) => el.querySelector(s);

/* ---------- utilities ---------- */
function pad(n) { return String(n).padStart(2, "0"); }
function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function parseYmd(s) { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); }
function parseLocal(s) {
  const [a, b] = s.split("T"); const [y, m, d] = a.split("-").map(Number);
  const [hh, mm] = (b || "00:00").split(":").map(Number); return new Date(y, m - 1, d, hh, mm);
}
function hhmm(min) { return `${pad(Math.floor(min / 60))}:${pad(min % 60)}`; }
function atMinutes(day, min) { return new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, min); }
function esc(s) { return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
function plain(s) { return String(s || "").replace(/\[([^\]]+)\]\((?:https?:)?[^)]*\)/g, "$1").replace(/(\*\*|__)(.+?)\1/g, "$2"); }
function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

const F = {
  time: new Intl.DateTimeFormat(undefined, { hour: "numeric", minute: "2-digit" }),
  hour: new Intl.DateTimeFormat(undefined, { hour: "numeric" }),
  wd: new Intl.DateTimeFormat(undefined, { weekday: "short" }),
  mon: new Intl.DateTimeFormat(undefined, { month: "short" }),
  dayShort: new Intl.DateTimeFormat(undefined, { weekday: "short", day: "numeric", month: "short" }),
  dayLong: new Intl.DateTimeFormat(undefined, { weekday: "long", day: "numeric", month: "long" }),
  dayFull: new Intl.DateTimeFormat(undefined, { weekday: "long", day: "numeric", month: "long", year: "numeric" }),
  range: new Intl.DateTimeFormat(undefined, { day: "numeric", month: "long", year: "numeric" }),
};
function timeRange(s, e) {
  try { return F.time.formatRange(s, e); } catch { return `${F.time.format(s)} – ${F.time.format(e)}`; }
}
function dateRange(a, b) {
  try { return F.range.formatRange(a, b); } catch { return `${F.range.format(a)} – ${F.range.format(b)}`; }
}
function prioClass(p) { return "pc" + (5 - clamp(p || 1, 1, 4)); }
const ICON = {
  check: '<svg viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 5.2l2 2L8 3"/></svg>',
  recur: '<svg class="recur" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-label="Recurring"><path d="M2 5a4 4 0 0 1 7-2.3M10 7a4 4 0 0 1-7 2.3"/><path d="M9.3 1v2h-2M2.7 11V9h2"/></svg>',
  chev: '<svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 2.5L8 6l-3.5 3.5"/></svg>',
};

async function api(path, body) {
  const opts = { method: body === undefined ? "GET" : "POST", headers: { "X-Planner-Token": TOKEN } };
  if (body !== undefined) { opts.headers["Content-Type"] = "application/json"; opts.body = JSON.stringify(body); }
  let res;
  try { res = await fetch(path, opts); }
  catch { throw new Error("The Planner app isn't running. Start it again, then reload this window."); }
  let data = {};
  try { data = await res.json(); } catch { /* empty body */ }
  if (!res.ok) throw new Error(data.error || `Something went wrong (HTTP ${res.status}).`);
  return data;
}

/* ---------- range ---------- */
function visibleDays() {
  if (state.view === "day") return [state.anchor];
  const ws = state.settings ? Number(state.settings.week_start) : 1;
  const start = addDays(state.anchor, -((state.anchor.getDay() - ws + 7) % 7));
  let days = Array.from({ length: 7 }, (_, i) => addDays(start, i));
  if (state.settings && !state.settings.show_weekends) days = days.filter(d => d.getDay() !== 0 && d.getDay() !== 6);
  return days;
}
function step(dir) {
  state.anchor = addDays(state.anchor, dir * (state.view === "day" ? 1 : 7));
  if (state.view === "day" && state.settings && !state.settings.show_weekends) {
    while (state.anchor.getDay() === 0 || state.anchor.getDay() === 6) state.anchor = addDays(state.anchor, dir);
  }
  render(); load();
}
function setView(v) { state.view = v; render(); load(); }
function goToday() { state.anchor = startOfDay(new Date()); render(); load(); scrollToStart(); }

/* ---------- loading ---------- */
async function load(force = false) {
  const days = visibleDays();
  const start = ymd(days[0]), end = ymd(addDays(days[days.length - 1], 1));
  const seq = ++loadSeq;
  state.loading = true; renderStatus();
  try {
    const d = await api(`/api/data?start=${start}&end=${end}&force=${force ? 1 : 0}`);
    if (seq !== loadSeq) return;
    state.events = d.events; state.tasks = d.tasks; state.errors = d.errors; state.demo = d.demo;
    state.updated = new Date();
  } catch (e) {
    if (seq !== loadSeq) return;
    state.errors = [{ source: "app", message: e.message }];
  } finally {
    if (seq === loadSeq) { state.loading = false; render(); }
  }
}

/* ---------- building the calendar ---------- */
function collect(days) {
  const keys = days.map(ymd), set = new Set(keys);
  const allDay = new Map(keys.map(k => [k, []])), timed = new Map(keys.map(k => [k, []]));
  for (const ev of state.events) {
    if (ev.allDay) {
      for (let d = parseYmd(ev.start), end = parseYmd(ev.end); d < end; d = addDays(d, 1)) {
        const k = ymd(d); if (set.has(k)) allDay.get(k).push({ kind: "event", ev });
      }
      continue;
    }
    const s = parseLocal(ev.start), e = parseLocal(ev.end);
    if (e - s >= 864e5) {
      for (let d = startOfDay(s); d < e; d = addDays(d, 1)) { const k = ymd(d); if (set.has(k)) allDay.get(k).push({ kind: "event", ev }); }
      continue;
    }
    let d = startOfDay(s);
    for (let guard = 0; guard < 3; guard++) {
      const next = addDays(d, 1), k = ymd(d);
      const segS = s > d ? s : d, segE = e < next ? e : next;
      if (set.has(k)) timed.get(k).push({ kind: "event", ev, s: (segS - d) / 6e4, e: (segE - d) / 6e4, cutTop: s < d, cutBottom: e > next });
      if (e <= next) break;
      d = next;
    }
  }
  const defMin = Number(state.settings?.default_task_minutes || 30);
  for (const t of state.tasks) {
    if (!t.due || !set.has(t.due.date)) continue;
    if (!t.due.time) { allDay.get(t.due.date).push({ kind: "task", t }); continue; }
    const [hh, mm] = t.due.time.split(":").map(Number), s = hh * 60 + mm;
    timed.get(t.due.date).push({ kind: "task", t, s, e: Math.min(1440, s + (t.duration || defMin)) });
  }
  for (const list of allDay.values()) list.sort((a, b) =>
    (a.kind === b.kind ? 0 : a.kind === "event" ? -1 : 1) || ((b.t?.priority || 0) - (a.t?.priority || 0)));
  return { allDay, timed };
}
function layout(items) {
  items.sort((a, b) => a.s - b.s || (b.e - b.s) - (a.e - a.s));
  let cluster = [], end = -1;
  const flush = () => { const cols = []; for (const it of cluster) { let i = 0; while (i < cols.length && cols[i] > it.s) i++; cols[i] = Math.max(it.e, it.s + 22); it.col = i; } cluster.forEach(it => it.ncol = cols.length); };
  for (const it of items) { if (cluster.length && it.s >= end) { flush(); cluster = []; end = -1; } cluster.push(it); end = Math.max(end, it.e, it.s + 22); }
  if (cluster.length) flush();
  return items;
}

function taskInner(t, extraMeta) {
  return `<button class="check" data-complete="${esc(t.id)}" aria-label="Complete: ${esc(plain(t.content))}" title="Complete">${ICON.check}</button>` +
    `<span class="t">${esc(plain(t.content))}${t.due?.recurring ? ICON.recur : ""}</span>${extraMeta || ""}`;
}
function blockHTML(it) {
  const top = it.s / 60 * HOUR_PX, h = Math.max((it.e - it.s) / 60 * HOUR_PX - 2, 18);
  const w = 100 / it.ncol, left = it.col * w;
  const style = `top:${top}px;height:${h}px;left:calc(${left}% + 2px);width:calc(${w}% - 5px)`;
  const size = h < 34 ? "short" : h >= 70 ? "tall" : "";
  if (it.kind === "event") {
    const ev = it.ev, s = parseLocal(ev.start);
    const meta = size === "short" ? F.time.format(s) : esc([timeRange(s, parseLocal(ev.end)), ev.location].filter(Boolean).join(", "));
    const cls = `blk ev busy-${esc(ev.busy)} ${size} ${ev.cancelled ? "cancelled" : ""} ${it.cutTop ? "cut-top" : ""} ${it.cutBottom ? "cut-bottom" : ""}`;
    return `<div class="${cls}" style="${style}" data-ev="${esc(ev.id)}" tabindex="0" role="button"><span class="t">${esc(ev.title)}</span><span class="m">${meta}</span></div>`;
  }
  const t = it.t, day = parseYmd(t.due.date), s = atMinutes(day, it.s), e = atMinutes(day, it.e);
  const meta = `<span class="m">${size === "short" ? F.time.format(s) : esc(timeRange(s, e) + (t.project ? ", " + t.project : ""))}</span>`;
  return `<div class="blk task ${prioClass(t.priority)} ${size}" style="${style}" data-task="${esc(t.id)}" tabindex="0" role="button">` +
    taskInner(t, meta) + `${t.due.recurring ? "" : '<span class="resize" aria-hidden="true"></span>'}</div>`;
}
function chipHTML(it) {
  if (it.kind === "event") {
    const ev = it.ev;
    return `<div class="chip ev busy-${esc(ev.busy)} ${ev.cancelled ? "cancelled" : ""}" data-ev="${esc(ev.id)}" tabindex="0" role="button"><span class="t">${esc(ev.title)}</span></div>`;
  }
  return `<div class="chip task ${prioClass(it.t.priority)}" data-task="${esc(it.t.id)}" tabindex="0" role="button">${taskInner(it.t)}</div>`;
}

/* ---------- render ---------- */
function render() {
  if (state.dragging) { state.pendingRender = true; return; }
  const days = visibleDays();
  document.documentElement.style.setProperty("--ncols", days.length);
  const first = days[0], last = days[days.length - 1];
  $("#range").textContent = state.view === "day" ? F.dayFull.format(first) : dateRange(first, last);
  document.querySelectorAll(".seg button").forEach(b => b.setAttribute("aria-pressed", String(b.dataset.view === state.view)));
  renderStatus(); renderBanner(); renderSide(); renderCal(days);
  if (state.firstPaint && state.settings) { state.firstPaint = false; scrollToStart(); }
}
function renderStatus() {
  const el = $("#status");
  const btn = $('[data-act="refresh"]');
  btn.classList.toggle("spin", state.loading);
  el.textContent = state.loading ? "Refreshing…" : state.updated ? `Updated ${F.time.format(state.updated)}` : "";
}
function renderBanner() {
  const el = $("#banner"), s = state.settings;
  const rows = [];
  const names = { calendar: "Outlook", todoist: "Todoist", app: "Planner", timezones: "Time zones" };
  for (const e of state.errors) rows.push(`<p><span class="src">${names[e.source] || "Problem"}</span><span>${esc(e.message)}</span>` +
    (e.source === "app" || e.source === "timezones" ? "" : ` <button data-act="settings">Open settings</button>`) + `</p>`);
  if (s && !state.demo) {
    if (!s.has_token) rows.push(`<p><span class="src">Todoist</span><span>Add your API token to see your tasks.</span> <button data-act="settings">Open settings</button></p>`);
    if (s.calendar_source === "ics" && !s.ics_url) rows.push(`<p><span class="src">Outlook</span><span>Add your published calendar link to see your meetings.</span> <button data-act="settings">Open settings</button></p>`);
  }
  if (state.demo) rows.push(`<p><span class="src">Demo</span><span>You're looking at sample data. Start the app without <code>--demo</code> to use your own calendar and tasks.</span></p>`);
  el.innerHTML = rows.join("");
  el.hidden = rows.length === 0;
  el.classList.toggle("demo", state.demo && state.errors.length === 0);
}
function renderSide() {
  const side = $("#side"), today = ymd(new Date());
  const overdue = state.tasks.filter(t => t.due && t.due.date < today)
    .sort((a, b) => a.due.date.localeCompare(b.due.date) || b.priority - a.priority);
  const nodate = state.tasks.filter(t => !t.due).sort((a, b) => b.priority - a.priority || a.content.localeCompare(b.content));
  const focused = document.activeElement && document.activeElement.classList.contains("filter");
  side.innerHTML =
    `<section><h2>Overdue <span class="count">${overdue.length || ""}</span></h2>` +
    (overdue.length ? `<ul>${overdue.map(t => itemHTML(t, true)).join("")}</ul>` : `<p class="empty">Nothing overdue.</p>`) + `</section>` +
    `<section><h2><button class="disclose" data-act="nodate" aria-expanded="${state.nodateOpen}">${ICON.chev}No date</button><span class="count">${nodate.length || ""}</span></h2>` +
    (state.nodateOpen ? `<input class="filter" type="search" placeholder="Filter these tasks" value="${esc(state.nodateFilter)}" aria-label="Filter tasks with no date"><ul id="nodateList"></ul>` : "") +
    `</section>` +
    `<p class="hint">Drag a task onto the calendar to give it a time, or onto the all-day row to give it a date. Double-click an empty slot to add a task there.</p>`;
  if (state.nodateOpen) {
    renderNodate(nodate);
    if (focused) { const f = $(".filter", side); f.focus(); f.setSelectionRange(f.value.length, f.value.length); }
  }
}
function renderNodate(list) {
  const q = state.nodateFilter.trim().toLowerCase();
  const items = (list || state.tasks.filter(t => !t.due)).filter(t => !q || t.content.toLowerCase().includes(q) || t.project.toLowerCase().includes(q));
  const ul = $("#nodateList"); if (!ul) return;
  ul.innerHTML = items.slice(0, 200).map(t => itemHTML(t, false)).join("") || `<li class="empty">${q ? "No matching tasks." : "Every task has a date."}</li>`;
}
function itemHTML(t, showDate) {
  const proj = t.project ? `<span><span class="dot" style="background:${esc(t.projectColor)}"></span>${esc(t.project)}</span>` : "";
  const date = showDate ? `<span class="late">${esc(F.dayShort.format(parseYmd(t.due.date)))}</span>` : "";
  return `<li class="item ${prioClass(t.priority)}" data-task="${esc(t.id)}" tabindex="0" role="button">${taskInner(t)}<span class="meta">${date}${proj}</span></li>`;
}
function renderCal(days) {
  const today = startOfDay(new Date()), tk = ymd(today);
  const { allDay, timed } = collect(days);
  const maxChips = Math.max(0, ...[...allDay.values()].map(l => l.length));
  const limit = state.adExpanded ? Infinity : 3;
  $("#calHead").innerHTML =
    `<div class="row days"><div></div>${days.map(d => {
      const k = ymd(d), cls = k === tk ? "today" : d < today ? "past" : "";
      const sub = (d.getDate() === 1 || d === days[0]) ? `<b>${esc(F.wd.format(d))}</b>${esc(F.mon.format(d))}` : `<b>${esc(F.wd.format(d))}</b>`;
      return `<button class="dayhead ${cls}" data-goto="${k}" title="Open ${esc(F.dayLong.format(d))}"><span class="num">${d.getDate()}</span><span class="wd">${sub}</span></button>`;
    }).join("")}</div>` +
    `<div class="row allday"><div class="gut">All day</div>${days.map(d => {
      const k = ymd(d), list = allDay.get(k), wk = d.getDay() === 0 || d.getDay() === 6 ? "wknd" : "";
      const shown = list.slice(0, limit).map(chipHTML).join("");
      const more = list.length > limit ? `<button class="more" data-act="adexpand">${list.length - limit} more</button>` :
        (state.adExpanded && maxChips > 3 && k === ymd(days[0]) ? `<button class="more" data-act="adexpand">Show less</button>` : "");
      return `<div class="ad-cell ${wk}" data-date="${k}">${shown}${more}</div>`;
    }).join("")}</div>`;

  let hours = "";
  for (let h = 1; h < 24; h++) hours += `<span style="top:${h * HOUR_PX}px">${esc(F.hour.format(new Date(2000, 0, 1, h)))}</span>`;
  const nowMin = (Date.now() - today) / 6e4;
  $("#grid").innerHTML = `<div class="hours">${hours}</div>` + days.map(d => {
    const k = ymd(d), cls = [k === tk ? "today" : "", d.getDay() === 0 || d.getDay() === 6 ? "wknd" : ""].join(" ");
    const blocks = layout(timed.get(k)).map(blockHTML).join("");
    const now = k === tk ? `<div class="now" style="top:${nowMin / 60 * HOUR_PX}px"></div>` : "";
    return `<div class="col ${cls}" data-date="${k}">${blocks}${now}</div>`;
  }).join("");
}
function scrollToStart() {
  const h = Number(state.settings?.day_start_hour ?? 7);
  $("#scroller").scrollTop = Math.max(0, h * HOUR_PX - 6);
}
function tickNow() {
  const el = $(".now"); if (!el) return;
  const today = startOfDay(new Date());
  if (!$(`.col[data-date="${ymd(today)}"]`)) return;
  el.style.top = `${(Date.now() - today) / 6e4 / 60 * HOUR_PX}px`;
}

/* ---------- lookups ---------- */
const taskById = id => state.tasks.find(t => t.id === id);
const eventById = id => state.events.find(e => e.id === id);

/* ---------- toast ---------- */
let toastTimer = null;
function toast(msg, opts = {}) {
  const el = $("#toast");
  el.className = "toast" + (opts.error ? " err" : "");
  el.innerHTML = `<span>${esc(msg)}</span>` + (opts.action ? `<button type="button">${esc(opts.action)}</button>` : "");
  el.hidden = false;
  if (opts.action) $("button", el).onclick = () => { el.hidden = true; opts.onAction(); };
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, opts.error ? 9000 : 5000);
}

/* ---------- task actions ---------- */
async function completeTask(id) {
  const t = taskById(id); if (!t) return;
  document.querySelectorAll(`[data-task="${CSS.escape(id)}"]`).forEach(el => el.classList.add("done"));
  closePop();
  const idx = state.tasks.indexOf(t);
  setTimeout(() => { state.tasks = state.tasks.filter(x => x !== t); render(); }, 450);
  try {
    await api("/api/task/close", { id });
    toast(`Completed "${plain(t.content).slice(0, 60)}"`, {
      action: "Undo", onAction: async () => {
        try { await api("/api/task/reopen", { id }); state.tasks.splice(Math.min(idx, state.tasks.length), 0, t); render(); load(); }
        catch (e) { toast(e.message, { error: true }); }
      }
    });
  } catch (e) {
    setTimeout(() => { if (!state.tasks.includes(t)) state.tasks.splice(idx, 0, t); render(); }, 500);
    toast(e.message, { error: true });
  }
}
function describeWhen(date, time) {
  const d = parseYmd(date);
  return time ? `${F.dayShort.format(d)}, ${F.time.format(atMinutes(d, toMin(time)))}` : F.dayShort.format(d);
}
function toMin(time) { const [h, m] = time.split(":").map(Number); return h * 60 + m; }
async function scheduleTask(t, date, time, duration) {
  const before = { due: t.due ? { ...t.due } : null, duration: t.duration };
  t.due = { ...(t.due || {}), date, time: time || null, recurring: false, string: "" };
  if (time && duration) t.duration = duration;
  render();
  try {
    await api("/api/task/schedule", { id: t.id, date, time: time || null, duration: time ? duration : null });
    toast(`Scheduled for ${describeWhen(date, time)}`);
    load();
  } catch (e) {
    t.due = before.due; t.duration = before.duration; render();
    toast(e.message, { error: true });
  }
}
async function createTask(content, date, time) {
  const duration = Number(state.settings?.default_task_minutes || 30);
  try {
    await api("/api/task/create", { content, date, time, duration: time ? duration : null });
    toast(`Added "${content.slice(0, 60)}" to your Todoist inbox`);
    load();
  } catch (e) { toast(e.message, { error: true }); }
}

/* ---------- popover ---------- */
const pop = $("#pop");
function openPop(html, anchor) {
  pop.innerHTML = `<button class="close" data-act="closepop" aria-label="Close">×</button>` + html;
  pop.hidden = false;
  const r = anchor.getBoundingClientRect(), pw = pop.offsetWidth, ph = pop.offsetHeight;
  let x = r.right + 8;
  if (x + pw > innerWidth - 8) x = r.left - pw - 8;
  if (x < 8) x = clamp(r.left, 8, innerWidth - pw - 8);
  let y = clamp(r.top, 8, Math.max(8, innerHeight - ph - 8));
  pop.style.left = `${x}px`; pop.style.top = `${y}px`;
  const focusable = pop.querySelector(".actions .btn"); if (focusable) focusable.focus({ preventScroll: true });
}
function closePop() { pop.hidden = true; pop.innerHTML = ""; }
const BUSY_LABEL = { busy: "Busy", tentative: "Tentative", free: "Free", oof: "Away", elsewhere: "Working elsewhere" };
function showEvent(ev, anchor) {
  let when;
  if (ev.allDay) {
    const s = parseYmd(ev.start), e = addDays(parseYmd(ev.end), -1);
    when = +s === +e ? `${F.dayLong.format(s)}, all day` : `${F.dayShort.format(s)} to ${F.dayShort.format(e)}, all day`;
  } else {
    const s = parseLocal(ev.start), e = parseLocal(ev.end);
    when = ymd(s) === ymd(e) || (e - startOfDay(e) === 0 && e - s <= 864e5)
      ? `${F.dayLong.format(s)}, ${timeRange(s, e)}` : `${F.dayShort.format(s)} ${F.time.format(s)} to ${F.dayShort.format(e)} ${F.time.format(e)}`;
  }
  const join = ev.url ? `<div class="actions"><a class="btn primary" href="${esc(ev.url)}" target="_blank" rel="noopener">Join meeting</a></div>` : "";
  openPop(`<h3>${esc(ev.title)}</h3><p>${esc(when)}</p>${ev.location ? `<p>${esc(ev.location)}</p>` : ""}` +
    `<p>${esc(BUSY_LABEL[ev.busy] || "Busy")}${ev.cancelled ? ", cancelled" : ""}</p>${join}` +
    `<p class="note">From Outlook. Edit it in Outlook.</p>`, anchor);
}
const PRIO_LABEL = { 4: "Priority 1", 3: "Priority 2", 2: "Priority 3", 1: "No priority" };
function showTask(t, anchor) {
  let when = "No date";
  if (t.due) {
    const d = parseYmd(t.due.date);
    if (t.due.time) {
      const s = atMinutes(d, toMin(t.due.time)), dur = t.duration || Number(state.settings?.default_task_minutes || 30);
      when = `${F.dayLong.format(d)}, ${timeRange(s, new Date(+s + dur * 6e4))}` + (t.duration ? "" : " (no length set)");
    } else when = `${F.dayLong.format(d)}, no time`;
  }
  const bits = [t.project ? `<span class="dot" style="background:${esc(t.projectColor)}"></span>${esc(t.project)}` : "", PRIO_LABEL[t.priority] || ""]
    .filter(Boolean).join(", ");
  const recur = t.due?.recurring ? `<p>Repeats: ${esc(t.due.string || "recurring")}</p>` : "";
  const unset = t.due?.time && !t.due.recurring ? `<button class="btn" data-act="untime" data-id="${esc(t.id)}">Remove time</button>` : "";
  openPop(`<h3>${esc(plain(t.content))}</h3><p>${esc(when)}</p><p>${bits}</p>${recur}` +
    (t.labels.length ? `<p>${t.labels.map(l => "@" + esc(l)).join(" ")}</p>` : "") +
    (t.description ? `<p class="desc">${esc(t.description)}</p>` : "") +
    `<div class="actions"><button class="btn primary" data-act="complete" data-id="${esc(t.id)}">Complete</button>${unset}` +
    `<a class="btn" href="${esc(t.url)}" target="_blank" rel="noopener">Open in Todoist</a></div>` +
    (t.due?.recurring ? `<p class="note">Recurring tasks are rescheduled in Todoist so their pattern stays intact.</p>` : ""), anchor);
}

/* ---------- drag to schedule / resize ---------- */
let drag = null, suppressClick = false, autoScroll = null;
const scroller = $("#scroller");

function onPointerDown(e) {
  if (e.button !== 0 || e.target.closest(".check")) return;
  const el = e.target.closest("[data-task]");
  if (!el || el.closest("#pop")) return;
  const t = taskById(el.dataset.task); if (!t) return;
  const resize = !!e.target.closest(".resize");
  const r = el.getBoundingClientRect();
  drag = { el, t, x0: e.clientX, y0: e.clientY, mode: resize ? "resize" : "move", started: false, target: null,
    grabMin: el.classList.contains("blk") && !resize ? ((e.clientY - r.top) / HOUR_PX) * 60 : 0 };
  el.setPointerCapture(e.pointerId);
  el.addEventListener("pointermove", onPointerMove);
  el.addEventListener("pointerup", onPointerUp);
  el.addEventListener("pointercancel", endDrag);
}
function startDrag() {
  const { t, el } = drag;
  if (t.due?.recurring) {
    toast("Recurring tasks keep their pattern in Todoist, so change their date or time there.");
    endDrag(); return false;
  }
  drag.started = true; state.dragging = true; closePop();
  document.body.classList.add("dragging");
  if (drag.mode === "move") {
    el.classList.add("lifted");
    const g = document.createElement("div");
    g.className = `ghost ${prioClass(t.priority)}`; g.textContent = plain(t.content);
    document.body.appendChild(g); drag.ghost = g;
  }
  return true;
}
function onPointerMove(e) {
  if (!drag) return;
  if (!drag.started) {
    if (Math.hypot(e.clientX - drag.x0, e.clientY - drag.y0) < 5) return;
    if (!startDrag()) return;
  }
  drag.lastX = e.clientX; drag.lastY = e.clientY;
  updateDrag();
  const r = scroller.getBoundingClientRect();
  const overX = e.clientX >= r.left && e.clientX <= r.right;
  const edge = !overX ? 0 : (e.clientY >= r.top && e.clientY < r.top + 28) ? -1 : (e.clientY > r.bottom - 28 && e.clientY < r.bottom + 40) ? 1 : 0;
  if (edge && !autoScroll) autoScroll = setInterval(() => { scroller.scrollTop += edge * 14; updateDrag(); }, 30);
  if (!edge && autoScroll) { clearInterval(autoScroll); autoScroll = null; }
}
function clearDropMarks() {
  document.querySelectorAll(".preview").forEach(p => p.remove());
  document.querySelectorAll(".ad-cell.drop").forEach(c => c.classList.remove("drop"));
}
function updateDrag() {
  const { t } = drag, x = drag.lastX, y = drag.lastY;
  const defMin = Number(state.settings?.default_task_minutes || 30);
  if (drag.mode === "resize") {
    const col = drag.el.closest(".col"); if (!col) return;
    const r = col.getBoundingClientRect(), start = toMin(t.due.time);
    const endMin = clamp(Math.round(((y - r.top) / HOUR_PX * 60) / SNAP) * SNAP, start + SNAP, 1440);
    drag.duration = endMin - start;
    drag.el.style.height = `${Math.max(drag.duration / 60 * HOUR_PX - 2, 18)}px`;
    const m = drag.el.querySelector(".m");
    if (m) m.textContent = timeRange(atMinutes(parseYmd(t.due.date), start), atMinutes(parseYmd(t.due.date), endMin));
    return;
  }
  if (drag.ghost) { drag.ghost.style.left = `${x + 12}px`; drag.ghost.style.top = `${y + 10}px`; }
  const under = document.elementFromPoint(x, y);
  const col = under && under.closest(".col"), cell = under && under.closest(".ad-cell");
  clearDropMarks();
  drag.target = null;
  if (col) {
    const r = col.getBoundingClientRect(), dur = t.duration || defMin;
    const min = clamp(Math.round((((y - r.top) / HOUR_PX) * 60 - drag.grabMin) / SNAP) * SNAP, 0, 1440 - SNAP);
    const p = document.createElement("div");
    p.className = "preview";
    p.style.top = `${min / 60 * HOUR_PX}px`;
    p.style.height = `${Math.max(Math.min(dur, 1440 - min) / 60 * HOUR_PX - 2, 18)}px`;
    p.textContent = F.time.format(atMinutes(parseYmd(col.dataset.date), min));
    col.appendChild(p);
    if (drag.ghost) drag.ghost.hidden = true;
    drag.target = { date: col.dataset.date, time: hhmm(min), duration: dur };
  } else if (cell) {
    cell.classList.add("drop");
    if (drag.ghost) drag.ghost.hidden = false;
    drag.target = { date: cell.dataset.date, time: null, duration: null };
  } else if (drag.ghost) drag.ghost.hidden = false;
}
function onPointerUp() {
  if (!drag) return;
  const d = drag;
  endDrag();
  if (!d.started) return;
  suppressClick = true; setTimeout(() => { suppressClick = false; }, 0);
  if (d.mode === "resize") {
    if (d.duration && d.duration !== d.t.duration) scheduleTask(d.t, d.t.due.date, d.t.due.time, d.duration);
    else render();
    return;
  }
  if (!d.target) return;
  const same = d.t.due && d.t.due.date === d.target.date && (d.t.due.time || null) === d.target.time;
  if (!same) scheduleTask(d.t, d.target.date, d.target.time, d.target.duration);
}
function endDrag() {
  if (!drag) return;
  const { el, ghost } = drag;
  el.removeEventListener("pointermove", onPointerMove);
  el.removeEventListener("pointerup", onPointerUp);
  el.removeEventListener("pointercancel", endDrag);
  el.classList.remove("lifted");
  if (ghost) ghost.remove();
  clearDropMarks();
  if (autoScroll) { clearInterval(autoScroll); autoScroll = null; }
  document.body.classList.remove("dragging");
  const wasDragging = state.dragging;
  drag = null; state.dragging = false;
  if (wasDragging && state.pendingRender) { state.pendingRender = false; render(); }
}

/* ---------- quick add by double-click ---------- */
function quickAdd(container, date, time) {
  document.querySelectorAll(".new-task").forEach(n => n.remove());
  const box = document.createElement("div");
  box.className = "new-task";
  if (time) box.style.top = `${toMin(time) / 60 * HOUR_PX}px`;
  const label = time ? `New task at ${F.time.format(atMinutes(parseYmd(date), toMin(time)))}` : `New task for ${F.dayShort.format(parseYmd(date))}`;
  box.innerHTML = `<input type="text" placeholder="${esc(label)}" aria-label="${esc(label)}" maxlength="500">`;
  container.appendChild(box);
  const input = $("input", box);
  input.focus();
  let done = false;
  const finish = (save) => {
    if (done) return; done = true;
    const v = input.value.trim(); box.remove();
    if (save && v) createTask(v, date, time);
  };
  input.addEventListener("keydown", ev => { if (ev.key === "Enter") finish(true); if (ev.key === "Escape") finish(false); ev.stopPropagation(); });
  input.addEventListener("blur", () => finish(true));
}

/* ---------- settings ---------- */
const dlg = $("#settings"), form = $("#settingsForm");
(function fillHours() {
  const sel = form.elements.day_start_hour;
  for (let h = 0; h <= 12; h++) sel.add(new Option(F.hour.format(new Date(2000, 0, 1, h)), h));
})();
function openSettings(intro = false) {
  const s = state.settings || {};
  if (state.demo) { toast("Settings are off in demo mode. Start the app without --demo to connect your accounts."); return; }
  $("#settingsIntro").hidden = !intro;
  $("#settingsError").hidden = true;
  form.elements.todoist_token.value = "";
  form.elements.todoist_token.placeholder = s.has_token ? `Saved (ends ${s.token_hint.replace("…", "")}). Paste a new one to replace it.` : "Paste your token";
  form.elements.todoist_filter.value = s.todoist_filter || "";
  form.elements.ics_url.value = s.ics_url || "";
  for (const r of form.elements.calendar_source) r.checked = r.value === (s.calendar_source || "ics");
  $("#desktopOption").hidden = s.platform !== "win32" && s.calendar_source !== "outlook_desktop";
  form.elements.week_start.value = String(s.week_start ?? 1);
  form.elements.day_start_hour.value = String(s.day_start_hour ?? 7);
  form.elements.default_task_minutes.value = String(s.default_task_minutes ?? 30);
  form.elements.refresh_minutes.value = String(s.refresh_minutes ?? 5);
  form.elements.show_weekends.checked = s.show_weekends !== false;
  syncIcsField();
  dlg.showModal();
}
function syncIcsField() {
  const ics = form.elements.calendar_source.value === "ics";
  $("#icsField").hidden = !ics; $("#icsHelp").hidden = !ics;
}
form.addEventListener("change", e => { if (e.target.name === "calendar_source") syncIcsField(); });
form.addEventListener("submit", async e => {
  if (e.submitter && e.submitter.value === "cancel") return;
  e.preventDefault();
  const f = form.elements, src = f.calendar_source.value;
  const url = f.ics_url.value.trim();
  const fail = msg => { const err = $("#settingsError"); err.textContent = msg; err.hidden = false; err.scrollIntoView({ block: "nearest" }); };
  if (src === "ics" && url && !/^(https?|webcal):\/\//i.test(url)) return fail("The ICS link should start with https://. Copy it from Outlook's Publish a calendar section.");
  if (src === "ics" && /\.html?(\?|#|$)/i.test(url)) return fail("That's the HTML link. Outlook shows two links when you publish; copy the ICS one (it ends in .ics).");
  const body = {
    todoist_token: f.todoist_token.value.trim(), todoist_filter: f.todoist_filter.value.trim(),
    calendar_source: src, ics_url: url, week_start: Number(f.week_start.value), day_start_hour: Number(f.day_start_hour.value),
    default_task_minutes: Number(f.default_task_minutes.value), refresh_minutes: Number(f.refresh_minutes.value),
    show_weekends: f.show_weekends.checked,
  };
  const btn = $("#saveSettings"); btn.disabled = true;
  try {
    state.settings = await api("/api/settings", body);
    dlg.close(); restartTimers(); render(); scrollToStart(); load(true);
    toast("Settings saved");
  } catch (err) {
    const el = $("#settingsError"); el.textContent = err.message; el.hidden = false;
  } finally { btn.disabled = false; }
});

/* ---------- events ---------- */
document.addEventListener("pointerdown", onPointerDown);
document.addEventListener("click", e => {
  if (suppressClick) return;
  const t = e.target;
  const complete = t.closest("[data-complete]");
  if (complete) { e.stopPropagation(); completeTask(complete.dataset.complete); return; }
  const act = t.closest("[data-act]");
  if (act) {
    const a = act.dataset.act;
    if (a === "prev") step(-1); else if (a === "next") step(1); else if (a === "today") goToday();
    else if (a === "refresh") load(true); else if (a === "settings") { closePop(); openSettings(); }
    else if (a === "closepop") closePop();
    else if (a === "complete") completeTask(act.dataset.id);
    else if (a === "untime") { const tk = taskById(act.dataset.id); closePop(); if (tk) scheduleTask(tk, tk.due.date, null, null); }
    else if (a === "nodate") { state.nodateOpen = !state.nodateOpen; renderSide(); if (state.nodateOpen) $(".filter")?.focus(); }
    else if (a === "adexpand") { state.adExpanded = !state.adExpanded; render(); }
    else if (a === "side") $("#side").classList.toggle("open");
    return;
  }
  const view = t.closest("[data-view]");
  if (view) { setView(view.dataset.view); return; }
  const go = t.closest("[data-goto]");
  if (go) { state.anchor = parseYmd(go.dataset.goto); setView("day"); return; }
  const tEl = t.closest("[data-task]");
  if (tEl && !tEl.closest("#pop")) { const tk = taskById(tEl.dataset.task); if (tk) showTask(tk, tEl); return; }
  const eEl = t.closest("[data-ev]");
  if (eEl) { const ev = eventById(eEl.dataset.ev); if (ev) showEvent(ev, eEl); return; }
  if (!t.closest("#pop")) closePop();
  if (!t.closest("#side") && !t.closest('[data-act="side"]')) $("#side").classList.remove("open");
});
document.addEventListener("dblclick", e => {
  if (e.target.closest("[data-task],[data-ev],.new-task")) return;
  const col = e.target.closest(".col");
  if (col) {
    const r = col.getBoundingClientRect();
    const min = clamp(Math.floor(((e.clientY - r.top) / HOUR_PX * 60) / 30) * 30, 0, 1410);
    quickAdd(col, col.dataset.date, hhmm(min)); return;
  }
  const cell = e.target.closest(".ad-cell");
  if (cell) quickAdd(cell, cell.dataset.date, null);
});
document.addEventListener("input", e => {
  if (e.target.classList.contains("filter")) { state.nodateFilter = e.target.value; renderNodate(); }
});
document.addEventListener("keydown", e => {
  if (e.key === "Escape") { closePop(); $("#side").classList.remove("open"); return; }
  if (dlg.open || e.target.closest("input,textarea,select") || e.metaKey || e.ctrlKey || e.altKey) return;
  const el = e.target.closest && e.target.closest("[data-task],[data-ev]");
  if (el && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); el.click(); return; }
  const k = e.key.toLowerCase();
  if (e.key === "ArrowLeft") step(-1); else if (e.key === "ArrowRight") step(1);
  else if (k === "t") goToday(); else if (k === "d") setView("day"); else if (k === "w") setView("week");
  else if (k === "r") load(true);
});

/* ---------- timers + lifecycle ---------- */
let refreshTimer = null;
function restartTimers() {
  clearInterval(refreshTimer);
  const mins = Number(state.settings?.refresh_minutes || 5);
  refreshTimer = setInterval(() => { if (!state.dragging && !document.hidden) load(); }, mins * 6e4);
}
setInterval(tickNow, 30e3);
setInterval(() => { api("/api/ping", {}).catch(() => {}); }, 20e3);
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) { api("/api/ping", {}).catch(() => {}); if (!state.updated || Date.now() - state.updated > 6e4) load(); }
});
addEventListener("pagehide", () => { try { navigator.sendBeacon(`/api/bye?t=${encodeURIComponent(TOKEN)}`); } catch { } });
addEventListener("resize", () => { if (!pop.hidden) closePop(); });
let lastDay = ymd(new Date());
setInterval(() => { const d = ymd(new Date()); if (d !== lastDay) { lastDay = d; render(); load(); } }, 60e3);

(async function init() {
  if (innerWidth < 820) state.view = "day";
  try { state.settings = await api("/api/settings"); state.demo = !!state.settings.demo; }
  catch (e) { state.errors = [{ source: "app", message: e.message }]; }
  if (state.view === "day" && state.settings && !state.settings.show_weekends) {
    while (state.anchor.getDay() === 0 || state.anchor.getDay() === 6) state.anchor = addDays(state.anchor, 1);
  }
  render(); restartTimers();
  const s = state.settings;
  if (s && !s.demo && !s.has_token && (s.calendar_source === "none" || (s.calendar_source === "ics" && !s.ics_url))) openSettings(true);
  load();
})();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    sys.exit(main())
