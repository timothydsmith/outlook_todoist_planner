"""Where calendar events and tasks come from.

PlannerApp holds one task source and picks one calendar source from the settings. Each source
caches its own data and can report:
  notices()        extra lines for the yellow bar, as {"source": ..., "message": ...}
  unknown_zones()  time-zone names it couldn't resolve (combined into one notice by PlannerApp)
  reset()          forget cached data (called when settings change)
  close()          release anything held open (called when the app stops)
"""
from __future__ import annotations

import datetime as dt
import socket
import threading
import time
import urllib.error

from .config import Config
from .demo import DEMO_PROJECTS, demo_ics, demo_tasks
from .errors import SourceError, _reason
from .ics import ICalendar
from .net import http_request
from .outlook import OutlookBridge, rows_to_events
from .todoist import TodoistClient, normalize_task, schedule_payload


def _ttl(cfg: Config) -> float:
    return 60.0 * float(cfg.get("refresh_minutes") or 5)


# --------------------------------------------------------------------------
# Calendar sources
# --------------------------------------------------------------------------

class CalendarSource:
    """No events. The other calendar sources override what they need."""

    def events(self, start: dt.date, end: dt.date, force: bool) -> list:
        return []

    def notices(self) -> list:
        return []

    def unknown_zones(self) -> set:
        return set()

    def reset(self) -> None:
        pass

    def close(self) -> None:
        pass


class NoCalendar(CalendarSource):
    """The "No calendar" setting: tasks only."""


class DemoCalendar(CalendarSource):
    def __init__(self, today: dt.date):
        self.cal = ICalendar(demo_ics(today))

    def events(self, start, end, force):
        return self.cal.events_between(start, end)


class IcsCalendar(CalendarSource):
    """A published calendar link, downloaded at most once per refresh interval."""

    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.lock = threading.Lock()
        self._cache = {"url": None, "fetched": 0.0, "cal": None}

    def events(self, start, end, force):
        url = (self.cfg.get("ics_url") or "").strip()
        if not url:
            return []
        if url.lower().startswith("webcal://"):
            url = "https://" + url[len("webcal://"):]
        age = time.time() - self._cache["fetched"]
        stale = (self._cache["url"] != url or self._cache["cal"] is None or age > _ttl(self.cfg)
                 or (force and age > 15))
        if stale:
            with self.lock:
                self._cache = {"url": url, "fetched": time.time(), "cal": self._download(url)}
        return self._cache["cal"].events_between(start, end)

    @staticmethod
    def _download(url: str) -> ICalendar:
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
        return ICalendar(text)

    def unknown_zones(self):
        cal = self._cache["cal"]
        return set(cal.unknown_zones) if cal is not None else set()

    def reset(self):
        self._cache["fetched"] = 0.0


class OutlookDesktopCalendar(CalendarSource):
    """Classic Outlook on this PC, run in the background (no window) through a long-lived helper.

    Results are kept per date range. A result read before Outlook had connected to the server
    isn't kept, so the next refresh tries again, and the yellow bar says it may be out of date.
    """

    NOT_CONNECTED = "Outlook hasn't connected to the server yet, so meetings may be out of date."

    def __init__(self, cfg: Config, bridge: OutlookBridge | None = None):
        self.cfg = cfg
        self.bridge = bridge or OutlookBridge()
        self._cache: dict = {}
        self._connected = True

    def events(self, start, end, force):
        key = (start, end)
        hit = self._cache.get(key)
        if hit and not force and time.time() - hit[0] < _ttl(self.cfg):
            self._connected = True
            return hit[1]
        resp = self.bridge.request("events", start=start.isoformat(), end=end.isoformat())
        evs = rows_to_events(resp.get("rows") or [])
        self._connected = bool(resp.get("connected", True))
        if self._connected:
            if len(self._cache) > 20:
                self._cache.clear()
            self._cache[key] = (time.time(), evs)
        return evs

    def notices(self):
        return [] if self._connected else [{"source": "calendar", "message": self.NOT_CONNECTED}]

    def reset(self):
        self._cache.clear()

    def close(self):
        self.bridge.close()


# --------------------------------------------------------------------------
# Task sources
# --------------------------------------------------------------------------

class TodoistTasks:
    """Your Todoist tasks, reused for 45 seconds unless something changes."""

    def __init__(self, cfg: Config):
        self.cfg = cfg
        self._cache = {"key": None, "fetched": 0.0, "tasks": None, "projects": {}, "zones": set()}

    def tasks(self, force: bool) -> list:
        token = self.cfg.get("todoist_token")
        if not token:
            return []
        filt = self.cfg.get("todoist_filter") or ""
        key = (token, filt)
        c = self._cache
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
        self._cache = {"key": key, "fetched": time.time(), "tasks": tasks, "projects": projects, "zones": zones}
        return tasks

    def _client(self) -> TodoistClient:
        token = self.cfg.get("todoist_token")
        if not token:
            raise SourceError("Add your Todoist API token in Settings first.")
        return TodoistClient(token)

    def close(self, task_id: str):
        self._client().close(task_id)
        self.reset()

    def reopen(self, task_id: str):
        self._client().reopen(task_id)
        self.reset()

    def schedule(self, task_id: str, date_str: str, time_str: str | None, minutes: int | None):
        self._client().update(task_id, schedule_payload(date_str, time_str, minutes))
        self.reset()

    def create(self, content: str, date_str: str | None, time_str: str | None, minutes: int | None):
        payload = {"content": content}
        if date_str:
            payload.update(schedule_payload(date_str, time_str, minutes))
        self._client().create(payload)
        self.reset()

    def notices(self) -> list:
        return []

    def unknown_zones(self) -> set:
        return set(self._cache["zones"]) if self.cfg.get("todoist_token") else set()

    def reset(self) -> None:
        self._cache["fetched"] = 0.0


class DemoTasks:
    """Sample tasks held in memory. Changes last until the app stops."""

    def __init__(self, today: dt.date):
        self._tasks = {t["id"]: t for t in demo_tasks(today)}
        self._closed: dict = {}
        self._seq = 100
        self._projects = {p["id"]: p for p in DEMO_PROJECTS}

    def tasks(self, force: bool) -> list:
        return [normalize_task(t, self._projects) for t in self._tasks.values()]

    def close(self, task_id: str):
        t = self._tasks.pop(task_id, None)
        if t:
            self._closed[task_id] = t

    def reopen(self, task_id: str):
        t = self._closed.pop(task_id, None)
        if t:
            self._tasks[task_id] = t

    def schedule(self, task_id: str, date_str: str, time_str: str | None, minutes: int | None):
        schedule_payload(date_str, time_str, minutes)  # same validation as for Todoist
        t = self._tasks.get(task_id)
        if t is None:
            raise SourceError("That task no longer exists.")
        due = dict(t.get("due") or {})
        due["date"] = date_str + (f"T{time_str}:00" if time_str else "")
        t["due"] = due
        if minutes and time_str:
            t["duration"] = {"amount": int(minutes), "unit": "minute"}

    def create(self, content: str, date_str: str | None, time_str: str | None, minutes: int | None):
        if date_str:
            schedule_payload(date_str, time_str, minutes)
        self._seq += 1
        tid = f"demo{self._seq}"
        t = {"id": tid, "content": content, "project_id": "w", "priority": 1, "labels": []}
        if date_str:
            t["due"] = {"date": date_str + (f"T{time_str}:00" if time_str else ""), "is_recurring": False}
            if minutes and time_str:
                t["duration"] = {"amount": int(minutes), "unit": "minute"}
        self._tasks[tid] = t

    def notices(self) -> list:
        return []

    def unknown_zones(self) -> set:
        return set()

    def reset(self) -> None:
        pass
