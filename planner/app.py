"""The app: fetching and caching calendar events and tasks, and task actions."""
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
from .outlook_desktop import outlook_desktop_events
from .timeutil import timezone_message
from .todoist import TodoistClient, normalize_task, schedule_payload

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
