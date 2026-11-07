"""Todoist: the API client and conversion between Todoist tasks and the calendar."""
from __future__ import annotations

import datetime as dt
import json
import os
import socket
import urllib.error
import urllib.parse

from .errors import SourceError, _reason
from .net import http_request
from .timeutil import parse_iso, to_local_naive, zone_by_name


TODOIST_API = os.environ.get("PLANNER_TODOIST_API", "https://api.todoist.com/api/v1")
TODOIST_TASK_URL = "https://app.todoist.com/app/task/{id}"


TODOIST_COLORS = {
    "berry_red": "#b8256f", "red": "#db4035", "orange": "#ff9933", "yellow": "#fad000",
    "olive_green": "#afb83b", "lime_green": "#7ecc49", "green": "#299438",
    "mint_green": "#6accbc", "teal": "#158fad", "sky_blue": "#14aaf5",
    "light_blue": "#96c3eb", "blue": "#4073ff", "grape": "#884dff", "violet": "#af38eb",
    "lavender": "#eb96eb", "magenta": "#e05194", "salmon": "#ff8d85",
    "charcoal": "#808080", "grey": "#b8b8b8", "taupe": "#ccac93",
}


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
