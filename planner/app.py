"""The app: picks where events and tasks come from, and combines them for the page."""
from __future__ import annotations

import datetime as dt

from .config import Config
from .errors import SourceError
from .sources import (CalendarSource, DemoCalendar, DemoTasks, IcsCalendar, NoCalendar,
                      OutlookDesktopCalendar, TodoistTasks)
from .timeutil import timezone_message


class PlannerApp:
    def __init__(self, cfg: Config, demo: bool = False):
        self.cfg = cfg
        self.demo = demo
        if demo:
            today = dt.date.today()
            self.task_source = DemoTasks(today)
            self._calendars = {"demo": DemoCalendar(today)}
        else:
            self.task_source = TodoistTasks(cfg)
            self._calendars = {"ics": IcsCalendar(cfg), "outlook_desktop": OutlookDesktopCalendar(cfg),
                               "none": NoCalendar()}

    @property
    def calendar(self) -> CalendarSource:
        """The calendar source the settings currently ask for."""
        if self.demo:
            return self._calendars["demo"]
        return self._calendars.get(self.cfg.get("calendar_source"), self._calendars["none"])

    def settings_changed(self) -> None:
        self.task_source.reset()
        for cal in self._calendars.values():
            cal.reset()

    def close(self) -> None:
        for cal in self._calendars.values():
            cal.close()

    # -- reading ---------------------------------------------------------------
    def events(self, start: dt.date, end: dt.date, force: bool) -> list:
        return self.calendar.events(start, end, force)

    def tasks(self, force: bool) -> list:
        return self.task_source.tasks(force)

    def data(self, start: dt.date, end: dt.date, force: bool) -> dict:
        cal, src = self.calendar, self.task_source
        errors = []
        events, tasks = [], []
        try:
            events = cal.events(start, end, force)
        except SourceError as exc:
            errors.append({"source": "calendar", "message": str(exc)})
        except Exception as exc:  # keep the app usable whatever the feed contains
            errors.append({"source": "calendar", "message": f"Couldn't read the calendar: {exc}"})
        try:
            tasks = src.tasks(force)
        except SourceError as exc:
            errors.append({"source": "todoist", "message": str(exc)})
        except Exception as exc:
            errors.append({"source": "todoist", "message": f"Couldn't read Todoist: {exc}"})
        errors += cal.notices() + src.notices()
        zones = sorted(cal.unknown_zones() | src.unknown_zones())
        if zones:
            errors.append({"source": "timezones", "message": timezone_message(zones)})
        return {"events": events, "tasks": tasks, "errors": errors, "demo": self.demo,
                "updated": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")}

    # -- task actions ----------------------------------------------------------
    def close_task(self, task_id: str):
        self.task_source.close(task_id)

    def reopen_task(self, task_id: str):
        self.task_source.reopen(task_id)

    def schedule_task(self, task_id: str, date_str: str, time_str: str | None, minutes: int | None):
        self.task_source.schedule(task_id, date_str, time_str, minutes)

    def create_task(self, content: str, date_str: str | None, time_str: str | None, minutes: int | None):
        content = content.strip()
        if not content:
            raise SourceError("Type a task name first.")
        self.task_source.create(content[:500], date_str, time_str, minutes)
