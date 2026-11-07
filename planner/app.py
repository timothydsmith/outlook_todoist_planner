"""The app: picks where events and tasks come from, and combines them for the page."""
from __future__ import annotations

import datetime as dt
import threading

from .config import Config
from .errors import SourceError
from .outlook import OutlookSync
from .sources import (CalendarSource, DemoCalendar, DemoTasks, IcsCalendar, NoCalendar,
                      OutlookDesktopCalendar, TodoistTasks)
from .timeutil import timezone_message


class PlannerApp:
    def __init__(self, cfg: Config, demo: bool = False, sync_state_path=None):
        self.cfg = cfg
        self.demo = demo
        self.sync_state_path = sync_state_path or cfg.path.with_name("outlook_sync.json")
        self._sync: OutlookSync | None = None
        self._sync_lock = threading.Lock()
        self._sync_notices: list = []     # from turning task events off
        self._sync_off: threading.Thread | None = None   # removing events after they were turned off
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

    # -- tasks in Outlook ------------------------------------------------------
    def _sync_wanted(self) -> bool:
        # A token is required: without one the task list is empty, which would look like "everything done".
        return (not self.demo and bool(self.cfg.get("outlook_writeback"))
                and self.cfg.get("calendar_source") == "outlook_desktop" and bool(self.cfg.get("todoist_token")))

    @property
    def sync(self) -> OutlookSync | None:
        """The Outlook sync, if the settings ask for it."""
        with self._sync_lock:
            off = self._sync_off
            if self._sync is None and self._sync_wanted() and not (off and off.is_alive()):
                self._sync = OutlookSync(self.cfg, self._calendars["outlook_desktop"].bridge, self.task_source,
                                         self.sync_state_path)
                self._sync_notices = []
            return self._sync

    def _request_sync(self) -> None:
        s = self.sync
        if s is not None:
            s.request()

    def _turn_sync_off(self, sync: OutlookSync) -> threading.Thread:
        """Remove the upcoming events Planner made, in the background, then stop syncing."""
        def work():
            sync.stop()
            try:
                sync.remove_all_upcoming()
            except SourceError as exc:
                self._sync_notices = [{"source": "outlook", "message": "Couldn't remove the task events Planner "
                                       f"added to Outlook: {getattr(exc, 'detail', exc)}. Turn \"Add timed tasks "
                                       "to my Outlook calendar\" on and off again to retry."}]
        t = self._sync_off = threading.Thread(target=work, name="outlook-sync-off", daemon=True)
        t.start()
        return t

    def settings_changed(self):
        self.task_source.reset()
        for cal in self._calendars.values():
            cal.reset()
        with self._sync_lock:
            old = self._sync
            if old is not None and not self._sync_wanted():
                self._sync = None
            else:
                old = None
        if old is not None:
            return self._turn_sync_off(old)
        self._request_sync()
        return None

    def close(self) -> None:
        with self._sync_lock:
            sync, self._sync = self._sync, None
        if sync is not None:
            sync.stop()
        for cal in self._calendars.values():
            cal.close()

    # -- reading ---------------------------------------------------------------
    def events(self, start: dt.date, end: dt.date, force: bool) -> list:
        return self.calendar.events(start, end, force)

    def tasks(self, force: bool) -> list:
        return self.task_source.tasks(force)

    def data(self, start: dt.date, end: dt.date, force: bool) -> dict:
        cal, src, sync = self.calendar, self.task_source, self.sync
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
        if sync is not None:
            events = [e for e in events if not e.get("taskId")]   # the task itself is drawn instead
            tasks = [dict(t, inOutlook=t["id"] in sync.tracked) for t in tasks]
            errors += sync.notices()
            sync.request()
        errors += cal.notices() + src.notices() + self._sync_notices
        zones = sorted(cal.unknown_zones() | src.unknown_zones())
        if zones:
            errors.append({"source": "timezones", "message": timezone_message(zones)})
        return {"events": events, "tasks": tasks, "errors": errors, "demo": self.demo,
                "updated": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")}

    # -- task actions ----------------------------------------------------------
    def close_task(self, task_id: str):
        self.task_source.close(task_id)
        self._request_sync()

    def reopen_task(self, task_id: str):
        self.task_source.reopen(task_id)
        self._request_sync()

    def schedule_task(self, task_id: str, date_str: str, time_str: str | None, minutes: int | None):
        self.task_source.schedule(task_id, date_str, time_str, minutes)
        self._request_sync()

    def create_task(self, content: str, date_str: str | None, time_str: str | None, minutes: int | None):
        content = content.strip()
        if not content:
            raise SourceError("Type a task name first.")
        self.task_source.create(content[:500], date_str, time_str, minutes)
        self._request_sync()
