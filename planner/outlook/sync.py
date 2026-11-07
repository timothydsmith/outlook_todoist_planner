"""Keeps one Outlook event per scheduled Todoist task, in both directions.

Each event Planner makes is tagged with the task's id (the PlannerTaskId property), and Planner only
ever changes or deletes tagged events. outlook_sync.json remembers, per task, the event and the last
time Planner wrote to it (the "baseline"). Comparing Outlook and Todoist with that baseline tells
which side moved:

  Outlook          Todoist          what happens
  unchanged        unchanged        nothing
  unchanged        moved            the event moves
  moved/resized    unchanged        the task is rescheduled in Todoist
  moved            moved            Todoist wins (with a notice)
  deleted          unchanged        the task keeps its date but loses its time
  moved            (repeating)      the event is put back (repeating tasks are moved in Todoist)

Only tasks with a time get an event. A new event is made when its task falls between today and
N days ahead; after that it follows the task wherever it goes. When a task is completed, deleted or
loses its time, its upcoming event is removed; one that has already started is kept as a record.

Nothing is written unless both Todoist and Outlook were read successfully, and state is saved only
after a sync finishes, so an interrupted sync simply repeats.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import threading
import time
from pathlib import Path

from ..errors import SourceError
from .bridge import HelperError

BUSY = {"free": 0, "tentative": 1, "busy": 2}
MAX_DELETES = 10          # more than this in one sync looks like a mistake, so none are made
SCAN_BACK_DAYS = 1        # how far back to look for tagged events Planner has lost track of
CLEANUP_DAYS = 400        # how far ahead "remove all upcoming" looks
NOT_CONNECTED = "Outlook hasn't connected to the server yet, so Planner will add tasks to it once it has."

_FMT = "%Y-%m-%dT%H:%M:%S"
_MD_LINK = re.compile(r"\[([^\]]+)\]\((?:https?:)?[^)]*\)")
_MD_BOLD = re.compile(r"(\*\*|__)(.+?)\1")


def _plain(text) -> str:
    """Todoist Markdown to plain text, as the page shows it."""
    return _MD_BOLD.sub(r"\2", _MD_LINK.sub(r"\1", str(text or "")))


def _iso(d: dt.datetime) -> str:
    return d.strftime(_FMT)


def _parse(s: str) -> dt.datetime:
    return dt.datetime.fromisoformat(s)


def wanted(task: dict, show_as: str = "busy", default_minutes: int = 30) -> dict | None:
    """The Outlook event a normalized task should have, or None if it shouldn't have one."""
    due = task.get("due") or {}
    if not due.get("date") or not due.get("time"):
        return None
    try:
        start = dt.datetime.fromisoformat(f"{due['date']}T{due['time']}")
        minutes = max(1, int(task.get("duration") or default_minutes or 30))
    except (TypeError, ValueError):
        return None
    subject = (_plain(task.get("content")).strip() or "(Untitled task)")[:255]
    desc = str(task.get("description") or "").strip()
    body = ((desc + "\n\n") if desc else "") + (
        f"Open in Todoist: {task.get('url') or ''}\n\n"
        "Added by Planner. Moving or resizing this event moves the task in Todoist too.")
    busy = BUSY.get(show_as, BUSY["busy"])
    sig = hashlib.sha1(json.dumps([subject, body, busy]).encode("utf-8")).hexdigest()[:16]
    return {"taskId": str(task["id"]), "subject": subject, "body": body, "busy": busy,
            "start": _iso(start), "end": _iso(start + dt.timedelta(minutes=minutes)),
            "date": due["date"], "recurring": bool(due.get("recurring")), "sig": sig}


class SyncPlan:
    def __init__(self):
        self.todoist: list = []   # (task_id, date, "HH:MM" or None, minutes or None)
        self.creates: list = []   # wanted() dicts
        self.updates: list = []   # (entry_id, wanted() dict with the final start/end)
        self.deletes: list = []   # entry ids
        self.state: dict = {}     # task id -> record, for tasks whose event is kept (creates add theirs when made)
        self.notices: list = []   # plain-English lines for the yellow bar


def reconcile(tasks: list, state: dict, current: dict, tagged: list, now: dt.datetime,
              window_end: dt.datetime, show_as: str = "busy", default_minutes: int = 30) -> SyncPlan:
    """Decide what to change. Pure: reads nothing and writes nothing.

    tasks    every open task the app shows (normalized), so a task missing here is done or gone
    state    task id -> {"entryId", "start", "end", "sig"} from the last finished sync
    current  entry id -> what Outlook says now about each event in state (the helper's `get`)
    tagged   tagged events found between yesterday and window_end (the helper's `tagged`)
    """
    plan = SyncPlan()
    want = {}
    for t in tasks:
        w = wanted(t, show_as, default_minutes)
        if w:
            want[w["taskId"]] = w
    today = dt.datetime.combine(now.date(), dt.time())
    delete_undo: dict = {}        # entry id -> (task id, record) to restore if deletions are held back

    # Where each remembered event stands now.
    alive: dict = {}              # task id -> (record, item)
    gone: dict = {}               # task id -> record, for events deleted in Outlook
    for tid, rec in state.items():
        item = current.get(rec.get("entryId"))
        if not item or item.get("tagged") is False:
            continue              # not ours any more: forget it
        if item.get("missing") or item.get("deleted"):
            gone[tid] = rec
        elif item.get("taskId") == tid:
            alive[tid] = (dict(rec), item)

    # Tagged events Planner isn't tracking (state lost, or a sync cut short): adopt one per task,
    # treating its current time as the baseline so nothing changes in Todoist; remove duplicates.
    alive_entries = {item["entryId"] for _, item in alive.values()}
    for it in sorted(tagged, key=lambda i: i.get("start") or ""):
        eid, tid = it.get("entryId"), it.get("taskId")
        if not eid or not tid or eid in alive_entries or it.get("deleted"):
            continue
        if tid in alive:
            if _parse(it["start"]) > now:
                plan.deletes.append(eid)
            continue
        alive[tid] = ({"entryId": eid, "start": it["start"], "end": it["end"], "sig": None}, it)
        alive_entries.add(eid)
        gone.pop(tid, None)

    for tid, (rec, item) in alive.items():
        w = want.get(tid)
        if w is None:             # completed, deleted, no time any more, or filtered out
            if _parse(item["start"]) > now:
                plan.deletes.append(item["entryId"])
                delete_undo[item["entryId"]] = (tid, rec)
            continue
        base = (rec["start"], rec["end"])
        here = (item["start"], item["end"])
        there = (w["start"], w["end"])
        outlook_moved = here != base and not item.get("allDay")
        todoist_moved = there != base
        final = there
        if outlook_moved and not todoist_moved:
            if w["recurring"]:
                plan.notices.append(f'"{w["subject"]}" repeats, so move it in Todoist. Planner put it back '
                                    "in Outlook.")
            else:
                final = here
                s, e = _parse(here[0]), _parse(here[1])
                plan.todoist.append((tid, s.date().isoformat(), s.strftime("%H:%M"),
                                     max(1, int((e - s).total_seconds() // 60))))
        elif outlook_moved and here != there:
            plan.notices.append(f'"{w["subject"]}" was moved in both Outlook and Todoist. Planner kept the '
                                "Todoist time.")
        if (here != final or item.get("allDay") or item.get("subject") != w["subject"]
                or item.get("busy") != w["busy"] or rec.get("sig") != w["sig"]):
            plan.updates.append((item["entryId"], dict(w, start=final[0], end=final[1])))
        plan.state[tid] = {"entryId": item["entryId"], "start": final[0], "end": final[1], "sig": w["sig"]}

    handled = set(plan.state)
    for tid, rec in gone.items():
        w = want.get(tid)
        if w is None or tid in handled:
            continue
        if w["recurring"]:
            plan.notices.append(f'"{w["subject"]}" repeats, so it can\'t be unscheduled from Outlook. '
                                "Planner added it back; change it in Todoist.")
        elif (w["start"], w["end"]) == (rec["start"], rec["end"]):
            plan.todoist.append((tid, w["date"], None, None))   # you took it off the calendar
            handled.add(tid)
        # otherwise Todoist moved it too: Todoist wins, and it's made again below

    for tid, w in want.items():
        if tid not in handled and today <= _parse(w["start"]) < window_end:
            plan.creates.append(w)

    if len(plan.deletes) > MAX_DELETES:
        plan.notices.append(f"Planner would have removed {len(plan.deletes)} task events from Outlook in one go, "
                            "which looks like a mistake (a changed Todoist filter, perhaps), so it removed none. "
                            "To clear them, turn off \"Add timed tasks to my Outlook calendar\" in Settings.")
        for eid in plan.deletes:
            if eid in delete_undo:
                tid, rec = delete_undo[eid]
                plan.state[tid] = rec
        plan.deletes = []
    return plan


def _put_args(w: dict) -> dict:
    return {k: w[k] for k in ("taskId", "subject", "body", "start", "end", "busy")}


class OutlookSync:
    """Runs reconcile() against the real (or a fake) helper and task source.

    request() asks for a sync in the background, a couple of seconds after the last request, so a
    burst of changes becomes one sync. run_now() syncs straight away (used by tests).
    """

    def __init__(self, cfg, bridge, task_source, state_path: Path, clock=dt.datetime.now, debounce: float = 2.0):
        self.cfg = cfg
        self.bridge = bridge
        self.task_source = task_source
        self.state_path = Path(state_path)
        self.clock = clock
        self.debounce = debounce
        self._lock = threading.Lock()
        self._wake = threading.Event()
        self._thread: threading.Thread | None = None
        self._stopped = False
        self._notices: list = []
        self._last_sig = None
        self._last_check = 0.0
        self.tracked: set = set(self.load_state())   # task ids that have an event in Outlook

    # -- settings and state ------------------------------------------------------
    def _settings(self):
        show_as = self.cfg.get("outlook_show_as") or "busy"
        minutes = int(self.cfg.get("default_task_minutes") or 30)
        days = int(self.cfg.get("outlook_writeback_days") or 14)
        ttl = 60.0 * float(self.cfg.get("refresh_minutes") or 5)
        return show_as, minutes, days, ttl

    def load_state(self) -> dict:
        try:
            data = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}
        tasks = data.get("tasks") if isinstance(data, dict) else None
        if not isinstance(tasks, dict):
            return {}
        return {str(k): v for k, v in tasks.items()
                if isinstance(v, dict) and v.get("entryId") and v.get("start") and v.get("end")}

    def _save_state(self, state: dict) -> None:
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.state_path.with_name(self.state_path.name + ".tmp")
        tmp.write_text(json.dumps({"version": 1, "tasks": state}, indent=1), encoding="utf-8")
        os.replace(tmp, self.state_path)

    def notices(self) -> list:
        return [{"source": "outlook", "message": m} for m in self._notices]

    # -- syncing -----------------------------------------------------------------
    def run_now(self, force: bool = False) -> bool:
        """Sync once. Returns True if it finished. Never raises: problems become notices."""
        with self._lock:
            try:
                self._run(force)
                return True
            except HelperError as exc:
                self._notices = [f"Couldn't update your Outlook calendar: {exc.detail}"]
            except SourceError as exc:
                self._notices = [f"Couldn't update your Outlook calendar. {exc}"]
            except Exception as exc:  # keep the app running whatever happens
                self._notices = [f"Couldn't update your Outlook calendar: {exc}"]
            self._last_sig = None
            return False

    @staticmethod
    def _signature(tasks: list, state: dict, show_as: str, minutes: int, days: int, today: dt.date) -> str:
        wants = sorted((w for w in (wanted(t, show_as, minutes) for t in tasks) if w), key=lambda w: w["taskId"])
        blob = json.dumps([wants, state, days, today.isoformat()], sort_keys=True)
        return hashlib.sha1(blob.encode("utf-8")).hexdigest()

    def _run(self, force: bool) -> None:
        show_as, minutes, days, ttl = self._settings()
        try:
            tasks = self.task_source.tasks(False)
        except SourceError:
            return                # the page already shows the Todoist problem; try again next time
        state = self.load_state()
        now = self.clock()
        today = now.date()
        sig = self._signature(tasks, state, show_as, minutes, days, today)
        if not force and sig == self._last_sig and time.monotonic() - self._last_check < ttl:
            return                # nothing changed here, and Outlook was checked recently
        window_end = dt.datetime.combine(today + dt.timedelta(days=days), dt.time())
        scan = self.bridge.request("tagged", start=(today - dt.timedelta(days=SCAN_BACK_DAYS)).isoformat(),
                                   end=window_end.date().isoformat())
        if not scan.get("connected", True):
            self._notices = [NOT_CONNECTED]
            self._last_sig = None
            return
        current = {}
        ids = [rec["entryId"] for rec in state.values()]
        if ids:
            current = {it.get("entryId"): it for it in self.bridge.request("get", ids=ids).get("items") or []}
        plan = reconcile(tasks, state, current, scan.get("items") or [], now, window_end, show_as, minutes)

        for tid, date, at, mins in plan.todoist:
            self.task_source.schedule(tid, date, at, mins)
        new_state = dict(plan.state)
        for eid in plan.deletes:
            self.bridge.request("remove", entryId=eid)
        for eid, w in plan.updates:
            r = self.bridge.request("put", entryId=eid, **_put_args(w))
            new_state[w["taskId"]].update(start=r.get("start") or w["start"], end=r.get("end") or w["end"])
        for w in plan.creates:
            r = self.bridge.request("put", **_put_args(w))
            new_state[w["taskId"]] = {"entryId": r["entryId"], "start": r.get("start") or w["start"],
                                      "end": r.get("end") or w["end"], "sig": w["sig"]}
        if plan.deletes or plan.updates or plan.creates:
            try:
                self.bridge.request("flush")
            except SourceError:
                pass              # Outlook sends the changes itself a little later
        self._save_state(new_state)
        self.tracked = set(new_state)
        self._notices = list(plan.notices)
        self._last_sig = None if plan.todoist else self._signature(tasks, new_state, show_as, minutes, days, today)
        self._last_check = time.monotonic()

    def remove_all_upcoming(self) -> int:
        """Delete every upcoming event Planner made (to Deleted Items) and forget them all. Returns how many."""
        with self._lock:
            now = self.clock()
            today = now.date()
            state = self.load_state()
            scan = self.bridge.request("tagged", start=today.isoformat(),
                                       end=(today + dt.timedelta(days=CLEANUP_DAYS)).isoformat())
            doomed = {it["entryId"] for it in scan.get("items") or []
                      if it.get("taskId") and not it.get("deleted") and _parse(it["start"]) > now}
            ids = [rec["entryId"] for rec in state.values() if rec["entryId"] not in doomed]
            if ids:
                for it in self.bridge.request("get", ids=ids).get("items") or []:
                    if it.get("taskId") and not it.get("deleted") and _parse(it["start"]) > now:
                        doomed.add(it["entryId"])
            for eid in sorted(doomed):
                self.bridge.request("remove", entryId=eid)
            if doomed:
                try:
                    self.bridge.request("flush")
                except SourceError:
                    pass
            try:
                self.state_path.unlink()
            except OSError:
                pass
            self.tracked = set()
            self._notices = []
            self._last_sig = None
            return len(doomed)

    # -- background ----------------------------------------------------------------
    def request(self) -> None:
        if self._stopped:
            return
        if self._thread is None or not self._thread.is_alive():
            self._thread = threading.Thread(target=self._loop, name="outlook-sync", daemon=True)
            self._thread.start()
        self._wake.set()

    def _loop(self) -> None:
        while not self._stopped:
            self._wake.wait()
            while not self._stopped:   # wait for a quiet moment
                self._wake.clear()
                if not self._wake.wait(self.debounce):
                    break
            if self._stopped:
                return
            self.run_now()

    def stop(self) -> None:
        self._stopped = True
        self._wake.set()
        if self._thread is not None:
            self._thread.join(timeout=30)
