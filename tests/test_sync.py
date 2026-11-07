"""Two-way sync between scheduled tasks and Outlook events: the rules (reconcile) and the runner (OutlookSync)."""
import datetime as dt
import json
import sys
import tempfile
import time
import unittest
from pathlib import Path

from planner.errors import SourceError
from planner.outlook import OutlookBridge
from planner.outlook.sync import MAX_DELETES, OutlookSync, reconcile, wanted

FAKE = Path(__file__).with_name("fake_helper.py")
NOW = dt.datetime(2026, 11, 10, 9, 0)            # a Tuesday morning
WINDOW_END = dt.datetime(2026, 11, 24)


def task(tid, date=None, time=None, duration=None, recurring=False, content="Write report", description=""):
    due = {"date": date, "time": time, "recurring": recurring, "string": ""} if date else None
    return {"id": tid, "content": content, "description": description, "due": due, "duration": duration,
            "url": f"https://app.todoist.com/app/task/{tid}", "priority": 1, "labels": []}


def item(eid, tid, start, end, **kw):
    return dict({"entryId": eid, "taskId": tid, "subject": "Write report", "start": start, "end": end,
                 "allDay": False, "busy": 2, "deleted": False}, **kw)


def tracked(t, eid="E1", start=None, end=None):
    """State and Outlook agreeing with task t as it is now: the 'nothing changed' starting point."""
    w = wanted(t)
    rec = {"entryId": eid, "start": start or w["start"], "end": end or w["end"], "sig": w["sig"]}
    return {t["id"]: rec}, {eid: item(eid, t["id"], rec["start"], rec["end"], subject=w["subject"])}


def run(tasks, state=None, current=None, tagged=(), **kw):
    return reconcile(tasks, state or {}, current or {}, list(tagged), NOW, WINDOW_END, **kw)


class WantedTests(unittest.TestCase):
    def test_timed_task(self):
        w = wanted(task("1", "2026-11-11", "10:00", 45, content="**Board** [papers](https://x.y)"))
        self.assertEqual((w["start"], w["end"], w["busy"]), ("2026-11-11T10:00:00", "2026-11-11T10:45:00", 2))
        self.assertEqual(w["subject"], "Board papers")
        self.assertIn("https://app.todoist.com/app/task/1", w["body"])

    def test_default_length_and_show_as(self):
        w = wanted(task("1", "2026-11-11", "10:00"), show_as="tentative", default_minutes=25)
        self.assertEqual((w["end"], w["busy"]), ("2026-11-11T10:25:00", 1))

    def test_no_event_without_a_time(self):
        self.assertIsNone(wanted(task("1", "2026-11-11")))
        self.assertIsNone(wanted(task("1")))


class ReconcileTests(unittest.TestCase):
    def test_creates_events_only_for_timed_tasks_in_the_window(self):
        plan = run([task("a", "2026-11-10", "13:00"), task("b", "2026-11-11"), task("c", "2026-11-30", "09:00"),
                    task("d", "2026-11-09", "09:00"), task("e", "2026-11-10", "08:00")])
        self.assertEqual(sorted(w["taskId"] for w in plan.creates), ["a", "e"])  # today counts, even if earlier
        self.assertEqual((plan.updates, plan.deletes, plan.todoist), ([], [], []))

    def test_nothing_changed(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        plan = run([t], state, current)
        self.assertEqual((plan.creates, plan.updates, plan.deletes, plan.todoist, plan.notices), ([], [], [], [], []))
        self.assertEqual(plan.state, state)

    def test_moved_in_todoist_moves_the_event(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        t["due"]["time"] = "14:00"
        plan = run([t], state, current)
        [(eid, w)] = plan.updates
        self.assertEqual((eid, w["start"], w["end"]), ("E1", "2026-11-11T14:00:00", "2026-11-11T14:30:00"))
        self.assertEqual(plan.todoist, [])
        self.assertEqual(plan.state["a"]["start"], "2026-11-11T14:00:00")

    def test_moved_in_outlook_moves_the_task(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        current["E1"].update(start="2026-11-12T15:00:00", end="2026-11-12T16:15:00")
        plan = run([t], state, current)
        self.assertEqual(plan.todoist, [("a", "2026-11-12", "15:00", 75)])
        self.assertEqual(plan.updates, [])  # the event is already where it should be
        self.assertEqual((plan.state["a"]["start"], plan.state["a"]["end"]), ("2026-11-12T15:00:00", "2026-11-12T16:15:00"))

    def test_moved_in_both_todoist_wins(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        current["E1"].update(start="2026-11-12T15:00:00", end="2026-11-12T15:30:00")
        t["due"]["time"] = "11:00"
        plan = run([t], state, current)
        self.assertEqual(plan.todoist, [])
        [(_, w)] = plan.updates
        self.assertEqual(w["start"], "2026-11-11T11:00:00")
        self.assertIn("moved in both", plan.notices[0])

    def test_moved_to_the_same_place_in_both_is_quiet(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        current["E1"].update(start="2026-11-11T11:00:00", end="2026-11-11T11:30:00")
        t["due"]["time"] = "11:00"
        plan = run([t], state, current)
        self.assertEqual((plan.todoist, plan.updates, plan.notices), ([], [], []))

    def test_deleted_in_outlook_removes_the_time(self):
        t = task("a", "2026-11-11", "10:00", 30)
        for gone in ({"deleted": True}, {"missing": True}):
            state, current = tracked(t)
            current["E1"].update(gone)
            plan = run([t], state, current)
            self.assertEqual(plan.todoist, [("a", "2026-11-11", None, None)])
            self.assertEqual((plan.creates, plan.state), ([], {}))

    def test_deleted_in_outlook_but_moved_in_todoist_is_made_again(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        current["E1"]["deleted"] = True
        t["due"]["time"] = "12:00"
        plan = run([t], state, current)
        self.assertEqual(plan.todoist, [])
        self.assertEqual([w["start"] for w in plan.creates], ["2026-11-11T12:00:00"])

    def test_repeating_task_moved_in_outlook_is_put_back(self):
        t = task("a", "2026-11-11", "16:30", 30, recurring=True)
        state, current = tracked(t)
        current["E1"].update(start="2026-11-11T17:00:00", end="2026-11-11T17:30:00")
        plan = run([t], state, current)
        self.assertEqual(plan.todoist, [])
        self.assertEqual(plan.updates[0][1]["start"], "2026-11-11T16:30:00")
        self.assertIn("repeats", plan.notices[0])

    def test_repeating_task_deleted_in_outlook_is_added_back(self):
        t = task("a", "2026-11-11", "16:30", 30, recurring=True)
        state, current = tracked(t)
        current["E1"]["deleted"] = True
        plan = run([t], state, current)
        self.assertEqual(plan.todoist, [])
        self.assertEqual(len(plan.creates), 1)
        self.assertIn("repeats", plan.notices[0])

    def test_completing_a_repeating_task_moves_its_event_on(self):
        t = task("a", "2026-11-11", "16:30", 30, recurring=True)
        state, current = tracked(t)
        t["due"]["date"] = "2026-11-12"   # Todoist moved it to the next occurrence
        plan = run([t], state, current)
        self.assertEqual(plan.updates[0][1]["start"], "2026-11-12T16:30:00")

    def test_all_day_or_renamed_in_outlook_is_put_back(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        current["E1"].update(allDay=True, start="2026-11-11T00:00:00", end="2026-11-12T00:00:00")
        plan = run([t], state, current)
        self.assertEqual(plan.todoist, [])
        self.assertEqual(plan.updates[0][1]["start"], "2026-11-11T10:00:00")
        state, current = tracked(t)
        current["E1"]["subject"] = "Something else"
        self.assertEqual(len(run([t], state, current).updates), 1)

    def test_description_or_show_as_change_updates_the_event(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        t["description"] = "New notes"
        self.assertEqual(len(run([t], state, current).updates), 1)
        state, current = tracked(t)
        self.assertEqual(run([t], state, current, show_as="tentative").updates[0][1]["busy"], 1)

    def test_completed_task_removes_upcoming_event_but_keeps_past_ones(self):
        later = task("a", "2026-11-11", "10:00", 30)
        earlier = task("b", "2026-11-10", "08:00", 30)
        s1, c1 = tracked(later, "E1")
        s2, c2 = tracked(earlier, "E2")
        plan = run([], {**s1, **s2}, {**c1, **c2})
        self.assertEqual(plan.deletes, ["E1"])
        self.assertEqual(plan.state, {})

    def test_task_without_a_time_any_more_loses_its_event(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        t["due"]["time"] = None
        self.assertEqual(run([t], state, current).deletes, ["E1"])

    def test_event_follows_its_task_beyond_the_window(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        t["due"]["date"] = "2027-01-15"
        plan = run([t], state, current)
        self.assertEqual((plan.deletes, plan.updates[0][1]["start"]), ([], "2027-01-15T10:00:00"))

    def test_too_many_deletions_are_held_back(self):
        state, current = {}, {}
        for i in range(MAX_DELETES + 1):
            s, c = tracked(task(f"t{i}", "2026-11-12", "10:00"), f"E{i}")
            state.update(s)
            current.update(c)
        plan = run([], state, current)
        self.assertEqual(plan.deletes, [])
        self.assertEqual(plan.state, state)       # still tracked, so a later sync can try again
        self.assertIn("removed none", plan.notices[0])

    def test_lost_state_adopts_tagged_events_without_touching_todoist(self):
        t = task("a", "2026-11-11", "10:00", 30)
        _, current = tracked(t)
        plan = run([t], {}, {}, tagged=current.values())
        self.assertEqual((plan.creates, plan.todoist, plan.deletes), ([], [], []))
        self.assertEqual(plan.state["a"]["entryId"], "E1")

    def test_lost_state_with_an_outlook_move_keeps_the_todoist_time(self):
        t = task("a", "2026-11-11", "10:00", 30)
        found = item("E1", "a", "2026-11-11T15:00:00", "2026-11-11T15:30:00")
        plan = run([t], {}, {}, tagged=[found])
        self.assertEqual(plan.todoist, [])
        self.assertEqual(plan.updates[0][1]["start"], "2026-11-11T10:00:00")

    def test_duplicate_tagged_events_are_removed(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, current = tracked(t)
        dupe = item("E9", "a", "2026-11-11T10:00:00", "2026-11-11T10:30:00")
        plan = run([t], state, current, tagged=[current["E1"], dupe])
        self.assertEqual(plan.deletes, ["E9"])

    def test_untagged_event_is_forgotten_and_the_task_gets_a_new_one(self):
        t = task("a", "2026-11-11", "10:00", 30)
        state, _ = tracked(t)
        plan = run([t], state, {"E1": {"entryId": "E1", "tagged": False}})
        self.assertEqual((len(plan.creates), plan.deletes, plan.todoist), (1, [], []))


class FakeTasks:
    """A task source holding normalized tasks; schedule() changes them as Todoist would."""

    def __init__(self, tasks):
        self.items = {t["id"]: t for t in tasks}
        self.calls = []
        self.fail = False

    def tasks(self, force):
        if self.fail:
            raise SourceError("Couldn't reach Todoist.")
        return [json.loads(json.dumps(t)) for t in self.items.values()]

    def notices(self):
        return []

    def unknown_zones(self):
        return set()

    def reset(self):
        pass

    def schedule(self, tid, date, at, minutes):
        self.calls.append((tid, date, at, minutes))
        t = self.items[tid]
        t["due"].update(date=date, time=at)
        if at and minutes:
            t["duration"] = minutes


class Cfg(dict):
    def get(self, key, default=None):
        return super().get(key, default)


class OutlookSyncTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.log = Path(self.dir.name) / "log.txt"
        self.state_path = Path(self.dir.name) / "outlook_sync.json"
        self.syncs = []

    def tearDown(self):
        for s in self.syncs:
            s.stop()
            s.bridge.close()
        self.dir.cleanup()

    def make(self, tasks, mode="ok", **cfg):
        self.tasks = FakeTasks(tasks)
        bridge = OutlookBridge(command=[sys.executable, str(FAKE), mode, str(self.log)])
        sync = OutlookSync(Cfg(cfg), bridge, self.tasks, self.state_path, clock=lambda: NOW, debounce=0.05)
        self.syncs.append(sync)
        return sync

    def cal(self):
        p = Path(str(self.log) + ".cal.json")
        return json.loads(p.read_text(encoding="utf-8"))["items"] if p.exists() else {}

    def edit_cal(self, eid, **changes):
        p = Path(str(self.log) + ".cal.json")
        data = json.loads(p.read_text(encoding="utf-8"))
        data["items"][eid].update(changes)
        p.write_text(json.dumps(data), encoding="utf-8")

    def ops(self):
        return self.log.read_text(encoding="utf-8").split() if self.log.exists() else []

    def test_round_trip(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30, content="Board papers")])
        self.assertTrue(sync.run_now())
        [(eid, ev)] = self.cal().items()
        self.assertEqual((ev["taskId"], ev["subject"], ev["start"], ev["busy"]), ("a", "Board papers", "2026-11-11T10:00:00", 2))
        self.assertEqual(sync.tracked, {"a"})
        self.assertIn("flush", self.ops())

        # You drag it in Outlook: Todoist follows.
        self.edit_cal(eid, start="2026-11-12T14:00:00", end="2026-11-12T15:00:00")
        self.assertTrue(sync.run_now(force=True))
        self.assertEqual(self.tasks.calls, [("a", "2026-11-12", "14:00", 60)])
        self.assertEqual(len(self.cal()), 1)

        # The next sync finds them agreeing and changes nothing.
        writes = self.ops().count("put")
        self.assertTrue(sync.run_now(force=True))
        self.assertEqual(self.ops().count("put"), writes)
        self.assertEqual(len(self.tasks.calls), 1)

        # Completed in Todoist: the upcoming event goes to Deleted Items.
        del self.tasks.items["a"]
        self.assertTrue(sync.run_now(force=True))
        self.assertTrue(self.cal()[eid]["deleted"])
        self.assertEqual(sync.tracked, set())

    def test_skips_when_nothing_changed(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)])
        sync.run_now()
        n = len(self.ops())
        sync.run_now()                       # not forced, same tasks, checked moments ago
        self.assertEqual(len(self.ops()), n)

    def test_state_survives_a_restart(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)])
        sync.run_now()
        again = OutlookSync(Cfg(), sync.bridge, self.tasks, self.state_path, clock=lambda: NOW)
        self.assertEqual(again.tracked, {"a"})
        again.run_now()
        self.assertEqual(len(self.cal()), 1)

    def test_lost_state_file_creates_no_duplicates(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)])
        sync.run_now()
        self.state_path.unlink()
        sync.run_now(force=True)
        self.assertEqual(len(self.cal()), 1)
        self.assertEqual(self.tasks.calls, [])

    def test_todoist_unreachable_changes_nothing(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)])
        sync.run_now()
        self.tasks.fail = True
        n = len(self.ops())
        self.assertTrue(sync.run_now(force=True))
        self.assertEqual(len(self.ops()), n)
        self.assertEqual(sync.notices(), [])

    def test_not_connected_writes_nothing(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)], mode="disconnected")
        sync.run_now()
        self.assertEqual(self.cal(), {})
        self.assertIn("hasn't connected", sync.notices()[0]["message"])

    def test_outlook_errors_become_a_notice(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)], mode="error")
        self.assertFalse(sync.run_now())
        self.assertIn("Class not registered", sync.notices()[0]["message"])
        self.assertFalse(self.state_path.exists())

    def test_show_as_setting(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)], outlook_show_as="free")
        sync.run_now()
        self.assertEqual(next(iter(self.cal().values()))["busy"], 0)

    def test_remove_all_upcoming(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30), task("b", "2026-11-12", "10:00", 30),
                          task("c", "2026-11-10", "08:00", 30)])
        sync.run_now()
        self.assertEqual(sync.remove_all_upcoming(), 2)   # c started at 08:00 today, so it stays
        deleted = sorted(v["taskId"] for v in self.cal().values() if v.get("deleted"))
        self.assertEqual(deleted, ["a", "b"])
        self.assertFalse(self.state_path.exists())
        self.assertEqual(sync.tracked, set())

    def test_never_touches_untagged_events(self):
        sync = self.make([])
        p = Path(str(self.log) + ".cal.json")
        p.write_text(json.dumps({"next": 2, "items": {"X1": {"subject": "Board", "start": "2026-11-11T10:00:00",
                                                                  "end": "2026-11-11T11:00:00"}}}), encoding="utf-8")
        sync.run_now()
        sync.remove_all_upcoming()
        self.assertNotIn("deleted", self.cal()["X1"])

    def test_requests_are_debounced_in_the_background(self):
        sync = self.make([task("a", "2026-11-11", "10:00", 30)])
        for _ in range(5):
            sync.request()
        deadline = time.monotonic() + 15
        while not self.cal() and time.monotonic() < deadline:
            time.sleep(0.05)
        time.sleep(0.3)
        self.assertEqual(len(self.cal()), 1)
        self.assertEqual(self.ops().count("tagged"), 1)   # five requests, one sync


if __name__ == "__main__":
    unittest.main()
