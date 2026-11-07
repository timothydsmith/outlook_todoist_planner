"""Tasks in Outlook, as wired into the app: when the sync runs, what the page gets, and turning it off."""
import datetime as dt
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from planner import app as planner_app
from planner import config
from planner.outlook import OutlookBridge
from planner.sources import OutlookDesktopCalendar

from .test_sync import FAKE, FakeTasks, task

TOMORROW = (dt.date.today() + dt.timedelta(days=1)).isoformat()


class AppSyncTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.log = Path(self.dir.name) / "log.txt"
        self.cfg = config.Config(Path("unused.json"), memory_only=True)
        self.cfg.update({"todoist_token": "secret-token", "calendar_source": "outlook_desktop"})
        self.app = planner_app.PlannerApp(self.cfg, sync_state_path=Path(self.dir.name) / "outlook_sync.json")
        self.tasks = FakeTasks([task("a", TOMORROW, "10:00", 30), task("b", TOMORROW)])
        self.app.task_source = self.tasks
        bridge = OutlookBridge(command=[sys.executable, str(FAKE), "ok", str(self.log)])
        self.app._calendars["outlook_desktop"] = OutlookDesktopCalendar(self.cfg, bridge)

    def tearDown(self):
        self.app.close()
        self.dir.cleanup()

    def cal(self):
        p = Path(str(self.log) + ".cal.json")
        return json.loads(p.read_text(encoding="utf-8"))["items"] if p.exists() else {}

    def test_off_unless_asked_for(self):
        self.assertIsNone(self.app.sync)
        self.cfg.update({"outlook_writeback": True})
        self.assertIsNotNone(self.app.sync)

    def test_needs_outlook_on_this_pc_a_token_and_real_data(self):
        self.cfg.update({"outlook_writeback": True, "calendar_source": "ics"})
        self.assertIsNone(self.app.sync)
        self.cfg.update({"calendar_source": "outlook_desktop", "todoist_token": "__clear__"})
        self.assertIsNone(self.app.sync)
        self.assertIsNone(planner_app.PlannerApp(self.cfg, demo=True).sync)

    def test_settings_are_checked(self):
        self.cfg.update({"outlook_show_as": "maybe", "outlook_writeback_days": 500})
        self.assertEqual((self.cfg.get("outlook_show_as"), self.cfg.get("outlook_writeback_days")), ("busy", 60))
        self.cfg.update({"outlook_writeback_days": 0, "outlook_show_as": "tentative"})
        self.assertEqual((self.cfg.get("outlook_show_as"), self.cfg.get("outlook_writeback_days")), ("tentative", 1))
        self.assertNotIn("todoist_token", self.cfg.public())

    def test_page_data_hides_planner_events_and_marks_tasks(self):
        self.cfg.update({"outlook_writeback": True})
        sync = self.app.sync
        sync.run_now()
        events = [{"id": "1", "title": "Board", "taskId": None}, {"id": "2", "title": "Write report", "taskId": "a"}]
        with mock.patch.object(self.app._calendars["outlook_desktop"], "events", return_value=events), \
                mock.patch.object(sync, "request") as request:
            d = self.app.data(dt.date.today(), dt.date.today() + dt.timedelta(days=7), False)
        self.assertEqual([e["title"] for e in d["events"]], ["Board"])
        self.assertEqual({t["id"]: t["inOutlook"] for t in d["tasks"]}, {"a": True, "b": False})
        request.assert_called_once()

    def test_events_are_shown_as_they_are_when_the_sync_is_off(self):
        events = [{"id": "2", "title": "Write report", "taskId": "a"}]
        with mock.patch.object(self.app._calendars["outlook_desktop"], "events", return_value=events):
            d = self.app.data(dt.date.today(), dt.date.today() + dt.timedelta(days=7), False)
        self.assertEqual(len(d["events"]), 1)
        self.assertNotIn("inOutlook", d["tasks"][0])

    def test_sync_notices_reach_the_page(self):
        self.cfg.update({"outlook_writeback": True})
        sync = self.app.sync
        sync._notices = ["Something to know"]
        with mock.patch.object(self.app._calendars["outlook_desktop"], "events", return_value=[]), \
                mock.patch.object(sync, "request"):
            d = self.app.data(dt.date.today(), dt.date.today(), False)
        self.assertIn({"source": "outlook", "message": "Something to know"}, d["errors"])

    def test_task_actions_ask_for_a_sync(self):
        self.tasks.close = self.tasks.reopen = lambda tid: None
        self.tasks.create = lambda *a: None
        self.cfg.update({"outlook_writeback": True})
        with mock.patch.object(self.app.sync, "request") as request:
            self.app.close_task("a")
            self.app.reopen_task("a")
            self.app.schedule_task("a", TOMORROW, "11:00", 30)
            self.app.create_task("New", None, None, None)
        self.assertEqual(request.call_count, 4)

    def test_turning_it_off_removes_upcoming_events(self):
        self.cfg.update({"outlook_writeback": True})
        self.app.sync.run_now()
        self.assertEqual(len(self.cal()), 1)
        self.cfg.update({"outlook_writeback": False})
        self.app.settings_changed().join(30)
        self.assertTrue(all(v.get("deleted") for v in self.cal().values()))
        self.assertIsNone(self.app.sync)
        self.assertFalse(self.app.sync_state_path.exists())

    def test_switching_calendar_source_also_turns_it_off(self):
        self.cfg.update({"outlook_writeback": True})
        self.app.sync.run_now()
        self.cfg.update({"calendar_source": "none"})
        self.app.settings_changed().join(30)
        self.assertTrue(all(v.get("deleted") for v in self.cal().values()))

    def test_other_settings_changes_keep_it_running(self):
        self.cfg.update({"outlook_writeback": True})
        sync = self.app.sync
        with mock.patch.object(sync, "request") as request:
            self.assertIsNone(self.app.settings_changed())
        self.assertIs(self.app.sync, sync)
        request.assert_called_once()


if __name__ == "__main__":
    unittest.main()
