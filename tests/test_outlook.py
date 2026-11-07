"""Classic Outlook through the long-running helper, with tests/fake_helper.py standing in for PowerShell."""
import datetime as dt
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from planner import config, errors, sources
from planner.outlook import OutlookBridge, bridge as bridge_mod, rows_to_events

FAKE = Path(__file__).with_name("fake_helper.py")


class RowsToEventsTests(unittest.TestCase):
    def test_converts_appointments(self):
        rows = [
            {"id": "E1|x", "title": "Board", "start": "2026-11-10T10:00:00", "end": "2026-11-10T11:00:00",
             "allDay": False, "location": "https://teams.microsoft.com/l/meetup-join/1 Room", "busy": 2},
            {"id": "E2|x", "title": "Leave", "start": "2026-11-12T00:00:00", "end": "2026-11-12T00:00:00",
             "allDay": True, "location": "", "busy": 3},
            {"id": "E3|x", "title": "Cancelled: Demo", "start": "2026-11-13T09:00:00", "end": "2026-11-13T09:30:00",
             "allDay": False, "location": "", "busy": 0},
            {"title": "Broken", "start": "garbage", "end": "2026-11-13T09:30:00"},
        ]
        board, leave, demo = rows_to_events(rows)
        self.assertEqual((board["start"], board["end"], board["busy"]), ("2026-11-10T10:00", "2026-11-10T11:00", "busy"))
        self.assertEqual(board["url"], "https://teams.microsoft.com/l/meetup-join/1")
        self.assertEqual((leave["allDay"], leave["start"], leave["end"], leave["busy"]),
                         (True, "2026-11-12", "2026-11-13", "oof"))
        self.assertTrue(demo["cancelled"])
        self.assertEqual(demo["busy"], "free")

    def test_missing_title_and_id(self):
        [ev] = rows_to_events([{"start": "2026-11-10T10:00:00", "end": "2026-11-10T10:30:00"}])
        self.assertEqual((ev["title"], ev["id"]), ("(No title)", "(No title)|2026-11-10T10:00:00"))


class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.log = Path(self.dir.name) / "log.txt"
        self.bridges = []

    def tearDown(self):
        for b in self.bridges:
            b.close()
        self.dir.cleanup()

    def bridge(self, mode, **kw):
        b = OutlookBridge(command=[sys.executable, str(FAKE), mode, str(self.log)], **kw)
        self.bridges.append(b)
        return b

    def logged(self):
        return self.log.read_text(encoding="utf-8").split() if self.log.exists() else []

    def test_starts_only_when_first_needed(self):
        b = self.bridge("ok")
        self.assertFalse(b.running)
        self.assertEqual(b.request("hello")["version"], 1)
        self.assertTrue(b.running)
        b.request("hello")
        self.assertEqual(self.logged(), ["start", "hello", "hello"])  # one process for both

    def test_events(self):
        rows = self.bridge("ok").request("events", start="2026-11-10", end="2026-11-11")["rows"]
        self.assertEqual(rows[0]["start"], "2026-11-10T10:00:00")

    def test_close_sends_quit_and_stops_the_helper(self):
        b = self.bridge("ok")
        b.request("hello")
        b.close()
        self.assertFalse(b.running)
        self.assertEqual(self.logged()[-1], "quit")
        b.close()  # harmless when already closed

    def test_close_without_ever_starting(self):
        b = self.bridge("ok")
        b.close()
        self.assertEqual(self.logged(), [])

    def test_restarts_once_after_a_crash(self):
        rows = self.bridge("crash_first").request("events", start="2026-11-10", end="2026-11-11")["rows"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(self.logged(), ["start", "events", "start", "events"])

    def test_gives_up_after_a_second_crash(self):
        with self.assertRaises(errors.SourceError) as cm:
            self.bridge("crash").request("events", start="2026-11-10", end="2026-11-11")
        self.assertIn("stopped unexpectedly", str(cm.exception))
        self.assertIn("boom: Outlook went away", str(cm.exception))
        self.assertEqual(self.logged().count("start"), 2)

    def test_hang_times_out(self):
        b = self.bridge("hang", first_timeout=0.5, timeout=0.5)
        t0 = time.monotonic()
        with self.assertRaisesRegex(errors.SourceError, "took too long"):
            b.request("events", start="2026-11-10", end="2026-11-11")
        self.assertLess(time.monotonic() - t0, 10)
        self.assertEqual(self.logged().count("start"), 2)  # retried once with a fresh helper

    def test_error_reply_is_explained_without_restarting(self):
        with self.assertRaises(errors.SourceError) as cm:
            self.bridge("error").request("events", start="2026-11-10", end="2026-11-11")
        self.assertIn("classic Outlook", str(cm.exception))
        self.assertIn("Class not registered", str(cm.exception))
        self.assertEqual(self.logged().count("start"), 1)

    def test_ignores_stray_output_and_stale_replies(self):
        resp = self.bridge("noise").request("hello")
        self.assertNotIn("stale", resp)
        self.assertEqual(resp["version"], 1)

    def test_missing_program(self):
        b = OutlookBridge(command=["definitely-not-a-real-program-1234"])
        with self.assertRaisesRegex(errors.SourceError, "Couldn't find PowerShell"):
            b.request("hello")

    def test_real_helper_only_on_windows(self):
        with mock.patch.object(bridge_mod.sys, "platform", "darwin"):
            with self.assertRaisesRegex(errors.SourceError, "only works on Windows"):
                OutlookBridge().request("hello")


class OutlookCalendarTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.log = Path(self.dir.name) / "log.txt"
        self.cfg = config.Config(Path("unused.json"), memory_only=True)

    def tearDown(self):
        self.cal.close()
        self.dir.cleanup()

    def make(self, mode):
        self.cal = sources.OutlookDesktopCalendar(
            self.cfg, OutlookBridge(command=[sys.executable, str(FAKE), mode, str(self.log)]))
        return self.cal

    def events_requests(self):
        return self.log.read_text(encoding="utf-8").split().count("events")

    def test_connected_results_are_cached(self):
        cal = self.make("ok")
        d = dt.date(2026, 11, 10)
        evs = cal.events(d, d + dt.timedelta(days=1), force=False)
        cal.events(d, d + dt.timedelta(days=1), force=False)
        self.assertEqual([e["title"] for e in evs], ["Board"])
        self.assertEqual(self.events_requests(), 1)
        self.assertEqual(cal.notices(), [])
        cal.events(d, d + dt.timedelta(days=1), force=True)
        self.assertEqual(self.events_requests(), 2)

    def test_not_connected_shows_a_notice_and_tries_again_next_time(self):
        cal = self.make("disconnected")
        d = dt.date(2026, 11, 10)
        self.assertEqual(len(cal.events(d, d + dt.timedelta(days=1), force=False)), 1)
        self.assertEqual(cal.notices(), [{"source": "calendar", "message": cal.NOT_CONNECTED}])
        cal.events(d, d + dt.timedelta(days=1), force=False)
        self.assertEqual(self.events_requests(), 2)

    def test_close_stops_the_helper(self):
        cal = self.make("ok")
        cal.events(dt.date(2026, 11, 10), dt.date(2026, 11, 11), force=False)
        cal.close()
        self.assertFalse(cal.bridge.running)
        self.assertEqual(self.log.read_text(encoding="utf-8").split()[-1], "quit")


if __name__ == "__main__":
    unittest.main()
