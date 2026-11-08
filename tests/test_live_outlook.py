"""Reads this week from your real classic Outlook through the real helper. Read-only; nothing is changed.

Skipped unless you ask for it:

    set PLANNER_LIVE_OUTLOOK=1 && py -m unittest tests.test_live_outlook -v       (cmd)
    $env:PLANNER_LIVE_OUTLOOK=1; py -m unittest tests.test_live_outlook -v       (PowerShell)
"""
import datetime as dt
import os
import subprocess
import sys
import time
import unittest

from planner.outlook import OutlookBridge, rows_to_events


def outlook_running() -> bool:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"], capture_output=True, text=True,
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
    return "OUTLOOK.EXE" in out.upper()


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("PLANNER_LIVE_OUTLOOK") == "1",
                     "set PLANNER_LIVE_OUTLOOK=1 to read your real Outlook (read-only)")
class LiveOutlookTests(unittest.TestCase):
    def test_read_this_week_then_close(self):
        bridge = OutlookBridge()
        today = dt.date.today()
        monday = today - dt.timedelta(days=today.weekday())
        t0 = time.monotonic()
        try:
            hello = bridge.request("hello")
            resp = bridge.request("events", start=monday.isoformat(), end=(monday + dt.timedelta(days=7)).isoformat())
            took = time.monotonic() - t0
        finally:
            bridge.close()
        self.assertEqual(hello["version"], 1)
        self.assertIsInstance(resp["rows"], list)
        events = rows_to_events(resp["rows"])
        self.assertEqual(len(events), len(resp["rows"]), "every row should convert")
        for ev in events:
            self.assertTrue(ev["start"] and ev["end"] and "title" in ev)
        print(f"\n  {len(events)} events this week; connected={resp['connected']} (mode {resp['connection']}); "
              f"{took:.1f}s including start-up; helper started Outlook: {hello['startedOutlook']}")
        if hello["startedOutlook"]:
            deadline = time.monotonic() + 30
            while outlook_running() and time.monotonic() < deadline:
                time.sleep(0.5)
            self.assertFalse(outlook_running(), "the background Outlook should close within 30s of quitting")


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("PLANNER_LIVE_OUTLOOK_WRITE") == "1",
                     "set PLANNER_LIVE_OUTLOOK_WRITE=1 to create, move and delete one test event in your real Outlook")
class LiveOutlookWriteTests(unittest.TestCase):
    """Writes ONE clearly labelled test event (shown as Free) three days ahead, moves it, then deletes it.

        $env:PLANNER_LIVE_OUTLOOK_WRITE=1; py -m unittest tests.test_live_outlook.LiveOutlookWriteTests -v
    """
    TASK_ID = "planner-live-test"

    def test_create_move_and_remove_a_tagged_event(self):
        day = dt.date.today() + dt.timedelta(days=3)
        at = lambda hh, mm: f"{day.isoformat()}T{hh:02d}:{mm:02d}:00"
        event = {"taskId": self.TASK_ID, "subject": "Planner test event (safe to delete)",
                 "body": "Made by Planner's live test. It deletes this event itself a moment later.", "busy": 0}
        bridge = OutlookBridge()
        eid = None

        def find():
            scan = bridge.request("tagged")
            self.assertTrue(scan["connected"])
            return {it["entryId"]: it for it in scan["items"]}.get(eid)

        try:
            made = bridge.request("put", start=at(7, 0), end=at(7, 15), **event)
            eid = made["entryId"]
            self.assertEqual((made["start"], made["end"]), (at(7, 0), at(7, 15)))
            got = find()
            self.assertEqual((got["taskId"], got["subject"], got["busy"], got["allDay"], got["start"]),
                             (self.TASK_ID, event["subject"], 0, False, at(7, 0)))

            moved = bridge.request("put", entryId=eid, start=at(7, 30), end=at(7, 50), **event)
            self.assertEqual((moved["entryId"], moved["start"], moved["end"]), (eid, at(7, 30), at(7, 50)))
            self.assertEqual((find()["start"], find()["end"]), (at(7, 30), at(7, 50)))

            bridge.request("remove", entryId=eid)
            self.assertIsNone(find(), "a deleted event must not be reported as still in the calendar")
            eid = None
            bridge.request("flush")
            print(f"\n  created, moved and removed a test event on {day}; it's now in Deleted Items")
        finally:
            if eid:
                try:
                    bridge.request("remove", entryId=eid)
                except Exception:
                    pass
            bridge.close()


if __name__ == "__main__":
    unittest.main()
