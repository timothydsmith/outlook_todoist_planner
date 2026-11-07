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


if __name__ == "__main__":
    unittest.main()
