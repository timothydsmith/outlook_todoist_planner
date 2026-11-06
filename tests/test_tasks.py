"""Dates, Todoist task conversion, and settings."""
import contextlib
import datetime as dt
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from planner import config, timeutil, todoist
from tests.support import has_zone, local_from_utc


class ParseIsoTests(unittest.TestCase):
    def test_date_only_gives_a_date(self):
        self.assertEqual(timeutil.parse_iso("2026-11-06"), dt.date(2026, 11, 6))

    def test_no_offset_gives_naive_datetime(self):
        v = timeutil.parse_iso("2026-11-06T09:30:00")
        self.assertEqual(v, dt.datetime(2026, 11, 6, 9, 30))
        self.assertIsNone(v.tzinfo)

    def test_z_means_utc(self):
        v = timeutil.parse_iso("2026-11-06T09:30:00Z")
        self.assertEqual(v.utcoffset(), dt.timedelta(0))

    def test_offsets_with_and_without_colon(self):
        for raw in ("2026-11-06T09:30:00+11:00", "2026-11-06T09:30:00+1100"):
            self.assertEqual(timeutil.parse_iso(raw).utcoffset(), dt.timedelta(hours=11), raw)
        self.assertEqual(timeutil.parse_iso("2026-11-06T09:30:00-05:30").utcoffset(),
                         -dt.timedelta(hours=5, minutes=30))

    def test_fractional_seconds_and_missing_seconds(self):
        self.assertEqual(timeutil.parse_iso("2026-11-06T09:30:15.123456Z").second, 15)
        self.assertEqual(timeutil.parse_iso("2026-11-06T09:30"), dt.datetime(2026, 11, 6, 9, 30))

    def test_rejects_rubbish(self):
        for raw in ("", "tomorrow", "06/11/2026", None):
            with self.assertRaises(ValueError):
                timeutil.parse_iso(raw)


class ToLocalNaiveTests(unittest.TestCase):
    def test_naive_is_left_alone(self):
        v = dt.datetime(2026, 1, 1, 8, 0)
        self.assertEqual(timeutil.to_local_naive(v), v)

    def test_aware_becomes_local_wall_time(self):
        v = dt.datetime(2026, 1, 1, 8, 0, tzinfo=dt.timezone.utc)
        out = timeutil.to_local_naive(v)
        self.assertIsNone(out.tzinfo)
        self.assertEqual(out.strftime("%Y-%m-%dT%H:%M"), local_from_utc(2026, 1, 1, 8, 0))


def raw_task(**over):
    t = {"id": "123", "content": "Write report", "description": "", "priority": 1,
         "project_id": "p1", "labels": [], "parent_id": None, "due": None, "duration": None}
    t.update(over)
    return t


PROJECTS = {"p1": {"id": "p1", "name": "Work", "color": "blue"}}


class NormalizeTaskTests(unittest.TestCase):
    def test_basic_fields(self):
        t = todoist.normalize_task(raw_task(priority=4, labels=["deep"]), PROJECTS)
        self.assertEqual(t["id"], "123")
        self.assertEqual(t["content"], "Write report")
        self.assertEqual(t["priority"], 4)
        self.assertEqual(t["project"], "Work")
        self.assertEqual(t["projectColor"], todoist.TODOIST_COLORS["blue"])
        self.assertEqual(t["labels"], ["deep"])
        self.assertIsNone(t["due"])
        self.assertIsNone(t["duration"])
        self.assertFalse(t["isSubtask"])
        self.assertEqual(t["url"], "https://app.todoist.com/app/task/123")

    def test_unknown_project_gets_grey(self):
        t = todoist.normalize_task(raw_task(project_id="nope"), PROJECTS)
        self.assertEqual(t["project"], "")
        self.assertEqual(t["projectColor"], "#9aa3ad")

    def test_subtask_and_long_description(self):
        t = todoist.normalize_task(raw_task(parent_id="9", description="x" * 5000), PROJECTS)
        self.assertTrue(t["isSubtask"])
        self.assertEqual(len(t["description"]), 1500)

    def test_date_only_due(self):
        t = todoist.normalize_task(raw_task(due={"date": "2026-11-10", "string": "10 Nov"}), PROJECTS)
        self.assertEqual(t["due"], {"date": "2026-11-10", "time": None, "recurring": False, "string": "10 Nov"})

    def test_floating_datetime_is_taken_as_local(self):
        due = {"date": "2026-11-10T09:30:00", "is_recurring": True, "string": "every tue 9:30"}
        t = todoist.normalize_task(raw_task(due=due), PROJECTS)
        self.assertEqual((t["due"]["date"], t["due"]["time"]), ("2026-11-10", "09:30"))
        self.assertTrue(t["due"]["recurring"])

    def test_utc_datetime_is_converted_to_local(self):
        due = {"date": "2026-11-09", "datetime": "2026-11-09T23:00:00Z", "timezone": "Australia/Melbourne"}
        t = todoist.normalize_task(raw_task(due=due), PROJECTS)
        self.assertEqual(f'{t["due"]["date"]}T{t["due"]["time"]}', local_from_utc(2026, 11, 9, 23, 0))

    @unittest.skipUnless(has_zone("Australia/Melbourne"), "needs a time-zone database (pip install tzdata)")
    def test_naive_datetime_with_named_zone(self):
        due = {"datetime": "2026-11-10T10:00:00", "timezone": "Australia/Melbourne"}  # 10:00 AEDT = 23:00Z
        t = todoist.normalize_task(raw_task(due=due), PROJECTS)
        self.assertEqual(f'{t["due"]["date"]}T{t["due"]["time"]}', local_from_utc(2026, 11, 9, 23, 0))

    def test_unknown_zone_is_reported(self):
        zones = set()
        t = todoist.normalize_task(raw_task(due={"datetime": "2026-11-10T10:00:00", "timezone": "Mars/Olympus"}),
                               PROJECTS, zones)
        self.assertEqual((t["due"]["date"], t["due"]["time"]), ("2026-11-10", "10:00"))  # shown as local
        self.assertEqual(zones, {"Mars/Olympus"})

    def test_zone_is_irrelevant_when_the_time_is_in_utc(self):
        zones = set()
        todoist.normalize_task(raw_task(due={"datetime": "2026-11-09T23:00:00Z", "timezone": "Mars/Olympus"}),
                           PROJECTS, zones)
        self.assertEqual(zones, set())

    def test_unreadable_due_is_dropped(self):
        t = todoist.normalize_task(raw_task(due={"date": "someday"}), PROJECTS)
        self.assertIsNone(t["due"])

    def test_durations(self):
        self.assertEqual(todoist.normalize_task(raw_task(duration={"amount": 45, "unit": "minute"}), PROJECTS)["duration"], 45)
        self.assertEqual(todoist.normalize_task(raw_task(duration={"amount": 2, "unit": "day"}), PROJECTS)["duration"], 2880)
        self.assertIsNone(todoist.normalize_task(raw_task(duration={"amount": "lots"}), PROJECTS)["duration"])


class SchedulePayloadTests(unittest.TestCase):
    def test_date_only(self):
        self.assertEqual(todoist.schedule_payload("2026-11-10", None, 30), {"due_date": "2026-11-10"})

    def test_time_without_duration(self):
        p = todoist.schedule_payload("2026-11-10", "09:00", None)
        self.assertEqual(set(p), {"due_datetime"})
        self.assertRegex(p["due_datetime"], r"^\d{4}-\d\d-\d\dT\d\d:\d\d:00Z$")

    def test_time_with_duration(self):
        p = todoist.schedule_payload("2026-11-10", "09:00", 45)
        self.assertEqual((p["duration"], p["duration_unit"]), (45, "minute"))

    def test_bad_input(self):
        with self.assertRaises(ValueError):
            todoist.schedule_payload("10/11/2026", None, None)
        with self.assertRaises(ValueError):
            todoist.schedule_payload("2026-11-10", "nine", None)

    def test_round_trip_through_todoist_shape(self):
        """Whatever time you drop a task at is the time it shows at after Todoist sends it back."""
        for date, time in [("2026-11-10", "09:00"), ("2026-11-10", "00:00"), ("2026-12-31", "23:45"),
                           ("2027-04-04", "12:15"), ("2026-07-01", "17:30")]:
            p = todoist.schedule_payload(date, time, 30)
            back = todoist.normalize_task(raw_task(due={"date": date, "datetime": p["due_datetime"]}), PROJECTS)
            self.assertEqual((back["due"]["date"], back["due"]["time"]), (date, time), p)


class TimezoneMessageTests(unittest.TestCase):
    def test_without_a_time_zone_database_it_says_how_to_fix_it(self):
        with mock.patch.object(timeutil, "has_tz_database", return_value=False), \
             mock.patch.object(timeutil.sys, "platform", "win32"):
            msg = timeutil.timezone_message(["Australia/Perth"])
        self.assertIn("(Australia/Perth)", msg)
        self.assertIn("py -m pip install tzdata", msg)

    def test_with_a_database_the_zone_is_just_unknown(self):
        with mock.patch.object(timeutil, "has_tz_database", return_value=True):
            msg = timeutil.timezone_message(["A", "B", "C", "D"])
        self.assertIn("(A, B, C and others)", msg)
        self.assertNotIn("tzdata", msg)


class ConfigTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.path = Path(self.dir.name) / "sub" / "settings.json"

    def tearDown(self):
        self.dir.cleanup()

    def test_defaults_when_no_file(self):
        cfg = config.Config(self.path)
        self.assertEqual(cfg.snapshot(), config.DEFAULTS)

    def test_update_saves_and_reloads(self):
        config.Config(self.path).update({"todoist_token": "  abcdefgh1234  ", "ics_url": " https://x/cal.ics "})
        again = config.Config(self.path)
        self.assertEqual(again.get("todoist_token"), "abcdefgh1234")
        self.assertEqual(again.get("ics_url"), "https://x/cal.ics")
        self.assertFalse(self.path.with_suffix(".tmp").exists())

    def test_values_are_clamped_and_unknown_keys_dropped(self):
        cfg = config.Config(self.path)
        cfg.update({"day_start_hour": 30, "default_task_minutes": 1, "refresh_minutes": 999,
                    "week_start": 5, "calendar_source": "gmail", "show_weekends": 0, "evil": "x"})
        s = cfg.snapshot()
        self.assertEqual((s["day_start_hour"], s["default_task_minutes"], s["refresh_minutes"]), (12, 5, 60))
        self.assertEqual((s["week_start"], s["calendar_source"], s["show_weekends"]), (1, "ics", False))
        self.assertNotIn("evil", json.loads(self.path.read_text()))

    def test_blank_token_keeps_old_one_and_clear_removes_it(self):
        cfg = config.Config(self.path)
        cfg.update({"todoist_token": "secret-token-1"})
        cfg.update({"todoist_token": ""})
        self.assertEqual(cfg.get("todoist_token"), "secret-token-1")
        cfg.update({"todoist_token": "__clear__"})
        self.assertEqual(cfg.get("todoist_token"), "")

    def test_public_never_reveals_the_token(self):
        cfg = config.Config(self.path)
        cfg.update({"todoist_token": "secret-token-WXYZ"})
        pub = cfg.public()
        self.assertNotIn("todoist_token", pub)
        self.assertTrue(pub["has_token"])
        self.assertEqual(pub["token_hint"], "…WXYZ")
        self.assertNotIn("secret", json.dumps(pub))

    def test_corrupt_file_falls_back_to_defaults(self):
        self.path.parent.mkdir(parents=True)
        self.path.write_text("{not json", encoding="utf-8")
        with contextlib.redirect_stdout(io.StringIO()) as said:
            cfg = config.Config(self.path)
        self.assertEqual(cfg.snapshot(), config.DEFAULTS)
        self.assertIn("Starting with defaults", said.getvalue())

    def test_memory_only_writes_nothing(self):
        cfg = config.Config(self.path, memory_only=True)
        cfg.update({"ics_url": "https://x"})
        self.assertEqual(cfg.get("ics_url"), "https://x")
        self.assertFalse(self.path.exists())


if __name__ == "__main__":
    unittest.main()
