"""Talking to Todoist (against a fake Todoist server) and reading desktop Outlook (with PowerShell faked)."""
import datetime as dt
import gc
import json
import subprocess
import threading
import unittest
import urllib.parse
import warnings
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest import mock

from planner import app as planner_app
from planner import config, errors, outlook_desktop, sources, todoist


class FakeTodoist(ThreadingHTTPServer):
    """Records each request and answers from `routes`: {(method, path): (status, body or callable)}."""

    def __init__(self):
        super().__init__(("127.0.0.1", 0), _FakeHandler)
        self.requests = []
        self.routes = {}
        threading.Thread(target=self.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True).start()

    @property
    def base(self):
        return f"http://127.0.0.1:{self.server_address[1]}/api/v1"

    def stop(self):
        self.shutdown()
        self.server_close()


class _FakeHandler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _handle(self):
        u = urllib.parse.urlsplit(self.path)
        body = self.rfile.read(int(self.headers.get("Content-Length") or 0))
        req = {"method": self.command, "path": u.path, "query": dict(urllib.parse.parse_qsl(u.query)),
               "headers": dict(self.headers), "json": json.loads(body) if body else None}
        self.server.requests.append(req)
        status, payload = self.server.routes.get((self.command, u.path), (404, {"error": "no route"}))
        if callable(payload):
            payload = payload(req)
        data = payload if isinstance(payload, bytes) else json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    do_GET = do_POST = _handle


class TodoistClientTests(unittest.TestCase):
    def setUp(self):
        self.fake = FakeTodoist()
        self.client = todoist.TodoistClient("tok123", base=self.fake.base)

    def tearDown(self):
        self.fake.stop()

    def test_sends_the_token(self):
        self.fake.routes[("GET", "/api/v1/projects")] = (200, {"results": [{"id": "1"}], "next_cursor": None})
        self.assertEqual(self.client.projects(), [{"id": "1"}])
        self.assertEqual(self.fake.requests[0]["headers"]["Authorization"], "Bearer tok123")

    def test_follows_pages(self):
        def page(req):
            if req["query"].get("cursor") == "c2":
                return {"results": [{"id": "3"}], "next_cursor": None}
            return {"results": [{"id": "1"}, {"id": "2"}], "next_cursor": "c2"}
        self.fake.routes[("GET", "/api/v1/tasks")] = (200, page)
        self.assertEqual([t["id"] for t in self.client.tasks()], ["1", "2", "3"])
        self.assertEqual([r["query"].get("cursor") for r in self.fake.requests], [None, "c2"])
        self.assertTrue(all(r["query"]["limit"] == "200" for r in self.fake.requests))

    def test_filter_uses_the_filter_endpoint(self):
        self.fake.routes[("GET", "/api/v1/tasks/filter")] = (200, {"results": [{"id": "9"}]})
        self.assertEqual(self.client.tasks("#Work & today"), [{"id": "9"}])
        self.assertEqual(self.fake.requests[0]["query"]["query"], "#Work & today")

    def test_tolerates_an_unpaged_list(self):
        self.fake.routes[("GET", "/api/v1/tasks")] = (200, [{"id": "1"}])
        self.assertEqual(self.client.tasks(), [{"id": "1"}])

    def test_task_actions(self):
        for path in ("/api/v1/tasks/6X7rM8997g3RQmvh/close", "/api/v1/tasks/7/reopen", "/api/v1/tasks/7",
                     "/api/v1/tasks"):
            self.fake.routes[("POST", path)] = (204, b"")
        self.client.close("6X7rM8997g3RQmvh")
        self.client.reopen("7")
        self.client.update("7", {"due_date": "2026-11-10"})
        self.client.create({"content": "New"})
        got = [(r["path"], r["json"]) for r in self.fake.requests]
        self.assertEqual(got, [("/api/v1/tasks/6X7rM8997g3RQmvh/close", None), ("/api/v1/tasks/7/reopen", None),
                               ("/api/v1/tasks/7", {"due_date": "2026-11-10"}), ("/api/v1/tasks", {"content": "New"})])
        self.assertEqual(self.fake.requests[2]["headers"]["Content-Type"], "application/json")

    def test_errors_are_explained(self):
        cases = {401: "didn't accept the API token", 403: "didn't accept the API token",
                 404: "couldn't find that task", 429: "rate-limiting", 500: "HTTP 500"}
        for status, words in cases.items():
            self.fake.routes[("GET", "/api/v1/projects")] = (status, {"error": "nope"})
            with self.assertRaises(errors.SourceError) as cm:
                self.client.projects()
            self.assertIn(words, str(cm.exception), status)

    def test_error_responses_are_closed(self):
        self.fake.routes[("GET", "/api/v1/projects")] = (500, {"error": "nope"})
        with warnings.catch_warnings(record=True) as seen:
            warnings.simplefilter("always", ResourceWarning)
            with self.assertRaises(errors.SourceError):
                self.client.projects()
            gc.collect()
        self.assertEqual([w for w in seen if issubclass(w.category, ResourceWarning)], [])

    def test_unreadable_response(self):
        self.fake.routes[("GET", "/api/v1/projects")] = (200, b"<html>gateway</html>")
        with self.assertRaisesRegex(errors.SourceError, "couldn't read"):
            self.client.projects()

    def test_unreachable(self):
        client = todoist.TodoistClient("t", base="http://127.0.0.1:9/api/v1")  # nothing listens on port 9
        with self.assertRaisesRegex(errors.SourceError, "Couldn't reach Todoist"):
            client.projects()


class PlannerAppTodoistTests(unittest.TestCase):
    """PlannerApp with a real token setting, pointed at the fake server."""

    def setUp(self):
        self.fake = FakeTodoist()
        self.fake.routes[("GET", "/api/v1/projects")] = (200, {"results": [{"id": "p", "name": "Work", "color": "red"}]})
        self.fake.routes[("GET", "/api/v1/tasks")] = (200, {"results": [
            {"id": "1", "content": "Open", "project_id": "p"},
            {"id": "2", "content": "Done", "project_id": "p", "checked": True},
            {"id": "3", "content": "Deleted", "project_id": "p", "is_deleted": True}]})
        self.fake.routes[("POST", "/api/v1/tasks/1")] = (200, {})
        # Point every client the app makes at the fake server.
        real = todoist.TodoistClient
        patcher2 = mock.patch.object(sources, "TodoistClient", lambda token: real(token, base=self.fake.base))
        patcher2.start()
        self.addCleanup(patcher2.stop)
        cfg = config.Config(Path("unused.json"), memory_only=True)
        cfg.update({"todoist_token": "tok", "calendar_source": "none"})
        self.app = planner_app.PlannerApp(cfg)

    def tearDown(self):
        self.fake.stop()

    def task_fetches(self):
        return sum(1 for r in self.fake.requests if r["path"] == "/api/v1/tasks")

    def test_only_open_tasks_with_project_names(self):
        tasks = self.app.tasks(force=False)
        self.assertEqual([(t["content"], t["project"]) for t in tasks], [("Open", "Work")])

    def test_tasks_are_cached_until_something_changes(self):
        self.app.tasks(force=False)
        self.app.tasks(force=False)
        self.assertEqual(self.task_fetches(), 1)
        self.app.schedule_task("1", "2026-11-10", "09:00", 30)
        self.app.tasks(force=False)
        self.assertEqual(self.task_fetches(), 2)
        self.app.tasks(force=True)
        self.assertEqual(self.task_fetches(), 3)

    def test_one_failing_source_does_not_hide_the_other(self):
        self.app.cfg.update({"calendar_source": "ics", "ics_url": self.fake.base + "/missing.ics"})
        out = self.app.data(dt.date(2026, 11, 9), dt.date(2026, 11, 16), force=False)
        self.assertEqual(len(out["tasks"]), 1)
        self.assertEqual([e["source"] for e in out["errors"]], ["calendar"])

    def test_no_notice_when_every_zone_is_known(self):
        out = self.app.data(dt.date(2026, 11, 9), dt.date(2026, 11, 16), force=False)
        self.assertEqual(out["errors"], [])

    def test_unknown_zones_from_tasks_and_calendar_are_reported_once(self):
        self.fake.routes[("GET", "/api/v1/tasks")] = (200, {"results": [
            {"id": "1", "content": "A", "project_id": "p", "due": {"datetime": "2026-11-10T10:00:00",
                                                                  "timezone": "Mars/Olympus"}}]})
        cal = ("BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:1\r\nSUMMARY:Hi\r\n"
               "DTSTART;TZID=Venus/Maxwell:20261110T090000\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n").encode()
        self.fake.routes[("GET", "/cal.ics")] = (200, cal)
        self.app.cfg.update({"calendar_source": "ics",
                             "ics_url": f"http://127.0.0.1:{self.fake.server_address[1]}/cal.ics"})
        out = self.app.data(dt.date(2026, 11, 9), dt.date(2026, 11, 16), force=False)
        [notice] = out["errors"]
        self.assertEqual(notice["source"], "timezones")
        self.assertIn("Mars/Olympus, Venus/Maxwell", notice["message"])
        # Switching the calendar off drops its part of the notice.
        self.app.cfg.update({"calendar_source": "none"})
        out = self.app.data(dt.date(2026, 11, 9), dt.date(2026, 11, 16), force=False)
        self.assertIn("(Mars/Olympus)", out["errors"][0]["message"])

    def test_no_token_means_no_tasks_and_a_clear_error_on_actions(self):
        cfg = config.Config(Path("unused.json"), memory_only=True)
        app = planner_app.PlannerApp(cfg)
        self.assertEqual(app.tasks(force=False), [])
        with self.assertRaisesRegex(errors.SourceError, "Add your Todoist API token"):
            app.close_task("1")


class IcsDownloadTests(unittest.TestCase):
    def setUp(self):
        self.fake = FakeTodoist()  # any small server will do
        cfg = config.Config(Path("unused.json"), memory_only=True)
        self.app = planner_app.PlannerApp(cfg)
        self.cfg = cfg

    def tearDown(self):
        self.fake.stop()

    def url(self, path):
        return f"http://127.0.0.1:{self.fake.server_address[1]}{path}"

    def test_downloads_and_caches(self):
        cal = ("BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:1\r\nSUMMARY:Hi\r\nDTSTART:20261110T090000\r\n"
               "END:VEVENT\r\nEND:VCALENDAR\r\n").encode()
        self.fake.routes[("GET", "/cal.ics")] = (200, cal)
        self.cfg.update({"ics_url": self.url("/cal.ics")})
        d0, d1 = dt.date(2026, 11, 10), dt.date(2026, 11, 11)
        self.assertEqual([e["title"] for e in self.app.events(d0, d1, force=False)], ["Hi"])
        self.app.events(d0, d1, force=False)
        self.app.events(d0, d1, force=True)  # a forced refresh within 15 seconds still uses the copy
        self.assertEqual(len(self.fake.requests), 1)

    def test_webcal_links_are_fetched_over_https(self):
        self.cfg.update({"ics_url": "webcal://outlook.office365.com/owa/calendar/x/calendar.ics"})
        with mock.patch.object(sources, "http_request", return_value=(200, b"BEGIN:VCALENDAR\r\nEND:VCALENDAR")) as req:
            self.app.events(dt.date(2026, 11, 10), dt.date(2026, 11, 11), force=False)
        self.assertEqual(req.call_args[0][1], "https://outlook.office365.com/owa/calendar/x/calendar.ics")

    def test_explains_bad_links(self):
        self.fake.routes[("GET", "/html")] = (200, b"<html>calendar.html</html>")
        cases = {self.url("/missing.ics"): "HTTP 404", self.url("/html"): "didn't return a calendar",
                 "not a url": "doesn't look like a web address"}
        for url, words in cases.items():
            self.cfg.update({"ics_url": url})
            with self.assertRaises(errors.SourceError) as cm:
                self.app.events(dt.date(2026, 11, 10), dt.date(2026, 11, 11), force=False)
            self.assertIn(words, str(cm.exception), url)


class SourceSelectionTests(unittest.TestCase):
    def setUp(self):
        self.cfg = config.Config(Path("unused.json"), memory_only=True)
        self.app = planner_app.PlannerApp(self.cfg)

    def test_calendar_follows_the_setting(self):
        for value, kind in (("ics", sources.IcsCalendar), ("outlook_desktop", sources.OutlookDesktopCalendar),
                            ("none", sources.NoCalendar)):
            self.cfg.update({"calendar_source": value})
            self.assertIsInstance(self.app.calendar, kind, value)
        self.assertIsInstance(self.app.task_source, sources.TodoistTasks)

    def test_demo_ignores_the_setting(self):
        app = planner_app.PlannerApp(self.cfg, demo=True)
        self.cfg.update({"calendar_source": "none"})
        self.assertIsInstance(app.calendar, sources.DemoCalendar)
        self.assertIsInstance(app.task_source, sources.DemoTasks)

    def test_changing_settings_forgets_cached_data(self):
        downloads = []

        def fake_get(*args, **kwargs):
            downloads.append(args[1])
            return 200, b"BEGIN:VCALENDAR\r\nEND:VCALENDAR"

        day = dt.date(2026, 11, 10)
        self.cfg.update({"calendar_source": "ics", "ics_url": "https://example.test/cal.ics"})
        with mock.patch.object(sources, "http_request", side_effect=fake_get):
            self.app.events(day, day + dt.timedelta(days=1), force=False)
            self.app.events(day, day + dt.timedelta(days=1), force=False)
            self.app.settings_changed()
            self.app.events(day, day + dt.timedelta(days=1), force=False)
        self.assertEqual(len(downloads), 2)

    def test_source_notices_reach_the_page(self):
        class Noisy(sources.CalendarSource):
            def notices(self):
                return [{"source": "calendar", "message": "Heads up"}]

        self.app._calendars["none"] = Noisy()
        self.cfg.update({"calendar_source": "none"})
        out = self.app.data(dt.date(2026, 11, 9), dt.date(2026, 11, 16), force=False)
        self.assertEqual(out["errors"], [{"source": "calendar", "message": "Heads up"}])

    def test_close_reaches_every_calendar_source(self):
        closed = []
        for name, cal in self.app._calendars.items():
            cal.close = lambda name=name: closed.append(name)
        self.app.close()
        self.assertEqual(sorted(closed), ["ics", "none", "outlook_desktop"])


def fake_powershell(stdout: str, returncode=0, stderr=""):
    return mock.patch.object(outlook_desktop.subprocess, "run", return_value=subprocess.CompletedProcess(
        [], returncode, stdout.encode("utf-8"), stderr.encode("utf-8")))


class OutlookDesktopTests(unittest.TestCase):
    def setUp(self):
        p = mock.patch.object(outlook_desktop.sys, "platform", "win32")
        p.start()
        self.addCleanup(p.stop)

    def read(self):
        return outlook_desktop.outlook_desktop_events(dt.date(2026, 11, 9), dt.date(2026, 11, 16))

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
        with fake_powershell("﻿" + json.dumps(rows)) as run:
            board, leave, demo = self.read()
        script = run.call_args[0][0][-1]
        self.assertIn("2026-11-09", script)
        self.assertIn("2026-11-16", script)
        self.assertEqual((board["start"], board["end"], board["busy"]), ("2026-11-10T10:00", "2026-11-10T11:00", "busy"))
        self.assertEqual(board["url"], "https://teams.microsoft.com/l/meetup-join/1")
        self.assertEqual((leave["allDay"], leave["start"], leave["end"], leave["busy"]),
                         (True, "2026-11-12", "2026-11-13", "oof"))
        self.assertTrue(demo["cancelled"])
        self.assertEqual(demo["busy"], "free")

    def test_explains_failures(self):
        with fake_powershell("", 1, "Retrieving the COM class factory failed\nmore"):
            with self.assertRaisesRegex(errors.SourceError, "classic Outlook.*COM class factory"):
                self.read()
        with mock.patch.object(outlook_desktop.subprocess, "run", side_effect=subprocess.TimeoutExpired("ps", 120)):
            with self.assertRaisesRegex(errors.SourceError, "took too long"):
                self.read()

    def test_only_on_windows(self):
        with mock.patch.object(outlook_desktop.sys, "platform", "darwin"):
            with self.assertRaisesRegex(errors.SourceError, "only works on Windows"):
                self.read()


if __name__ == "__main__":
    unittest.main()
