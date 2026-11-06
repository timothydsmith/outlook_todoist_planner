"""The local web server, run in demo mode: security checks, the data endpoint, and task actions."""
import datetime as dt
import http.client
import json
import threading
import unittest

import outlook_todoist_planner as otp

TOKEN = "test-token-abc"


class ServerTests(unittest.TestCase):
    def setUp(self):
        cfg = otp.Config(otp.Path("unused.json"), memory_only=True)
        self.server = otp.PlannerServer(("127.0.0.1", 0), otp.PlannerApp(cfg, demo=True), TOKEN, keep_running=True)
        self.port = self.server.server_address[1]
        threading.Thread(target=self.server.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True).start()
        today = dt.date.today()
        self.monday = today - dt.timedelta(days=today.weekday())

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()

    def call(self, method, path, body=None, token=TOKEN, host=None):
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=10)
        headers = {"Host": host or f"127.0.0.1:{self.port}"}
        if token:
            headers["X-Planner-Token"] = token
        data = None
        if body is not None:
            data = json.dumps(body).encode()
            headers["Content-Type"] = "application/json"
        conn.request(method, path, body=data, headers=headers)
        resp = conn.getresponse()
        raw = resp.read()
        conn.close()
        ctype = resp.getheader("Content-Type", "")
        return resp.status, (json.loads(raw) if ctype.startswith("application/json") else raw.decode())

    def week(self):
        start, end = self.monday, self.monday + dt.timedelta(days=7)
        status, data = self.call("GET", f"/api/data?start={start}&end={end}")
        self.assertEqual(status, 200)
        return data

    def task(self, data, tid):
        return next((t for t in data["tasks"] if t["id"] == tid), None)

    # -- security ----------------------------------------------------------------
    def test_page_gets_the_token_written_in(self):
        status, html = self.call("GET", "/", token=None)
        self.assertEqual(status, 200)
        self.assertIn(f'const TOKEN = "{TOKEN}"', html)
        self.assertNotIn("__PLANNER_TOKEN__", html)

    def test_localhost_name_is_allowed(self):
        self.assertEqual(self.call("GET", "/health", token=None, host=f"localhost:{self.port}")[0], 200)

    def test_other_host_names_are_refused(self):  # blocks DNS-rebinding from a web page
        for method, path in (("GET", "/"), ("GET", "/api/data"), ("POST", "/api/task/close")):
            self.assertEqual(self.call(method, path, body={}, host="evil.test")[0], 403, path)

    def test_api_needs_the_token(self):
        path = f"/api/data?start={self.monday}&end={self.monday + dt.timedelta(days=7)}"
        self.assertEqual(self.call("GET", path, token=None)[0], 404)
        self.assertEqual(self.call("GET", path, token="wrong")[0], 404)
        self.assertEqual(self.call("GET", path + f"&t={TOKEN}", token=None)[0], 200)
        self.assertEqual(self.call("POST", "/api/task/close", body={"id": "demo1"}, token=None)[0], 403)

    def test_settings_never_include_the_token(self):
        status, s = self.call("GET", "/api/settings")
        self.assertEqual(status, 200)
        self.assertTrue(s["demo"])
        self.assertNotIn("todoist_token", s)

    def test_settings_cannot_change_in_demo(self):
        status, body = self.call("POST", "/api/settings", body={"ics_url": "https://x"})
        self.assertEqual(status, 400)
        self.assertIn("demo mode", body["error"])

    # -- data --------------------------------------------------------------------
    def test_week_data(self):
        data = self.week()
        self.assertTrue(data["demo"])
        self.assertEqual(data["errors"], [])
        self.assertEqual(len(data["tasks"]), 13)
        self.assertTrue(any(e["title"] == "Leadership meeting" for e in data["events"]))

    def test_bad_date_ranges(self):
        for q in ("start=2026-11-09&end=2026-11-09", "start=2026-11-09&end=2027-01-20", "start=x&end=y", ""):
            self.assertEqual(self.call("GET", f"/api/data?{q}")[0], 400, q)

    # -- task actions ------------------------------------------------------------
    def test_complete_then_undo(self):
        self.assertEqual(self.call("POST", "/api/task/close", body={"id": "demo1"})[0], 200)
        self.assertIsNone(self.task(self.week(), "demo1"))
        self.assertEqual(self.call("POST", "/api/task/reopen", body={"id": "demo1"})[0], 200)
        self.assertIsNotNone(self.task(self.week(), "demo1"))

    def test_schedule_an_undated_task(self):
        day = str(self.monday + dt.timedelta(days=2))
        status, _ = self.call("POST", "/api/task/schedule",
                              body={"id": "demo11", "date": day, "time": "14:15", "duration": 45})
        self.assertEqual(status, 200)
        t = self.task(self.week(), "demo11")
        self.assertEqual((t["due"]["date"], t["due"]["time"], t["duration"]), (day, "14:15", 45))

    def test_move_a_task_to_all_day(self):
        day = str(self.monday + dt.timedelta(days=4))
        self.call("POST", "/api/task/schedule", body={"id": "demo3", "date": day})
        t = self.task(self.week(), "demo3")
        self.assertEqual((t["due"]["date"], t["due"]["time"]), (day, None))

    def test_create_a_task(self):
        day = str(self.monday + dt.timedelta(days=1))
        status, _ = self.call("POST", "/api/task/create",
                              body={"content": "  Call Sam  ", "date": day, "time": "10:00", "duration": 15})
        self.assertEqual(status, 200)
        [t] = [t for t in self.week()["tasks"] if t["content"] == "Call Sam"]
        self.assertEqual((t["due"]["date"], t["due"]["time"], t["duration"]), (day, "10:00", 15))

    def test_action_errors(self):
        status, body = self.call("POST", "/api/task/schedule", body={"id": "nope", "date": str(self.monday)})
        self.assertEqual((status, "no longer exists" in body["error"]), (502, True))
        self.assertEqual(self.call("POST", "/api/task/close", body={})[0], 400)
        self.assertEqual(self.call("POST", "/api/task/create", body={"content": "   "})[0], 502)
        self.assertEqual(self.call("POST", "/api/task/schedule", body={"id": "demo3", "date": "soon"})[0], 400)
        self.assertEqual(self.call("POST", "/api/nothing", body={})[0], 404)

    # -- window lifecycle --------------------------------------------------------
    def test_bye_only_counts_without_keep_running(self):
        self.call("POST", "/api/bye", body={})
        self.assertIsNone(self.server.bye_at)
        self.server.keep_running = False
        self.call("POST", "/api/bye", body={})
        self.assertIsNotNone(self.server.bye_at)
        self.call("POST", "/api/ping", body={})  # the window came back (e.g. a reload)
        self.assertIsNone(self.server.bye_at)


if __name__ == "__main__":
    unittest.main()
