"""Stands in for planner/outlook/helper.ps1 in tests. Speaks the same one-JSON-line-per-request protocol.

    python fake_helper.py MODE LOGFILE

Each line of LOGFILE records "start" or a request's op. MODE:
  ok            behaves like a healthy helper with Outlook connected
  disconnected  answers events with connected=false
  error         answers events with ok=false (as when classic Outlook isn't installed)
  crash         exits on every events request
  crash_first   exits on the first events request only (needs LOGFILE to remember)
  hang          never answers events
  noise         prints stray output and a stale reply before each real reply
"""
import json
import sys
import time

mode, logfile = sys.argv[1], sys.argv[2]


def log(text):
    with open(logfile, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def send(obj):
    sys.stdout.write(json.dumps(obj) + "\n")
    sys.stdout.flush()


with open(logfile, "a", encoding="utf-8"):
    pass
starts = sum(1 for ln in open(logfile, encoding="utf-8") if ln.strip() == "start")
log("start")

for line in sys.stdin:
    req = json.loads(line)
    op, rid = req["op"], req["id"]
    log(op)
    if mode == "noise":
        sys.stdout.write("WARNING: something PowerShell-ish\n")
        send({"id": rid - 1, "ok": True, "stale": True})
    if op == "hello":
        send({"id": rid, "ok": True, "version": 1, "connection": 700, "startedOutlook": False})
    elif op == "quit":
        send({"id": rid, "ok": True})
        break
    elif op == "events":
        if mode == "crash" or (mode == "crash_first" and starts == 0):
            sys.stderr.write("boom: Outlook went away\n")
            sys.stderr.flush()
            sys.exit(3)
        if mode == "hang":
            time.sleep(60)
            continue
        if mode == "error":
            send({"id": rid, "ok": False, "error": "Class not registered (0x80040154)"})
            continue
        row = {"id": "E1|x", "title": "Board", "start": req["start"] + "T10:00:00",
               "end": req["start"] + "T11:00:00", "allDay": False, "location": "", "busy": 2}
        send({"id": rid, "ok": True, "connected": mode != "disconnected",
              "connection": 400 if mode == "disconnected" else 700, "rows": [row]})
    else:
        send({"id": rid, "ok": False, "error": f"Unknown request '{op}'."})
