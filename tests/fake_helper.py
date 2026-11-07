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
  disconnected also answers tagged with connected=false

Planner's own events (tagged/get/put/remove/flush) live in LOGFILE + ".cal.json", so tests can
read it, and edit it to play the part of someone changing the event in Outlook:
  {"next": 1, "items": {entryId: {taskId, subject, body, start, end, allDay, busy, deleted}}}
An item with no taskId is an ordinary (untagged) event.
"""
import json
import os
import sys
import time

mode, logfile = sys.argv[1], sys.argv[2]
calfile = logfile + ".cal.json"


def load_cal():
    if not os.path.exists(calfile):
        return {"next": 1, "items": {}}
    with open(calfile, encoding="utf-8") as f:
        return json.load(f)


def save_cal(cal):
    with open(calfile, "w", encoding="utf-8") as f:
        json.dump(cal, f, indent=1)


def row(eid, it):
    return {"entryId": eid, "taskId": it.get("taskId"), "subject": it.get("subject", ""), "start": it["start"],
            "end": it["end"], "allDay": bool(it.get("allDay")), "busy": it.get("busy", 2),
            "deleted": bool(it.get("deleted"))}


def calendar_op(req):
    op = req["op"]
    cal = load_cal()
    items = cal["items"]
    if op == "tagged":
        s, e = req["start"], req["end"]
        found = [row(k, v) for k, v in items.items()
                 if v.get("taskId") and not v.get("deleted") and v["start"] < e and v["end"] > s]
        return {"ok": True, "connected": mode != "disconnected", "connection": 400 if mode == "disconnected" else 700,
                "items": found}
    if op == "get":
        out = []
        for eid in req["ids"]:
            it = items.get(eid)
            if it is None:
                out.append({"entryId": eid, "missing": True})
            elif not it.get("taskId"):
                out.append({"entryId": eid, "tagged": False})
            else:
                out.append(row(eid, it))
        return {"ok": True, "items": out}
    if op == "put":
        if req["end"] <= req["start"]:
            return {"ok": False, "error": "The end time must be after the start time."}
        eid = req.get("entryId")
        if eid:
            it = items.get(eid)
            if it is None:
                return {"ok": False, "error": "That Outlook event no longer exists."}
            if it.get("taskId") != req["taskId"]:
                return {"ok": False, "error": "That Outlook event wasn't made by Planner for this task."}
        else:
            eid = f"F{cal['next']}"
            cal["next"] += 1
            it = items[eid] = {"taskId": req["taskId"]}
        it.update(subject=req["subject"], body=req["body"], start=req["start"], end=req["end"],
                  busy=req["busy"], allDay=False)
        save_cal(cal)
        return {"ok": True, "entryId": eid, "start": it["start"], "end": it["end"]}
    if op == "remove":
        it = items.get(req["entryId"])
        if it is not None:
            if not it.get("taskId"):
                return {"ok": False, "error": "That Outlook event wasn't made by Planner."}
            it["deleted"] = True
            save_cal(cal)
        return {"ok": True}
    return {"ok": True}  # flush


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
    elif op in ("tagged", "get", "put", "remove", "flush"):
        if mode == "error":
            send({"id": rid, "ok": False, "error": "Class not registered (0x80040154)"})
            continue
        send(dict(calendar_op(req), id=rid))
    else:
        send({"id": rid, "ok": False, "error": f"Unknown request '{op}'."})
