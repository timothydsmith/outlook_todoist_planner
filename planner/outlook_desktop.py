"""Outlook desktop (Windows, classic Outlook) via PowerShell + COM. Read-only."""
from __future__ import annotations

import datetime as dt
import json
import subprocess
import sys

from .errors import SourceError
from .ics import _MEETING_URL

_PS_SCRIPT = (
    "$ErrorActionPreference='Stop';"
    "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
    "$ic=[Globalization.CultureInfo]::InvariantCulture;"
    "$s=[datetime]::ParseExact('@START@','yyyy-MM-dd',$ic);"
    "$e=[datetime]::ParseExact('@END@','yyyy-MM-dd',$ic);"
    "$ol=New-Object -ComObject Outlook.Application;"
    "$items=$ol.GetNamespace('MAPI').GetDefaultFolder(9).Items;"
    "$items.Sort('[Start]');"
    "$items.IncludeRecurrences=$true;"
    "$f='[Start] < '''+$e.ToString('g')+''' AND [End] > '''+$s.ToString('g')+'''';"
    "try{$res=$items.Restrict($f)}catch{$res=$items};"
    "$out=New-Object System.Collections.ArrayList;"
    "foreach($a in $res){"
    "if($a.Start -ge $e){break};"
    "if($a.End -le $s){continue};"
    "[void]$out.Add([pscustomobject]@{id=[string]$a.EntryID+'|'+$a.Start.ToString('s');"
    "title=[string]$a.Subject;start=$a.Start.ToString('s');end=$a.End.ToString('s');"
    "allDay=[bool]$a.AllDayEvent;location=[string]$a.Location;busy=[int]$a.BusyStatus});"
    "if($out.Count -ge 3000){break}"
    "};"
    "ConvertTo-Json -InputObject @($out) -Compress"
)

_OUTLOOK_BUSY = {0: "free", 1: "tentative", 2: "busy", 3: "oof", 4: "elsewhere"}


def outlook_desktop_events(start_date: dt.date, end_date: dt.date) -> list:
    if sys.platform != "win32":
        raise SourceError("Reading from the Outlook desktop app only works on Windows. "
                          "Use the published calendar (ICS) link instead.")
    script = _PS_SCRIPT.replace("@START@", start_date.isoformat()).replace("@END@", end_date.isoformat())
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, timeout=120, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except FileNotFoundError:
        raise SourceError("Couldn't find PowerShell on this PC.")
    except subprocess.TimeoutExpired:
        raise SourceError("Outlook took too long to answer. Make sure Outlook is open and not showing a dialog.")
    text = proc.stdout.decode("utf-8", "replace").lstrip("﻿").strip()
    if proc.returncode != 0 or not text.startswith("["):
        err = proc.stderr.decode("utf-8", "replace").strip().splitlines()
        detail = next((ln for ln in err if ln.strip()), "no details")
        raise SourceError("Couldn't read the Outlook desktop app. This needs classic Outlook for Windows "
                          f"(the new Outlook can't be read this way). Details: {detail[:240]}")
    rows = json.loads(text or "[]")
    out = []
    for r in rows:
        try:
            s = dt.datetime.fromisoformat(r["start"])
            e = dt.datetime.fromisoformat(r["end"])
        except (KeyError, ValueError, TypeError):
            continue
        busy = _OUTLOOK_BUSY.get(r.get("busy"), "busy")
        title = r.get("title") or "(No title)"
        loc = r.get("location") or ""
        m = _MEETING_URL.search(loc)
        base = {"id": r.get("id") or f"{title}|{r['start']}", "title": title, "location": loc, "busy": busy,
                "cancelled": title.lower().startswith(("canceled:", "cancelled:")), "url": m.group(0) if m else ""}
        if r.get("allDay"):
            sd, ed = s.date(), e.date()
            if ed <= sd:
                ed = sd + dt.timedelta(days=1)
            base.update(allDay=True, start=sd.isoformat(), end=ed.isoformat())
        else:
            base.update(allDay=False, start=s.strftime("%Y-%m-%dT%H:%M"), end=e.strftime("%Y-%m-%dT%H:%M"))
        out.append(base)
    return out
