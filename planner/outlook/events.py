"""Turn appointment rows from the Outlook helper into the event shape the page uses."""
from __future__ import annotations

import datetime as dt

from ..categories import category_color, split_categories
from ..ics import _MEETING_URL

_OUTLOOK_BUSY = {0: "free", 1: "tentative", 2: "busy", 3: "oof", 4: "elsewhere"}


def rows_to_events(rows: list, category_colors: dict | None = None) -> list:
    """`category_colors` is Outlook's master category list, name -> preset colour number."""
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
                "cancelled": title.lower().startswith(("canceled:", "cancelled:")), "url": m.group(0) if m else "",
                "taskId": r.get("taskId") or None}  # set on events Planner made for a task
        cats = split_categories(r.get("categories"))
        base.update(categories=cats, color=category_color(cats, category_colors))
        if r.get("allDay"):
            sd, ed = s.date(), e.date()
            if ed <= sd:
                ed = sd + dt.timedelta(days=1)
            base.update(allDay=True, start=sd.isoformat(), end=ed.isoformat())
        else:
            base.update(allDay=False, start=s.strftime("%Y-%m-%dT%H:%M"), end=e.strftime("%Y-%m-%dT%H:%M"))
        out.append(base)
    return out
