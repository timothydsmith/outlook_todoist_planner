"""Demo data (used with --demo; also handy for trying the app before setup)."""
from __future__ import annotations

import datetime as dt

def demo_ics(today: dt.date) -> str:
    monday = today - dt.timedelta(days=today.weekday())

    def stamp(day_offset, hh, mm):
        d = monday + dt.timedelta(days=day_offset)
        return f"{d:%Y%m%d}T{hh:02d}{mm:02d}00"

    def day(day_offset):
        return f"{monday + dt.timedelta(days=day_offset):%Y%m%d}"

    ev = []

    def add(lines):
        ev.append("BEGIN:VEVENT\r\n" + "\r\n".join(lines) + "\r\nEND:VEVENT")

    add(["UID:demo-standup", f"DTSTART:{stamp(-14, 9, 0)}", f"DTEND:{stamp(-14, 9, 15)}",
         "RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR", f"EXDATE:{stamp(3, 9, 0)}",
         "SUMMARY:Team check-in", "LOCATION:Microsoft Teams Meeting",
         "DESCRIPTION:Join: https://teams.microsoft.com/l/meetup-join/demo"])
    add(["UID:demo-standup", f"RECURRENCE-ID:{stamp(1, 9, 0)}", f"DTSTART:{stamp(1, 9, 30)}",
         f"DTEND:{stamp(1, 9, 45)}", "SUMMARY:Team check-in (moved)", "LOCATION:Microsoft Teams Meeting"])
    add(["UID:demo-ops", f"DTSTART:{stamp(-7, 10, 0)}", f"DTEND:{stamp(-7, 11, 0)}",
         "RRULE:FREQ=WEEKLY;INTERVAL=1;BYDAY=MO", "SUMMARY:Operations weekly", "LOCATION:Room 3.20"])
    add(["UID:demo-abc", f"DTSTART:{stamp(1, 13, 30)}", f"DTEND:{stamp(1, 15, 30)}",
         "SUMMARY:Leadership meeting", "LOCATION:Boardroom"])
    add(["UID:demo-tent", f"DTSTART:{stamp(1, 14, 0)}", f"DTEND:{stamp(1, 15, 0)}",
         "SUMMARY:Process review working group", "X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE"])
    add(["UID:demo-1on1", f"DTSTART:{stamp(2, 11, 0)}", f"DTEND:{stamp(2, 11, 30)}",
         "SUMMARY:1:1 with Priya", "LOCATION:Microsoft Teams Meeting"])
    add(["UID:demo-lunch", f"DTSTART:{stamp(2, 12, 30)}", f"DTEND:{stamp(2, 13, 30)}",
         "SUMMARY:Lunch", "TRANSP:TRANSPARENT", "X-MICROSOFT-CDO-BUSYSTATUS:FREE"])
    add(["UID:demo-plan", f"DTSTART:{stamp(3, 9, 30)}", f"DTEND:{stamp(3, 12, 0)}",
         "SUMMARY:Planning workshop: 2027 roadmap", "LOCATION:Level 2 meeting room"])
    add(["UID:demo-leave", f"DTSTART;VALUE=DATE:{day(4)}", f"DTEND;VALUE=DATE:{day(5)}",
         "SUMMARY:Working from the other office", "X-MICROSOFT-CDO-BUSYSTATUS:WORKINGELSEWHERE"])
    add(["UID:demo-exec", f"DTSTART:{stamp(4, 15, 0)}", f"DTEND:{stamp(4, 16, 0)}",
         "SUMMARY:Executive briefing prep", "LOCATION:Level 4 office"])
    add(["UID:demo-conf", f"DTSTART;VALUE=DATE:{day(2)}", f"DTEND;VALUE=DATE:{day(4)}",
         "SUMMARY:Industry conference (Wed–Thu)"])
    add(["UID:demo-cancel", f"DTSTART:{stamp(0, 15, 0)}", f"DTEND:{stamp(0, 15, 30)}",
         "SUMMARY:Canceled: Vendor demo"])
    return "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Planner demo//EN\r\n" + "\r\n".join(ev) + "\r\nEND:VCALENDAR\r\n"


def demo_tasks(today: dt.date) -> list:
    monday = today - dt.timedelta(days=today.weekday())

    def d(off):
        return (monday + dt.timedelta(days=off)).isoformat()

    def task(i, content, project, priority=1, due=None, duration=None, recurring=False, description=""):
        t = {"id": f"demo{i}", "content": content, "project_id": project, "priority": priority,
             "description": description, "labels": [], "parent_id": None}
        if due:
            t["due"] = {"date": due, "string": "every weekday" if recurring else "", "is_recurring": recurring}
        if duration:
            t["duration"] = {"amount": duration, "unit": "minute"}
        return t

    yesterday = (today - dt.timedelta(days=1)).isoformat()
    return [
        task(1, "Send leadership meeting papers to the committee", "w", 4, yesterday),
        task(2, "Reply to the team about the schedule change", "w", 3, (today - dt.timedelta(days=3)).isoformat()),
        task(3, "Draft briefing note on the service review", "w", 4, d(0) + "T13:00:00", 90,
             description="Two pages. Lead with the demand trend, then the three options."),
        task(4, "Review steering committee minutes", "w", 2, d(1) + "T11:00:00", 45),
        task(5, "Inbox to zero", "w", 1, d(2) + "T16:30:00", 30, recurring=True),
        task(6, "Book travel for the industry conference", "p", 2, d(1)),
        task(7, "Prep questions for 1:1 with Priya", "w", 1, d(2) + "T10:30:00", 30),
        task(8, "Professional development plan: write the goals section", "p", 3, d(4) + "T08:00:00", 60),
        task(9, "Pick up dry cleaning", "p", 1, d(5)),
        task(10, "Check in on the 2027 roadmap survey responses", "w", 2, d(3)),
        task(11, "Read the industry report on remote work", "w", 1),
        task(12, "Update the risk register", "w", 3),
        task(13, "Book a session with the writing group", "p", 1),
    ]


DEMO_PROJECTS = [{"id": "w", "name": "Work", "color": "blue"}, {"id": "p", "name": "Personal", "color": "green"}]
