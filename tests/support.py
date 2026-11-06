"""Shared helpers for the tests."""
import datetime as dt

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None


def local_from_utc(y, mo, d, hh=0, mm=0) -> str:
    """The local wall time ('YYYY-MM-DDTHH:MM') of a UTC instant, on whatever machine runs the tests."""
    utc = dt.datetime(y, mo, d, hh, mm, tzinfo=dt.timezone.utc)
    return utc.astimezone().replace(tzinfo=None).strftime("%Y-%m-%dT%H:%M")


def has_zone(name: str) -> bool:
    """True when Python has a time-zone database (on Windows that needs the 'tzdata' package)."""
    if ZoneInfo is None:
        return False
    try:
        ZoneInfo(name)
        return True
    except Exception:
        return False


def ics(*events: str, timezones: str = "") -> str:
    """Wrap VEVENT bodies (one property per line, \\n-separated) into a calendar file."""
    parts = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//tests//EN"]
    if timezones:
        parts.append(timezones.strip())
    for body in events:
        parts += ["BEGIN:VEVENT", body.strip(), "END:VEVENT"]
    parts.append("END:VCALENDAR")
    return "\r\n".join(line.strip() for part in parts for line in part.splitlines()) + "\r\n"


# The VTIMEZONE block Outlook writes for Melbourne/Sydney (copied in shape from a real feed).
AUS_EASTERN = """
BEGIN:VTIMEZONE
TZID:AUS Eastern Standard Time
BEGIN:STANDARD
DTSTART:16010101T030000
TZOFFSETFROM:+1100
TZOFFSETTO:+1000
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=1SU;BYMONTH=4
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:16010101T020000
TZOFFSETFROM:+1000
TZOFFSETTO:+1100
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=1SU;BYMONTH=10
END:DAYLIGHT
END:VTIMEZONE
"""
