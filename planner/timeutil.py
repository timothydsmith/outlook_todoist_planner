"""Dates, local time, and named time zones."""
from __future__ import annotations

import datetime as dt
import re
import sys

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:  # pragma: no cover
    ZoneInfo = None


_ISO_RE = re.compile(
    r"^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?\s*(Z|[+-]\d{2}:?\d{2})?$")


def parse_iso(raw: str):
    """'2026-11-06' -> date; '2026-11-06T09:00:00' -> naive; with Z/offset -> aware."""
    m = _ISO_RE.match((raw or "").strip())
    if not m:
        raise ValueError(f"not an ISO date: {raw!r}")
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if m.group(4) is None:
        return dt.date(y, mo, d)
    val = dt.datetime(y, mo, d, int(m.group(4)), int(m.group(5)), int(m.group(6) or 0))
    tz = m.group(7)
    if tz:
        if tz == "Z":
            val = val.replace(tzinfo=dt.timezone.utc)
        else:
            sign = -1 if tz[0] == "-" else 1
            digits = tz[1:].replace(":", "")
            off = dt.timedelta(hours=int(digits[:2]), minutes=int(digits[2:4]))
            val = val.replace(tzinfo=dt.timezone(sign * off))
    return val


def to_local_naive(value: dt.datetime) -> dt.datetime:
    """Aware -> this computer's local wall time (naive). Naive values are already local."""
    if value.tzinfo is None:
        return value
    return value.astimezone().replace(tzinfo=None)


def zone_by_name(name: str):
    """A named time zone ('Australia/Melbourne'), or None if unknown.

    Python on Windows has no time-zone database unless the 'tzdata' package is installed,
    so there every name is unknown.
    """
    if ZoneInfo is None or not name:
        return None
    try:
        return ZoneInfo(name)
    except Exception:
        return None


def has_tz_database() -> bool:
    return zone_by_name("UTC") is not None


def timezone_message(names: list) -> str:
    shown = ", ".join(names[:3]) + (" and others" if len(names) > 3 else "")
    if has_tz_database():
        return (f"Some items use time zones Planner doesn't recognise ({shown}), "
                "so they're shown as if they were in your own time zone.")
    cmd = "py -m pip install tzdata" if sys.platform == "win32" else "python3 -m pip install tzdata"
    return (f"Some items use time zones this computer's Python doesn't know ({shown}), so they're shown "
            f"as if they were in your own time zone. To fix it, close Planner, run {cmd}, then start it again.")
