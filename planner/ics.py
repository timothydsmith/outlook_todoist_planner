"""iCalendar (.ics) parsing with time zones and recurrence."""
from __future__ import annotations

import calendar as _calendar
import datetime as dt
import re

from .categories import split_categories
from .timeutil import to_local_naive, zone_by_name


WEEKDAY_CODES = {"MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4, "SA": 5, "SU": 6}

# Fallback only: Outlook feeds normally carry their own VTIMEZONE definitions.
WINDOWS_TZ = {
    "AUS Eastern Standard Time": "Australia/Sydney", "E. Australia Standard Time": "Australia/Brisbane",
    "Cen. Australia Standard Time": "Australia/Adelaide", "AUS Central Standard Time": "Australia/Darwin",
    "W. Australia Standard Time": "Australia/Perth", "Tasmania Standard Time": "Australia/Hobart",
    "New Zealand Standard Time": "Pacific/Auckland", "UTC": "UTC", "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik", "W. Europe Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris", "Central Europe Standard Time": "Europe/Budapest",
    "Eastern Standard Time": "America/New_York", "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver", "Pacific Standard Time": "America/Los_Angeles",
    "Singapore Standard Time": "Asia/Singapore", "China Standard Time": "Asia/Shanghai",
    "Tokyo Standard Time": "Asia/Tokyo", "India Standard Time": "Asia/Kolkata",
    "Korea Standard Time": "Asia/Seoul", "SE Asia Standard Time": "Asia/Bangkok",
}


def unfold_lines(text: str) -> list:
    out: list = []
    for line in re.split(r"\r\n|\n|\r", text):
        if line[:1] in (" ", "\t") and out:
            out[-1] += line[1:]
        elif line:
            out.append(line)
    return out


def _split_unquoted(s: str, sep: str) -> list:
    parts, buf, in_q = [], [], False
    for ch in s:
        if ch == '"':
            in_q = not in_q
            buf.append(ch)
        elif ch == sep and not in_q:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf))
    return parts


def parse_content_line(line: str):
    in_q = False
    colon = -1
    for i, ch in enumerate(line):
        if ch == '"':
            in_q = not in_q
        elif ch == ":" and not in_q:
            colon = i
            break
    if colon <= 0:
        return None
    head, value = line[:colon], line[colon + 1:]
    parts = _split_unquoted(head, ";")
    name = parts[0].strip().upper()
    if "." in name:  # strip vCard-style group prefixes
        name = name.rsplit(".", 1)[1]
    params = {}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            params[k.strip().upper()] = v.strip().strip('"')
    return name, params, value


class Component:
    __slots__ = ("name", "props", "children")

    def __init__(self, name: str):
        self.name = name
        self.props: list = []
        self.children: list = []

    def first(self, key: str):
        for n, p, v in self.props:
            if n == key:
                return p, v
        return None

    def every(self, key: str) -> list:
        return [(p, v) for n, p, v in self.props if n == key]

    def walk(self, name: str):
        for c in self.children:
            if c.name == name:
                yield c
            yield from c.walk(name)


def parse_components(text: str) -> Component:
    root = Component("ROOT")
    stack = [root]
    for line in unfold_lines(text):
        cl = parse_content_line(line)
        if not cl:
            continue
        n, p, v = cl
        if n == "BEGIN":
            c = Component(v.strip().upper())
            stack[-1].children.append(c)
            stack.append(c)
        elif n == "END":
            if len(stack) > 1:
                stack.pop()
        else:
            stack[-1].props.append(cl)
    return root


def unescape_text(v: str) -> str:
    return re.sub(r"\\([\\;,nN])", lambda m: "\n" if m.group(1) in "nN" else m.group(1), v)


_DT_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$")


def parse_dt_value(value: str, params: dict, tzres=None):
    """ICS DATE / DATE-TIME -> date | aware datetime | naive (floating) datetime."""
    m = _DT_RE.match(value.strip())
    if not m:
        raise ValueError(f"bad ICS date {value!r}")
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if m.group(4) is None:
        return dt.date(y, mo, d)
    naive = dt.datetime(y, mo, d, int(m.group(4)), int(m.group(5)), int(m.group(6) or 0))
    if m.group(7):
        return naive.replace(tzinfo=dt.timezone.utc)
    tzid = params.get("TZID")
    if tzid and tzres is not None:
        tz = tzres(tzid)
        if tz is not None:
            return naive.replace(tzinfo=tz)
    return naive


def parse_offset(s: str) -> dt.timedelta:
    s = s.strip()
    sign = -1 if s.startswith("-") else 1
    digits = s.lstrip("+-")
    h, m = int(digits[0:2]), int(digits[2:4] or 0)
    sec = int(digits[4:6] or 0) if len(digits) >= 6 else 0
    return sign * dt.timedelta(hours=h, minutes=m, seconds=sec)


def parse_duration(s: str) -> dt.timedelta:
    m = re.match(r"^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$", s.strip())
    if not m:
        raise ValueError(f"bad duration {s!r}")
    sign = -1 if m.group(1) == "-" else 1
    w, d, h, mi, se = (int(x or 0) for x in m.groups()[1:])
    return sign * dt.timedelta(weeks=w, days=d, hours=h, minutes=mi, seconds=se)


# ---- recurrence rules ------------------------------------------------------

def parse_rrule(s: str) -> dict:
    r = {"FREQ": None, "INTERVAL": 1, "COUNT": None, "UNTIL": None, "BYDAY": [], "BYMONTHDAY": [],
         "BYMONTH": [], "BYSETPOS": [], "BYYEARDAY": [], "WKST": 0}
    for part in s.split(";"):
        if "=" not in part:
            continue
        k, v = part.split("=", 1)
        k, v = k.strip().upper(), v.strip()
        try:
            if k == "FREQ":
                r["FREQ"] = v.upper()
            elif k == "INTERVAL":
                r["INTERVAL"] = max(1, int(v))
            elif k == "COUNT":
                r["COUNT"] = max(0, int(v))
            elif k == "UNTIL":
                r["UNTIL"] = v
            elif k == "BYDAY":
                for x in v.split(","):
                    mm = re.match(r"^([+-]?\d{1,2})?(MO|TU|WE|TH|FR|SA|SU)$", x.strip().upper())
                    if mm:
                        r["BYDAY"].append((int(mm.group(1)) if mm.group(1) else None, WEEKDAY_CODES[mm.group(2)]))
            elif k in ("BYMONTHDAY", "BYMONTH", "BYSETPOS", "BYYEARDAY"):
                r[k] = [int(x) for x in v.split(",") if x.strip()]
            elif k == "WKST":
                r["WKST"] = WEEKDAY_CODES.get(v.upper(), 0)
        except ValueError:
            continue
    return r


def _add_months(d: dt.date, n: int) -> dt.date:
    idx = d.year * 12 + (d.month - 1) + n
    return dt.date(idx // 12, idx % 12 + 1, 1)


def _weekdays_in(first: dt.date, last: dt.date, wd: int) -> list:
    d = first + dt.timedelta(days=(wd - first.weekday()) % 7)
    out = []
    while d <= last:
        out.append(d)
        d += dt.timedelta(days=7)
    return out


def _pick(days: list, n: int | None) -> list:
    if n is None:
        return days
    if n > 0:
        return [days[n - 1]] if n <= len(days) else []
    return [days[n]] if -n <= len(days) else []


def _monthday_dates(y: int, m: int, mdays: list) -> list:
    last = _calendar.monthrange(y, m)[1]
    out = []
    for md in mdays:
        day = md if md > 0 else last + md + 1
        if 1 <= day <= last:
            out.append(dt.date(y, m, day))
    return out


def _period_candidates(rule: dict, pstart: dt.date, first: dt.date) -> list:
    freq = rule["FREQ"]
    bm, bmd, bd, byd = rule["BYMONTH"], rule["BYMONTHDAY"], rule["BYDAY"], rule["BYYEARDAY"]
    wdset = {w for _, w in bd}
    days: list = []
    if freq == "DAILY":
        d = pstart
        if bm and d.month not in bm:
            return []
        if bmd and d not in _monthday_dates(d.year, d.month, bmd):
            return []
        if bd and d.weekday() not in wdset:
            return []
        return [d]
    if freq == "WEEKLY":
        if bd:
            days = sorted({pstart + dt.timedelta(days=(w - pstart.weekday()) % 7) for w in wdset})
        else:
            days = [pstart + dt.timedelta(days=(first.weekday() - pstart.weekday()) % 7)]
        if bm:
            days = [d for d in days if d.month in bm]
        return days
    if freq == "MONTHLY":
        y, m = pstart.year, pstart.month
        if bm and m not in bm:
            return []
        if bmd:
            days = _monthday_dates(y, m, bmd)
            if bd:
                days = [d for d in days if d.weekday() in wdset]
        elif bd:
            last = dt.date(y, m, _calendar.monthrange(y, m)[1])
            s = set()
            for n, w in bd:
                s.update(_pick(_weekdays_in(dt.date(y, m, 1), last, w), n))
            days = sorted(s)
        else:
            days = _monthday_dates(y, m, [first.day]) if first.day <= _calendar.monthrange(y, m)[1] else []
        return sorted(set(days))
    if freq == "YEARLY":
        y = pstart.year
        ylen = 366 if _calendar.isleap(y) else 365
        if byd:
            for yd in byd:
                n = yd if yd > 0 else ylen + yd + 1
                if 1 <= n <= ylen:
                    days.append(dt.date(y, 1, 1) + dt.timedelta(days=n - 1))
            if bm:
                days = [d for d in days if d.month in bm]
            if bmd:
                days = [d for d in days if d in _monthday_dates(d.year, d.month, bmd)]
            if bd:
                days = [d for d in days if d.weekday() in wdset]
        elif bmd:
            for m in (bm or range(1, 13)):
                days.extend(_monthday_dates(y, m, bmd))
            if bd:
                days = [d for d in days if d.weekday() in wdset]
        elif bd:
            s = set()
            if bm:
                for m in bm:
                    last = dt.date(y, m, _calendar.monthrange(y, m)[1])
                    for n, w in bd:
                        s.update(_pick(_weekdays_in(dt.date(y, m, 1), last, w), n))
            else:
                for n, w in bd:
                    s.update(_pick(_weekdays_in(dt.date(y, 1, 1), dt.date(y, 12, 31), w), n))
            days = list(s)
        else:
            for m in (bm or [first.month]):
                if first.day <= _calendar.monthrange(y, m)[1]:
                    days.append(dt.date(y, m, first.day))
        return sorted(set(days))
    return []


def _apply_setpos(days: list, setpos: list) -> list:
    if not setpos:
        return days
    out = set()
    for p in setpos:
        if p > 0 and p <= len(days):
            out.add(days[p - 1])
        elif p < 0 and -p <= len(days):
            out.add(days[p])
    return sorted(out)


def expand_rrule(dtstart: dt.datetime, rule: dict, win_start: dt.datetime, win_end: dt.datetime,
                 until: dt.datetime | None = None, limit: int = 5000) -> list:
    """Occurrences (naive wall-clock datetimes) of a rule that fall in [win_start, win_end).

    DTSTART always counts as the first occurrence, as RFC 5545 requires.
    """
    freq = rule.get("FREQ")
    out: list = []
    if dtstart < win_end and dtstart >= win_start and (until is None or dtstart <= until):
        out.append(dtstart)
    if freq not in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
        return out
    interval = rule["INTERVAL"]
    count = rule["COUNT"]
    if count is not None and count <= 1:
        return out
    first = dtstart.date()
    tpart = dtstart.time()
    if freq == "DAILY":
        p0 = first
    elif freq == "WEEKLY":
        p0 = first - dt.timedelta(days=(first.weekday() - rule["WKST"]) % 7)
    elif freq == "MONTHLY":
        p0 = dt.date(first.year, first.month, 1)
    else:
        p0 = dt.date(first.year, 1, 1)

    k = 0
    if count is None:  # jump straight to the window; COUNT rules must be walked from the start
        target = win_start.date()
        if target > first:
            if freq == "DAILY":
                n = (target - p0).days // interval
            elif freq == "WEEKLY":
                n = (target - p0).days // (7 * interval)
            elif freq == "MONTHLY":
                n = ((target.year - p0.year) * 12 + target.month - p0.month) // interval
            else:
                n = (target.year - p0.year) // interval
            k = max(0, n - 1)

    produced = 1  # DTSTART
    end_date = win_end.date()
    for _ in range(200000):
        if freq == "DAILY":
            ps = p0 + dt.timedelta(days=k * interval)
        elif freq == "WEEKLY":
            ps = p0 + dt.timedelta(days=7 * k * interval)
        elif freq == "MONTHLY":
            if p0.year * 12 + k * interval > 9998 * 12:
                break
            ps = _add_months(p0, k * interval)
        else:
            if p0.year + k * interval > 9998:
                break
            ps = dt.date(p0.year + k * interval, 1, 1)
        k += 1
        if ps > end_date:
            break
        if until is not None and ps > until.date():
            break
        days = _apply_setpos(_period_candidates(rule, ps, first), rule["BYSETPOS"])
        for d in days:
            c = dt.datetime.combine(d, tpart)
            if c <= dtstart:
                continue
            if until is not None and c > until:
                return out
            produced += 1
            if count is not None and produced > count:
                return out
            if c >= win_end:
                return out
            if c >= win_start:
                out.append(c)
                if len(out) >= limit:
                    return out
    return out


# ---- time zones defined inside the feed -----------------------------------

class VTimezone(dt.tzinfo):
    """tzinfo built from an ICS VTIMEZONE block (what Outlook feeds use)."""

    def __init__(self, tzid: str, comp: Component):
        self.tzid = tzid
        self.observances = []
        for c in comp.children:
            if c.name not in ("STANDARD", "DAYLIGHT"):
                continue
            start, to = c.first("DTSTART"), c.first("TZOFFSETTO")
            if not start or not to:
                continue
            try:
                onset = parse_dt_value(start[1], {})
                if isinstance(onset, dt.date) and not isinstance(onset, dt.datetime):
                    onset = dt.datetime.combine(onset, dt.time())
                onset = onset.replace(tzinfo=None)
                off_to = parse_offset(to[1])
                frm = c.first("TZOFFSETFROM")
                off_from = parse_offset(frm[1]) if frm else off_to
            except ValueError:
                continue
            rr = c.first("RRULE")
            rule = parse_rrule(rr[1]) if rr else None
            rdates = []
            for p, v in c.every("RDATE"):
                for item in v.split(","):
                    try:
                        rd = parse_dt_value(item, {})
                        if isinstance(rd, dt.datetime):
                            rdates.append(rd.replace(tzinfo=None))
                    except ValueError:
                        pass
            self.observances.append((c.name, onset, off_from, off_to, rule, rdates))
        self._cache: dict = {}

    def _onsets(self, year: int) -> list:
        hit = self._cache.get(year)
        if hit is not None:
            return hit
        ws, we = dt.datetime(year - 1, 1, 1), dt.datetime(year + 1, 1, 1)
        lst = []
        for kind, onset, off_from, off_to, rule, rdates in self.observances:
            if rule:
                until = None
                if rule.get("UNTIL"):
                    try:
                        u = parse_dt_value(rule["UNTIL"], {})
                        until = (u if isinstance(u, dt.datetime) else dt.datetime.combine(u, dt.time(23, 59, 59)))
                        until = until.replace(tzinfo=None) + off_from
                    except ValueError:
                        until = None
                for occ in expand_rrule(onset, rule, ws, we, until):
                    lst.append((occ, off_to, off_from, kind))
            elif onset < we:
                lst.append((onset, off_to, off_from, kind))
            for rd in rdates:
                if rd < we:
                    lst.append((rd, off_to, off_from, kind))
        lst.sort(key=lambda x: x[0])
        if len(self._cache) > 64:
            self._cache.clear()
        self._cache[year] = lst
        return lst

    def _lookup(self, naive: dt.datetime):
        lst = self._onsets(naive.year)
        cur = None
        for onset, off_to, off_from, kind in lst:
            if onset <= naive:
                cur = (off_to, kind)
            else:
                break
        if cur is None:
            if lst:
                return lst[0][2], "STANDARD"
            if self.observances:
                return self.observances[0][3], self.observances[0][0]
            return dt.timedelta(0), "STANDARD"
        return cur

    def utcoffset(self, d):
        return self._lookup(d.replace(tzinfo=None))[0] if d is not None else None

    def dst(self, d):
        return dt.timedelta(0)

    def tzname(self, d):
        return self.tzid

    def fromutc(self, d):
        naive = d.replace(tzinfo=None)
        guess = naive + self._lookup(naive)[0]
        return (naive + self._lookup(guess)[0]).replace(tzinfo=self)

    def __repr__(self):
        return f"VTimezone({self.tzid!r})"


def make_tz_resolver(vtimezones: dict, unknown: set | None = None):
    """unknown, if given, collects the TZIDs that couldn't be resolved."""
    cache: dict = {}

    def resolve(tzid: str):
        key = (tzid or "").strip().strip('"')
        if key in cache:
            return cache[key]
        tz = vtimezones.get(key)
        if tz is None and key.upper() in ("UTC", "GMT", "Z", "ETC/UTC", "COORDINATED UNIVERSAL TIME"):
            tz = dt.timezone.utc
        if tz is None:
            tz = zone_by_name(key) or zone_by_name(WINDOWS_TZ.get(key, ""))
        if tz is None and key and unknown is not None:
            unknown.add(key)
        cache[key] = tz  # None -> treat as local time
        return tz

    return resolve


# ---- events ----------------------------------------------------------------

_MEETING_URL = re.compile(
    r"https://(?:[\w-]+\.)?(?:teams\.microsoft\.com|teams\.live\.com|zoom\.us|[\w-]+\.zoom\.us|meet\.google\.com|"
    r"[\w-]+\.webex\.com)/[^\s<>\"')\]]+", re.I)

_BUSY_MAP = {"FREE": "free", "TENTATIVE": "tentative", "BUSY": "busy", "OOF": "oof",
             "WORKINGELSEWHERE": "elsewhere"}


def _instance_key(value, all_day: bool) -> str:
    if isinstance(value, dt.datetime):
        if all_day:
            return value.date().isoformat()
        if value.tzinfo is None:
            value = value.astimezone()  # floating = this computer's time
        return value.astimezone(dt.timezone.utc).strftime("%Y%m%dT%H%M%S")
    return value.isoformat()


class ICalendar:
    def __init__(self, text: str):
        root = parse_components(text)
        vtz = {}
        for c in root.walk("VTIMEZONE"):
            tzid = c.first("TZID")
            if tzid:
                vtz[tzid[1].strip().strip('"')] = VTimezone(tzid[1].strip(), c)
        self.unknown_zones: set = set()
        self.tzres = make_tz_resolver(vtz, self.unknown_zones)
        self.masters: list = []
        self.singles: list = []
        self.overrides: dict = {}
        self.skipped = 0
        for c in root.walk("VEVENT"):
            try:
                ev = self._read_event(c)
            except (ValueError, TypeError, OverflowError):
                self.skipped += 1
                continue
            if ev is None:
                continue
            if ev["rid"] is not None:
                self.overrides.setdefault(ev["uid"], {})[ev["rid"]] = ev
            elif ev["rule"] or ev["rdates"]:
                self.masters.append(ev)
            else:
                self.singles.append(ev)

    def _read_event(self, c: Component):
        s = c.first("DTSTART")
        if not s:
            return None
        start = parse_dt_value(s[1], s[0], self.tzres)
        all_day = not isinstance(start, dt.datetime)
        e = c.first("DTEND")
        if e:
            end = parse_dt_value(e[1], e[0], self.tzres)
            if all_day and isinstance(end, dt.datetime):
                end = end.date()
            if not all_day and not isinstance(end, dt.datetime):
                end = dt.datetime.combine(end, dt.time()).replace(tzinfo=start.tzinfo)
            if all_day:
                duration = end - start
            elif (start.tzinfo is None) != (end.tzinfo is None):
                duration = dt.timedelta(0)
            elif start.tzinfo is not None and end.tzinfo is not start.tzinfo:
                duration = end.astimezone(dt.timezone.utc) - start.astimezone(dt.timezone.utc)
            else:
                duration = end - start  # same zone: wall-clock length, right for recurring events
        else:
            du = c.first("DURATION")
            duration = parse_duration(du[1]) if du else (dt.timedelta(days=1) if all_day else dt.timedelta(0))
        if duration < dt.timedelta(0):
            duration = dt.timedelta(0)
        if all_day and duration < dt.timedelta(days=1):
            duration = dt.timedelta(days=1)

        def text(key):
            v = c.first(key)
            return unescape_text(v[1]).strip() if v else ""

        status = text("STATUS").upper()
        summary = text("SUMMARY") or "(No title)"
        cancelled = status == "CANCELLED" or summary.lower().startswith(("canceled:", "cancelled:"))
        busy = _BUSY_MAP.get(text("X-MICROSOFT-CDO-BUSYSTATUS").upper())
        if busy is None:
            busy = "free" if text("TRANSP").upper() == "TRANSPARENT" else ("tentative" if status == "TENTATIVE" else "busy")
        location = text("LOCATION")
        categories = []
        for _p, v in c.every("CATEGORIES"):  # a comma-separated list; "\," is a comma inside a name
            categories += [unescape_text(x).strip() for x in re.split(r"(?<!\\),", v)]
        categories = split_categories(categories)
        url = ""
        for key in ("X-MICROSOFT-SKYPETEAMSMEETINGURL", "X-MICROSOFT-ONLINEMEETINGEXTERNALLINK",
                    "X-GOOGLE-CONFERENCE", "URL"):
            val = text(key)
            if val.startswith("https://"):
                url = val
                break
        if not url:
            m = _MEETING_URL.search(location + "\n" + text("DESCRIPTION"))
            if m:
                url = m.group(0)

        rule = None
        rr = c.first("RRULE")
        if rr:
            rule = parse_rrule(rr[1])
            if rule.get("FREQ") not in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
                rule = None
        exdates = set()
        for p, v in c.every("EXDATE"):
            for item in v.split(","):
                try:
                    exdates.add(_instance_key(parse_dt_value(item, p, self.tzres), all_day))
                except ValueError:
                    pass
        rdates = []
        for p, v in c.every("RDATE"):
            if p.get("VALUE", "").upper() == "PERIOD":
                continue
            for item in v.split(","):
                try:
                    rdates.append(parse_dt_value(item, p, self.tzres))
                except ValueError:
                    pass
        rid = None
        r = c.first("RECURRENCE-ID")
        if r:
            rid = _instance_key(parse_dt_value(r[1], r[0], self.tzres), all_day)
        uid = text("UID") or f"nouid-{id(c)}"
        return {"uid": uid, "start": start, "duration": duration, "all_day": all_day, "title": summary,
                "location": location, "busy": busy, "cancelled": cancelled, "status": status, "url": url,
                "categories": categories, "rule": rule, "exdates": exdates, "rdates": rdates, "rid": rid}

    # -- expansion ------------------------------------------------------------
    def _occurrences(self, ev: dict, ws: dt.datetime, we: dt.datetime) -> list:
        """Start values (same type as ev['start']) of a master event near [ws, we)."""
        start, dur = ev["start"], ev["duration"]
        margin = dur + dt.timedelta(days=1)
        if ev["all_day"]:
            s0 = dt.datetime.combine(start, dt.time())
            lo = dt.datetime.combine(ws.date(), dt.time()) - margin
            hi = dt.datetime.combine(we.date(), dt.time()) + dt.timedelta(days=1)
            until = self._until(ev, None, all_day=True)
            occ = [o.date() for o in expand_rrule(s0, ev["rule"], lo, hi, until)] if ev["rule"] else []
            occ += [r if not isinstance(r, dt.datetime) else r.date() for r in ev["rdates"]]
            return occ
        tz = start.tzinfo
        frame = (lambda x: x.astimezone(tz).replace(tzinfo=None)) if tz else (lambda x: to_local_naive(x))
        lo, hi = frame(ws) - margin, frame(we) + dt.timedelta(days=1)
        until = self._until(ev, tz)
        occ = expand_rrule(start.replace(tzinfo=None), ev["rule"], lo, hi, until) if ev["rule"] else []
        out = [o.replace(tzinfo=tz) if tz else o for o in occ]
        for r in ev["rdates"]:
            if isinstance(r, dt.datetime):
                out.append(r.astimezone(tz) if (tz and r.tzinfo) else r)
        return out

    def _until(self, ev: dict, tz, all_day: bool = False):
        rule = ev["rule"]
        if not rule or not rule.get("UNTIL"):
            return None
        try:
            u = parse_dt_value(rule["UNTIL"], {}, self.tzres)
        except ValueError:
            return None
        if all_day:
            d = u.date() if isinstance(u, dt.datetime) else u
            return dt.datetime.combine(d, dt.time(23, 59, 59))
        if not isinstance(u, dt.datetime):
            return dt.datetime.combine(u, dt.time(23, 59, 59))
        if u.tzinfo is not None:
            return u.astimezone(tz).replace(tzinfo=None) if tz else to_local_naive(u)
        return u

    def events_between(self, start_date: dt.date, end_date: dt.date) -> list:
        """Event instances overlapping [start_date, end_date) in this computer's local time."""
        ws = dt.datetime.combine(start_date, dt.time()).astimezone()
        we = dt.datetime.combine(end_date, dt.time()).astimezone()
        out = []
        for ev in self.singles:
            item = self._emit(ev, ev["start"], start_date, end_date, ws, we)
            if item:
                out.append(item)
        for ev in self.masters:
            overrides = self.overrides.get(ev["uid"], {})
            for occ in self._occurrences(ev, ws, we):
                key = _instance_key(occ, ev["all_day"])
                if key in ev["exdates"] or key in overrides:
                    continue
                item = self._emit(ev, occ, start_date, end_date, ws, we, key)
                if item:
                    out.append(item)
        for uid, by_rid in self.overrides.items():
            for rid, ev in by_rid.items():
                item = self._emit(ev, ev["start"], start_date, end_date, ws, we, rid)
                if item:
                    out.append(item)
        out.sort(key=lambda x: (x["start"], x["title"]))
        return out

    def _emit(self, ev, occ_start, start_date, end_date, ws, we, key=None):
        if ev["status"] == "CANCELLED":
            return None
        if ev["all_day"]:
            s = occ_start
            e = s + dt.timedelta(days=max(1, ev["duration"].days))
            if not (s < end_date and e > start_date):
                return None
            return {"id": f"{ev['uid']}|{key or s.isoformat()}", "title": ev["title"], "allDay": True,
                    "start": s.isoformat(), "end": e.isoformat(), "location": ev["location"],
                    "busy": ev["busy"], "cancelled": ev["cancelled"], "url": ev["url"],
                    "categories": list(ev["categories"]), "color": None}
        s = occ_start
        if s.tzinfo is not None:
            e = (s.replace(tzinfo=None) + ev["duration"]).replace(tzinfo=s.tzinfo)
            s_loc, e_loc = to_local_naive(s), to_local_naive(e)
        else:
            s_loc, e_loc = s, s + ev["duration"]
        lo, hi = to_local_naive(ws), to_local_naive(we)
        if not (s_loc < hi and (e_loc > lo or (e_loc == s_loc and s_loc >= lo))):
            return None
        return {"id": f"{ev['uid']}|{key or _instance_key(s, False)}", "title": ev["title"], "allDay": False,
                "start": s_loc.strftime("%Y-%m-%dT%H:%M"), "end": e_loc.strftime("%Y-%m-%dT%H:%M"),
                "location": ev["location"], "busy": ev["busy"], "cancelled": ev["cancelled"], "url": ev["url"],
                "categories": list(ev["categories"]), "color": None}  # published feeds don't carry colours
