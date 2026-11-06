"""The iCalendar (.ics) reader: line handling, repeat rules, time zones, and events in a date range."""
import datetime as dt
import unittest

from planner import demo
from planner import ics as ical
from tests.support import AUS_EASTERN, has_zone, ics, local_from_utc

D = dt.datetime


class LowLevelParsingTests(unittest.TestCase):
    def test_unfold_lines(self):
        text = "SUMMARY:Long\r\n  title\r\n\tcontinued\r\n\r\nLOCATION:Room\nDESCRIPTION:x\rEND:VEVENT"
        self.assertEqual(ical.unfold_lines(text),
                         ["SUMMARY:Long titlecontinued", "LOCATION:Room", "DESCRIPTION:x", "END:VEVENT"])

    def test_content_line_with_params_and_quoted_colon(self):
        name, params, value = ical.parse_content_line(
            'DTSTART;TZID="(UTC+10:00) Canberra, Melbourne";VALUE=DATE-TIME:20261110T090000')
        self.assertEqual(name, "DTSTART")
        self.assertEqual(params, {"TZID": "(UTC+10:00) Canberra, Melbourne", "VALUE": "DATE-TIME"})
        self.assertEqual(value, "20261110T090000")

    def test_content_line_edge_cases(self):
        self.assertEqual(ical.parse_content_line("item1.summary:Hi")[0], "SUMMARY")
        self.assertEqual(ical.parse_content_line("URL:https://x.test/a:b")[2], "https://x.test/a:b")
        self.assertIsNone(ical.parse_content_line("no colon here"))
        self.assertIsNone(ical.parse_content_line(":starts with colon"))

    def test_unescape_text(self):
        self.assertEqual(ical.unescape_text(r"Budget\, planning\; Q1\nRoom 2\\3"), "Budget, planning; Q1\nRoom 2\\3")

    def test_parse_dt_value(self):
        self.assertEqual(ical.parse_dt_value("20261110", {}), dt.date(2026, 11, 10))
        self.assertEqual(ical.parse_dt_value("20261110T090000Z", {}).utcoffset(), dt.timedelta(0))
        floating = ical.parse_dt_value("20261110T0900", {})
        self.assertEqual(floating, D(2026, 11, 10, 9, 0))
        self.assertIsNone(floating.tzinfo)
        tz = dt.timezone(dt.timedelta(hours=10))
        self.assertIs(ical.parse_dt_value("20261110T090000", {"TZID": "X"}, lambda _: tz).tzinfo, tz)
        with self.assertRaises(ValueError):
            ical.parse_dt_value("2026-11-10", {})

    def test_parse_duration(self):
        self.assertEqual(ical.parse_duration("PT1H30M"), dt.timedelta(minutes=90))
        self.assertEqual(ical.parse_duration("P1D"), dt.timedelta(days=1))
        self.assertEqual(ical.parse_duration("P2W"), dt.timedelta(weeks=2))
        self.assertEqual(ical.parse_duration("-PT15M"), -dt.timedelta(minutes=15))
        with self.assertRaises(ValueError):
            ical.parse_duration("1 hour")

    def test_parse_offset(self):
        self.assertEqual(ical.parse_offset("+1100"), dt.timedelta(hours=11))
        self.assertEqual(ical.parse_offset("-0530"), -dt.timedelta(hours=5, minutes=30))
        self.assertEqual(ical.parse_offset("+013045"), dt.timedelta(hours=1, minutes=30, seconds=45))

    def test_parse_rrule(self):
        r = ical.parse_rrule("FREQ=monthly;INTERVAL=2;BYDAY=2TU,-1FR,MO;COUNT=5;BYSETPOS=-1;WKST=SU;JUNK")
        self.assertEqual(r["FREQ"], "MONTHLY")
        self.assertEqual(r["INTERVAL"], 2)
        self.assertEqual(r["COUNT"], 5)
        self.assertEqual(r["BYDAY"], [(2, 1), (-1, 4), (None, 0)])
        self.assertEqual(r["BYSETPOS"], [-1])
        self.assertEqual(r["WKST"], 6)

    def test_bad_rrule_parts_are_ignored(self):
        r = ical.parse_rrule("FREQ=DAILY;INTERVAL=abc;COUNT=x")
        self.assertEqual((r["FREQ"], r["INTERVAL"], r["COUNT"]), ("DAILY", 1, None))


def expand(start, rule, ws, we, until=None):
    return ical.expand_rrule(start, ical.parse_rrule(rule), ws, we, until)


class RepeatRuleTests(unittest.TestCase):
    """Dates checked against a calendar: 2 Nov 2026 is a Monday."""

    def test_daily_count(self):
        out = expand(D(2026, 11, 2, 9), "FREQ=DAILY;COUNT=3", D(2026, 1, 1), D(2027, 1, 1))
        self.assertEqual(out, [D(2026, 11, 2, 9), D(2026, 11, 3, 9), D(2026, 11, 4, 9)])

    def test_weekly_on_several_days(self):
        out = expand(D(2026, 11, 2, 9), "FREQ=WEEKLY;BYDAY=MO,WE,FR", D(2026, 11, 2), D(2026, 11, 16))
        self.assertEqual([o.day for o in out], [2, 4, 6, 9, 11, 13])

    def test_fortnightly(self):
        out = expand(D(2026, 11, 3, 9), "FREQ=WEEKLY;INTERVAL=2;BYDAY=TU", D(2026, 11, 1), D(2026, 12, 1))
        self.assertEqual(out, [D(2026, 11, 3, 9), D(2026, 11, 17, 9)])

    def test_second_tuesday_of_the_month(self):
        out = expand(D(2026, 11, 10, 14), "FREQ=MONTHLY;BYDAY=2TU", D(2026, 11, 1), D(2027, 3, 1))
        self.assertEqual([o.date() for o in out],
                         [dt.date(2026, 11, 10), dt.date(2026, 12, 8), dt.date(2027, 1, 12), dt.date(2027, 2, 9)])

    def test_last_day_of_the_month(self):
        out = expand(D(2026, 11, 30, 9), "FREQ=MONTHLY;BYMONTHDAY=-1", D(2026, 11, 1), D(2027, 3, 1))
        self.assertEqual([o.date() for o in out],
                         [dt.date(2026, 11, 30), dt.date(2026, 12, 31), dt.date(2027, 1, 31), dt.date(2027, 2, 28)])

    def test_monthly_on_the_31st_skips_short_months(self):
        out = expand(D(2026, 10, 31, 9), "FREQ=MONTHLY", D(2026, 10, 1), D(2027, 4, 1))
        self.assertEqual([o.date() for o in out],
                         [dt.date(2026, 10, 31), dt.date(2026, 12, 31), dt.date(2027, 1, 31), dt.date(2027, 3, 31)])

    def test_last_weekday_of_the_month(self):
        out = expand(D(2026, 11, 30, 16), "FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-1",
                     D(2026, 11, 1), D(2027, 2, 1))
        self.assertEqual([o.date() for o in out], [dt.date(2026, 11, 30), dt.date(2026, 12, 31), dt.date(2027, 1, 29)])

    def test_yearly_on_29_february(self):
        out = expand(D(2024, 2, 29, 9), "FREQ=YEARLY", D(2024, 1, 1), D(2033, 1, 1))
        self.assertEqual([o.year for o in out], [2024, 2028, 2032])

    def test_until_is_inclusive(self):
        out = expand(D(2026, 11, 2, 9), "FREQ=DAILY", D(2026, 1, 1), D(2027, 1, 1), until=D(2026, 11, 5, 9))
        self.assertEqual([o.day for o in out], [2, 3, 4, 5])

    def test_old_series_jumps_to_the_window(self):
        out = expand(D(2020, 1, 6, 9), "FREQ=WEEKLY;BYDAY=MO", D(2026, 11, 2), D(2026, 11, 9))
        self.assertEqual(out, [D(2026, 11, 2, 9)])

    def test_count_includes_the_first_date_even_off_pattern(self):
        out = expand(D(2026, 11, 1, 10), "FREQ=WEEKLY;BYDAY=MO;COUNT=3", D(2026, 1, 1), D(2027, 1, 1))
        self.assertEqual(out, [D(2026, 11, 1, 10), D(2026, 11, 2, 10), D(2026, 11, 9, 10)])

    def test_unsupported_frequency_gives_only_the_first_date(self):
        out = expand(D(2026, 11, 2, 9), "FREQ=HOURLY", D(2026, 1, 1), D(2027, 1, 1))
        self.assertEqual(out, [D(2026, 11, 2, 9)])


class FeedTimezoneTests(unittest.TestCase):
    def setUp(self):
        comp = next(ical.parse_components(AUS_EASTERN).walk("VTIMEZONE"))
        self.tz = ical.VTimezone("AUS Eastern Standard Time", comp)

    def off(self, *args):
        return self.tz.utcoffset(D(*args)) / dt.timedelta(hours=1)

    def test_summer_and_winter(self):
        self.assertEqual(self.off(2026, 1, 15, 12), 11)
        self.assertEqual(self.off(2026, 7, 15, 12), 10)

    def test_daylight_saving_ends_first_sunday_of_april(self):  # 5 Apr 2026, 3am -> 2am
        self.assertEqual(self.off(2026, 4, 5, 2, 59), 11)
        self.assertEqual(self.off(2026, 4, 5, 3, 0), 10)

    def test_daylight_saving_starts_first_sunday_of_october(self):  # 4 Oct 2026, 2am -> 3am
        self.assertEqual(self.off(2026, 10, 4, 1, 59), 10)
        self.assertEqual(self.off(2026, 10, 4, 2, 0), 11)

    def test_converting_from_utc(self):
        utc = D(2026, 11, 9, 23, 0, tzinfo=dt.timezone.utc)
        self.assertEqual(utc.astimezone(self.tz).replace(tzinfo=None), D(2026, 11, 10, 10, 0))


def between(text, start, end):
    return ical.ICalendar(text).events_between(start, end)


class EventsBetweenTests(unittest.TestCase):
    def test_single_event(self):
        [ev] = between(ics("UID:a\nSUMMARY:Planning\nDTSTART:20261110T090000\nDTEND:20261110T100000\n"
                           "LOCATION:Room 1"), dt.date(2026, 11, 10), dt.date(2026, 11, 11))
        self.assertEqual((ev["title"], ev["start"], ev["end"]), ("Planning", "2026-11-10T09:00", "2026-11-10T10:00"))
        self.assertEqual((ev["allDay"], ev["location"], ev["busy"], ev["cancelled"]), (False, "Room 1", "busy", False))

    def test_window_edges(self):
        text = ics("UID:before\nSUMMARY:Ends at midnight\nDTSTART:20261109T230000\nDTEND:20261110T000000",
                   "UID:across\nSUMMARY:Crosses midnight\nDTSTART:20261110T233000\nDTEND:20261111T003000",
                   "UID:zero\nSUMMARY:Reminder\nDTSTART:20261110T000000")
        titles = [e["title"] for e in between(text, dt.date(2026, 11, 10), dt.date(2026, 11, 11))]
        self.assertEqual(sorted(titles), ["Crosses midnight", "Reminder"])
        self.assertEqual([e["title"] for e in between(text, dt.date(2026, 11, 11), dt.date(2026, 11, 12))],
                         ["Crosses midnight"])

    def test_all_day_events(self):
        text = ics("UID:conf\nSUMMARY:Conference\nDTSTART;VALUE=DATE:20261111\nDTEND;VALUE=DATE:20261113",
                   "UID:one\nSUMMARY:Holiday\nDTSTART;VALUE=DATE:20261113")
        [conf] = between(text, dt.date(2026, 11, 12), dt.date(2026, 11, 13))
        self.assertEqual((conf["allDay"], conf["start"], conf["end"]), (True, "2026-11-11", "2026-11-13"))
        [hol] = between(text, dt.date(2026, 11, 13), dt.date(2026, 11, 14))
        self.assertEqual((hol["title"], hol["end"]), ("Holiday", "2026-11-14"))

    def test_series_with_skipped_moved_and_cancelled_meetings(self):
        text = ics("UID:s\nSUMMARY:Check-in\nDTSTART:20261102T090000\nDTEND:20261102T091500\n"
                   "RRULE:FREQ=WEEKLY;BYDAY=MO,WE\nEXDATE:20261104T090000",
                   "UID:s\nRECURRENCE-ID:20261109T090000\nSUMMARY:Check-in (moved)\n"
                   "DTSTART:20261109T140000\nDTEND:20261109T141500",
                   "UID:s\nRECURRENCE-ID:20261111T090000\nSTATUS:CANCELLED\nSUMMARY:Check-in\n"
                   "DTSTART:20261111T090000\nDTEND:20261111T091500")
        got = [(e["start"], e["title"]) for e in between(text, dt.date(2026, 11, 2), dt.date(2026, 11, 17))]
        self.assertEqual(got, [("2026-11-02T09:00", "Check-in"),
                               ("2026-11-09T14:00", "Check-in (moved)"),
                               ("2026-11-16T09:00", "Check-in")])

    def test_ids_are_unique_per_occurrence(self):
        text = ics("UID:s\nSUMMARY:Daily\nDTSTART:20261102T090000\nDTEND:20261102T091500\nRRULE:FREQ=DAILY")
        ids = [e["id"] for e in between(text, dt.date(2026, 11, 2), dt.date(2026, 11, 9))]
        self.assertEqual(len(ids), 7)
        self.assertEqual(len(set(ids)), 7)

    def test_cancelled_events(self):
        text = ics("UID:gone\nSUMMARY:Gone\nSTATUS:CANCELLED\nDTSTART:20261110T090000\nDTEND:20261110T100000",
                   "UID:prefix\nSUMMARY:Canceled: Vendor demo\nDTSTART:20261110T110000\nDTEND:20261110T120000")
        [ev] = between(text, dt.date(2026, 11, 10), dt.date(2026, 11, 11))
        self.assertEqual(ev["title"], "Canceled: Vendor demo")
        self.assertTrue(ev["cancelled"])

    def test_free_busy_status(self):
        text = ics("UID:1\nSUMMARY:Leave\nX-MICROSOFT-CDO-BUSYSTATUS:OOF\nDTSTART:20261110T090000",
                   "UID:2\nSUMMARY:Lunch\nTRANSP:TRANSPARENT\nDTSTART:20261110T120000",
                   "UID:3\nSUMMARY:Maybe\nSTATUS:TENTATIVE\nDTSTART:20261110T130000",
                   "UID:4\nSUMMARY:Elsewhere\nX-MICROSOFT-CDO-BUSYSTATUS:WORKINGELSEWHERE\nDTSTART:20261110T140000",
                   "UID:5\nSUMMARY:Normal\nDTSTART:20261110T150000")
        busy = {e["title"]: e["busy"] for e in between(text, dt.date(2026, 11, 10), dt.date(2026, 11, 11))}
        self.assertEqual(busy, {"Leave": "oof", "Lunch": "free", "Maybe": "tentative",
                                "Elsewhere": "elsewhere", "Normal": "busy"})

    def test_meeting_links(self):
        text = ics("UID:1\nSUMMARY:Teams\nX-MICROSOFT-SKYPETEAMSMEETINGURL:https://teams.microsoft.com/l/abc\n"
                   "DESCRIPTION:also https://zoom.us/j/999\nDTSTART:20261110T090000",
                   "UID:2\nSUMMARY:Zoom\nDESCRIPTION:Join at https://example.zoom.us/j/123?pwd=x for the call\n"
                   "DTSTART:20261110T100000",
                   "UID:3\nSUMMARY:Plain\nDESCRIPTION:Agenda at https://example.com/agenda\nDTSTART:20261110T110000")
        urls = {e["title"]: e["url"] for e in between(text, dt.date(2026, 11, 10), dt.date(2026, 11, 11))}
        self.assertEqual(urls, {"Teams": "https://teams.microsoft.com/l/abc",
                                "Zoom": "https://example.zoom.us/j/123?pwd=x", "Plain": ""})

    def test_escaped_text_and_missing_title(self):
        text = ics(r"UID:1\nSUMMARY:Budget\, planning\nDTSTART:20261110T090000".replace(r"\n", "\n"),
                   "UID:2\nDTSTART:20261110T100000")
        titles = sorted(e["title"] for e in between(text, dt.date(2026, 11, 10), dt.date(2026, 11, 11)))
        self.assertEqual(titles, ["(No title)", "Budget, planning"])

    def test_broken_events_are_skipped_not_fatal(self):
        cal = ical.ICalendar(ics("UID:bad\nSUMMARY:Bad\nDTSTART:tomorrow",
                                "UID:ok\nSUMMARY:Fine\nDTSTART:20261110T090000"))
        self.assertEqual(cal.skipped, 1)
        self.assertEqual([e["title"] for e in cal.events_between(dt.date(2026, 11, 10), dt.date(2026, 11, 11))],
                         ["Fine"])


class EventsInFeedTimezoneTests(unittest.TestCase):
    """Times written in Outlook's 'AUS Eastern Standard Time', checked as UTC instants so the
    tests pass wherever they run."""

    def test_event_in_feed_timezone(self):
        text = ics("UID:1\nSUMMARY:Board\nDTSTART;TZID=AUS Eastern Standard Time:20261110T100000\n"
                   "DTEND;TZID=AUS Eastern Standard Time:20261110T110000", timezones=AUS_EASTERN)
        [ev] = between(text, dt.date(2026, 11, 8), dt.date(2026, 11, 12))
        self.assertEqual((ev["start"], ev["end"]), (local_from_utc(2026, 11, 9, 23), local_from_utc(2026, 11, 10, 0)))

    def test_weekly_meeting_keeps_its_local_time_across_daylight_saving(self):
        # Mondays 10:00 Melbourne time. Daylight saving ends on Sunday 4 April 2027.
        text = ics("UID:1\nSUMMARY:Weekly\nDTSTART;TZID=AUS Eastern Standard Time:20270329T100000\n"
                   "DTEND;TZID=AUS Eastern Standard Time:20270329T110000\nRRULE:FREQ=WEEKLY;BYDAY=MO",
                   timezones=AUS_EASTERN)
        starts = [e["start"] for e in between(text, dt.date(2027, 3, 27), dt.date(2027, 4, 8))]
        self.assertEqual(starts, [local_from_utc(2027, 3, 28, 23), local_from_utc(2027, 4, 5, 0)])

    def test_until_in_utc_and_exdate_in_feed_timezone(self):
        text = ics("UID:1\nSUMMARY:Weekly\nDTSTART;TZID=AUS Eastern Standard Time:20261102T100000\n"
                   "DTEND;TZID=AUS Eastern Standard Time:20261102T110000\n"
                   "RRULE:FREQ=WEEKLY;BYDAY=MO;UNTIL=20261122T230000Z\n"
                   "EXDATE;TZID=AUS Eastern Standard Time:20261109T100000", timezones=AUS_EASTERN)
        starts = [e["start"] for e in between(text, dt.date(2026, 10, 30), dt.date(2026, 12, 31))]
        # 2 Nov and 16 Nov run, 9 Nov is skipped, 23 Nov 10:00 local (= 22 Nov 23:00 UTC) is the last one.
        self.assertEqual(starts, [local_from_utc(2026, 11, 1, 23), local_from_utc(2026, 11, 15, 23),
                                  local_from_utc(2026, 11, 22, 23)])

    def test_unknown_timezone_is_treated_as_local_and_reported(self):
        cal = ical.ICalendar(ics("UID:1\nSUMMARY:X\nDTSTART;TZID=Mars/Olympus:20261110T090000"))
        [ev] = cal.events_between(dt.date(2026, 11, 10), dt.date(2026, 11, 11))
        self.assertEqual(ev["start"], "2026-11-10T09:00")
        self.assertEqual(cal.unknown_zones, {"Mars/Olympus"})

    def test_zones_defined_in_the_feed_or_utc_are_not_reported(self):
        cal = ical.ICalendar(ics("UID:1\nSUMMARY:X\nDTSTART;TZID=AUS Eastern Standard Time:20261110T090000",
                                "UID:2\nSUMMARY:Y\nDTSTART;TZID=UTC:20261110T090000\nDTEND:20261110T100000Z",
                                timezones=AUS_EASTERN))
        self.assertEqual(cal.unknown_zones, set())

    @unittest.skipUnless(has_zone("Australia/Sydney"), "needs a time-zone database (pip install tzdata)")
    def test_windows_zone_name_without_vtimezone(self):
        [ev] = between(ics("UID:1\nSUMMARY:X\nDTSTART;TZID=AUS Eastern Standard Time:20261110T100000"),
                       dt.date(2026, 11, 8), dt.date(2026, 11, 12))
        self.assertEqual(ev["start"], local_from_utc(2026, 11, 9, 23))


class DemoCalendarTests(unittest.TestCase):
    def test_demo_week(self):
        today = dt.date(2026, 11, 6)
        monday = today - dt.timedelta(days=today.weekday())
        evs = ical.ICalendar(demo.demo_ics(today)).events_between(monday, monday + dt.timedelta(days=7))
        titles = [e["title"] for e in evs]
        self.assertEqual(titles.count("Team check-in"), 3)  # Mon, Wed, Fri: Tue moved, Thu skipped
        self.assertEqual(titles.count("Team check-in (moved)"), 1)
        self.assertTrue(next(e for e in evs if e["title"].startswith("Canceled"))["cancelled"])


if __name__ == "__main__":
    unittest.main()
