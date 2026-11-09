"""Outlook category colours: splitting category strings, choosing an event's colour, and where colours appear."""
import datetime as dt
import unittest

from planner import categories, demo
from planner import ics as ical
from planner.outlook import rows_to_events
from planner.sources import DemoCalendar
from tests.support import ics


class SplitTests(unittest.TestCase):
    def test_outlook_strings(self):
        self.assertEqual(categories.split_categories("Team, Projects"), ["Team", "Projects"])
        self.assertEqual(categories.split_categories("Team;Projects"), ["Team", "Projects"])
        self.assertEqual(categories.split_categories("  Team ,, team, "), ["Team"])
        self.assertEqual(categories.split_categories(""), [])
        self.assertEqual(categories.split_categories(None), [])
        self.assertEqual(categories.split_categories(["A", " B "]), ["A", "B"])


class ColorTests(unittest.TestCase):
    COLORS = {"Team": 8, "Notes": 0, "Projects": 16, "Odd": "x", "Big": 99}

    def test_first_category_with_a_colour_wins(self):
        self.assertEqual(categories.category_color(["Notes", "Team", "Projects"], self.COLORS), "#9DB7E8")
        self.assertEqual(categories.category_color(["Projects", "Team"], self.COLORS), "#C11A25")

    def test_names_match_regardless_of_case(self):
        self.assertEqual(categories.category_color(["team"], self.COLORS), "#9DB7E8")

    def test_no_colour(self):
        for names in ([], ["Notes"], ["Unknown"], ["Odd"], ["Big"]):
            self.assertIsNone(categories.category_color(names, self.COLORS), names)
        self.assertIsNone(categories.category_color(["Team"], None))

    def test_every_preset_is_a_hex_colour(self):
        self.assertEqual(sorted(categories.PRESET_COLORS), list(range(1, 26)))
        for v in categories.PRESET_COLORS.values():
            self.assertRegex(v, r"^#[0-9A-F]{6}$")


class OutlookRowTests(unittest.TestCase):
    def test_rows_carry_categories_and_colour(self):
        rows = [{"start": "2026-11-10T10:00:00", "end": "2026-11-10T11:00:00", "title": "A",
                 "categories": "Notes, Team"},
                {"start": "2026-11-10T12:00:00", "end": "2026-11-10T13:00:00", "title": "B", "categories": ""},
                {"start": "2026-11-10T14:00:00", "end": "2026-11-10T15:00:00", "title": "C"}]
        a, b, c = rows_to_events(rows, {"Team": 8, "Notes": 0})
        self.assertEqual((a["categories"], a["color"]), (["Notes", "Team"], "#9DB7E8"))
        self.assertEqual((b["categories"], b["color"]), ([], None))
        self.assertEqual((c["categories"], c["color"]), ([], None))

    def test_without_the_category_list_there_is_no_colour(self):
        [a] = rows_to_events([{"start": "2026-11-10T10:00:00", "end": "2026-11-10T11:00:00", "categories": "Team"}])
        self.assertEqual((a["categories"], a["color"]), (["Team"], None))


class IcsTests(unittest.TestCase):
    def test_categories_are_read_but_feeds_have_no_colours(self):
        text = ics("UID:1\nSUMMARY:A\nCATEGORIES:Team,Big\\, small\nCATEGORIES:Extra\nDTSTART:20261110T090000",
                   "UID:2\nSUMMARY:B\nDTSTART;VALUE=DATE:20261110")
        evs = {e["title"]: e for e in ical.ICalendar(text).events_between(dt.date(2026, 11, 10), dt.date(2026, 11, 11))}
        self.assertEqual(evs["A"]["categories"], ["Team", "Big, small", "Extra"])
        self.assertEqual(evs["B"]["categories"], [])
        self.assertIsNone(evs["A"]["color"])
        self.assertIsNone(evs["B"]["color"])


class DemoTests(unittest.TestCase):
    def test_demo_events_are_coloured_by_category(self):
        today = dt.date(2026, 11, 6)
        monday = today - dt.timedelta(days=today.weekday())
        evs = DemoCalendar(today).events(monday, monday + dt.timedelta(days=7), False)
        by_title = {e["title"]: e for e in evs}
        self.assertEqual(by_title["Leadership meeting"]["color"],
                         categories.PRESET_COLORS[demo.DEMO_CATEGORY_COLORS["Leadership"]])
        self.assertEqual(by_title["Planning workshop: 2027 roadmap"]["categories"], ["Projects", "Leadership"])
        self.assertEqual(by_title["Planning workshop: 2027 roadmap"]["color"],
                         categories.PRESET_COLORS[demo.DEMO_CATEGORY_COLORS["Projects"]])
        self.assertIsNone(by_title["1:1 with Priya"]["color"])


if __name__ == "__main__":
    unittest.main()
