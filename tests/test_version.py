"""The version number follows semantic versioning and matches the newest release in CHANGELOG.md."""
import re
import unittest
from pathlib import Path

from planner import VERSION

SEMVER = re.compile(r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$")
CHANGELOG = Path(__file__).resolve().parent.parent / "CHANGELOG.md"


def released_versions():
    return re.findall(r"^## \[(\d[^\]]*)\]", CHANGELOG.read_text(encoding="utf-8"), re.MULTILINE)


class VersionTests(unittest.TestCase):
    def test_version_is_semver(self):
        self.assertRegex(VERSION, SEMVER)

    def test_changelog_has_an_unreleased_section_first(self):
        first = re.search(r"^## \[([^\]]+)\]", CHANGELOG.read_text(encoding="utf-8"), re.MULTILINE)
        self.assertEqual(first.group(1), "Unreleased")

    def test_version_matches_the_newest_changelog_release(self):
        self.assertEqual(released_versions()[0], VERSION,
                         "update VERSION in planner/__init__.py and CHANGELOG.md together")

    def test_every_release_has_a_compare_link(self):
        text = CHANGELOG.read_text(encoding="utf-8")
        for v in released_versions():
            self.assertRegex(text, rf"(?m)^\[{re.escape(v)}\]: https://")


if __name__ == "__main__":
    unittest.main()
