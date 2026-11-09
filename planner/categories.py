"""Outlook category colours.

Outlook gives each category one of 25 preset colours, numbered 1 (red) to 25 (dark maroon), with 0 for
none (OlCategoryColor). Microsoft doesn't publish the exact shades, so these are close matches to
the ones Outlook shows. An event takes the colour of its first category that has one, as in Outlook.
"""
from __future__ import annotations

import re

PRESET_COLORS = {
    1: "#E7A1A2",   # red
    2: "#F9BA89",   # orange
    3: "#F7DD8F",   # peach
    4: "#FCFA90",   # yellow
    5: "#78D168",   # green
    6: "#9FDCC9",   # teal
    7: "#C6D2B0",   # olive
    8: "#9DB7E8",   # blue
    9: "#B5A1E2",   # purple
    10: "#DAAEC2",  # maroon
    11: "#DAD9DC",  # steel
    12: "#6B7994",  # dark steel
    13: "#BFBFBF",  # grey
    14: "#6F6F6F",  # dark grey
    15: "#4F4F4F",  # black
    16: "#C11A25",  # dark red
    17: "#E2620D",  # dark orange
    18: "#C79930",  # dark peach
    19: "#B9B300",  # dark yellow
    20: "#368F2B",  # dark green
    21: "#329B7A",  # dark teal
    22: "#778B45",  # dark olive
    23: "#2858A5",  # dark blue
    24: "#5C3FA3",  # dark purple
    25: "#93446B",  # dark maroon
}


def split_categories(text) -> list:
    """Outlook keeps an item's categories as one string, separated by commas (or semicolons in some locales)."""
    if isinstance(text, (list, tuple)):
        parts = text
    else:
        parts = re.split(r"[,;]", str(text or ""))
    seen, out = set(), []
    for p in parts:
        name = str(p).strip()
        if name and name.lower() not in seen:
            seen.add(name.lower())
            out.append(name)
    return out


def category_color(names: list, colors: dict | None):
    """The hex colour of the first category with a preset colour, or None. Names match regardless of case."""
    if not names or not colors:
        return None
    lookup = {str(k).lower(): v for k, v in colors.items()}
    for name in names:
        try:
            hexcol = PRESET_COLORS.get(int(lookup.get(name.lower(), 0)))
        except (TypeError, ValueError):
            hexcol = None
        if hexcol:
            return hexcol
    return None
