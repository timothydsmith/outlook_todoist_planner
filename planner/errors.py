"""Errors worded for the person using the app."""
from __future__ import annotations



class SourceError(Exception):
    """A problem with Todoist or the calendar, worded for the person using the app."""


def _reason(exc: Exception) -> str:
    r = getattr(exc, "reason", None) or exc
    return str(r)[:160]
