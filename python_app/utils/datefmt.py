"""Helpers to format and parse dates consistently across the app."""
from __future__ import annotations

from datetime import date, datetime, timedelta

DISPLAY_DATE_FORMAT = "%d/%m/%Y"
_INPUT_FORMATS = (
    DISPLAY_DATE_FORMAT,
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d.%m.%Y",
)


def format_display_date(value: date | datetime | None) -> str:
    """Return a ``JJ/MM/AAAA`` string or ``""`` when ``value`` is falsy."""

    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value.strftime(DISPLAY_DATE_FORMAT)
    return ""


def parse_date_value(value) -> date | None:
    """Best-effort parsing for user-entered or Excel-style dates."""

    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(value))
    text = str(value).strip()
    if not text:
        return None
    for fmt in _INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        serial = float(text)
    except ValueError:
        return None
    return date(1899, 12, 30) + timedelta(days=int(serial))


__all__ = ["DISPLAY_DATE_FORMAT", "format_display_date", "parse_date_value"]
