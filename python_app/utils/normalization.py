"""Shared normalization helpers."""

from __future__ import annotations

import math


def normalize_integer_value(value) -> int | None:
    """Convert any spreadsheet-like value to a clean integer.

    Returns ``None`` for blanks, booleans, non-numeric values, and non-integer
    numbers. Floats that represent whole numbers (e.g., ``3.0``) are accepted.
    """

    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            return None
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(text)
    except (TypeError, ValueError):
        try:
            number = float(text)
        except (TypeError, ValueError):
            return None
        if not number.is_integer():
            return None
        return int(number)


__all__ = ["normalize_integer_value"]
