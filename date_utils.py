"""Date parsing helpers used across the app."""

from __future__ import annotations

import re
from datetime import datetime
from typing import Optional


def _expand_two_digit_year(two_digit_year: int) -> int:
    """Convert a two-digit year into a four-digit year with a rolling pivot.

    Uses a pivot of current year + 20 years (capped at 99) to keep future-dated
    assets in the 2000s while sending older archival dates to the 1900s.
    """
    pivot = min(99, (datetime.now().year % 100) + 20)
    return 2000 + two_digit_year if two_digit_year <= pivot else 1900 + two_digit_year


def normalize_date_string(date_str: str) -> str:
    """Normalize common slash-separated dates to MM/D/YYYY with a four-digit year.

    Returns the original string when the input doesn't look like a slash date
    or cannot be parsed into a valid calendar date.
    """
    if not date_str:
        return date_str

    if not isinstance(date_str, str):
        date_str = str(date_str)

    match = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", date_str)
    if not match:
        return date_str.strip()

    month, day, year = (int(part) for part in match.groups())
    if year < 100:
        year = _expand_two_digit_year(year)

    try:
        parsed = datetime(year, month, day)
    except ValueError:
        return date_str.strip()

    return parsed.strftime("%m/%d/%Y")


def parse_flexible_date(date_str: str) -> Optional[datetime]:
    """Parse a date string that might use two-digit years or ISO formats.

    Returns None when parsing fails.
    """
    if not date_str:
        return None

    raw = str(date_str).strip()
    normalized = raw

    # If we have a slash-style date with optional time, normalize the date portion
    if "/" in raw:
        date_part, *rest = raw.split(None, 1)
        normalized_date = normalize_date_string(date_part)
        normalized = f"{normalized_date} {rest[0]}" if rest else normalized_date

    formats = [
        "%m/%d/%Y",          # Slash dates
        "%m/%d/%Y %I:%M %p", # Slash dates with 12-hour time
        "%m/%d/%Y %H:%M:%S", # Slash dates with 24-hour time
        "%Y-%m-%d",          # ISO date
        "%Y-%m-%dT%H:%M:%S", # ISO with "T"
        "%Y-%m-%d %H:%M:%S", # ISO with space
    ]

    for fmt in formats:
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(normalized.replace("Z", "+00:00"))
    except Exception:
        return None
