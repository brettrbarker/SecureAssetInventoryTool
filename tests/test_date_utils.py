"""
Tests for date_utils: normalize_date_string, parse_flexible_date,
and the two-digit year pivot logic.
"""

import sys
import os
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from date_utils import normalize_date_string, parse_flexible_date, _expand_two_digit_year


class TestExpandTwoDigitYear(unittest.TestCase):

    def test_recent_year_maps_to_2000s(self):
        # 25 should be 2025 (assuming current year is 2026, pivot is ~46)
        result = _expand_two_digit_year(25)
        self.assertEqual(result, 2025)

    def test_far_old_year_maps_to_1900s(self):
        # 99 should map to 1999 for any foreseeable pivot
        result = _expand_two_digit_year(99)
        self.assertEqual(result, 1999)

    def test_zero_maps_to_2000s(self):
        result = _expand_two_digit_year(0)
        self.assertEqual(result, 2000)


class TestNormalizeDateString(unittest.TestCase):

    def test_standard_mm_dd_yyyy(self):
        self.assertEqual(normalize_date_string("03/10/2026"), "03/10/2026")

    def test_single_digit_month_and_day(self):
        result = normalize_date_string("1/5/2023")
        self.assertEqual(result, "01/05/2023")

    def test_two_digit_year_recent(self):
        # "3/10/26" should normalize to 03/10/2026
        result = normalize_date_string("3/10/26")
        self.assertEqual(result, "03/10/2026")

    def test_non_date_string_returned_as_is(self):
        result = normalize_date_string("not-a-date")
        self.assertEqual(result, "not-a-date")

    def test_empty_string_returned_as_is(self):
        result = normalize_date_string("")
        self.assertEqual(result, "")

    def test_none_returns_none(self):
        result = normalize_date_string(None)
        self.assertIsNone(result)

    def test_strips_whitespace(self):
        result = normalize_date_string("  03/10/2026  ")
        self.assertEqual(result, "03/10/2026")

    def test_invalid_calendar_date_returned_as_is(self):
        # Month 13 is not valid
        result = normalize_date_string("13/01/2026")
        self.assertEqual(result, "13/01/2026")


class TestParseFlexibleDate(unittest.TestCase):

    def test_standard_slash_date(self):
        dt = parse_flexible_date("03/10/2026")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.year, 2026)
        self.assertEqual(dt.month, 3)
        self.assertEqual(dt.day, 10)

    def test_iso_date(self):
        dt = parse_flexible_date("2026-03-10")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.year, 2026)

    def test_iso_datetime_with_T(self):
        dt = parse_flexible_date("2026-03-10T14:30:00")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.hour, 14)

    def test_iso_datetime_with_space(self):
        dt = parse_flexible_date("2026-03-10 09:15:00")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.minute, 15)

    def test_slash_date_with_12h_time(self):
        dt = parse_flexible_date("03/10/2026 02:30 PM")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.hour, 14)

    def test_two_digit_year_in_slash_date(self):
        dt = parse_flexible_date("3/10/26")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.year, 2026)

    def test_empty_string_returns_none(self):
        self.assertIsNone(parse_flexible_date(""))

    def test_none_returns_none(self):
        self.assertIsNone(parse_flexible_date(None))

    def test_garbage_returns_none(self):
        self.assertIsNone(parse_flexible_date("not-a-date-at-all"))

    def test_sentinel_date_parseable(self):
        # The app uses "1901-01-01 00:00:00" as a sentinel
        dt = parse_flexible_date("1901-01-01 00:00:00")
        self.assertIsInstance(dt, datetime)
        self.assertEqual(dt.year, 1901)


if __name__ == "__main__":
    unittest.main()
