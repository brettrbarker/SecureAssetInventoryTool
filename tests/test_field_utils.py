"""
Tests for field_utils: compute_db_fields_from_template and compute_dropdown_fields.
Uses a lightweight fake database and config so no real SQLite file is needed.
"""

import csv
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from field_utils import compute_db_fields_from_template, compute_dropdown_fields


# ---------------------------------------------------------------------------
# Fake objects
# ---------------------------------------------------------------------------

class FakeDB:
    """Minimal AssetDatabase stand-in."""

    def __init__(self, columns: list, mapping: dict):
        self._columns = columns
        self._mapping = mapping

    def get_table_columns(self):
        return self._columns

    def get_dynamic_column_mapping(self, template_path=None):
        return self._mapping

    def _column_to_header(self, col):
        return col.replace("_", " ").title()


class FakeConfig:
    def __init__(self, template_path="", excluded_fields=None, dropdown_fields=None):
        self.default_template_path = template_path
        self.excluded_fields = excluded_fields or []
        self.dropdown_fields = dropdown_fields or []


# ---------------------------------------------------------------------------
# Tests: compute_db_fields_from_template
# ---------------------------------------------------------------------------

class TestComputeDbFieldsFromTemplate(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        headers = ["*Asset Type", "*Manufacturer", "*Model", "Serial Number", "Notes"]
        self.template_path = os.path.join(self.tmp.name, "tmpl.csv")
        with open(self.template_path, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(headers)

        table_columns = [
            "id", "created_date", "modified_date", "modified_by", "created_by",
            "is_deleted", "asset_type", "manufacturer", "model", "serial_number", "notes"
        ]
        mapping = {
            "*Asset Type": "asset_type",
            "*Manufacturer": "manufacturer",
            "*Model": "model",
            "Serial Number": "serial_number",
            "Notes": "notes",
        }
        self.db = FakeDB(table_columns, mapping)

    def tearDown(self):
        self.tmp.cleanup()

    def test_returns_list_of_dicts(self):
        config = FakeConfig(self.template_path)
        fields = compute_db_fields_from_template(self.db, config)
        self.assertIsInstance(fields, list)
        self.assertTrue(all("db_name" in f and "display_name" in f for f in fields))

    def test_system_columns_excluded(self):
        config = FakeConfig(self.template_path)
        fields = compute_db_fields_from_template(self.db, config)
        db_names = [f["db_name"] for f in fields]
        for sys_col in ("id", "created_date", "modified_date", "is_deleted"):
            self.assertNotIn(sys_col, db_names)

    def test_excluded_fields_not_returned(self):
        config = FakeConfig(self.template_path, excluded_fields=["Notes"])
        fields = compute_db_fields_from_template(self.db, config)
        display_names = [f["display_name"] for f in fields]
        self.assertNotIn("Notes", display_names)

    def test_template_fields_present(self):
        config = FakeConfig(self.template_path)
        fields = compute_db_fields_from_template(self.db, config)
        db_names = [f["db_name"] for f in fields]
        self.assertIn("manufacturer", db_names)
        self.assertIn("serial_number", db_names)

    def test_column_not_in_table_skipped(self):
        """Fields in the template that have no matching DB column are dropped."""
        mapping_with_missing = dict(self.db._mapping)
        mapping_with_missing["Ghost Field"] = "ghost_field"
        db2 = FakeDB(self.db._columns, mapping_with_missing)

        headers = ["*Asset Type", "*Manufacturer", "Ghost Field"]
        tp2 = os.path.join(self.tmp.name, "tmpl2.csv")
        with open(tp2, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(headers)

        config = FakeConfig(tp2)
        fields = compute_db_fields_from_template(db2, config)
        db_names = [f["db_name"] for f in fields]
        self.assertNotIn("ghost_field", db_names)


# ---------------------------------------------------------------------------
# Tests: compute_dropdown_fields
# ---------------------------------------------------------------------------

class TestComputeDropdownFields(unittest.TestCase):

    def _db_fields(self):
        return [
            {"db_name": "asset_type", "display_name": "*Asset Type"},
            {"db_name": "manufacturer", "display_name": "*Manufacturer"},
            {"db_name": "serial_number", "display_name": "Serial Number"},
            {"db_name": "location", "display_name": "Location"},
        ]

    def test_returns_only_configured_dropdowns(self):
        config = FakeConfig(dropdown_fields=["*Asset Type", "*Manufacturer"])
        result = compute_dropdown_fields(self._db_fields(), config)
        display_names = [f["display_name"] for f in result]
        self.assertIn("*Asset Type", display_names)
        self.assertIn("*Manufacturer", display_names)
        self.assertNotIn("Serial Number", display_names)

    def test_field_not_in_db_fields_skipped(self):
        config = FakeConfig(dropdown_fields=["*Asset Type", "Nonexistent"])
        result = compute_dropdown_fields(self._db_fields(), config)
        display_names = [f["display_name"] for f in result]
        self.assertNotIn("Nonexistent", display_names)

    def test_empty_dropdown_config_returns_empty(self):
        config = FakeConfig(dropdown_fields=[])
        result = compute_dropdown_fields(self._db_fields(), config)
        self.assertEqual(result, [])

    def test_all_fields_as_dropdown(self):
        all_names = ["*Asset Type", "*Manufacturer", "Serial Number", "Location"]
        config = FakeConfig(dropdown_fields=all_names)
        result = compute_dropdown_fields(self._db_fields(), config)
        self.assertEqual(len(result), 4)


if __name__ == "__main__":
    unittest.main()
