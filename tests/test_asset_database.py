"""
Tests for AssetDatabase CRUD operations, schema management, and audit logging.
All tests use an isolated in-memory (or temp-file) SQLite database so they
never touch the real database.
"""

import contextlib
import csv
import os
import sqlite3
import tempfile
import unittest
from datetime import datetime

# ---------------------------------------------------------------------------
# Path bootstrap – allow running from repo root or from tests/
# ---------------------------------------------------------------------------
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from asset_database import AssetDatabase


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_db(template_path: str = None) -> AssetDatabase:
    """Return an AssetDatabase backed by a fresh temp file."""
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    db = AssetDatabase.__new__(AssetDatabase)
    db.db_path = path
    db.ensure_database_exists(template_path)
    return db


def _make_template(headers: list, tmp_dir: str = None) -> str:
    """Write a minimal CSV template and return its path."""
    if tmp_dir is None:
        tmp_dir = tempfile.gettempdir()
    path = os.path.join(tmp_dir, "test_template.csv")
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(headers)
    return path


STANDARD_HEADERS = [
    "Asset No.", "*Asset Type", "*Manufacturer", "*Model",
    "Serial Number", "Status", "Location", "Room", "Cubicle",
    "Notes", "IP Address", "MAC Address",
]


# ---------------------------------------------------------------------------
# Schema & initialisation tests
# ---------------------------------------------------------------------------

class TestDatabaseInit(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        try:
            os.remove(self.db.db_path)
        except FileNotFoundError:
            pass
        self.tmp.cleanup()

    def test_assets_table_created(self):
        tables = self.db.get_database_tables()
        self.assertIn("assets", tables)

    def test_audit_log_table_created(self):
        tables = self.db.get_database_tables()
        self.assertIn("asset_audit_log", tables)

    def test_system_columns_present(self):
        cols = self.db.get_table_columns()
        for col in ("id", "created_date", "modified_date", "is_deleted",
                    "created_by", "modified_by", "data_source"):
            self.assertIn(col, cols, f"System column '{col}' missing")

    def test_template_columns_present(self):
        cols = self.db.get_table_columns()
        # "Serial Number" -> "serial_number"
        self.assertIn("serial_number", cols)
        self.assertIn("location", cols)
        self.assertIn("notes", cols)

    def test_idempotent_init(self):
        """Calling ensure_database_exists twice must not raise or corrupt."""
        self.db.ensure_database_exists(self.template)
        tables = self.db.get_database_tables()
        self.assertIn("assets", tables)


# ---------------------------------------------------------------------------
# CRUD – add_asset / get_asset_by_id
# ---------------------------------------------------------------------------

class TestAddAsset(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def _sample_asset(self, **overrides):
        data = {
            "asset_type": "Laptop",
            "manufacturer": "Dell",
            "model": "XPS 15",
            "serial_number": "SN-001",
            "status": "Active",
            "location": "Building A",
        }
        data.update(overrides)
        return data

    # ---- basic insert -------------------------------------------------------

    def test_add_returns_integer_id(self):
        asset_id = self.db.add_asset(self._sample_asset())
        self.assertIsInstance(asset_id, int)
        self.assertGreater(asset_id, 0)

    def test_get_by_id_returns_correct_data(self):
        asset_id = self.db.add_asset(self._sample_asset(serial_number="SN-ABC"))
        result = self.db.get_asset_by_id(asset_id)
        self.assertIsNotNone(result)
        self.assertEqual(result["serial_number"], "SN-ABC")
        self.assertEqual(result["manufacturer"], "Dell")

    def test_get_by_id_sqlite_directly(self):
        """Verify data persisted correctly by querying SQLite outside the ORM."""
        asset_id = self.db.add_asset(self._sample_asset(serial_number="DIRECT-CHECK"))
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row["serial_number"], "DIRECT-CHECK")

    def test_created_date_populated(self):
        before = datetime.now().isoformat()
        asset_id = self.db.add_asset(self._sample_asset())
        after = datetime.now().isoformat()
        result = self.db.get_asset_by_id(asset_id)
        self.assertIsNotNone(result["created_date"])
        self.assertGreaterEqual(result["created_date"], before)
        self.assertLessEqual(result["created_date"], after)

    def test_modified_date_defaults_to_sentinel(self):
        asset_id = self.db.add_asset(self._sample_asset())
        result = self.db.get_asset_by_id(asset_id)
        # Default modified_date should be the 1901 sentinel, not current time
        self.assertIn("1901", str(result["modified_date"]))

    def test_auto_asset_number_assigned(self):
        asset_id = self.db.add_asset(self._sample_asset())
        result = self.db.get_asset_by_id(asset_id)
        self.assertIsNotNone(result.get("asset_no"))
        self.assertTrue(str(result["asset_no"]).startswith("AST"))

    def test_explicit_asset_number_preserved(self):
        asset_id = self.db.add_asset(self._sample_asset(asset_no="MYTAG-999"))
        result = self.db.get_asset_by_id(asset_id)
        self.assertEqual(result["asset_no"], "MYTAG-999")

    def test_multiple_assets_get_unique_ids(self):
        id1 = self.db.add_asset(self._sample_asset(serial_number="S1"))
        id2 = self.db.add_asset(self._sample_asset(serial_number="S2"))
        self.assertNotEqual(id1, id2)

    def test_data_source_recorded(self):
        asset_id = self.db.add_asset(self._sample_asset(), data_source="import")
        result = self.db.get_asset_by_id(asset_id)
        self.assertEqual(result["data_source"], "import")

    # ---- get_asset_by_serial ------------------------------------------------

    def test_get_by_serial_returns_asset(self):
        self.db.add_asset(self._sample_asset(serial_number="SERIAL-XYZ"))
        result = self.db.get_asset_by_serial("SERIAL-XYZ")
        self.assertIsNotNone(result)
        self.assertEqual(result["serial_number"], "SERIAL-XYZ")

    def test_get_by_serial_returns_none_for_missing(self):
        result = self.db.get_asset_by_serial("DOES-NOT-EXIST")
        self.assertIsNone(result)

    # ---- get_asset_by_id edge cases -----------------------------------------

    def test_get_by_id_returns_none_for_missing(self):
        self.assertIsNone(self.db.get_asset_by_id(999999))


# ---------------------------------------------------------------------------
# CRUD – update_asset
# ---------------------------------------------------------------------------

class TestUpdateAsset(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)
        self.asset_id = self.db.add_asset({
            "manufacturer": "HP",
            "model": "EliteBook",
            "serial_number": "HP-001",
            "status": "Active",
        })

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_update_returns_true(self):
        result = self.db.update_asset(self.asset_id, {"status": "Retired"})
        self.assertTrue(result)

    def test_update_value_persisted(self):
        self.db.update_asset(self.asset_id, {"status": "Retired", "location": "Warehouse"})
        updated = self.db.get_asset_by_id(self.asset_id)
        self.assertEqual(updated["status"], "Retired")
        self.assertEqual(updated["location"], "Warehouse")

    def test_update_sqlite_directly(self):
        self.db.update_asset(self.asset_id, {"model": "ProBook"})
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT model FROM assets WHERE id=?", (self.asset_id,)).fetchone()
        self.assertEqual(row["model"], "ProBook")

    def test_update_sets_modified_date(self):
        before = datetime.now().isoformat()
        self.db.update_asset(self.asset_id, {"status": "In Repair"})
        result = self.db.get_asset_by_id(self.asset_id)
        self.assertGreaterEqual(result["modified_date"], before)

    def test_update_does_not_change_created_date(self):
        original = self.db.get_asset_by_id(self.asset_id)["created_date"]
        self.db.update_asset(self.asset_id, {"status": "Storage"})
        updated = self.db.get_asset_by_id(self.asset_id)
        self.assertEqual(updated["created_date"], original)

    def test_update_nonexistent_asset_returns_false(self):
        result = self.db.update_asset(99999, {"status": "Active"})
        self.assertFalse(result)

    def test_update_multiple_fields_at_once(self):
        self.db.update_asset(self.asset_id, {
            "status": "Decommissioned",
            "location": "Storage Room",
            "model": "EliteBook G9",
        })
        a = self.db.get_asset_by_id(self.asset_id)
        self.assertEqual(a["status"], "Decommissioned")
        self.assertEqual(a["location"], "Storage Room")
        self.assertEqual(a["model"], "EliteBook G9")


# ---------------------------------------------------------------------------
# CRUD – delete_asset
# ---------------------------------------------------------------------------

class TestDeleteAsset(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_delete_returns_true(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell", "model": "T40"})
        self.assertTrue(self.db.delete_asset(asset_id))

    def test_deleted_asset_not_retrievable(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell", "model": "T40"})
        self.db.delete_asset(asset_id)
        self.assertIsNone(self.db.get_asset_by_id(asset_id))

    def test_deleted_row_absent_in_sqlite(self):
        asset_id = self.db.add_asset({"manufacturer": "Lenovo", "model": "ThinkPad"})
        self.db.delete_asset(asset_id)
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            row = conn.execute("SELECT id FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertIsNone(row)

    def test_delete_nonexistent_returns_false(self):
        self.assertFalse(self.db.delete_asset(99999))

    def test_delete_does_not_affect_other_assets(self):
        id1 = self.db.add_asset({"serial_number": "S1"})
        id2 = self.db.add_asset({"serial_number": "S2"})
        self.db.delete_asset(id1)
        self.assertIsNotNone(self.db.get_asset_by_id(id2))


# ---------------------------------------------------------------------------
# search_assets
# ---------------------------------------------------------------------------

class TestSearchAssets(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

        self.db.add_asset({"manufacturer": "Dell", "model": "XPS", "location": "NYC", "status": "Active"})
        self.db.add_asset({"manufacturer": "HP", "model": "EliteBook", "location": "NYC", "status": "Retired"})
        self.db.add_asset({"manufacturer": "Apple", "model": "MacBook", "location": "LA", "status": "Active"})

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_search_no_filters_returns_all(self):
        results = self.db.search_assets()
        self.assertEqual(len(results), 3)

    def test_search_by_status_filter(self):
        results = self.db.search_assets({"status": "Active"})
        self.assertEqual(len(results), 2)
        for r in results:
            self.assertEqual(r["status"], "Active")

    def test_search_by_location_partial_match(self):
        # location uses LIKE internally
        results = self.db.search_assets({"location": "NYC"})
        self.assertEqual(len(results), 2)

    def test_search_returns_no_results_for_unknown(self):
        results = self.db.search_assets({"status": "DOES_NOT_EXIST"})
        self.assertEqual(len(results), 0)

    def test_search_result_contains_all_fields(self):
        results = self.db.search_assets({"manufacturer": "Dell"})
        self.assertEqual(len(results), 1)
        r = results[0]
        self.assertIn("manufacturer", r)
        self.assertIn("model", r)
        self.assertIn("id", r)

    def test_search_assets_by_field(self):
        results = self.db.search_assets_by_field("manufacturer", "Dell")
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].manufacturer, "Dell")

    def test_search_assets_by_field_partial(self):
        # "Book" matches EliteBook and MacBook
        results = self.db.search_assets_by_field("model", "Book")
        self.assertEqual(len(results), 2)

    def test_search_assets_by_field_case_insensitive_partial(self):
        results = self.db.search_assets_by_field("manufacturer", "dell")
        self.assertEqual(len(results), 1)


# ---------------------------------------------------------------------------
# Unique values
# ---------------------------------------------------------------------------

class TestUniqueValues(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)
        self.db.add_asset({"manufacturer": "Dell", "status": "Active"})
        self.db.add_asset({"manufacturer": "HP", "status": "Active"})
        self.db.add_asset({"manufacturer": "Dell", "status": "Retired"})

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_get_unique_values_manufacturer(self):
        values = self.db.get_unique_values("manufacturer")
        self.assertEqual(sorted(values), ["Dell", "HP"])

    def test_get_unique_values_status(self):
        values = self.db.get_unique_values("status")
        self.assertEqual(sorted(values), ["Active", "Retired"])

    def test_get_unique_field_values_excludes_empty(self):
        self.db.add_asset({"manufacturer": "", "status": "Active"})
        values = self.db.get_unique_field_values("manufacturer")
        self.assertNotIn("", values)

    def test_get_unique_field_values_sorted(self):
        values = self.db.get_unique_field_values("manufacturer")
        self.assertEqual(values, sorted(values))


# ---------------------------------------------------------------------------
# Audit logging
# ---------------------------------------------------------------------------

class TestAuditLog(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_insert_creates_audit_entry(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        history = self.db.get_audit_history(asset_id)
        self.assertGreater(len(history), 0)
        self.assertEqual(history[0]["action"], "INSERT")

    def test_update_creates_audit_entry(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        self.db.update_asset(asset_id, {"status": "Retired"})
        history = self.db.get_audit_history(asset_id)
        actions = [h["action"] for h in history]
        self.assertIn("UPDATE", actions)

    def test_delete_creates_audit_entry(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        self.db.delete_asset(asset_id)
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            rows = conn.execute(
                "SELECT action FROM asset_audit_log WHERE asset_id=?", (asset_id,)
            ).fetchall()
        actions = [r[0] for r in rows]
        self.assertIn("DELETE", actions)

    def test_audit_history_ordered_desc(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        self.db.update_asset(asset_id, {"status": "A"})
        self.db.update_asset(asset_id, {"status": "B"})
        history = self.db.get_audit_history(asset_id)
        dates = [h["change_date"] for h in history]
        self.assertEqual(dates, sorted(dates, reverse=True))

    def test_audit_records_changed_by(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"}, changed_by="tester")
        self.db.update_asset(asset_id, {"status": "X"}, changed_by="tester2")
        history = self.db.get_audit_history(asset_id)
        changed_by_values = {h["changed_by"] for h in history}
        self.assertIn("tester", changed_by_values)
        self.assertIn("tester2", changed_by_values)


# ---------------------------------------------------------------------------
# Label requests
# ---------------------------------------------------------------------------

class TestLabelRequest(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_request_label_sets_date(self):
        before = datetime.now().isoformat()
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        self.db.request_label(asset_id)
        result = self.db.get_asset_by_id(asset_id)
        self.assertIsNotNone(result["label_requested_date"])
        self.assertGreaterEqual(result["label_requested_date"], before)

    def test_request_label_returns_true(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        self.assertTrue(self.db.request_label(asset_id))

    def test_label_date_persisted_in_sqlite(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell"})
        self.db.request_label(asset_id)
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT label_requested_date FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertIsNotNone(row["label_requested_date"])
        self.assertNotIn("1901", str(row["label_requested_date"]))


# ---------------------------------------------------------------------------
# Schema management
# ---------------------------------------------------------------------------

class TestSchemaManagement(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(["Asset No.", "*Manufacturer", "*Model"], self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_update_schema_adds_new_column(self):
        # Write a new template with an extra column
        extended = os.path.join(self.tmp.name, "extended.csv")
        with open(extended, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(["Asset No.", "*Manufacturer", "*Model", "New Custom Field"])
        result = self.db.update_schema_for_template(extended)
        self.assertTrue(result)
        cols = self.db.get_table_columns()
        self.assertIn("new_custom_field", cols)

    def test_update_schema_idempotent(self):
        """Running update_schema_for_template twice must not raise."""
        result1 = self.db.update_schema_for_template(self.template)
        result2 = self.db.update_schema_for_template(self.template)
        self.assertTrue(result1)
        self.assertTrue(result2)

    def test_verify_template_compatibility_existing_fields(self):
        compat = self.db.verify_template_compatibility(self.template)
        self.assertNotIn("error", compat)
        self.assertGreater(compat["mapped_fields"], 0)

    def test_verify_template_compatibility_new_fields(self):
        new_template = os.path.join(self.tmp.name, "new.csv")
        with open(new_template, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(["Asset No.", "*Manufacturer", "*Model", "Brand New Column"])
        compat = self.db.verify_template_compatibility(new_template)
        self.assertGreater(compat["new_fields"], 0)

    def test_verify_template_compatibility_missing_file(self):
        compat = self.db.verify_template_compatibility("/nonexistent/file.csv")
        self.assertIn("error", compat)

    def test_generate_safe_column_name(self):
        self.assertEqual(self.db._generate_safe_column_name("Serial Number"), "serial_number")
        self.assertEqual(self.db._generate_safe_column_name("*Asset Type"), "asset_type")
        self.assertEqual(self.db._generate_safe_column_name("IP Address"), "ip_address")

    def test_get_dynamic_column_mapping_returns_dict(self):
        mapping = self.db.get_dynamic_column_mapping(self.template)
        self.assertIsInstance(mapping, dict)
        self.assertGreater(len(mapping), 0)

    def test_get_dynamic_column_mapping_no_template(self):
        mapping = self.db.get_dynamic_column_mapping(None)
        self.assertIsInstance(mapping, dict)


# ---------------------------------------------------------------------------
# CSV import
# ---------------------------------------------------------------------------

class TestCsvImport(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.headers = ["Asset No.", "*Asset Type", "*Manufacturer", "*Model",
                        "Serial Number", "Status", "Location"]
        self.template = _make_template(self.headers, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def _write_csv(self, name, rows):
        path = os.path.join(self.tmp.name, name)
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=self.headers)
            w.writeheader()
            w.writerows(rows)
        return path

    def test_import_csv_template_count(self):
        path = self._write_csv("import.csv", [
            {"Asset No.": "A1", "*Asset Type": "Laptop", "*Manufacturer": "Dell",
             "*Model": "XPS", "Serial Number": "S1", "Status": "Active", "Location": "NYC"},
            {"Asset No.": "A2", "*Asset Type": "Desktop", "*Manufacturer": "HP",
             "*Model": "Z440", "Serial Number": "S2", "Status": "Active", "Location": "LA"},
        ])
        count = self.db.import_csv_template(path)
        self.assertEqual(count, 2)

    def test_import_csv_data_in_sqlite(self):
        path = self._write_csv("import2.csv", [
            {"Asset No.": "A3", "*Asset Type": "Server", "*Manufacturer": "Dell",
             "*Model": "R740", "Serial Number": "SRV-01", "Status": "Active", "Location": "DC"},
        ])
        self.db.import_csv_template(path)
        result = self.db.get_asset_by_serial("SRV-01")
        self.assertIsNotNone(result)
        self.assertEqual(result["model"], "R740")

    def test_import_from_csv_skips_duplicate_by_default(self):
        path = self._write_csv("dup.csv", [
            {"Asset No.": "A4", "*Asset Type": "Laptop", "*Manufacturer": "Dell",
             "*Model": "XPS", "Serial Number": "SAME-SN", "Status": "Active", "Location": "NYC"},
        ])
        self.db.import_csv_template(path)
        # Import again – no callback → skip
        count2 = self.db.import_from_csv(path)
        # Row count in DB should still be 1
        assets = self.db.search_assets()
        sns = [a.get("serial_number") for a in assets]
        self.assertEqual(sns.count("SAME-SN"), 1)

    def test_import_from_csv_overwrites_when_callback_says_overwrite(self):
        path = self._write_csv("ow.csv", [
            {"Asset No.": "A5", "*Asset Type": "Laptop", "*Manufacturer": "Dell",
             "*Model": "Original", "Serial Number": "OW-SN", "Status": "Active", "Location": "NYC"},
        ])
        self.db.import_csv_template(path)

        # Create a second CSV with updated model
        path2 = self._write_csv("ow2.csv", [
            {"Asset No.": "A5", "*Asset Type": "Laptop", "*Manufacturer": "Dell",
             "*Model": "Updated", "Serial Number": "OW-SN", "Status": "Active", "Location": "NYC"},
        ])
        self.db.import_from_csv(path2, overwrite_callback=lambda *_: "overwrite")
        result = self.db.get_asset_by_serial("OW-SN")
        self.assertEqual(result["model"], "Updated")


# ---------------------------------------------------------------------------
# Export to CSV
# ---------------------------------------------------------------------------

class TestExportCsv(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.headers = ["Asset No.", "*Manufacturer", "*Model", "Serial Number", "Location"]
        self.template = _make_template(self.headers, self.tmp.name)
        self.db = _make_db(self.template)
        self.db.add_asset({"manufacturer": "Dell", "model": "XPS", "serial_number": "E1", "location": "NYC"})
        self.db.add_asset({"manufacturer": "HP", "model": "Z6", "serial_number": "E2", "location": "LA"})

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_export_returns_correct_count(self):
        out = os.path.join(self.tmp.name, "export.csv")
        count = self.db.export_to_csv(out, template_path=self.template)
        self.assertEqual(count, 2)

    def test_export_file_exists(self):
        out = os.path.join(self.tmp.name, "export2.csv")
        self.db.export_to_csv(out, template_path=self.template)
        self.assertTrue(os.path.exists(out))

    def test_export_file_has_header_row(self):
        out = os.path.join(self.tmp.name, "export3.csv")
        self.db.export_to_csv(out, template_path=self.template)
        with open(out, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            header_row = next(reader)
        self.assertIn("*Manufacturer", header_row)

    def test_export_file_contains_data_rows(self):
        out = os.path.join(self.tmp.name, "export4.csv")
        self.db.export_to_csv(out, template_path=self.template)
        with open(out, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
        serials = [r.get("Serial Number", "") for r in rows]
        self.assertIn("E1", serials)
        self.assertIn("E2", serials)

    def test_export_returns_zero_for_empty_db(self):
        empty_db = _make_db(self.template)
        out = os.path.join(self.tmp.name, "empty_export.csv")
        count = empty_db.export_to_csv(out, template_path=self.template)
        self.assertEqual(count, 0)
        os.remove(empty_db.db_path)


# ---------------------------------------------------------------------------
# Database stats
# ---------------------------------------------------------------------------

class TestDatabaseStats(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_stats_empty_db(self):
        stats = self.db.get_database_stats()
        self.assertEqual(stats["total_assets"], 0)

    def test_stats_after_add(self):
        self.db.add_asset({"manufacturer": "Dell", "location": "NYC"})
        self.db.add_asset({"manufacturer": "HP", "location": "LA"})
        stats = self.db.get_database_stats()
        self.assertEqual(stats["total_assets"], 2)
        self.assertEqual(stats["unique_locations"], 2)

    def test_stats_include_db_path(self):
        stats = self.db.get_database_stats()
        self.assertEqual(stats["database_path"], self.db.db_path)

    def test_stats_include_audit_entries(self):
        self.db.add_asset({"manufacturer": "Dell"})
        stats = self.db.get_database_stats()
        self.assertGreater(stats["total_audit_entries"], 0)


# ---------------------------------------------------------------------------
# Unique field conflict checking
# ---------------------------------------------------------------------------

class TestUniqueFieldConflicts(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.template = _make_template(STANDARD_HEADERS, self.tmp.name)
        self.db = _make_db(self.template)
        self.db.add_asset({"serial_number": "CONFLICT-SN", "ip_address": "10.0.0.1"})

    def tearDown(self):
        os.remove(self.db.db_path)
        self.tmp.cleanup()

    def test_conflict_detected_on_serial(self):
        conflicts = self.db.check_unique_field_conflicts(
            {"serial_number": "CONFLICT-SN"},
            ["Serial Number"],
            self.template,
        )
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts[0]["field_name"], "Serial Number")

    def test_no_conflict_for_unique_serial(self):
        conflicts = self.db.check_unique_field_conflicts(
            {"serial_number": "BRAND-NEW"},
            ["Serial Number"],
            self.template,
        )
        self.assertEqual(len(conflicts), 0)

    def test_conflict_detected_on_ip(self):
        conflicts = self.db.check_unique_field_conflicts(
            {"ip_address": "10.0.0.1"},
            ["IP Address"],
            self.template,
        )
        self.assertEqual(len(conflicts), 1)


if __name__ == "__main__":
    unittest.main()
