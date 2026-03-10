"""
Integration tests that exercise the full end-to-end flow:
  add asset → search → update → verify in SQLite → delete → verify gone.

These tests use a real (temp file) SQLite database opened directly via
sqlite3 for cross-checks, validating that the ORM layer and raw DB are
always in sync.
"""

import contextlib
import csv
import os
import sqlite3
import tempfile
import unittest
from datetime import datetime

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from asset_database import AssetDatabase


HEADERS = [
    "Asset No.", "*Asset Type", "*Manufacturer", "*Model",
    "Serial Number", "Status", "Location", "Room", "Cubicle",
    "Notes", "IP Address", "MAC Address",
]


def _make_db() -> AssetDatabase:
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    tmp = tempfile.mkdtemp()
    tmpl = os.path.join(tmp, "tmpl.csv")
    with open(tmpl, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(HEADERS)
    db = AssetDatabase.__new__(AssetDatabase)
    db.db_path = path
    db.ensure_database_exists(tmpl)
    db._test_template = tmpl
    db._test_tmp = tmp
    return db


class TestEndToEndWorkflow(unittest.TestCase):

    def setUp(self):
        self.db = _make_db()

    def tearDown(self):
        import shutil
        os.remove(self.db.db_path)
        shutil.rmtree(self.db._test_tmp, ignore_errors=True)

    # ---- Add → Get ---------------------------------------------------------

    def test_add_and_retrieve(self):
        asset_id = self.db.add_asset({
            "manufacturer": "Dell",
            "model": "XPS 15",
            "serial_number": "E2E-001",
            "status": "Active",
            "location": "HQ",
        })
        asset = self.db.get_asset_by_id(asset_id)
        self.assertIsNotNone(asset)
        self.assertEqual(asset["serial_number"], "E2E-001")
        self.assertEqual(asset["location"], "HQ")

    # ---- Add → SQLite direct verify ----------------------------------------

    def test_add_persisted_raw_sqlite(self):
        asset_id = self.db.add_asset({
            "manufacturer": "Lenovo",
            "model": "ThinkPad",
            "serial_number": "E2E-RAW",
        })
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row["serial_number"], "E2E-RAW")
        self.assertEqual(row["manufacturer"], "Lenovo")

    # ---- Update → Get → SQLite verify --------------------------------------

    def test_update_reflected_everywhere(self):
        asset_id = self.db.add_asset({
            "manufacturer": "HP",
            "model": "EliteBook",
            "serial_number": "E2E-UPDATE",
            "status": "Active",
        })
        self.db.update_asset(asset_id, {"status": "Retired", "location": "Archive"})

        # Verify via ORM
        asset = self.db.get_asset_by_id(asset_id)
        self.assertEqual(asset["status"], "Retired")
        self.assertEqual(asset["location"], "Archive")

        # Verify via raw SQLite
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT status, location FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertEqual(row["status"], "Retired")
        self.assertEqual(row["location"], "Archive")

    # ---- Delete → Not findable ---------------------------------------------

    def test_delete_removes_from_db(self):
        asset_id = self.db.add_asset({"serial_number": "E2E-DEL"})
        self.db.delete_asset(asset_id)

        self.assertIsNone(self.db.get_asset_by_id(asset_id))
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            row = conn.execute("SELECT id FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertIsNone(row)

    # ---- Full lifecycle: add → update × 2 → delete → audit log check -------

    def test_full_lifecycle_audit_trail(self):
        asset_id = self.db.add_asset(
            {"manufacturer": "Apple", "model": "MacBook Pro", "serial_number": "E2E-LIFE"},
            changed_by="alice",
        )
        self.db.update_asset(asset_id, {"status": "In Repair"}, changed_by="bob")
        self.db.update_asset(asset_id, {"status": "Active"}, changed_by="alice")
        self.db.delete_asset(asset_id, changed_by="admin")

        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            rows = conn.execute(
                "SELECT action, changed_by FROM asset_audit_log WHERE asset_id=? ORDER BY id",
                (asset_id,),
            ).fetchall()

        actions = [r[0] for r in rows]
        self.assertIn("INSERT", actions)
        self.assertIn("UPDATE", actions)
        self.assertIn("DELETE", actions)

        changed_by_values = {r[1] for r in rows}
        self.assertIn("alice", changed_by_values)
        self.assertIn("bob", changed_by_values)
        self.assertIn("admin", changed_by_values)

    # ---- Search after multiple inserts -------------------------------------

    def test_search_after_multiple_inserts(self):
        self.db.add_asset({"manufacturer": "Dell", "status": "Active", "location": "NYC"})
        self.db.add_asset({"manufacturer": "HP", "status": "Retired", "location": "NYC"})
        self.db.add_asset({"manufacturer": "Apple", "status": "Active", "location": "SF"})

        active = self.db.search_assets({"status": "Active"})
        self.assertEqual(len(active), 2)
        for a in active:
            self.assertEqual(a["status"], "Active")

    # ---- Label request lifecycle -------------------------------------------

    def test_label_request_lifecycle(self):
        asset_id = self.db.add_asset({"manufacturer": "Dell", "serial_number": "LBL-001"})

        # Initially no label date
        asset = self.db.get_asset_by_id(asset_id)
        # label_requested_date defaults to NULL
        self.assertFalse(
            asset.get("label_requested_date") and "1901" not in str(asset.get("label_requested_date", ""))
        )

        # Request label
        before = datetime.now().isoformat()
        self.db.request_label(asset_id)
        after = datetime.now().isoformat()

        asset = self.db.get_asset_by_id(asset_id)
        lrd = asset.get("label_requested_date", "")
        self.assertIsNotNone(lrd)
        self.assertGreaterEqual(lrd, before)
        self.assertLessEqual(lrd, after)

        # Verify in raw SQLite
        with contextlib.closing(sqlite3.connect(self.db.db_path)) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT label_requested_date FROM assets WHERE id=?", (asset_id,)).fetchone()
        self.assertIsNotNone(row["label_requested_date"])

    # ---- Multiple assets, no cross-contamination ---------------------------

    def test_updates_isolated_to_target_asset(self):
        id1 = self.db.add_asset({"serial_number": "ISO-1", "status": "Active"})
        id2 = self.db.add_asset({"serial_number": "ISO-2", "status": "Active"})

        self.db.update_asset(id1, {"status": "Retired"})

        asset2 = self.db.get_asset_by_id(id2)
        self.assertEqual(asset2["status"], "Active")

    # ---- Import → verify ---------------------------------------------------

    def test_csv_import_roundtrip(self):
        """Import a CSV, then verify each row is retrievable by serial number."""
        import_path = os.path.join(self.db._test_tmp, "roundtrip.csv")
        rows = [
            {"Asset No.": "RT1", "*Asset Type": "Laptop", "*Manufacturer": "Dell",
             "*Model": "XPS", "Serial Number": "RT-S1", "Status": "Active",
             "Location": "NYC", "Room": "", "Cubicle": "", "Notes": "",
             "IP Address": "", "MAC Address": ""},
            {"Asset No.": "RT2", "*Asset Type": "Desktop", "*Manufacturer": "HP",
             "*Model": "Z6", "Serial Number": "RT-S2", "Status": "Active",
             "Location": "LA", "Room": "", "Cubicle": "", "Notes": "",
             "IP Address": "", "MAC Address": ""},
        ]
        with open(import_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=HEADERS)
            w.writeheader()
            w.writerows(rows)

        count = self.db.import_csv_template(import_path)
        self.assertEqual(count, 2)

        a1 = self.db.get_asset_by_serial("RT-S1")
        a2 = self.db.get_asset_by_serial("RT-S2")
        self.assertIsNotNone(a1)
        self.assertIsNotNone(a2)
        self.assertEqual(a1["manufacturer"], "Dell")
        self.assertEqual(a2["manufacturer"], "HP")


if __name__ == "__main__":
    unittest.main()
