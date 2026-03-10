"""
Unit tests for ConfigManager and AppConfig.

Tests run against a temp directory to avoid reading or writing to the real
assets/config.json on disk. The ConfigManager uses a module-level singleton;
each test resets that singleton so tests are independent.
"""

import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config_manager import AppConfig, ConfigManager


def _reset_singleton():
    """Reset the ConfigManager singleton so the next instantiation is fresh."""
    ConfigManager._instance = None
    ConfigManager._config = None


def _make_manager(tmp_dir: str) -> ConfigManager:
    """Create a ConfigManager whose config_path points into tmp_dir."""
    _reset_singleton()
    mgr = ConfigManager()
    mgr.config_path = os.path.join(tmp_dir, "config.json")
    # Force a fresh load from the new path
    mgr._config = mgr._load_config()
    return mgr


class TestAppConfigDefaults(unittest.TestCase):

    def test_default_theme(self):
        cfg = AppConfig()
        self.assertEqual(cfg.theme, "dark")

    def test_default_database_path(self):
        cfg = AppConfig()
        self.assertEqual(cfg.database_path, "assets/asset_database.db")

    def test_default_output_directory(self):
        cfg = AppConfig()
        self.assertEqual(cfg.output_directory, "assets/output_files")

    def test_dropdown_fields_populated_by_default(self):
        cfg = AppConfig()
        self.assertIsNotNone(cfg.dropdown_fields)
        self.assertGreater(len(cfg.dropdown_fields), 0)

    def test_required_fields_populated_by_default(self):
        cfg = AppConfig()
        self.assertIsNotNone(cfg.required_fields)
        self.assertGreater(len(cfg.required_fields), 0)

    def test_unique_fields_populated_by_default(self):
        cfg = AppConfig()
        self.assertIsInstance(cfg.unique_fields, list)
        self.assertIn("Serial Number", cfg.unique_fields)

    def test_bulk_update_presets_is_dict(self):
        cfg = AppConfig()
        self.assertIsInstance(cfg.bulk_update_presets, dict)

    def test_saved_searches_empty_dict(self):
        cfg = AppConfig()
        self.assertEqual(cfg.saved_searches, {})

    def test_label_output_fields_populated(self):
        cfg = AppConfig()
        self.assertIsNotNone(cfg.label_output_fields)
        self.assertGreater(len(cfg.label_output_fields), 0)


class TestAppConfigAccessPatterns(unittest.TestCase):
    """AppConfig has backward-compat dict-like access."""

    def setUp(self):
        self.cfg = AppConfig()

    def test_get_existing_key(self):
        self.assertEqual(self.cfg.get("theme"), "dark")

    def test_get_missing_key_returns_default(self):
        self.assertIsNone(self.cfg.get("nonexistent_key"))

    def test_get_missing_key_returns_given_default(self):
        self.assertEqual(self.cfg.get("nonexistent_key", "fallback"), "fallback")

    def test_in_operator_existing(self):
        self.assertIn("theme", self.cfg)

    def test_in_operator_missing(self):
        self.assertNotIn("totally_fake_field_xyz", self.cfg)

    def test_bracket_read(self):
        self.assertEqual(self.cfg["theme"], "dark")

    def test_bracket_write(self):
        self.cfg["theme"] = "light"
        self.assertEqual(self.cfg.theme, "light")

    def test_bracket_read_missing_raises_key_error(self):
        with self.assertRaises(KeyError):
            _ = self.cfg["definitely_not_a_field"]

    def test_to_dict_returns_dict(self):
        d = self.cfg.to_dict()
        self.assertIsInstance(d, dict)
        self.assertIn("theme", d)

    def test_to_dict_contains_all_fields(self):
        d = self.cfg.to_dict()
        for field in ("theme", "database_path", "output_directory", "dropdown_fields"):
            self.assertIn(field, d)


class TestAppConfigConstruction(unittest.TestCase):

    def test_construct_with_overrides(self):
        cfg = AppConfig(theme="light", database_path="/custom/db.db")
        self.assertEqual(cfg.theme, "light")
        self.assertEqual(cfg.database_path, "/custom/db.db")

    def test_construct_with_custom_dropdown_fields(self):
        custom = ["FieldA", "FieldB"]
        cfg = AppConfig(dropdown_fields=custom)
        self.assertEqual(cfg.dropdown_fields, custom)

    def test_construct_from_dict_roundtrip(self):
        original = AppConfig(theme="light")
        d = original.to_dict()
        reconstructed = AppConfig(**d)
        self.assertEqual(reconstructed.theme, "light")
        self.assertEqual(reconstructed.database_path, original.database_path)


class TestConfigManagerLoadSave(unittest.TestCase):

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp_dir, ignore_errors=True)
        _reset_singleton()

    def test_load_creates_default_when_no_file(self):
        mgr = _make_manager(self.tmp_dir)
        cfg = mgr.get_config()
        self.assertIsNotNone(cfg)
        self.assertEqual(cfg.theme, "dark")

    def test_save_creates_file(self):
        mgr = _make_manager(self.tmp_dir)
        result = mgr.save_config()
        self.assertTrue(result)
        self.assertTrue(os.path.exists(mgr.config_path))

    def test_save_then_load_roundtrip(self):
        mgr = _make_manager(self.tmp_dir)
        mgr.get_config().theme = "light"
        mgr.save_config()

        # Fresh manager from same path
        _reset_singleton()
        mgr2 = ConfigManager()
        mgr2.config_path = mgr.config_path
        mgr2._config = mgr2._load_config()

        self.assertEqual(mgr2.get_config().theme, "light")

    def test_save_writes_valid_json(self):
        mgr = _make_manager(self.tmp_dir)
        mgr.save_config()
        with open(mgr.config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertIn("theme", data)
        self.assertIn("database_path", data)

    def test_load_restores_custom_fields(self):
        data = AppConfig(theme="light", database_path="/tmp/test.db").to_dict()
        cfg_path = os.path.join(self.tmp_dir, "config.json")
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(data, f)

        mgr = _make_manager(self.tmp_dir)
        mgr.config_path = cfg_path
        mgr._config = mgr._load_config()

        self.assertEqual(mgr.get_config().theme, "light")
        self.assertEqual(mgr.get_config().database_path, "/tmp/test.db")

    def test_load_falls_back_to_default_on_corrupt_json(self):
        cfg_path = os.path.join(self.tmp_dir, "config.json")
        with open(cfg_path, "w", encoding="utf-8") as f:
            f.write("NOT VALID JSON {{{")

        mgr = _make_manager(self.tmp_dir)
        mgr.config_path = cfg_path
        mgr._config = mgr._load_config()

        # Should still return a valid AppConfig with defaults
        cfg = mgr.get_config()
        self.assertIsNotNone(cfg)
        self.assertEqual(cfg.theme, "dark")

    def test_save_returns_false_on_bad_path(self):
        mgr = _make_manager(self.tmp_dir)
        # Force the file write to raise OSError regardless of platform
        with patch("builtins.open", side_effect=OSError("mocked write error")):
            result = mgr.save_config()
        self.assertFalse(result)


class TestConfigManagerUpdateConfig(unittest.TestCase):

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.mgr = _make_manager(self.tmp_dir)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp_dir, ignore_errors=True)
        _reset_singleton()

    def test_update_config_single_key(self):
        self.mgr.update_config(theme="light")
        self.assertEqual(self.mgr.get_config().theme, "light")

    def test_update_config_multiple_keys(self):
        self.mgr.update_config(theme="light", database_path="/tmp/x.db")
        self.assertEqual(self.mgr.get_config().theme, "light")
        self.assertEqual(self.mgr.get_config().database_path, "/tmp/x.db")

    def test_update_config_unknown_key_ignored(self):
        # Should not raise for an unknown key
        try:
            self.mgr.update_config(totally_fake_field="value")
        except Exception as exc:
            self.fail(f"update_config raised unexpectedly: {exc}")

    def test_get_database_path_helper(self):
        self.mgr.update_config(database_path="/db/custom.db")
        self.assertEqual(self.mgr.get_database_path(), "/db/custom.db")

    def test_get_template_path_helper(self):
        self.mgr.update_config(default_template_path="/tpl/t.csv")
        self.assertEqual(self.mgr.get_template_path(), "/tpl/t.csv")


class TestConfigManagerSingleton(unittest.TestCase):
    """Verify singleton behaviour across multiple instantiations."""

    def tearDown(self):
        _reset_singleton()

    def test_same_instance_returned(self):
        mgr1 = ConfigManager()
        mgr2 = ConfigManager()
        self.assertIs(mgr1, mgr2)

    def test_mutation_visible_across_references(self):
        mgr1 = ConfigManager()
        mgr1.get_config().theme = "light"
        mgr2 = ConfigManager()
        self.assertEqual(mgr2.get_config().theme, "light")


if __name__ == "__main__":
    unittest.main()
