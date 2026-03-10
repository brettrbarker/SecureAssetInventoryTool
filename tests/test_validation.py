"""
Tests for AssetValidator and ValidationResult.
"""

import sys
import os
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from validation import AssetValidator, ValidationResult


class TestValidationResult(unittest.TestCase):

    def test_default_is_valid(self):
        r = ValidationResult()
        self.assertTrue(r.is_valid)

    def test_add_error_marks_invalid(self):
        r = ValidationResult()
        r.add_error("Something is wrong")
        self.assertFalse(r.is_valid)
        self.assertIn("Something is wrong", r.errors)

    def test_add_warning_does_not_invalidate(self):
        r = ValidationResult()
        r.add_warning("Just a heads-up")
        self.assertTrue(r.is_valid)
        self.assertIn("Just a heads-up", r.warnings)

    def test_get_all_messages_includes_errors_and_warnings(self):
        r = ValidationResult()
        r.add_error("Bad field")
        r.add_warning("Odd value")
        msg = r.get_all_messages()
        self.assertIn("Bad field", msg)
        self.assertIn("Odd value", msg)

    def test_explicit_invalid_at_construction(self):
        r = ValidationResult(is_valid=False, errors=["pre-existing error"])
        self.assertFalse(r.is_valid)


class TestAssetValidatorRequiredFields(unittest.TestCase):

    def setUp(self):
        self.validator = AssetValidator()

    def test_missing_required_field_raises_error(self):
        result = self.validator.validate_asset(
            {},
            template_config={"required_fields": ["manufacturer"]}
        )
        self.assertFalse(result.is_valid)
        self.assertTrue(any("manufacturer" in e for e in result.errors))

    def test_present_required_field_passes(self):
        result = self.validator.validate_asset(
            {"manufacturer": "Dell"},
            template_config={"required_fields": ["manufacturer"]}
        )
        self.assertTrue(result.is_valid)

    def test_empty_string_required_field_fails(self):
        result = self.validator.validate_asset(
            {"manufacturer": "   "},
            template_config={"required_fields": ["manufacturer"]}
        )
        self.assertFalse(result.is_valid)

    def test_no_required_fields_always_passes(self):
        result = self.validator.validate_asset({})
        self.assertTrue(result.is_valid)


class TestAssetValidatorFieldFormats(unittest.TestCase):

    def setUp(self):
        self.validator = AssetValidator()

    # ---- IP address --------------------------------------------------------

    def test_valid_ip_passes(self):
        result = self.validator.validate_asset({"IP_Address": "192.168.1.100"})
        self.assertTrue(result.is_valid)

    def test_invalid_ip_fails(self):
        result = self.validator.validate_asset({"IP_Address": "999.999.999.999"})
        self.assertFalse(result.is_valid)

    def test_invalid_ip_alpha_fails(self):
        result = self.validator.validate_asset({"IP_Address": "abc.def.ghi.jkl"})
        self.assertFalse(result.is_valid)

    # ---- MAC address -------------------------------------------------------

    def test_valid_mac_colon_passes(self):
        result = self.validator.validate_asset({"MAC_Address": "AA:BB:CC:DD:EE:FF"})
        self.assertTrue(result.is_valid)

    def test_valid_mac_dash_passes(self):
        result = self.validator.validate_asset({"MAC_Address": "AA-BB-CC-DD-EE-FF"})
        self.assertTrue(result.is_valid)

    def test_invalid_mac_fails(self):
        result = self.validator.validate_asset({"MAC_Address": "ZZZZ-NOPE"})
        self.assertFalse(result.is_valid)

    # ---- email -------------------------------------------------------------

    def test_valid_email_passes(self):
        result = self.validator.validate_asset({"Email": "user@example.com"})
        self.assertTrue(result.is_valid)

    def test_invalid_email_fails(self):
        result = self.validator.validate_asset({"Email": "not-an-email"})
        self.assertFalse(result.is_valid)

    # ---- serial number -----------------------------------------------------

    def test_valid_serial_number_passes(self):
        result = self.validator.validate_asset({"Serial_Number": "SN-12345-ABC"})
        self.assertTrue(result.is_valid)

    def test_serial_number_with_spaces_fails(self):
        result = self.validator.validate_asset({"Serial_Number": "SN 123"})
        self.assertFalse(result.is_valid)


class TestAssetValidatorBusinessRules(unittest.TestCase):

    def setUp(self):
        self.validator = AssetValidator()

    def test_short_serial_number_warns(self):
        result = self.validator.validate_asset({"Serial Number": "AB"})
        # Should be valid but have a warning
        self.assertTrue(result.is_valid)
        self.assertTrue(len(result.warnings) > 0)

    def test_long_enough_serial_no_warning(self):
        result = self.validator.validate_asset({"Serial Number": "ABC123"})
        serial_warnings = [w for w in result.warnings if "Serial Number" in w]
        self.assertEqual(len(serial_warnings), 0)

    def test_valid_date_passes(self):
        result = self.validator.validate_asset({"Audit Date": "03/10/2026"})
        self.assertTrue(result.is_valid)

    def test_invalid_date_fails(self):
        result = self.validator.validate_asset({"Audit Date": "not-a-date"})
        self.assertFalse(result.is_valid)

    def test_negative_monetary_value_warns(self):
        result = self.validator.validate_asset({"Purchase_Price": "-500"})
        self.assertTrue(len(result.warnings) > 0)

    def test_invalid_monetary_value_fails(self):
        result = self.validator.validate_asset({"Purchase_Price": "one hundred"})
        self.assertFalse(result.is_valid)


class TestAssetValidatorFilePath(unittest.TestCase):

    def setUp(self):
        self.validator = AssetValidator()

    def test_valid_existing_file_passes(self):
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False) as f:
            path = f.name
        try:
            result = self.validator.validate_file_path(path)
            self.assertTrue(result.is_valid)
        finally:
            os.remove(path)

    def test_nonexistent_file_fails(self):
        result = self.validator.validate_file_path("/no/such/file.csv")
        self.assertFalse(result.is_valid)

    def test_empty_path_fails(self):
        result = self.validator.validate_file_path("")
        self.assertFalse(result.is_valid)


if __name__ == "__main__":
    unittest.main()
