"""
Comprehensive validation framework for asset data.
Provides consistent validation rules and user feedback.
"""

import re
from typing import Dict, List, Any, Tuple
from datetime import datetime
import os

class ValidationResult:
    """Result of a validation operation."""
    
    def __init__(self, is_valid: bool = True, errors: List[str] = None, warnings: List[str] = None):
        self.is_valid = is_valid
        self.errors = errors or []
        self.warnings = warnings or []
    
    def add_error(self, error: str):
        """Add an error message."""
        self.errors.append(error)
        self.is_valid = False
    
    def add_warning(self, warning: str):
        """Add a warning message."""
        self.warnings.append(warning)
    
    def get_all_messages(self) -> str:
        """Get all validation messages as a formatted string."""
        messages = []
        if self.errors:
            messages.append("Errors:")
            messages.extend([f"• {error}" for error in self.errors])
        if self.warnings:
            messages.append("Warnings:")
            messages.extend([f"• {warning}" for warning in self.warnings])
        return "\n".join(messages)

class AssetValidator:
    """Validates asset data according to business rules."""
    
    def __init__(self):
        # Common validation patterns
        self.patterns = {
            'email': re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'),
            'phone': re.compile(r'^\+?1?[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}$'),
            'ip_address': re.compile(r'^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'),
            'mac_address': re.compile(r'^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$'),
            'serial_number': re.compile(r'^[A-Za-z0-9\-_]+$'),
        }
    
    def validate_asset(self, asset_data: Dict[str, Any], template_config: Dict[str, Any] = None) -> ValidationResult:
        """Validate complete asset data."""
        result = ValidationResult()
        
        # Basic required field validation
        self._validate_required_fields(asset_data, result, template_config)
        
        # Field-specific validation
        self._validate_field_formats(asset_data, result)
        
        # Business rule validation
        self._validate_business_rules(asset_data, result)
        
        return result
    
    def _validate_required_fields(self, asset_data: Dict[str, Any], result: ValidationResult, 
                                 template_config: Dict[str, Any] = None):
        """Validate required fields based on template configuration."""
        required_fields = []  # Start with empty list - will be populated from config
        
        if template_config and 'required_fields' in template_config:
            required_fields = template_config['required_fields']
        
        # Only validate if we have configured required fields
        for field in required_fields:
            if field not in asset_data or not str(asset_data[field]).strip():
                result.add_error(f"Required field '{field}' is missing or empty")
    
    def _validate_field_formats(self, asset_data: Dict[str, Any], result: ValidationResult):
        """Validate field formats using regex patterns."""
        field_validations = {
            'Email': 'email',
            'Contact_Email': 'email',
            'IP_Address': 'ip_address',
            'Network_IP': 'ip_address',
            'MAC_Address': 'mac_address',
            'Serial_Number': 'serial_number',
            'Phone': 'phone',
            'Contact_Phone': 'phone'
        }
        
        for field, pattern_name in field_validations.items():
            if field in asset_data and asset_data[field]:
                value = str(asset_data[field]).strip()
                if value and not self.patterns[pattern_name].match(value):
                    result.add_error(f"Invalid format for '{field}': {value}")
    
    def _validate_business_rules(self, asset_data: Dict[str, Any], result: ValidationResult):
        """Validate business-specific rules."""
        # Serial Number validation (common identifier)
        if 'Serial Number' in asset_data:
            serial_number = str(asset_data['Serial Number']).strip()
            if serial_number and len(serial_number) < 3:
                result.add_warning("Serial Number should be at least 3 characters long for better tracking")
        
        # Date validations
        date_fields = ['Purchase_Date', 'Warranty_End', 'Last_Maintenance', 'Audit Date']
        for field in date_fields:
            if field in asset_data and asset_data[field]:
                if not self._validate_date_format(asset_data[field]):
                    result.add_error(f"Invalid date format for '{field}'. Expected format: MM/DD/YYYY")
        
        # Monetary value validation
        money_fields = ['Purchase_Price', 'Current_Value']
        for field in money_fields:
            if field in asset_data and asset_data[field]:
                try:
                    value = float(str(asset_data[field]).replace('$', '').replace(',', ''))
                    if value < 0:
                        result.add_warning(f"Negative value for '{field}' may indicate an error")
                except ValueError:
                    result.add_error(f"Invalid monetary value for '{field}': {asset_data[field]}")
    
    def _validate_date_format(self, date_value: Any) -> bool:
        """Validate date format. Only accepts MM/D/YYYY format."""
        if not date_value:
            return True  # Empty dates are OK
        
        if isinstance(date_value, str):
            date_str = date_value.strip()
            if not date_str:
                return True
            
            # Only accept MM/D/YYYY format (handles both MM/DD/YYYY and M/D/YYYY)
            try:
                datetime.strptime(date_str, '%m/%d/%Y')
                return True
            except ValueError:
                return False
        
        return True
    
    def validate_file_path(self, file_path: str, must_exist: bool = True) -> ValidationResult:
        """Validate file path."""
        result = ValidationResult()
        
        if not file_path:
            result.add_error("File path is required")
            return result
        
        if must_exist and not os.path.exists(file_path):
            result.add_error(f"File does not exist: {file_path}")
        
        if must_exist and not os.access(file_path, os.R_OK):
            result.add_error(f"File is not readable: {file_path}")
        
        return result
    
    def validate_template_compatibility(self, csv_headers: List[str], 
                                      template_fields: List[str]) -> ValidationResult:
        """Validate CSV template compatibility."""
        result = ValidationResult()
        
        missing_fields = set(template_fields) - set(csv_headers)
        extra_fields = set(csv_headers) - set(template_fields)
        
        if missing_fields:
            result.add_error(f"Missing required fields: {', '.join(missing_fields)}")
        
        if extra_fields:
            result.add_warning(f"Extra fields in CSV: {', '.join(extra_fields)}")
        
        return result

class FormValidator:
    """Validates form input in real-time."""
    
    def __init__(self):
        self.asset_validator = AssetValidator()
    
    def validate_field(self, field_name: str, value: Any) -> Tuple[bool, str]:
        """Validate a single field and return (is_valid, message)."""
        if not value or not str(value).strip():
            return True, ""  # Empty fields are handled by required field validation
        
        # Quick validation for common fields
        if field_name in ['Email', 'Contact_Email']:
            if not self.asset_validator.patterns['email'].match(str(value)):
                return False, "Invalid email format"
        
        elif field_name in ['IP_Address', 'Network_IP']:
            if not self.asset_validator.patterns['ip_address'].match(str(value)):
                return False, "Invalid IP address format"
        
        elif field_name in ['MAC_Address']:
            if not self.asset_validator.patterns['mac_address'].match(str(value)):
                return False, "Invalid MAC address format (XX:XX:XX:XX:XX:XX)"
        
        elif field_name in ['Phone', 'Contact_Phone']:
            if not self.asset_validator.patterns['phone'].match(str(value)):
                return False, "Invalid phone number format"
        
        elif 'Date' in field_name:
            try:
                if isinstance(value, str) and value.strip():
                    datetime.strptime(value, '%m/%d/%Y')
            except ValueError:
                return False, "Invalid date format (MM/DD/YYYY)"
        
        return True, ""

# Global validator instances
asset_validator = AssetValidator()
form_validator = FormValidator()
