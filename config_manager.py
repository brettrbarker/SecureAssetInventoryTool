"""
Centralized configuration management for the Secure Asset Inventory Tool.
Handles loading, saving, and providing consistent access to configuration data.
"""

import os
import sys
import json
import shutil
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict

@dataclass
class AppConfig:
    """Configuration data class with type hints and defaults."""
    theme: str = "dark"
    default_template_path: str = ""
    output_directory: str = "assets/output_files"
    database_path: str = "assets/asset_database.db"
    dropdown_fields: list = None
    required_fields: list = None
    excluded_fields: list = None
    unique_fields: list = None
    monitor_primary_fields: list = None
    monitor_secondary_fields: list = None
    monitor_tertiary_fields: list = None
    label_output_fields: list = None
    hmr_fields: list = None
    destruction_report_fields: list = None
    bulk_update_presets: dict = None
    saved_searches: dict = None
    
    def __post_init__(self):
        """Set default field lists if not provided."""
        if self.dropdown_fields is None:
            self.dropdown_fields = ["System Name", "*Asset Type", "*Manufacturer", "*Model", 
                                   "Status", "Location", "Room", "Cubicle", "Child Asset? (Y/N)"]
        if self.required_fields is None:
            self.required_fields = ["System Name", "*Asset Type", "*Manufacturer", "*Model", 
                                   "Status", "Location", "Room", "Serial Number"]
        if self.excluded_fields is None:
            self.excluded_fields = ["Asset No.", "Version", "Client (user names, semicolon delimited)",
                                   "Service Contract? (Y/N)", "Contract Expiration Date", "Billing Rate Name",
                                   "Warranty Type", "Multi-Install? (Y/N, Child Assets only)", 
                                   "Install Count (Child Assets only)", "Reservable? (Y/N)",
                                   "Discovered Serial Number", "Discovery Sync ID", "Delete? (Y/N)",
                                   "NOTE: * = Field required for new records."]
        if self.unique_fields is None:
            self.unique_fields = ["Serial Number", "IP Address", "MAC Address", "Phone Number",
                                 "Media Control#", "TSCO Control#", "Tamper Seal", "Network Name"]
        if self.monitor_primary_fields is None:
            self.monitor_primary_fields = ["Serial Number", "Asset No."]
        if self.monitor_secondary_fields is None:
            self.monitor_secondary_fields = ["*Manufacturer", "*Model"]
        if self.monitor_tertiary_fields is None:
            self.monitor_tertiary_fields = ["Room", "Cubicle", "System Name"]
        if self.label_output_fields is None:
            self.label_output_fields = ["Asset No.", "Serial Number", "*Manufacturer", "*Model"]
        if self.hmr_fields is None:
            self.hmr_fields = ["Asset No.", "Serial Number", "*Manufacturer", "*Model", "Location", "Room"]
        if self.destruction_report_fields is None:
            self.destruction_report_fields = ["Asset No.", "Serial Number", "*Manufacturer", "*Model", "Status", "Location"]
        if self.bulk_update_presets is None:
            self.bulk_update_presets = {
                "Update Audit Date": {
                    "type": "system",
                    "fields": [
                        {
                            "field": "Audit Date",
                            "operation": "replace",
                            "value": "current_date"
                        }
                    ]
                }
            }
        if self.saved_searches is None:
            self.saved_searches = {}
    
    def get(self, key: str, default=None):
        """Backward compatibility method to work like dict.get()."""
        return getattr(self, key, default)
    
    def __contains__(self, key: str) -> bool:
        """Support 'in' operator for backward compatibility."""
        return hasattr(self, key)
    
    def __getitem__(self, key: str):
        """Support bracket notation for backward compatibility."""
        if hasattr(self, key):
            return getattr(self, key)
        raise KeyError(key)
    
    def __setitem__(self, key: str, value):
        """Support bracket notation assignment for backward compatibility."""
        setattr(self, key, value)
    
    def to_dict(self) -> dict:
        """Convert AppConfig to dictionary for JSON serialization."""
        return asdict(self)

class ConfigManager:
    """Centralized configuration manager with singleton pattern."""
    
    _instance = None
    _config = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if self._config is None:
            self.config_path = os.path.join("assets", "config.json")
            self._config = self._load_config()
    
    def _load_config(self) -> AppConfig:
        """Load configuration from file or create default."""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return AppConfig(**data)
            except (json.JSONDecodeError, TypeError, OSError):
                pass
        
        # Create default config and save it
        config = AppConfig()
        self.save_config(config)
        return config
    
    def get_config(self) -> AppConfig:
        """Get current configuration."""
        return self._config
    
    def update_config(self, **kwargs) -> None:
        """Update configuration with new values."""
        for key, value in kwargs.items():
            if hasattr(self._config, key):
                setattr(self._config, key, value)
    
    def save_config(self, config: AppConfig = None) -> bool:
        """Save configuration to file."""
        if config:
            self._config = config
            
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(asdict(self._config), f, indent=4)
            return True
        except OSError:
            return False
    
    def get_database_path(self) -> str:
        """Get the current database path."""
        return self._config.database_path
    
    def get_template_path(self) -> str:
        """Get the current template path."""
        return self._config.default_template_path
    
    def _get_bundled_path(self, relative_path: str) -> str:
        """
        Get the absolute path to a bundled resource.
        Works for both development and PyInstaller --onefile mode.
        """
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            # _MEIPASS is where PyInstaller extracts bundled files
            base_path = sys._MEIPASS
        else:
            # Running in development
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        return os.path.join(base_path, relative_path)
    
    def _copy_bundled_file(self, src_relative: str, dest_path: str) -> bool:
        """
        Copy a bundled file from PyInstaller temp directory to working directory.
        Only copies if destination doesn't exist.
        """
        try:
            # Don't overwrite existing files
            if os.path.exists(dest_path):
                return True
            
            # Get the bundled file path
            src_path = self._get_bundled_path(src_relative)
            
            # Check if source exists
            if not os.path.exists(src_path):
                return False
            
            # Create destination directory if needed
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            
            # Copy the file
            shutil.copy2(src_path, dest_path)
            return True
        except Exception as e:
            print(f"Error copying bundled file {src_relative}: {e}")
            return False
    
    def _copy_bundled_directory(self, src_relative: str, dest_dir: str) -> bool:
        """
        Copy a bundled directory from PyInstaller temp to working directory.
        Only copies files that don't exist.
        """
        try:
            src_dir = self._get_bundled_path(src_relative)
            
            # Check if source directory exists
            if not os.path.exists(src_dir):
                return False
            
            # Create destination directory
            os.makedirs(dest_dir, exist_ok=True)
            
            # Walk through source directory and copy files
            for root, dirs, files in os.walk(src_dir):
                # Calculate relative path from source directory
                rel_path = os.path.relpath(root, src_dir)
                dest_subdir = os.path.join(dest_dir, rel_path) if rel_path != '.' else dest_dir
                
                # Create subdirectories
                os.makedirs(dest_subdir, exist_ok=True)
                
                # Copy files
                for file in files:
                    src_file = os.path.join(root, file)
                    dest_file = os.path.join(dest_subdir, file)
                    
                    # Only copy if destination doesn't exist
                    if not os.path.exists(dest_file):
                        shutil.copy2(src_file, dest_file)
            
            return True
        except Exception as e:
            print(f"Error copying bundled directory {src_relative}: {e}")
            return False
    
    def ensure_directories(self) -> None:
        """Ensure all configured directories exist and copy bundled assets."""
        # Create output directory structure
        os.makedirs(self._config.output_directory, exist_ok=True)
        os.makedirs(os.path.join(self._config.output_directory, "exports"), exist_ok=True)
        os.makedirs(os.path.join(self._config.output_directory, "labels"), exist_ok=True)
        os.makedirs(os.path.join(self._config.output_directory, "reports"), exist_ok=True)
        
        # Create database directory
        os.makedirs(os.path.dirname(self._config.database_path), exist_ok=True)
        
        # Create templates directory
        templates_dir = os.path.join("assets", "templates")
        os.makedirs(templates_dir, exist_ok=True)
        
        # Create fonts directory
        fonts_dir = os.path.join("assets", "fonts")
        os.makedirs(fonts_dir, exist_ok=True)
        
        # Create auto_backups directory
        backups_dir = os.path.join("assets", "auto_backups")
        os.makedirs(backups_dir, exist_ok=True)
        
        # Copy bundled assets from PyInstaller temp directory (if running as exe)
        if getattr(sys, 'frozen', False):
            # Copy default template
            self._copy_bundled_file(
                "assets/templates/default_template.csv",
                os.path.join(templates_dir, "default_template.csv")
            )
            
            # Copy config.json if it doesn't exist
            config_dest = os.path.join("assets", "config.json")
            if not os.path.exists(config_dest):
                self._copy_bundled_file("assets/config.json", config_dest)
            
            # Copy entire fonts directory
            self._copy_bundled_directory("assets/fonts", fonts_dir)
    
    def get_output_directory(self) -> str:
        """Get the configured output directory path."""
        return self._config.output_directory
    
    def get_suggested_filepath(self, filename: str, file_type: str = "export") -> str:
        """
        Get a suggested full file path in the output directory.
        
        Args:
            filename: The base filename (with or without extension)
            file_type: Type of file (export, report, barcode, etc.) for organization
            
        Returns:
            str: Full suggested file path in output directory
        """
        # Ensure output directory exists
        self.ensure_directories()
        
        # Optionally create subdirectories for organization
        subdir_map = {
            "export": "exports",
            "report": "reports", 
            "barcode": "barcodes",
            "audit": "audit_reports",
            "analysis": "analysis"
        }
        
        subdir = subdir_map.get(file_type, "")
        if subdir:
            full_output_dir = os.path.join(self._config.output_directory, subdir)
            os.makedirs(full_output_dir, exist_ok=True)
        else:
            full_output_dir = self._config.output_directory
            
        return os.path.join(full_output_dir, filename)
