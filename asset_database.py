"""
SQLite database manager for asset management system.
Handles database creation, migrations, and all CRUD operations.
"""

import sqlite3
import csv
import os
import json
import re
import getpass
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any, Set
from contextlib import contextmanager

# Generated from AI prompt to convert to sqlite DB
class AssetDatabase:
    """Manages SQLite database operations for asset management."""
    
    def __init__(self, db_path: str = None):
        """Initialize database connection and ensure schema exists."""
        # Use configured database path if no path provided
        if db_path is None:
            db_path = self._get_configured_database_path()
        
        self.db_path = db_path
        
        # Load default template from config if available
        default_template = self._get_default_template_path()
        self.ensure_database_exists(default_template)
    
    def _get_configured_database_path(self) -> str:
        """Get the database path from config.json, with fallback to default."""
        config_path = os.path.join("assets", "config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    db_path = config.get("database_path")
                    if db_path and os.path.dirname(db_path):
                        return db_path
            except (json.JSONDecodeError, IOError):
                pass
        
        return "assets/asset_database.db"  # Default fallback
    
    def _get_default_template_path(self) -> Optional[str]:
        """Get the default template path from config.json if available."""
        # Determine the assets directory - if db_path has no directory, assume 'assets'
        db_dir = os.path.dirname(self.db_path)
        if not db_dir:
            db_dir = "assets"
        
        config_path = os.path.join(db_dir, "config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    template_path = config.get("default_template_path")
                    if template_path and os.path.exists(template_path):
                        return template_path
            except (json.JSONDecodeError, IOError): 
                pass
        
        # Fallback to common template files in the assets directory
        for template_name in ["default_template.csv", "sample.csv", "template.csv"]:
            template_path = os.path.join(db_dir, template_name)
            if os.path.exists(template_path):
                return template_path
        
        return None

    def _get_current_user(self) -> str:
        """Get the current logged-in user."""
        try:
            return getpass.getuser()
        except Exception:
            # Fallback to 'system' if unable to get user
            return 'system'

    @contextmanager
    def get_connection(self):
        """Context manager for database connections."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row  # Enable column access by name
        try:
            yield conn
        finally:
            conn.close()
    
    def ensure_database_exists(self, template_path: str = None):
        """Create database and tables if they don't exist. If template_path is provided, use its headers for columns."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        
        # Determine columns from template if provided
        template_headers = []
        if template_path and os.path.exists(template_path):
            try:
                with open(template_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    template_headers = next(reader)
            except Exception:
                pass
        
        # System columns always present
        system_columns = [
            'id INTEGER PRIMARY KEY AUTOINCREMENT',
            'created_date DATETIME DEFAULT CURRENT_TIMESTAMP',
            "modified_date DATETIME DEFAULT '1901-01-01 00:00:00'",
            "label_requested_date DATETIME DEFAULT NULL",
            "created_by TEXT DEFAULT 'system'",
            "modified_by TEXT DEFAULT 'system'",
            "data_source TEXT DEFAULT 'manual'",
            "is_deleted INTEGER DEFAULT 0"
        ]
        # Template columns - use consistent mapping with field type detection
        template_columns = []
        if template_headers:
            column_mapping = self.get_dynamic_column_mapping(template_path)
            field_types = self._detect_field_types(template_path, template_headers)
            for header in template_headers:
                if header.strip():
                    col_name = column_mapping.get(header, self._generate_safe_column_name(header))
                    field_type = field_types.get(header, 'TEXT')
                    template_columns.append(f'{col_name} {field_type}')
        
        all_columns = system_columns + template_columns
        columns_sql = ',\n    '.join(all_columns)
        
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS assets (
                    {columns_sql}
                )
            """)
            # Create audit log table for tracking changes
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS asset_audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    asset_id INTEGER,
                    action TEXT NOT NULL,
                    field_name TEXT,
                    old_value TEXT,
                    new_value TEXT,
                    changed_by TEXT DEFAULT 'system',
                    change_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (asset_id) REFERENCES assets (id)
                )
            """)
            conn.commit()
    
    def update_schema_for_template(self, csv_path: str) -> bool:
        """Update database schema to accommodate new template fields."""
        if not os.path.exists(csv_path):
            return False
        
        try:
            # Read template headers
            with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                headers = next(reader)
            
            # Get current database columns
            existing_columns = self.get_table_columns()
            
            # Check for new fields that need columns
            new_fields = []
            for header in headers:
                if header.strip():
                    # Generate a safe column name
                    safe_column = self._generate_safe_column_name(header)
                    # Check if this column actually exists in the database
                    if safe_column not in existing_columns:
                        new_fields.append((header, safe_column))
            
            if new_fields:
                with self.get_connection() as conn:
                    cursor = conn.cursor()
                    
                    for header, column_name in new_fields:
                        try:
                            # Add new column to assets table
                            cursor.execute(f"ALTER TABLE assets ADD COLUMN {column_name} TEXT")
                            print(f"Added column '{column_name}' for header '{header}'")
                        except Exception as e:
                            print(f"Warning: Could not add column {column_name}: {e}")
                    
                    conn.commit()
                
                # Update the column mapping
                self._update_column_mapping(new_fields)
                
            return True
            
        except Exception as e:
            print(f"Error updating schema: {e}")
            return False
    
    def _detect_field_types(self, template_path: str, headers: List[str]) -> Dict[str, str]:
        """Detect field types based on template content and field names."""
        field_types = {}
        
        # Default field type mappings based on field names
        multiline_field_names = ['notes', 'description', 'comments', 'remarks', 'details']
        
        for header in headers:
            header_lower = header.lower().replace('*', '').strip()
            
            # Check if field name suggests multiline content
            if any(keyword in header_lower for keyword in multiline_field_names):
                field_types[header] = 'TEXT'  # Still TEXT but we'll track it's multiline
                continue
            
            # Default to TEXT for most fields
            field_types[header] = 'TEXT'
        
        # If template file exists, scan for actual multiline content
        if template_path and os.path.exists(template_path):
            try:
                with open(template_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        for header, value in row.items():
                            if header in headers and value and '\n' in str(value):
                                field_types[header] = 'TEXT'  # Mark as multiline
                                break
            except Exception:
                pass
        
        return field_types

    def _detect_multiline_fields_from_data(self) -> Set[str]:
        """Detect which fields contain multiline data in existing database."""
        multiline_fields = set()
        
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get all table columns
                columns = self.get_table_columns()
                system_columns = {'id', 'created_date', 'modified_date', 'label_requested_date', 'created_by', 'modified_by', 'data_source', 'is_deleted'}
                data_columns = [col for col in columns if col not in system_columns]
                
                # Check each column for newline characters
                for column in data_columns:
                    cursor.execute(f"SELECT {column} FROM assets WHERE {column} LIKE '%\n%' LIMIT 1")
                    if cursor.fetchone():
                        multiline_fields.add(column)
                        
        except Exception as e:
            print(f"Error detecting multiline fields: {e}")
        
        return multiline_fields

    def get_field_metadata(self, template_path: str = None) -> Dict[str, Dict[str, Any]]:
        """Get metadata about fields including whether they should be multiline."""
        metadata = {}
        
        # Get multiline fields from existing data
        multiline_fields = self._detect_multiline_fields_from_data()
        
        # Get column mapping
        column_mapping = self.get_dynamic_column_mapping(template_path) if template_path else {}
        
        # Get all columns
        columns = self.get_table_columns()
        system_columns = {'id', 'created_date', 'modified_date', 'label_requested_date', 'created_by', 'modified_by', 'data_source', 'is_deleted'}
        
        for column in columns:
            if column in system_columns:
                continue
                
            # Find the corresponding header name
            header_name = None
            for header, db_col in column_mapping.items():
                if db_col == column:
                    header_name = header
                    break
            
            if not header_name:
                header_name = self._column_to_header(column)
            
            # Determine if field should be multiline
            is_multiline = (
                column in multiline_fields or
                any(keyword in column.lower() for keyword in ['notes', 'description', 'comments', 'remarks', 'details'])
            )
            
            metadata[header_name] = {
                'db_column': column,
                'is_multiline': is_multiline,
                'field_type': 'multiline_text' if is_multiline else 'text'
            }
        
        return metadata

    def _generate_safe_column_name(self, header: str) -> str:
        """Generate a safe database column name from header."""
        import re
        # Remove special characters and convert to lowercase
        safe_name = re.sub(r'[^a-zA-Z0-9\s]', '', header)
        safe_name = re.sub(r'\s+', '_', safe_name.strip())
        safe_name = safe_name.lower()
        
        # Ensure it doesn't conflict with existing names
        if safe_name in ['id', 'created_date', 'modified_date', 'label_requested_date', 'created_by', 'modified_by']:
            safe_name = f"field_{safe_name}"
        
        return safe_name
    
    def _update_column_mapping(self, new_fields: List[Tuple[str, str]]):
        """Update the internal column mapping with new fields."""
        # This would ideally be stored in a configuration table
        # For now, we'll extend the existing mapping method
        pass
    
    def get_dynamic_column_mapping(self, csv_path: str = None) -> Dict[str, str]:
        """Get column mapping using purely dynamic generation from CSV headers."""
        mapping = {}
        
        if csv_path and os.path.exists(csv_path):
            try:
                with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    headers = next(reader)
                
                # Generate mappings for all fields dynamically
                for header in headers:
                    if header.strip():
                        safe_column = self._generate_safe_column_name(header)
                        mapping[header] = safe_column
            except Exception:
                pass
        
        return mapping
    
    def get_table_columns(self) -> List[str]:
        """Get list of all columns in the assets table."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(assets)")
            columns = cursor.fetchall()
            return [col[1] for col in columns]  # Column name is at index 1

    def get_database_tables(self) -> List[str]:
        """Get list of all tables in the database."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
            tables = cursor.fetchall()
            return [table[0] for table in tables]
    
    def verify_template_compatibility(self, csv_path: str) -> Dict[str, Any]:
        """Verify template compatibility and return status information."""
        if not os.path.exists(csv_path):
            return {"error": "File not found"}
        
        try:
            # Read template headers
            with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                headers = next(reader)
            
            # Get current columns and mapping
            table_columns = self.get_table_columns()
            column_mapping = self.get_dynamic_column_mapping(csv_path)
            
            # Analyze compatibility
            mapped_fields = []
            new_fields = []
            
            for header in headers:
                if header.strip():
                    db_column = column_mapping.get(header)
                    if db_column and db_column in table_columns:
                        mapped_fields.append({"header": header, "column": db_column, "status": "exists"})
                    else:
                        safe_column = self._generate_safe_column_name(header)
                        new_fields.append({"header": header, "column": safe_column, "status": "will_create"})
            
            return {
                "total_fields": len(headers),
                "mapped_fields": len(mapped_fields),
                "new_fields": len(new_fields),
                "field_details": mapped_fields + new_fields,
                "table_columns": len(table_columns)
            }
            
        except Exception as e:
            return {"error": str(e)}
    
    def import_csv_template(self, csv_path: str) -> int:
        """Import CSV template data into database. Returns number of records imported."""
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV file not found: {csv_path}")
        
        # First, update schema to accommodate any new fields
        self.update_schema_for_template(csv_path)
        
        column_mapping = self.get_dynamic_column_mapping(csv_path)
        imported_count = 0
        
        with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                for row in reader:
                    # Skip empty rows or notes
                    if not any(row.values()) or row.get("Asset No.", "").startswith("NOTE:"):
                        continue
                    
                    # Map CSV columns to database columns
                    db_row = {}
                    for csv_col, value in row.items():
                        db_col = column_mapping.get(csv_col)
                        if db_col and value and value.strip():
                            db_row[db_col] = value.strip()
                    
                    if db_row:  # Only insert if we have some data
                        self._insert_asset(cursor, db_row, 'import')
                        imported_count += 1
                
                conn.commit()
        
        return imported_count
    
    def import_from_csv(self, csv_path: str, overwrite_callback=None) -> int:
        """
        Import CSV data with duplicate checking and user confirmation.
        
        Args:
            csv_path: Path to CSV file to import
            overwrite_callback: Function to call when duplicate found (serial, asset_no, existing_data, new_data) -> 'overwrite', 'skip', 'overwrite_all', 'skip_all'
        
        Returns:
            Number of records imported
        """
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV file not found: {csv_path}")
        
        # First, update schema to accommodate any new fields
        self.update_schema_for_template(csv_path)
        
        column_mapping = self.get_dynamic_column_mapping(csv_path)
        imported_count = 0
        overwrite_mode = None  # 'all', 'none', or None for ask each time
        
        with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                for row in reader:
                    # Skip empty rows or notes
                    if not any(row.values()) or row.get("Asset No.", "").startswith("NOTE:"):
                        continue
                    
                    # Map CSV columns to database columns
                    db_row = {}
                    for csv_col, value in row.items():
                        db_col = column_mapping.get(csv_col)
                        if db_col and value and value.strip():
                            db_row[db_col] = value.strip()
                    
                    if not db_row:  # Skip if no data
                        continue
                    
                    # Check for duplicates by serial number and asset number
                    existing_asset = None
                    duplicate_type = None
                    
                    # Check serial number first (using current connection)
                    if db_row.get('serial_number'):
                        cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (db_row['serial_number'],))
                        result = cursor.fetchone()
                        if result:
                            existing_asset = dict(result)
                            duplicate_type = 'serial_number'
                    
                    # Check asset number if no serial duplicate found (using current connection)
                    if not existing_asset and db_row.get('asset_no'):
                        cursor.execute("SELECT * FROM assets WHERE asset_no = ?", (db_row['asset_no'],))
                        result = cursor.fetchone()
                        if result:
                            existing_asset = dict(result)
                            duplicate_type = 'asset_no'
                    
                    # Handle duplicate if found
                    if existing_asset:
                        action = None
                        
                        if overwrite_mode == 'all':
                            action = 'overwrite'
                        elif overwrite_mode == 'none':
                            action = 'skip'
                        elif overwrite_callback:
                            action = overwrite_callback(
                                duplicate_type,
                                db_row.get(duplicate_type, ''),
                                existing_asset,
                                db_row
                            )
                        else:
                            action = 'skip'  # Default behavior if no callback
                        
                        # Handle global overwrite decisions
                        if action == 'overwrite_all':
                            overwrite_mode = 'all'
                            action = 'overwrite'
                        elif action == 'skip_all':
                            overwrite_mode = 'none'
                            action = 'skip'
                        
                        if action == 'overwrite':
                            # Update existing asset using current connection
                            if self._update_asset_with_cursor(cursor, existing_asset['id'], db_row):
                                imported_count += 1
                        # If action is 'skip', just continue to next record
                        
                    else:
                        # No duplicate found, insert new asset
                        if not db_row.get('asset_no'):
                            db_row['asset_no'] = self._generate_asset_number(cursor)
                        
                        self._insert_asset(cursor, db_row, 'import')
                        imported_count += 1
                
                conn.commit()
        
        return imported_count
    
    def _insert_asset(self, cursor, asset_data: Dict[str, Any], data_source: str = 'manual', changed_by: Optional[str] = None) -> int:
        """Insert a single asset record. Returns the new asset ID."""
        # Generate asset number if not provided
        if not asset_data.get('asset_no'):
            asset_data['asset_no'] = self._generate_asset_number(cursor)
        
        # Set data source
        asset_data['data_source'] = data_source
        
        # Set created_by if not already set
        if 'created_by' not in asset_data or not asset_data['created_by']:
            asset_data['created_by'] = changed_by if changed_by else self._get_current_user()
        
        # Prepare columns and values
        columns = list(asset_data.keys())
        placeholders = ['?' for _ in columns]
        values = list(asset_data.values())
        
        # Add timestamp - created_date as current time, modified_date stays at default '1901-01-01'
        columns.append('created_date')
        placeholders.append('?')
        now = datetime.now().isoformat()
        values.append(now)
        # Don't set modified_date - let it use the default '1901-01-01 00:00:00'
        
        query = f"""
            INSERT OR REPLACE INTO assets ({', '.join(columns)})
            VALUES ({', '.join(placeholders)})
        """
        
        cursor.execute(query, values)
        asset_id = cursor.lastrowid
        
        # Log the creation
        self._log_audit_action(cursor, asset_id, 'INSERT', None, None, str(asset_data), changed_by)
        
        return asset_id
    
    def _generate_asset_number(self, cursor) -> str:
        """Generate a unique asset number."""
        cursor.execute("SELECT COUNT(*) FROM assets")
        count = cursor.fetchone()[0]
        return f"AST{count + 1:06d}"
    
    def _log_audit_action(self, cursor, asset_id: int, action: str, field_name: Optional[str] = None, 
                         old_value: Optional[str] = None, new_value: Optional[str] = None, 
                         changed_by: Optional[str] = None):
        """Log an action to the audit table."""
        if changed_by is None:
            changed_by = self._get_current_user()
        
        cursor.execute("""
            INSERT INTO asset_audit_log (asset_id, action, field_name, old_value, new_value, changed_by)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (asset_id, action, field_name, old_value, new_value, changed_by))
    
    def _update_asset_with_cursor(self, cursor, asset_id: int, updates: Dict[str, Any], changed_by: Optional[str] = None) -> bool:
        """Update an asset using the provided cursor (for transaction safety)."""
        if not updates:
            return False
        
        # Get changed_by user or default
        if changed_by is None:
            changed_by = self._get_current_user()
        
        # Build SET clause
        set_clauses = []
        values = []
        
        for field, value in updates.items():
            if field not in ['id', 'created_date', 'created_by']:  # Don't update these fields
                set_clauses.append(f"{field} = ?")
                values.append(value)
        
        if not set_clauses:
            return False
        
        # Add modified timestamp and user
        set_clauses.append("modified_date = ?")
        set_clauses.append("modified_by = ?")
        values.extend([datetime.now().isoformat(), changed_by])
        values.append(asset_id)  # For WHERE clause
        
        query = f"UPDATE assets SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        
        # Log the update
        self._log_audit_action(cursor, asset_id, 'UPDATE', None, None, str(updates), changed_by)
        
        return cursor.rowcount > 0
    
    def add_asset(self, asset_data: Dict[str, Any], data_source: str = 'manual', changed_by: Optional[str] = None) -> int:
        """Add a new asset to the database. Returns the new asset ID."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            asset_id = self._insert_asset(cursor, asset_data, data_source, changed_by)
            conn.commit()
            return asset_id
    
    def get_asset_by_id(self, asset_id: int) -> Optional[Dict[str, Any]]:
        """Get an asset by its ID."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM assets WHERE id = ?", (asset_id,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def get_asset_by_serial(self, serial_number: str) -> Optional[Dict[str, Any]]:
        """Get an asset by its serial number."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (serial_number,))
            row = cursor.fetchone()
            return dict(row) if row else None
    
    def check_unique_field_conflicts(self, asset_data: Dict[str, Any], unique_fields: List[str], template_path: str = None) -> List[Dict[str, Any]]:
        """Check for conflicts in unique fields with existing assets.
        
        Args:
            asset_data: The asset data to check
            unique_fields: List of field names that should be unique
            template_path: Path to template CSV file for column mapping
            
        Returns:
            List of dictionaries containing conflict information:
            [{'field_name': str, 'field_value': str, 'conflicting_asset': dict}, ...]
        """
        conflicts = []
        
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Get available table columns to avoid referencing non-existent columns
            available_columns = self.get_table_columns()
            
            # Get dynamic column mapping with template path
            column_mapping = self.get_dynamic_column_mapping(template_path)
            
            for field_name in unique_fields:
                # Get the database column name for this field
                db_column = column_mapping.get(field_name)
                if not db_column or db_column not in available_columns:
                    # If mapping failed, try direct column name matching
                    potential_db_column = self._generate_safe_column_name(field_name)
                    if potential_db_column in available_columns:
                        db_column = potential_db_column
                    else:
                        continue
                    
                # Get the value from asset_data
                field_value = asset_data.get(db_column)
                if not field_value or not field_value.strip():
                    continue  # Skip empty values
                    
                # Check if this value already exists in the database
                cursor.execute(f"SELECT * FROM assets WHERE {db_column} = ? AND is_deleted = 0", (field_value.strip(),))
                conflicting_row = cursor.fetchone()
                
                if conflicting_row:
                    conflicts.append({
                        'field_name': field_name,
                        'field_value': field_value.strip(),
                        'conflicting_asset': dict(conflicting_row)
                    })
        
        return conflicts
    
    def search_assets(self, filters: Dict[str, Any] = None, limit: int = 1000) -> List[Dict[str, Any]]:
        """Search assets with optional filters."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Get available table columns to avoid referencing non-existent columns
            available_columns = self.get_table_columns()
            
            query = "SELECT * FROM assets WHERE 1=1"
            params = []
            
            if filters:
                for field, value in filters.items():
                    if value and field in available_columns:
                        # Use LIKE for text fields that might contain partial matches
                        text_fields = ['notes', 'description', 'manufacturer', 'model', 'location', 'system_name', 'serial_number']
                        if field in text_fields:
                            query += f" AND {field} LIKE ?"
                            params.append(f"%{value}%")
                        else:
                            query += f" AND {field} = ?"
                            params.append(value)
            
            query += " ORDER BY modified_date DESC LIMIT ?"
            params.append(limit)
            
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
    
    def update_asset(self, asset_id: int, updates: Dict[str, Any], changed_by: Optional[str] = None) -> bool:
        """Update an existing asset. Returns True if successful."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Get current values for audit logging
            cursor.execute("SELECT * FROM assets WHERE id = ?", (asset_id,))
            current_asset = cursor.fetchone()
            if not current_asset:
                return False
            
            # Get changed_by user or default
            if changed_by is None:
                changed_by = self._get_current_user()
            
            # Update the asset
            set_clauses = []
            params = []
            for field, value in updates.items():
                set_clauses.append(f"{field} = ?")
                params.append(value)
                
                # Log the change
                old_value = current_asset[field] if field in current_asset.keys() else None
                self._log_audit_action(cursor, asset_id, 'UPDATE', field, str(old_value), str(value), changed_by)
            
            # Add modified timestamp and modified_by
            set_clauses.append("modified_date = ?")
            set_clauses.append("modified_by = ?")
            params.append(datetime.now().isoformat())
            params.append(changed_by)
            params.append(asset_id)
            
            query = f"UPDATE assets SET {', '.join(set_clauses)} WHERE id = ?"
            cursor.execute(query, params)
            
            conn.commit()
            return cursor.rowcount > 0
    
    def request_label(self, asset_id: int, changed_by: Optional[str] = None) -> bool:
        """Request a label for an asset by updating the label_requested_date field. Returns True if successful."""
        current_datetime = datetime.now().isoformat()
        updates = {'label_requested_date': current_datetime}
        return self.update_asset(asset_id, updates, changed_by)
    
    def delete_asset(self, asset_id: int, changed_by: Optional[str] = None) -> bool:
        """Delete an asset. Returns True if successful."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Log the deletion
            self._log_audit_action(cursor, asset_id, 'DELETE', changed_by=changed_by)
            
            # Delete the asset
            cursor.execute("DELETE FROM assets WHERE id = ?", (asset_id,))
            conn.commit()
            return cursor.rowcount > 0
    
    def get_audit_history(self, asset_id: int) -> List[Dict[str, Any]]:
        """Get audit history for a specific asset."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    action,
                    field_name,
                    old_value,
                    new_value,
                    changed_by,
                    change_date
                FROM asset_audit_log 
                WHERE asset_id = ? 
                ORDER BY change_date DESC
            """, (asset_id,))
            
            history = []
            for row in cursor.fetchall():
                action, field_name, old_value, new_value, changed_by, change_date = row
                history.append({
                    'action': action,
                    'field_name': field_name,
                    'old_value': old_value,
                    'new_value': new_value,
                    'changed_by': changed_by,
                    'change_date': change_date
                })
            
            return history
    
    def get_unique_values(self, field: str) -> List[str]:
        """Get unique values for a field to populate dropdowns."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT DISTINCT {field} FROM assets WHERE {field} IS NOT NULL AND {field} != '' ORDER BY {field}")
            return [row[0] for row in cursor.fetchall()]
    
    def get_recent_changes(self, days: int = 7) -> List[Dict[str, Any]]:
        """Get assets modified in the last N days."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
            
            cursor.execute("""
                SELECT * FROM assets 
                WHERE modified_date >= ? 
                ORDER BY modified_date DESC
            """, (cutoff_date.isoformat(),))
            
            return [dict(row) for row in cursor.fetchall()]
    
    def export_to_csv(self, output_path: str, filters: Dict[str, Any] = None, template_path: str = None) -> int:
        """Export assets to CSV format. Returns number of records exported."""
        assets = self.search_assets(filters)
        
        if not assets:
            return 0
        
        # If no template path provided, try to find one
        if not template_path:
            template_path = self._get_default_template_path()
        
        # Get column mapping and headers from template
        if template_path and os.path.exists(template_path):
            try:
                with open(template_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    csv_headers = next(reader)
                column_mapping = self.get_dynamic_column_mapping(template_path)
            except Exception:
                # Fallback: derive headers from database columns
                csv_headers, column_mapping = self._derive_headers_from_database()
        else:
            # Fallback: derive headers from database columns
            csv_headers, column_mapping = self._derive_headers_from_database()
        
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
        
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=csv_headers)
            writer.writeheader()
            
            for asset in assets:
                csv_row = {}
                for header in csv_headers:
                    db_column = column_mapping.get(header)
                    if db_column and db_column in asset:
                        csv_row[header] = asset[db_column] or ""
                    else:
                        csv_row[header] = ""
                writer.writerow(csv_row)
        
        return len(assets)
    
    def _derive_headers_from_database(self) -> Tuple[List[str], Dict[str, str]]:
        """Derive CSV headers from database columns when no template available."""
        table_columns = self.get_table_columns()
        
        # Filter out system columns for CSV export
        system_columns = {'id', 'created_date', 'modified_date', 'label_requested_date', 'created_by', 'modified_by', 'is_deleted'}
        export_columns = [col for col in table_columns if col not in system_columns]
        
        # Create reverse mapping and human-readable headers
        headers = []
        mapping = {}
        
        for db_column in export_columns:
            # Convert database column back to a human-readable header
            header = self._column_to_header(db_column)
            headers.append(header)
            mapping[header] = db_column
        
        return headers, mapping
    
    def _column_to_header(self, db_column: str) -> str:
        """Convert database column name back to human-readable header."""
        # Simple conversion: replace underscores with spaces and title case
        header = db_column.replace('_', ' ').title()
        
        # Handle some special cases for better readability
        replacements = {
            'Asset No': 'Asset No.',
            'Ip Address': 'IP Address',
            'Mac Address': 'MAC Address',
            'Po Number': 'PO Number',
            'Hmr Entrance': 'HMR# (Entrance)',
            'Hmr Exit': 'HMR# (Exit)',
            'Media Control Number': 'Media Control#',
            'Tsco Control Number': 'TSCO Control#',
            'Child Asset Yn': 'Child Asset? (Y/N)',
            'Service Contract': 'Service Contract? (Y/N)',
            'Multi Install': 'Multi-Install? (Y/N, Child Assets only)',
            'Reservable': 'Reservable? (Y/N)',
            'Delete Flag': 'Delete? (Y/N)'
        }
        
        return replacements.get(header, header)
    
    def get_database_stats(self) -> Dict[str, Any]:
        """Get database statistics."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*) FROM assets")
            total_assets = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM asset_audit_log")
            total_audit_entries = cursor.fetchone()[0]
            
            cursor.execute("SELECT MAX(modified_date) FROM assets")
            last_modified = cursor.fetchone()[0]
            
            # Get available columns to check which ones exist
            available_columns = self.get_table_columns()
            
            unique_manufacturers = 0
            unique_locations = 0
            
            # Only query columns if they exist
            if 'manufacturer' in available_columns:
                cursor.execute("SELECT COUNT(DISTINCT manufacturer) FROM assets WHERE manufacturer IS NOT NULL AND manufacturer != ''")
                unique_manufacturers = cursor.fetchone()[0]
            
            if 'location' in available_columns:
                cursor.execute("SELECT COUNT(DISTINCT location) FROM assets WHERE location IS NOT NULL AND location != ''")
                unique_locations = cursor.fetchone()[0]
            
            return {
                "total_assets": total_assets,
                "total_audit_entries": total_audit_entries,
                "last_modified": last_modified,
                "unique_manufacturers": unique_manufacturers,
                "unique_locations": unique_locations,
                "database_path": self.db_path,
                "database_size_mb": os.path.getsize(self.db_path) / (1024 * 1024) if os.path.exists(self.db_path) else 0
            }

    def search_assets_by_field(self, field_name: str, field_value: str) -> List[Any]:
        """Search for assets by a specific field value."""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # Use LIKE for partial matching and case-insensitive search
                cursor.execute(f"""
                    SELECT * FROM assets 
                    WHERE {field_name} LIKE ? AND is_deleted = 0
                    ORDER BY id DESC
                """, (f"%{field_value}%",))
                
                rows = cursor.fetchall()
                
                # Convert to objects with attribute access
                assets = []
                for row in rows:
                    asset = type('Asset', (), {})()
                    for key in row.keys():
                        setattr(asset, key, row[key])
                    assets.append(asset)
                
                return assets
                
        except Exception as e:
            print(f"Error searching assets by field {field_name}: {e}")
            return []

    def get_unique_field_values(self, field_name: str) -> List[str]:
        """Get unique values for a specific field from the database."""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute(f"""
                    SELECT DISTINCT {field_name} 
                    FROM assets 
                    WHERE {field_name} IS NOT NULL 
                    AND {field_name} != '' 
                    AND is_deleted = 0
                    ORDER BY {field_name}
                """)
                
                rows = cursor.fetchall()
                return [row[0] for row in rows if row[0]]
                
        except Exception as e:
            print(f"Error getting unique values for field {field_name}: {e}")
            return []

    def should_field_be_multiline(self, field_name: str, template_path: str = None) -> bool:
        """Determine if a field should be rendered as multiline based on content and name."""
        # Get field metadata
        metadata = self.get_field_metadata(template_path)
        
        # Check if we have specific metadata for this field
        if field_name in metadata:
            return metadata[field_name]['is_multiline']
        
        # Fallback: check field name for common multiline indicators
        field_lower = field_name.lower().replace('*', '').strip()
        multiline_keywords = ['notes', 'description', 'comments', 'remarks', 'details', 'observations']
        
        return any(keyword in field_lower for keyword in multiline_keywords)

    def get_field_content_sample(self, field_name: str, template_path: str = None) -> Optional[str]:
        """Get a sample of content from a field to help determine if it's multiline."""
        # Get the database column name for this field
        column_mapping = self.get_dynamic_column_mapping(template_path) if template_path else {}
        db_column = column_mapping.get(field_name)
        
        if not db_column:
            db_column = self._generate_safe_column_name(field_name)
        
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get a sample of non-empty values from this field
                cursor.execute(f"""
                    SELECT {db_column} 
                    FROM assets 
                    WHERE {db_column} IS NOT NULL 
                    AND {db_column} != '' 
                    AND is_deleted = 0
                    LIMIT 5
                """)
                
                rows = cursor.fetchall()
                if rows:
                    # Return the first non-empty value as a sample
                    for row in rows:
                        if row[0] and row[0].strip():
                            return row[0]
                            
        except Exception as e:
            print(f"Error getting field content sample for {field_name}: {e}")
        
        return None


# Global database instance
db = AssetDatabase()


def migrate_existing_csvs(csv_directory: str = "assets/output_files") -> Dict[str, int]:
    """Migrate existing CSV files to the database."""
    results = {}
    
    if not os.path.exists(csv_directory):
        return results
    
    for filename in os.listdir(csv_directory):
        if filename.endswith('.csv'):
            csv_path = os.path.join(csv_directory, filename)
            try:
                count = db.import_csv_template(csv_path)
                results[filename] = count
            except Exception as e:
                results[filename] = f"Error: {str(e)}"
    
    return results


if __name__ == "__main__":
    # Test the database functionality
    print("Testing Asset Database...")
    
    # Create database
    db = AssetDatabase("test_assets.db")
    
    # Test adding an asset
    test_asset = {
        "asset_type": "Computer",
        "manufacturer": "Dell",
        "model": "OptiPlex 7090",
        "serial_number": "TEST123456",
        "status": "Active",
        "location": "Building A",
        "room": "101",
        "system_name": "WORKSTATION-01"
    }
    
    asset_id = db.add_asset(test_asset)
    print(f"Added test asset with ID: {asset_id}")
    
    # Test searching
    assets = db.search_assets({"manufacturer": "Dell"})
    print(f"Found {len(assets)} Dell assets")
    
    # Test stats
    stats = db.get_database_stats()
    print(f"Database stats: {stats}")
    
    print("Database test completed successfully!")
