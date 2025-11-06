"""
Database service layer providing high-level operations for the asset management system.
Abstracts common database patterns and provides reusable methods.
"""

from typing import Dict, List, Optional, Any, Callable, Tuple, Union
from datetime import datetime, timedelta
import csv
import os
import shutil
from asset_database import AssetDatabase
from config_manager import ConfigManager

class DatabaseService:
    """High-level database service providing common operations."""
    
    def __init__(self):
        self.config = ConfigManager().get_config()
        self.db = AssetDatabase(self.config.database_path)
    
    def get_database_instance(self) -> AssetDatabase:
        """Get the underlying database instance."""
        return self.db
    
    def ensure_template_compatibility(self, template_path: str) -> bool:
        """Ensure database schema matches template requirements."""
        if not os.path.exists(template_path):
            return False
        
        try:
            return self.db.update_schema_for_template(template_path)
        except Exception as e:
            print(f"Warning: Could not update database schema: {e}")
            return False
    
    def get_dropdown_values(self, template_path: str, dropdown_fields: List[str]) -> Dict[str, List[str]]:
        """Get unique values for dropdown fields from database."""
        unique_values = {}
        
        try:
            column_mapping = self.db.get_dynamic_column_mapping(template_path)
            
            for field in dropdown_fields:
                db_column = column_mapping.get(field)
                if db_column:
                    values = self.db.get_unique_values(db_column)
                    if values:
                        filtered_values = [v for v in values if v and v.strip()]
                        if filtered_values:
                            unique_values[field] = sorted(filtered_values)
        except Exception as e:
            print(f"Warning: Could not load dropdown values: {e}")
        
        return unique_values
    
    def add_asset_from_form(self, form_data: Dict[str, str], template_path: str) -> Optional[int]:
        """Add an asset from form data using template mapping."""
        try:
            column_mapping = self.db.get_dynamic_column_mapping(template_path)
            asset_data = {}
            
            for header, value in form_data.items():
                if value and value.strip():
                    db_column = column_mapping.get(header)
                    if db_column:
                        asset_data[db_column] = value.strip()
            
            return self.db.add_asset(asset_data)
        except Exception as e:
            print(f"Error adding asset: {e}")
            return None
    
    def search_assets_with_filters(self, filters: Dict[str, Any] = None, 
                                  limit: int = 1000) -> List[Dict[str, Any]]:
        """Search assets with optional filters."""
        try:
            return self.db.search_assets(filters or {}, limit=limit)
        except Exception as e:
            print(f"Error searching assets: {e}")
            return []
    
    def get_recently_modified_assets(self, days: Union[int, float] = 30, 
                                   exclude_new: bool = True) -> List[Dict[str, Any]]:
        """Get assets modified in the last N days."""
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_str = cutoff_date.isoformat()
        
        try:
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                if exclude_new:
                    # Exclude assets where modified_date = created_date
                    query = """
                    SELECT * FROM assets 
                    WHERE modified_date >= ? 
                    AND modified_date != created_date
                    ORDER BY modified_date DESC, id
                    """
                else:
                    query = """
                    SELECT * FROM assets 
                    WHERE modified_date >= ?
                    ORDER BY modified_date DESC, id
                    """
                
                cursor.execute(query, (cutoff_str,))
                results = cursor.fetchall()
                
                columns = [desc[0] for desc in cursor.description]
                return [dict(zip(columns, row)) for row in results]
        except Exception as e:
            print(f"Error getting recently modified assets: {e}")
            return []
    
    def get_recently_added_assets(self, days: Union[int, float] = 30) -> List[Dict[str, Any]]:
        """Get assets added in the last N days (manual entries only, excludes imports)."""
        cutoff_date = datetime.now() - timedelta(days=days)
        cutoff_str = cutoff_date.isoformat()
        
        try:
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                query = """
                SELECT * FROM assets 
                WHERE created_date >= ? 
                AND data_source = 'manual'
                AND is_deleted = 0
                ORDER BY created_date DESC, id
                """
                cursor.execute(query, (cutoff_str,))
                results = cursor.fetchall()
                
                columns = [desc[0] for desc in cursor.description]
                return [dict(zip(columns, row)) for row in results]
        except Exception as e:
            print(f"Error getting recently added assets: {e}")
            return []
    
    def export_assets_to_csv(self, assets: List[Dict[str, Any]], 
                           file_path: str, template_path: str = None) -> bool:
        """Export assets to CSV with optional template formatting."""
        try:
            if template_path and os.path.exists(template_path):
                # Export using template headers
                with open(template_path, 'r', newline='', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    template_headers = next(reader, [])
                
                if template_headers:
                    column_mapping = self.db.get_dynamic_column_mapping(template_path)
                    
                    with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                        writer = csv.DictWriter(csvfile, fieldnames=template_headers)
                        writer.writeheader()
                        
                        for asset in assets:
                            row = {}
                            for header in template_headers:
                                db_column = column_mapping.get(header, 
                                    header.lower().replace(' ', '_').replace('*', ''))
                                row[header] = asset.get(db_column, '')
                            writer.writerow(row)
                    return True
            
            # Export all fields if no template or template not found
            if assets:
                all_columns = list(assets[0].keys())
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=all_columns)
                    writer.writeheader()
                    for asset in assets:
                        writer.writerow(asset)
                return True
            
        except Exception as e:
            print(f"Error exporting assets: {e}")
            
        return False
    
    def import_assets_from_csv(self, csv_path: str, 
                             duplicate_handler: Callable = None) -> int:
        """Import assets from CSV with duplicate handling."""
        try:
            return self.db.import_from_csv(csv_path, duplicate_handler)
        except Exception as e:
            print(f"Error importing assets: {e}")
            return 0
    
    def backup_database(self, backup_path: str = None) -> Optional[str]:
        """Create a database backup."""
        if not backup_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = f"{self.config.database_path}_backup_{timestamp}"
        
        try:
            shutil.copy2(self.config.database_path, backup_path)
            return backup_path
        except Exception as e:
            print(f"Error creating backup: {e}")
            return None

    def create_automatic_backup(self, max_backups: int = 5, silent: bool = True) -> bool:
        """
        Create automatic rotating backup of database.
        
        Args:
            max_backups: Maximum number of auto-backups to keep (default: 5)
            silent: If True, only log to console. If False, could raise exceptions.
            
        Returns:
            bool: True if backup was successful, False otherwise
        """
        try:
            db_path = self.config.database_path
            
            # Skip if no database path configured or file doesn't exist
            if not db_path or not os.path.exists(db_path):
                if not silent:
                    print("No database file found for automatic backup")
                return False
            
            # Get the directory and base name for backup files
            db_dir = os.path.dirname(db_path)
            db_name = os.path.basename(db_path)
            db_base, db_ext = os.path.splitext(db_name)
            
            # Create backup directory if it doesn't exist
            backup_dir = os.path.join(db_dir, "auto_backups")
            os.makedirs(backup_dir, exist_ok=True)
            
            # Rotate existing backups (max -> delete, max-1 -> max, ..., 1 -> 2)
            for i in range(max_backups, 0, -1):
                backup_file = os.path.join(backup_dir, f"{db_base}_autobackup_{i}{db_ext}")
                
                if i == max_backups:
                    # Delete the oldest backup
                    if os.path.exists(backup_file):
                        os.remove(backup_file)
                        if not silent:
                            print(f"Deleted oldest backup: {backup_file}")
                else:
                    # Move backup_i to backup_i+1
                    next_backup = os.path.join(backup_dir, f"{db_base}_autobackup_{i+1}{db_ext}")
                    if os.path.exists(backup_file):
                        shutil.move(backup_file, next_backup)
                        if not silent:
                            print(f"Rotated backup: {backup_file} -> {next_backup}")
            
            # Create new backup as autobackup_1
            new_backup = os.path.join(backup_dir, f"{db_base}_autobackup_1{db_ext}")
            shutil.copy2(db_path, new_backup)
            
            # Log timestamp info
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if not silent:
                print(f"Created automatic backup: {new_backup} at {timestamp}")
            
            return True
            
        except Exception as e:
            error_msg = f"Failed to create automatic backup: {e}"
            if silent:
                print(f"Warning: {error_msg}")
                return False
            else:
                raise RuntimeError(error_msg)

    def get_database_statistics(self) -> Dict[str, Any]:
        """Get comprehensive database statistics."""
        try:
            stats = self.db.get_database_stats()
            
            # Add additional statistics
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get table size info
                cursor.execute("PRAGMA page_count")
                page_count = cursor.fetchone()[0]
                cursor.execute("PRAGMA page_size")
                page_size = cursor.fetchone()[0]
                
                stats.update({
                    'database_size_bytes': page_count * page_size,
                    'database_size_mb': (page_count * page_size) / (1024 * 1024),
                    'page_count': page_count,
                    'page_size': page_size
                })
                
                # Get recent activity
                cursor.execute("""
                    SELECT COUNT(*) FROM assets 
                    WHERE created_date >= datetime('now', '-7 days')
                """)
                stats['assets_added_last_week'] = cursor.fetchone()[0]
                
                cursor.execute("""
                    SELECT COUNT(*) FROM assets 
                    WHERE modified_date >= datetime('now', '-7 days')
                    AND modified_date != created_date
                """)
                stats['assets_modified_last_week'] = cursor.fetchone()[0]
            
            return stats
        except Exception as e:
            print(f"Error getting database statistics: {e}")
            return {}

# Singleton instance for global access
database_service = DatabaseService()
