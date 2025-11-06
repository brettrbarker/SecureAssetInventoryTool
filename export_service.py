"""
Export service providing common export functionality for the asset management system.
Centralizes export operations to avoid code duplication between main menu and settings menu.
"""

import csv
import os
import sqlite3
from datetime import datetime
from tkinter import messagebox, filedialog
from typing import Optional, Dict, Any, List
import customtkinter as ctk

from asset_database import AssetDatabase
from config_manager import ConfigManager
from error_handling import error_handler
from performance_monitoring import performance_monitor
from database_service import database_service


class ExportService:
    """Centralized export service for asset data."""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.config = self.config_manager.get_config()
    
    @performance_monitor("Export Database Template")
    def export_database_template(self, parent_window=None) -> bool:
        """
        Export database using template fields with filtering options.
        
        Args:
            parent_window: Parent window for dialogs (optional)
            
        Returns:
            bool: True if export was successful, False otherwise
        """
        try:
            # First, show filtering options
            filter_option = self._show_export_filter_dialog(parent_window)
            if not filter_option:
                return False
            
            # Get template path and check existence
            template_path = self.config.default_template_path
            if not os.path.exists(template_path):
                messagebox.showerror("Error", f"Template file not found: {template_path}", parent=parent_window)
                return False
            
            # Get assets based on filter option
            current_db = self.config.database_path
            db = AssetDatabase(current_db)
            assets = self._get_filtered_assets(db, filter_option, parent_window)
            if assets is None:  # User cancelled
                return False
            
            if not assets:
                messagebox.showinfo("No Data", "No assets found matching the selected criteria.", parent=parent_window)
                return False
            
            # Get export options from filter_option
            export_options = filter_option.get("export_options", {})
            filter_asset_number_for_non_imported = export_options.get("filter_asset_number_for_non_imported", True)
            filter_sync_keys_from_imported = export_options.get("filter_sync_keys_from_imported", True)
            filter_manufacturer_from_imported = export_options.get("filter_manufacturer_from_imported", True)
            
            # Get save location for export file with filter type in filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filter_type = filter_option.get("type", "all")
            
            # Map filter types to filename suffixes
            filter_suffix_map = {
                "all": "all",
                "modified": "modified", 
                "added": "added",
                "both": "modified_and_added"
            }
            
            filter_suffix = filter_suffix_map.get(filter_type, "all")
            default_filename = f"asset_export_{filter_suffix}_{timestamp}.csv"
            
            # Create exports subdirectory in output directory
            exports_dir = os.path.join(self.config.output_directory, "exports")
            os.makedirs(exports_dir, exist_ok=True)
            
            export_path = filedialog.asksaveasfilename(
                title="Save Export File",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=default_filename,
                initialdir=exports_dir,
                parent=parent_window
            )
            
            if not export_path:
                return False
            
            # Get template headers to filter export columns
            template_headers = []
            try:
                with open(template_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    template_headers = next(reader, [])
            except Exception as e:
                messagebox.showerror("Template Error", f"Could not read template file:\n{str(e)}", parent=parent_window)
                return False
            
            # Write to CSV with template columns only
            try:
                with open(export_path, 'w', newline='', encoding='utf-8') as csvfile:
                    if template_headers:
                        writer = csv.DictWriter(csvfile, fieldnames=template_headers)
                        writer.writeheader()
                        
                        # Get column mapping from template headers to database columns
                        column_mapping = db.get_dynamic_column_mapping(template_path)
                        
                        # Export each asset, but only include template columns
                        for asset in assets:
                            asset_dict = dict(asset) if hasattr(asset, '_asdict') else asset
                            
                            # Check if this is a manually added asset for Asset Number export option
                            is_manually_added = asset_dict.get('data_source', '') == 'manual'
                            
                            # Map template headers to database columns and get values
                            filtered_asset = {}
                            for header in template_headers:
                                db_column = column_mapping.get(header, header)  # fallback to header name
                                asset_value = asset_dict.get(db_column, '')
                                
                                # Special handling for Asset Number export option
                                if header.lower() in ['asset no.', 'asset no', 'asset number', 'assetno']:
                                    if filter_asset_number_for_non_imported:
                                        # Filter out Asset Numbers for non-imported (manually added) assets when checked
                                        if not is_manually_added:
                                            filtered_asset[header] = asset_value  # Export for imported assets
                                        else:
                                            filtered_asset[header] = ''  # Leave blank for manually added assets
                                    else:
                                        # Export all Asset Numbers when unchecked
                                        filtered_asset[header] = asset_value
                                # Special handling for Related Asset Sync Keys filtering
                                elif header.lower() in ['related asset sync keys', 'related asset sync key']:
                                    if filter_sync_keys_from_imported:
                                        # Only export sync keys for manually added assets when checked
                                        if is_manually_added:
                                            filtered_asset[header] = asset_value
                                        else:
                                            filtered_asset[header] = ''  # Leave blank for imported assets
                                    else:
                                        # Normal behavior - export all sync keys
                                        filtered_asset[header] = asset_value
                                # Special handling for Manufacturer filtering as safeguard
                                elif header.lower() in ['*manufacturer', 'manufacturer']:
                                    if filter_manufacturer_from_imported:
                                        # Only export manufacturer for manually added assets when checked
                                        if is_manually_added:
                                            filtered_asset[header] = asset_value
                                        else:
                                            filtered_asset[header] = ''  # Leave blank for imported assets
                                    else:
                                        # Normal behavior - export all manufacturers
                                        filtered_asset[header] = asset_value
                                else:
                                    filtered_asset[header] = asset_value
                                    
                            writer.writerow(filtered_asset)
                    else:
                        messagebox.showerror("Template Error", "Template file appears to be empty", parent=parent_window)
                        return False
                        
                messagebox.showinfo("Export Complete", 
                                  f"Successfully exported {len(assets)} assets to:\n{export_path}", 
                                  parent=parent_window)
                return True
                
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to write export file:\n{str(e)}", parent=parent_window)
                return False
                
        except Exception as e:
            error_handler.logger.error(f"Error in exporting database template: {str(e)}", exception=e)
            messagebox.showerror("Export Error", f"Failed to export database template:\n{str(e)}", parent=parent_window)
            return False
    
    def _show_export_filter_dialog(self, parent_window=None):
        """Show dialog to select export filtering options."""
        dialog = ctk.CTkToplevel(parent_window if parent_window else None)
        dialog.title("Export Filter Options")
        dialog.geometry("800x550")
        dialog.minsize(800, 550)
        if parent_window:
            dialog.transient(parent_window)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (800 // 2)
        y = (dialog.winfo_screenheight() // 2) - (550 // 2)
        dialog.geometry(f"800x550+{x}+{y}")
        
        result = {"choice": None}
        
        # Create main container with fixed button area at bottom
        main_container = ctk.CTkFrame(dialog)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Content frame with two columns (no longer scrollable)
        content_frame = ctk.CTkFrame(main_container)
        content_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Left column - Filter options
        filter_frame = ctk.CTkFrame(content_frame)
        filter_frame.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        
        filter_title = ctk.CTkLabel(filter_frame, text="Filter Options", 
                                   font=ctk.CTkFont(size=14, weight="bold"))
        filter_title.pack(pady=(15, 10))
        
        # Right column - Export options
        export_options_frame = ctk.CTkFrame(content_frame)
        export_options_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        
        export_options_title = ctk.CTkLabel(export_options_frame, text="Export Options", 
                                           font=ctk.CTkFont(size=14, weight="bold"))
        export_options_title.pack(pady=(15, 10))
        
        # Configure grid weights for equal column distribution
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)
        
        # Radio button variable
        selection_var = ctk.StringVar(value="both")
        
        # All assets option
        all_radio = ctk.CTkRadioButton(filter_frame, text="All Assets", 
                                      variable=selection_var, value="all")
        all_radio.pack(anchor="w", padx=20, pady=10)
        
        # Recently modified option with day counter
        modified_frame = ctk.CTkFrame(filter_frame)
        modified_frame.pack(fill="x", padx=20, pady=8)
        
        modified_radio = ctk.CTkRadioButton(modified_frame, text="Recently Modified Assets", 
                                           variable=selection_var, value="modified")
        modified_radio.pack(anchor="w", padx=10, pady=10)
        
        # Day counter for modified
        modified_days_frame = ctk.CTkFrame(modified_frame)
        modified_days_frame.pack(fill="x", padx=30, pady=5)
        
        ctk.CTkLabel(modified_days_frame, text="Last").pack(side="left", padx=(10, 5))
        modified_days_var = ctk.StringVar(value="0.5")
        modified_days_spinbox = ctk.CTkFrame(modified_days_frame)
        modified_days_spinbox.pack(side="left", padx=5)
        
        modified_days_entry = ctk.CTkEntry(modified_days_spinbox, textvariable=modified_days_var, width=50)
        modified_days_entry.pack(side="left")
        
        def increment_modified_days():
            try:
                current = float(modified_days_var.get())
                if current == 0.5:
                    new_val = 1  # From 0.5 to 1
                else:
                    new_val = current + 1  # Normal increment
                
                # Format the value appropriately
                if new_val == int(new_val):
                    modified_days_var.set(str(int(new_val)))
                else:
                    modified_days_var.set(str(new_val))
            except ValueError:
                modified_days_var.set("1")
        
        def decrement_modified_days():
            try:
                current = float(modified_days_var.get())
                if current > 1:
                    new_val = current - 1
                elif current == 1:
                    new_val = 0.5  # Allow going down to half day (12 hours)
                else:
                    new_val = 0.5  # Don't go below 0.5
                
                # Format the value appropriately
                if new_val == int(new_val):
                    modified_days_var.set(str(int(new_val)))
                else:
                    modified_days_var.set(str(new_val))
            except ValueError:
                modified_days_var.set("1")
        
        up_button = ctk.CTkButton(modified_days_spinbox, text="▲", width=25, height=20, 
                                 command=increment_modified_days)
        up_button.pack(side="right", pady=(0, 1))
        down_button = ctk.CTkButton(modified_days_spinbox, text="▼", width=25, height=20, 
                                   command=decrement_modified_days)
        down_button.pack(side="right")
        
        ctk.CTkLabel(modified_days_frame, text="days").pack(side="left", padx=(5, 10))
        
        # Recently added option with day counter
        added_frame = ctk.CTkFrame(filter_frame)
        added_frame.pack(fill="x", padx=20, pady=8)
        
        added_radio = ctk.CTkRadioButton(added_frame, text="Recently Added Assets (Manual Only)", 
                                        variable=selection_var, value="added")
        added_radio.pack(anchor="w", padx=10, pady=10)
        
        # Day counter for added
        added_days_frame = ctk.CTkFrame(added_frame)
        added_days_frame.pack(fill="x", padx=30, pady=5)
        
        ctk.CTkLabel(added_days_frame, text="Last").pack(side="left", padx=(10, 5))
        added_days_var = ctk.StringVar(value="0.5")
        added_days_spinbox = ctk.CTkFrame(added_days_frame)
        added_days_spinbox.pack(side="left", padx=5)
        
        added_days_entry = ctk.CTkEntry(added_days_spinbox, textvariable=added_days_var, width=50)
        added_days_entry.pack(side="left")
        
        def increment_added_days():
            try:
                current = float(added_days_var.get())
                if current == 0.5:
                    new_val = 1  # From 0.5 to 1
                else:
                    new_val = current + 1  # Normal increment
                
                # Format the value appropriately
                if new_val == int(new_val):
                    added_days_var.set(str(int(new_val)))
                else:
                    added_days_var.set(str(new_val))
            except ValueError:
                added_days_var.set("1")
        
        def decrement_added_days():
            try:
                current = float(added_days_var.get())
                if current > 1:
                    new_val = current - 1
                elif current == 1:
                    new_val = 0.5  # Allow going down to half day (12 hours)
                else:
                    new_val = 0.5  # Don't go below 0.5
                
                # Format the value appropriately
                if new_val == int(new_val):
                    added_days_var.set(str(int(new_val)))
                else:
                    added_days_var.set(str(new_val))
            except ValueError:
                added_days_var.set("1")
        
        up_button2 = ctk.CTkButton(added_days_spinbox, text="▲", width=25, height=20, 
                                  command=increment_added_days)
        up_button2.pack(side="right", pady=(0, 1))
        down_button2 = ctk.CTkButton(added_days_spinbox, text="▼", width=25, height=20, 
                                    command=decrement_added_days)
        down_button2.pack(side="right")
        
        ctk.CTkLabel(added_days_frame, text="days").pack(side="left", padx=(5, 10))
        
        # Both recently modified and added option with day counter
        both_frame = ctk.CTkFrame(filter_frame)
        both_frame.pack(fill="x", padx=20, pady=8)
        
        both_radio = ctk.CTkRadioButton(both_frame, text="Both Recently Modified and Added Assets", 
                                       variable=selection_var, value="both")
        both_radio.pack(anchor="w", padx=10, pady=10)
        
        # Day counter for both
        both_days_frame = ctk.CTkFrame(both_frame)
        both_days_frame.pack(fill="x", padx=30, pady=5)
        
        ctk.CTkLabel(both_days_frame, text="Last").pack(side="left", padx=(10, 5))
        both_days_var = ctk.StringVar(value="0.5")
        both_days_spinbox = ctk.CTkFrame(both_days_frame)
        both_days_spinbox.pack(side="left", padx=5)
        
        both_days_entry = ctk.CTkEntry(both_days_spinbox, textvariable=both_days_var, width=50)
        both_days_entry.pack(side="left")
        
        def increment_both_days():
            try:
                current = float(both_days_var.get())
                if current == 0.5:
                    new_val = 1  # From 0.5 to 1
                else:
                    new_val = current + 1  # Normal increment
                
                # Format the value appropriately
                if new_val == int(new_val):
                    both_days_var.set(str(int(new_val)))
                else:
                    both_days_var.set(str(new_val))
            except ValueError:
                both_days_var.set("1")
        
        def decrement_both_days():
            try:
                current = float(both_days_var.get())
                if current > 1:
                    new_val = current - 1
                elif current == 1:
                    new_val = 0.5  # Allow going down to half day (12 hours)
                else:
                    new_val = 0.5  # Don't go below 0.5
                
                # Format the value appropriately
                if new_val == int(new_val):
                    both_days_var.set(str(int(new_val)))
                else:
                    both_days_var.set(str(new_val))
            except ValueError:
                both_days_var.set("1")
        
        up_button3 = ctk.CTkButton(both_days_spinbox, text="▲", width=25, height=20, 
                                  command=increment_both_days)
        up_button3.pack(side="right", pady=(0, 1))
        down_button3 = ctk.CTkButton(both_days_spinbox, text="▼", width=25, height=20, 
                                    command=decrement_both_days)
        down_button3.pack(side="right")
        
        ctk.CTkLabel(both_days_frame, text="days").pack(side="left", padx=(5, 10))
        
        # Export options in the right column (already created above)
        # Export Asset Number for Manually Added Assets checkbox
        export_asset_number_var = ctk.BooleanVar(value=True)
        export_asset_number_cb = ctk.CTkCheckBox(export_options_frame, 
                                                text="Filter Asset Number for Non-Imported Assets *",
                                                variable=export_asset_number_var)
        export_asset_number_cb.pack(anchor="w", padx=20, pady=10)
        
        # Filter Related Asset Sync Keys from Imported checkbox
        filter_sync_keys_var = ctk.BooleanVar(value=True)
        filter_sync_keys_cb = ctk.CTkCheckBox(export_options_frame, 
                                            text="Filter Related Asset Sync Keys from Imported *",
                                            variable=filter_sync_keys_var)
        filter_sync_keys_cb.pack(anchor="w", padx=20, pady=10)
        
        # Filter Manufacturer from Imported as Safeguard checkbox
        filter_manufacturer_var = ctk.BooleanVar(value=True)
        filter_manufacturer_cb = ctk.CTkCheckBox(export_options_frame, 
                                               text="Filter Manufacturer from Imported as Safeguard *",
                                               variable=filter_manufacturer_var)
        filter_manufacturer_cb.pack(anchor="w", padx=20, pady=10)
        
        # Add footnote for the asterisked options
        footnote_label = ctk.CTkLabel(export_options_frame, 
                                     text="* Recommended if using SW Help Desk Asset System",
                                     font=ctk.CTkFont(size=12),
                                     text_color="gray")
        footnote_label.pack(anchor="w", padx=20, pady=(15, 10))
        
        # Add some bottom padding to balance the layout
        ctk.CTkLabel(export_options_frame, text="").pack(pady=(10, 0))
        
        # Fixed buttons frame at bottom (like settings menu)
        buttons_frame = ctk.CTkFrame(main_container)
        buttons_frame.pack(fill="x", side="bottom", pady=(0, 0))
        buttons_frame.pack_propagate(False)  # Prevent frame from shrinking
        buttons_frame.configure(height=60)  # Fixed height
        
        def on_ok():
            choice = selection_var.get()
            
            # Capture export options
            export_options = {
                "filter_asset_number_for_non_imported": export_asset_number_var.get(),
                "filter_sync_keys_from_imported": filter_sync_keys_var.get(),
                "filter_manufacturer_from_imported": filter_manufacturer_var.get()
            }
            
            if choice in ["modified", "added", "both"]:
                # Get the day count for the selected option
                if choice == "modified":
                    try:
                        days = float(modified_days_var.get())
                    except ValueError:
                        days = 1
                elif choice == "added":
                    try:
                        days = float(added_days_var.get())
                    except ValueError:
                        days = 1
                else:  # both
                    try:
                        days = float(both_days_var.get())
                    except ValueError:
                        days = 1
                
                result["choice"] = {
                    "type": choice,
                    "days": days,
                    "export_options": export_options
                }
            else:
                result["choice"] = {
                    "type": choice,
                    "export_options": export_options
                }
            dialog.destroy()
        
        def on_cancel():
            result["choice"] = None
            dialog.destroy()
        
        ok_button = ctk.CTkButton(buttons_frame, text="OK", command=on_ok, width=100)
        ok_button.pack(side="left", padx=(15, 10), pady=15)
        
        cancel_button = ctk.CTkButton(buttons_frame, text="Cancel", command=on_cancel, width=100)
        cancel_button.pack(side="right", padx=(10, 15), pady=15)
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result["choice"]
    
    def _get_filtered_assets(self, db: AssetDatabase, filter_option: Dict[str, Any], parent_window=None):
        """Get assets based on filter criteria using existing database service methods."""
        if not filter_option:
            return None
        
        try:
            filter_type = filter_option["type"]
            
            if filter_type == "all":
                return db.search_assets(limit=999999)  # Get all assets
            
            elif filter_type in ["modified", "added", "both"]:
                days = filter_option.get("days", 1)
                
                # Use existing database service methods
                if filter_type == "modified":
                    filtered_assets = database_service.get_recently_modified_assets(days)
                elif filter_type == "added":
                    filtered_assets = database_service.get_recently_added_assets(days)
                elif filter_type == "both":
                    # Get both modified and added assets, then combine and deduplicate
                    modified_assets = database_service.get_recently_modified_assets(days)
                    added_assets = database_service.get_recently_added_assets(days)
                    
                    # Combine lists and deduplicate by asset ID
                    combined_assets = []
                    seen_ids = set()
                    
                    # Add modified assets first
                    for asset in modified_assets:
                        asset_id = asset.get('id')
                        if asset_id not in seen_ids:
                            combined_assets.append(asset)
                            seen_ids.add(asset_id)
                    
                    # Add added assets if not already included
                    for asset in added_assets:
                        asset_id = asset.get('id')
                        if asset_id not in seen_ids:
                            combined_assets.append(asset)
                            seen_ids.add(asset_id)
                    
                    # Sort by most recent activity (created_date or modified_date)
                    def get_most_recent_date(asset):
                        created = asset.get('created_date', '1900-01-01')
                        modified = asset.get('modified_date', '1901-01-01')
                        # Return the more recent of the two dates
                        return max(created, modified)
                    
                    combined_assets.sort(key=get_most_recent_date, reverse=True)
                    filtered_assets = combined_assets
                
                if not filtered_assets:
                    filter_text = "modified or added" if filter_type == "both" else filter_type
                    # Format the days value appropriately for display
                    if days == 0.5:
                        time_text = "12 hours"
                    elif days == 1:
                        time_text = "1 day"
                    else:
                        time_text = f"{days} days"
                    
                    messagebox.showinfo("No Assets Found", 
                                      f"No assets found {filter_text} in the last {time_text}.", 
                                      parent=parent_window)
                    return []
                
                filter_text = "modified or added" if filter_type == "both" else filter_type
                # Format the days value appropriately for display
                if days == 0.5:
                    time_text = "12 hours"
                elif days == 1:
                    time_text = "1 day"
                else:
                    time_text = f"{days} days"
                
                messagebox.showinfo("Filtering Applied", 
                                  f"Found {len(filtered_assets)} assets {filter_text} in the last {time_text}.", 
                                  parent=parent_window)
                
                return filtered_assets
            
            else:
                return db.search_assets(limit=999999)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to retrieve assets:\n{str(e)}", parent=parent_window)
            return None


# Global export service instance
export_service = ExportService()
