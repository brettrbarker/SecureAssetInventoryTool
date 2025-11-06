"""
Bulk Update Assets window for updating multiple fields of existing assets.
"""

import customtkinter as ctk
import os
import csv
from datetime import datetime
from tkinter import messagebox
import tkinter as tk
from typing import Dict, List, Any, Optional
from asset_database import AssetDatabase
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from performance_monitoring import performance_monitor
from ui_components import SearchableDropdown, DatePicker, EmbeddedAssetDetail
from field_utils import compute_db_fields_from_template, compute_dropdown_fields, compute_date_fields
from edit_asset import EditAssetWindow


def _today_audit_date_str() -> str:
    """Return today's date in requested format: MM/D/YYYY (month zero-padded, day without leading zero)."""
    now = datetime.now()
    return f"{now:%m}/{now.day}/{now:%Y}"


class BulkUpdateWindow:
    """Window for bulk updating asset fields."""
    
    def __init__(self, parent, config=None):
        """Initialize the bulk update window."""
        self.parent = parent
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = config or self.config_manager.get_config()
        
        # Create database instance using the configured database path
        self.db = AssetDatabase(self.config.database_path)
        
        # Get template path for multiline field detection
        self.template_path = self.config.default_template_path
        
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Bulk Update Assets")
        self.window.geometry("1200x700")
        self.window.minsize(900, 700)  # Set minimum window size
        self.window.transient(parent)
        # Removed grab_set() to allow interaction with other windows like Monitor
        # self.window.grab_set()
        
        # Center the window
        self._center_window()
        
        # Get database fields constrained to current template and config
        self.db_fields = compute_db_fields_from_template(self.db, self.config)
        self.dropdown_fields = compute_dropdown_fields(self.db_fields, self.config)
        self.date_fields = compute_date_fields(self.db_fields)
        
        # Variables for bulk changes
        self.bulk_change_rows = []
        self.selected_asset_data = None
        
        # Search variables
        self.search_field = tk.StringVar()
        self.search_value = tk.StringVar()
        
        # Set default search field to one containing "serial"
        default_field = self._find_default_search_field()
        if default_field:
            self.search_field.set(default_field)
        
        self._create_widgets()
        self._setup_key_bindings()
        
        # Initialize status indicator as hidden
        self._update_status_indicator(found=None)
        
        # Focus handling
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.window.focus_force()
    
    def _center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        
        # Get screen dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calculate center position
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Set window position
        self.window.geometry(f"1200x700+{x}+{y}")
    
    # Legacy helper methods removed in favor of shared field_utils helpers
    
    def _find_default_search_field(self):
        """Find a field containing 'serial' for default search field."""
        for field in self.db_fields:
            if 'serial' in field['display_name'].lower() or 'serial' in field['db_name'].lower():
                return field['display_name']
        
        # Fallback to first field if no serial field found
        if self.db_fields:
            return self.db_fields[0]['display_name']
        
        return ""
    
    def _create_widgets(self):
        """Create all window widgets."""
        # Create main scrollable frame
        self.main_frame = ctk.CTkScrollableFrame(self.window)
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=(10, 5))  # Reduced padding
        
        # Create sections
        self._create_bulk_changes_section()
        self._create_search_section()
        self._create_asset_display_section()
        
        # Create bottom action frame (anchored to bottom)
        self._create_action_section()
    
    def _create_bulk_changes_section(self):
        """Create the bulk changes section at the top."""
        # Bulk Changes Section
        changes_frame = ctk.CTkFrame(self.main_frame)
        changes_frame.pack(fill="x", pady=(0, 10))  # Reduced from pady=(0, 20)
        
        # Title and Presets row
        title_row = ctk.CTkFrame(changes_frame, fg_color="transparent")
        title_row.pack(fill="x", pady=(10, 5), padx=20)  # Reduced from pady=(15, 10)
        
        # Section title
        title_label = ctk.CTkLabel(title_row, 
                                  text="Bulk Changes", 
                                  font=ctk.CTkFont(size=18, weight="bold"))  # Reduced from size=20
        title_label.pack(side="left")
        
        # Presets dropdown
        presets_frame = ctk.CTkFrame(title_row, fg_color="transparent")
        presets_frame.pack(side="right")
        
        presets_label = ctk.CTkLabel(presets_frame, text="Presets:", font=ctk.CTkFont(size=12))  # Reduced from size=14
        presets_label.pack(side="left", padx=(0, 10))
        
        self.presets_var = ctk.StringVar(value="Select Preset...")
        preset_options = ["Select Preset...", "Save Current as Preset..."] + list(self.config.bulk_update_presets.keys())
        
        self.presets_dropdown = ctk.CTkComboBox(presets_frame,
                                               variable=self.presets_var,
                                               values=preset_options,
                                               command=self._on_preset_selected,
                                               width=200)
        self.presets_dropdown.pack(side="left")
        
        # Container for change rows
        self.changes_container = ctk.CTkFrame(changes_frame)
        self.changes_container.pack(fill="x", padx=20, pady=(0, 8))  # Reduced from pady=(0, 15)
        
        # Add first row by default
        self._add_change_row()
        
        # Add Row button
        add_row_btn = ctk.CTkButton(changes_frame, 
                                   text="+ Add Field to Change", 
                                   command=self._add_change_row,
                                   width=200,
                                   height=28)  # Reduced height from default
        add_row_btn.pack(pady=(0, 8))  # Reduced from pady=(0, 15)
    
    def _create_search_section(self):
        """Create the search section in the middle."""
        # Search Section
        search_frame = ctk.CTkFrame(self.main_frame)
        search_frame.pack(fill="x", pady=(0, 10))  # Reduced from pady=(0, 20)
        
        # Section title
        title_label = ctk.CTkLabel(search_frame, 
                                  text="Find Asset", 
                                  font=ctk.CTkFont(size=18, weight="bold"))  # Reduced from size=20
        title_label.pack(pady=(10, 5))  # Reduced from pady=(15, 10)
        
        # Search controls frame
        search_controls = ctk.CTkFrame(search_frame)
        search_controls.pack(fill="x", padx=20, pady=(0, 8))  # Reduced from pady=(0, 15)
        
        # Search field dropdown (use SearchableDropdown for consistency)
        field_label = ctk.CTkLabel(search_controls, text="Search Field:")
        field_label.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="w")  # Reduced from pady=10
        
        field_names = [field['display_name'] for field in self.db_fields]
        self.search_field_dropdown = SearchableDropdown(search_controls,
                                                       values=field_names,
                                                       variable=self.search_field,
                                                       width=200,
                                                       height=28)  # Reduced height
        self.search_field_dropdown.grid(row=0, column=1, padx=5, pady=5)  # Reduced from pady=10
        
        # Search value entry
        value_label = ctk.CTkLabel(search_controls, text="Search Value:")
        value_label.grid(row=0, column=2, padx=(20, 5), pady=5, sticky="w")  # Reduced from pady=10
        
        self.search_entry = ctk.CTkEntry(search_controls, 
                                        textvariable=self.search_value,
                                        width=300,
                                        height=28,  # Reduced height
                                        placeholder_text="Enter value to search for...")
        self.search_entry.grid(row=0, column=3, padx=5, pady=5)  # Reduced from pady=10
        
        # Search button
        search_btn = ctk.CTkButton(search_controls, 
                                  text="Search", 
                                  command=self._search_asset,
                                  width=100,
                                  height=28)  # Reduced height
        search_btn.grid(row=0, column=4, padx=(10, 10), pady=5)  # Reduced from pady=10
        
        # Configure grid weights
        search_controls.grid_columnconfigure(3, weight=1)
    
    def _create_asset_display_section(self):
        """Create the asset display section at the bottom."""
        # Asset Display Section
        display_frame = ctk.CTkFrame(self.main_frame)
        display_frame.pack(fill="both", expand=True)
        
        # Title frame to hold title and status indicator
        title_frame = ctk.CTkFrame(display_frame, fg_color="transparent")
        title_frame.pack(pady=(8, 5))  # Reduced from pady=(15, 10)
        
        # Section title
        title_label = ctk.CTkLabel(title_frame, 
                                  text="Selected Asset Details", 
                                  font=ctk.CTkFont(size=18, weight="bold"))  # Reduced from size=20
        title_label.pack(side="left", padx=(0, 10))
        
        # Status indicator (initially hidden)
        self.status_indicator = ctk.CTkLabel(title_frame, 
                                           text="", 
                                           font=ctk.CTkFont(size=20, weight="bold"))  # Reduced from size=24
        self.status_indicator.pack(side="left")
        
        # Button to open full Asset Detail Window (initially hidden)
        self.open_detail_btn = ctk.CTkButton(title_frame, 
                                            text="📋",  # Clipboard/document icon
                                            width=32, 
                                            height=28,
                                            command=self._open_full_asset_detail,
                                            fg_color="#1f538d", 
                                            hover_color="#14375e",
                                            font=ctk.CTkFont(size=16))
        self.open_detail_btn.pack(side="left", padx=(5, 0))
        self.open_detail_btn.pack_forget()  # Hide initially
        
        # Asset details container (regular frame to avoid nested scrollbars)
        self.asset_details_frame = ctk.CTkFrame(display_frame)
        self.asset_details_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))  # Reduced from pady=(0, 15)
        
        # Initially show "No asset selected" message
        self.no_asset_label = ctk.CTkLabel(self.asset_details_frame, 
                                          text="No asset selected. Use the search above to find an asset.",
                                          font=ctk.CTkFont(size=14))
        self.no_asset_label.pack(pady=30)  # Reduced from pady=50

    def _create_action_section(self):
        """Create the bottom action buttons section."""
        # Bottom action frame (anchored to bottom of window)
        action_frame = ctk.CTkFrame(self.window)
        action_frame.pack(fill="x", padx=20, pady=(0, 15))  # Reduced from pady=(0, 20)
        
        # Instructions
        instructions_label = ctk.CTkLabel(action_frame, 
                                        text="💡 Tips: Enter to search • Ctrl+Enter to apply changes • Ctrl+Backspace to clear search • Add rows to update multiple fields",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(pady=(8, 0))
        
        # Apply Changes button
        self.apply_btn = ctk.CTkButton(action_frame, 
                                      text="Apply Changes", 
                                      command=self._apply_changes,
                                      width=200,
                                      height=35,  # Reduced from height=40
                                      font=ctk.CTkFont(size=14, weight="bold"),  # Reduced from size=16
                                      state="disabled")
        self.apply_btn.pack(pady=8)  # Reduced from pady=10

    def _setup_key_bindings(self):
        """Set up keyboard shortcuts for the window."""
        # Return/Enter = Search
        self.window.bind('<Return>', lambda event: self._search_asset())
        self.window.bind('<KP_Enter>', lambda event: self._search_asset())
        
        # Ctrl+Enter = Apply Changes
        self.window.bind('<Control-Return>', lambda event: self._apply_changes())
        self.window.bind('<Control-KP_Enter>', lambda event: self._apply_changes())
        
        # Ctrl+Backspace = Clear search field and focus
        self.window.bind('<Control-BackSpace>', lambda event: self._clear_and_focus_search())
        
        # Make sure the window can receive focus for key events
        self.window.focus_set()

    def _clear_and_focus_search(self):
        """Clear the search field and set focus to it."""
        self.search_value.set("")
        self.search_entry.focus_set()

    def _update_status_indicator(self, found=None):
        """Update the status indicator next to the title.
        
        Args:
            found (bool or None): True for asset found (green check), 
                                 False for no asset found (red X),
                                 None to hide indicator
        """
        if found is True:
            # Green check mark in circle
            self.status_indicator.configure(text="✅", text_color="green")
            # Show the button to open full detail window
            self.open_detail_btn.pack(side="left", padx=(5, 0))
        elif found is False:
            # Red X
            self.status_indicator.configure(text="❌", text_color="red")
            # Hide the button
            self.open_detail_btn.pack_forget()
        else:
            # Hide indicator and button
            self.status_indicator.configure(text="")
            self.open_detail_btn.pack_forget()
    
    def _add_change_row(self):
        """Add a new row for specifying field changes."""
        row_frame = ctk.CTkFrame(self.changes_container)
        row_frame.pack(fill="x", pady=1)
        
        # Field dropdown (use SearchableDropdown and trace changes)
        field_label = ctk.CTkLabel(row_frame, text="Field:")
        field_label.grid(row=0, column=0, padx=(10, 5), pady=3, sticky="w")  # Reduced from pady=5
        
        field_names = [field['display_name'] for field in self.db_fields]
        field_var = tk.StringVar()
        field_dropdown = SearchableDropdown(row_frame,
                                           values=field_names,
                                           variable=field_var,
                                           width=200,
                                           height=26)  # Reduced height
        field_dropdown.grid(row=0, column=1, padx=5, pady=3)  # Reduced from pady=5
        
        # When the field selection changes, update the value widget type if needed
        # Capture the current row index at creation time (same behavior as previous command lambda)
        current_row_index = len(self.bulk_change_rows)
        def _on_field_var_change(*args, _var=field_var, _row=current_row_index):
            try:
                self._on_field_change(_row, _var.get())
            except Exception:
                pass
        field_var.trace_add('write', _on_field_var_change)
        
        # Action dropdown (Replace/Append)
        action_label = ctk.CTkLabel(row_frame, text="Action:")
        action_label.grid(row=0, column=2, padx=(20, 5), pady=3, sticky="w")  # Reduced from pady=5
        
        action_var = tk.StringVar(value="Replace")
        action_dropdown = ctk.CTkComboBox(row_frame, 
                                         variable=action_var,
                                         values=["Replace", "Append to"],
                                         width=120,
                                         height=26)  # Reduced height
        action_dropdown.grid(row=0, column=3, padx=5, pady=3)  # Reduced from pady=5
        
        # Value entry/dropdown (will be created dynamically)
        value_label = ctk.CTkLabel(row_frame, text="New Value:")
        value_label.grid(row=0, column=4, padx=(20, 5), pady=3, sticky="w")  # Reduced from pady=5
        
        value_var = tk.StringVar()
        value_widget = ctk.CTkEntry(row_frame, textvariable=value_var, width=250, height=26)  # Reduced height
        value_widget.grid(row=0, column=5, padx=5, pady=3, sticky="ew")  # Reduced from pady=5
        
        # Remove button
        remove_btn = ctk.CTkButton(row_frame, 
                                  text="✕", 
                                  command=lambda: self._remove_change_row(row_frame),
                                  width=26,  # Reduced from width=30
                                  height=26)  # Reduced from height=30
        remove_btn.grid(row=0, column=6, padx=(10, 10), pady=3)  # Reduced from pady=5
        
        # Configure grid weights
        row_frame.grid_columnconfigure(5, weight=1)
        
        # Store row data
        row_data = {
            'frame': row_frame,
            'field_var': field_var,
            'action_var': action_var,
            'value_var': value_var,
            'value_widget': value_widget,
            'field_dropdown': field_dropdown
        }
        
        self.bulk_change_rows.append(row_data)
    
    def _on_field_change(self, row_index, field_name):
        """Handle field selection change to update value widget if needed."""
        if row_index >= len(self.bulk_change_rows):
            return
        
        row_data = self.bulk_change_rows[row_index]
        
        # Find the database field name
        db_field_name = None
        for field in self.db_fields:
            if field['display_name'] == field_name:
                db_field_name = field['db_name']
                break
        
        # Check if this field should use a dropdown
        should_use_dropdown = any(df['db_name'] == db_field_name for df in self.dropdown_fields)
        
        # Check if this field should use a date picker
        should_use_datepicker = any(df['db_name'] == db_field_name for df in self.date_fields)
        
        if should_use_dropdown:
            # Replace entry with searchable dropdown
            current_value = row_data['value_var'].get()
            row_data['value_widget'].destroy()
            
            # Get unique values for this field from database
            try:
                unique_values = self.db.get_unique_field_values(db_field_name)
                if not unique_values:
                    unique_values = []
            except Exception as e:
                print(f"Error getting unique values for {db_field_name}: {e}")
                unique_values = []
            
            # Create SearchableDropdown widget
            value_dropdown = SearchableDropdown(row_data['frame'], 
                                              values=unique_values,
                                              variable=row_data['value_var'],
                                              width=250,
                                              height=26)  # Reduced height
            value_dropdown.grid(row=0, column=5, padx=5, pady=3, sticky="ew")  # Reduced pady
            row_data['value_widget'] = value_dropdown
            
            # Restore value if it exists
            if current_value:
                row_data['value_var'].set(current_value)
        
        elif should_use_datepicker:
            # Replace entry with date picker
            current_value = row_data['value_var'].get()
            row_data['value_widget'].destroy()
            
            # Create DatePicker widget
            value_datepicker = DatePicker(row_data['frame'], 
                                        variable=row_data['value_var'],
                                        width=250,
                                        height=26)  # Reduced height
            value_datepicker.grid(row=0, column=5, padx=5, pady=3, sticky="ew")  # Reduced pady
            row_data['value_widget'] = value_datepicker
            
            # Restore value if it exists
            if current_value:
                row_data['value_var'].set(current_value)
        else:
            # Use regular entry widget
            current_widget = row_data['value_widget']
            if not isinstance(current_widget, ctk.CTkEntry):
                current_value = ""
                if isinstance(current_widget, SearchableDropdown):
                    current_value = row_data['value_var'].get()
                elif isinstance(current_widget, DatePicker):
                    current_value = row_data['value_var'].get()
                else:
                    current_value = row_data['value_var'].get()
                current_widget.destroy()
                
                value_entry = ctk.CTkEntry(row_data['frame'], 
                                         textvariable=row_data['value_var'], 
                                         width=250,
                                         height=26)  # Reduced height
                value_entry.grid(row=0, column=5, padx=5, pady=3)  # Reduced pady
                row_data['value_widget'] = value_entry
                row_data['value_var'].set(current_value)
    
    def _remove_change_row(self, row_frame):
        """Remove a change row."""
        # Find and remove the row data
        for i, row_data in enumerate(self.bulk_change_rows):
            if row_data['frame'] == row_frame:
                row_frame.destroy()
                self.bulk_change_rows.pop(i)
                break
        
        # Ensure at least one row exists
        if not self.bulk_change_rows:
            self._add_change_row()
    
    def _search_asset(self):
        """Search for an asset based on the selected field and value."""
        search_field_name = self.search_field.get()
        search_value = self.search_value.get().strip()
        
        if not search_field_name or not search_value:
            messagebox.showwarning("Search Error", "Please select a field and enter a search value.")
            return
        
        # Find the database field name
        db_field_name = None
        for field in self.db_fields:
            if field['display_name'] == search_field_name:
                db_field_name = field['db_name']
                break
        
        if not db_field_name:
            messagebox.showerror("Search Error", "Invalid search field selected.")
            return
        
        try:
            # Search for the asset
            assets = self.db.search_assets_by_field(db_field_name, search_value)
            
            if not assets:
                # Show custom dialog with child asset option
                self._show_add_new_asset_dialog(search_field_name, search_value)
                return
            
            if len(assets) > 1:
                # Multiple assets found - show selection dialog
                selected_asset = self._show_asset_selection_dialog(assets, search_field_name, search_value)
                if selected_asset:
                    self._display_asset(selected_asset)
            else:
                # Single asset found
                self._display_asset(assets[0])
                
        except Exception as e:
            messagebox.showerror("Search Error", f"Error searching for asset: {str(e)}")
            print(f"Search error: {e}")
    
    def _show_asset_selection_dialog(self, assets, search_field, search_value):
        """Show dialog to select from multiple matching assets."""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Multiple Assets Found")
        dialog.geometry("800x400")
        dialog.transient(self.window)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 800) // 2
        y = (dialog.winfo_screenheight() - 400) // 2
        dialog.geometry(f"800x400+{x}+{y}")
        
        selected_asset = None
        
        # Title
        title_label = ctk.CTkLabel(dialog, 
                                  text=f"Multiple assets found with {search_field} = '{search_value}'",
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=(20, 10))
        
        # Asset list frame
        list_frame = ctk.CTkFrame(dialog)
        list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Create scrollable frame for assets
        assets_scroll = ctk.CTkScrollableFrame(list_frame)
        assets_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        def select_asset(asset):
            nonlocal selected_asset
            selected_asset = asset
            dialog.destroy()
        
        # Display each asset as a button
        for i, asset in enumerate(assets):
            asset_text = f"Asset {i+1}: "
            if hasattr(asset, 'serial_number') and asset.serial_number:
                asset_text += f"Serial: {asset.serial_number}"
            if hasattr(asset, 'asset_no') and asset.asset_no:
                asset_text += f" | Asset No: {asset.asset_no}"
            if hasattr(asset, 'manufacturer') and asset.manufacturer:
                asset_text += f" | {asset.manufacturer}"
            if hasattr(asset, 'model') and asset.model:
                asset_text += f" {asset.model}"
            
            asset_btn = ctk.CTkButton(assets_scroll, 
                                     text=asset_text,
                                     command=lambda a=asset: select_asset(a),
                                     height=40)
            asset_btn.pack(fill="x", pady=5)
        
        # Cancel button
        cancel_btn = ctk.CTkButton(dialog, 
                                  text="Cancel", 
                                  command=dialog.destroy)
        cancel_btn.pack(pady=(0, 20))
        
        # Wait for dialog to close
        dialog.wait_window()
        return selected_asset
    
    def _display_asset(self, asset):
        """Display the selected asset details using the reusable EmbeddedAssetDetail component."""
        self.selected_asset_data = asset
        
        # Update status indicator to show asset found
        self._update_status_indicator(found=True)
        
        # Clear previous display
        for widget in self.asset_details_frame.winfo_children():
            widget.destroy()
        
        # Convert asset object to dictionary for EmbeddedAssetDetail
        asset_dict = {}
        for field in self.db_fields:
            db_name = field['db_name']
            value = getattr(asset, db_name, None)
            if value is not None:
                asset_dict[db_name] = value
        
        # Add any missing standard fields
        if hasattr(asset, 'id') and asset.id:
            asset_dict['id'] = asset.id
        if hasattr(asset, 'asset_no') and asset.asset_no:
            asset_dict['asset_no'] = asset.asset_no
        
        # Add system fields for audit/history display
        system_fields = ['created_by', 'created_date', 'data_source', 'modified_by', 'modified_date']
        for field in system_fields:
            if hasattr(asset, field):
                value = getattr(asset, field)
                if value is not None:
                    asset_dict[field] = value
            
        # Create embedded asset detail component
        try:
            self.embedded_detail = EmbeddedAssetDetail(
                parent_frame=self.asset_details_frame,
                asset=asset_dict,
                on_edit_callback=self._on_asset_edited,
                show_edit_button=True
            )
        except Exception as e:
            # Fallback to simple display if embedded component fails
            error_label = ctk.CTkLabel(self.asset_details_frame, 
                                      text=f"Error displaying asset details: {e}",
                                      font=ctk.CTkFont(size=14))
            error_label.pack(pady=20)
            print(f"Error creating embedded asset detail: {e}")
        
        # Enable apply button
        self.apply_btn.configure(state="normal")
    
    def _on_asset_edited(self):
        """Callback when asset is edited from the embedded detail view."""
        # Refresh the display after editing
        if self.selected_asset_data:
            # Re-search for the asset to get updated data
            try:
                asset_id = self.selected_asset_data.id
                updated_asset_dict = self.db.get_asset_by_id(asset_id)
                if updated_asset_dict:
                    # Convert dict back to object format that _display_asset expects
                    class AssetObj:
                        def __init__(self, asset_dict):
                            for key, value in asset_dict.items():
                                setattr(self, key, value)
                    
                    updated_asset = AssetObj(updated_asset_dict)
                    self._display_asset(updated_asset)
                else:
                    self._clear_asset_display()
            except Exception as e:
                print(f"Error refreshing asset after edit: {e}")
                self._clear_asset_display()
    
    def _clear_asset_display(self):
        """Clear the asset display area."""
        self.selected_asset_data = None
        
        # Update status indicator to show no asset found
        self._update_status_indicator(found=False)
        
        # Clear previous display
        for widget in self.asset_details_frame.winfo_children():
            widget.destroy()
        
        # Show "no asset selected" message
        self.no_asset_label = ctk.CTkLabel(self.asset_details_frame, 
                                          text="No asset selected. Use the search above to find an asset.",
                                          font=ctk.CTkFont(size=14))
        self.no_asset_label.pack(pady=50)
    
    def _open_full_asset_detail(self):
        """Open the full Asset Detail Window for the currently selected asset."""
        if not self.selected_asset_data:
            messagebox.showwarning("No Asset Selected", "Please select an asset first.")
            return
        
        try:
            # Import AssetDetailWindow
            from ui_components import AssetDetailWindow
            
            # Convert asset object to dictionary if needed
            if hasattr(self.selected_asset_data, '__dict__'):
                # It's an object, convert to dict
                asset_dict = {}
                for attr in dir(self.selected_asset_data):
                    if not attr.startswith('_'):
                        value = getattr(self.selected_asset_data, attr, None)
                        if value is not None and not callable(value):
                            asset_dict[attr] = value
            else:
                # Already a dictionary
                asset_dict = self.selected_asset_data
            
            # Open the Asset Detail Window
            AssetDetailWindow(
                parent=self.window,
                asset=asset_dict,
                on_edit_callback=self._on_asset_edited  # Refresh embedded view when edited from full window
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Asset Detail Window:\n{str(e)}")
            print(f"Error opening full asset detail: {e}")
        
        # Disable apply button
        self.apply_btn.configure(state="disabled")
    
    def _apply_changes(self):
        """Apply the bulk changes to the selected asset."""
        if not self.selected_asset_data:
            messagebox.showwarning("No Asset", "Please search for and select an asset first.")
            return
        
        # Collect changes to apply
        changes_to_apply = {}
        
        for row_data in self.bulk_change_rows:
            field_name = row_data['field_var'].get()
            action = row_data['action_var'].get()
            
            # Get value from the appropriate widget
            if isinstance(row_data['value_widget'], SearchableDropdown):
                new_value = row_data['value_var'].get().strip()
            elif isinstance(row_data['value_widget'], DatePicker):
                new_value = row_data['value_var'].get().strip()
            else:
                new_value = row_data['value_var'].get().strip()
            
            # Skip empty rows
            if not field_name or not new_value:
                continue
            
            # Find database field name
            db_field_name = None
            for field in self.db_fields:
                if field['display_name'] == field_name:
                    db_field_name = field['db_name']
                    break
            
            if not db_field_name:
                continue
            
            # Apply action
            if action == "Replace":
                changes_to_apply[db_field_name] = new_value
            elif action == "Append to":
                current_value = getattr(self.selected_asset_data, db_field_name, '') or ''
                if current_value:
                    # Check if this is a multiline field (notes, description, etc.)
                    if self.db.should_field_be_multiline(field_name, self.template_path):
                        # For multiline fields, add a newline before appending
                        changes_to_apply[db_field_name] = f"{current_value}\n{new_value}"
                    else:
                        # For single-line fields, add a space
                        changes_to_apply[db_field_name] = f"{current_value} {new_value}"
                else:
                    changes_to_apply[db_field_name] = new_value
        
        if not changes_to_apply:
            messagebox.showwarning("No Changes", "Please specify at least one field to change.")
            return
        
        # Confirm changes
        changes_text = "\n".join([f"• {self._get_display_name(field)}: {value}" 
                                  for field, value in changes_to_apply.items()])
        
        asset_identifier = ""
        if hasattr(self.selected_asset_data, 'serial_number') and self.selected_asset_data.serial_number:
            asset_identifier = f"Serial: {self.selected_asset_data.serial_number}"
        elif hasattr(self.selected_asset_data, 'asset_no') and self.selected_asset_data.asset_no:
            asset_identifier = f"Asset No: {self.selected_asset_data.asset_no}"
        else:
            asset_identifier = f"Asset ID: {self.selected_asset_data.id}"
        
        confirm_message = f"Apply the following changes to {asset_identifier}?\n\n{changes_text}"
        
        if not messagebox.askyesno("Confirm Changes", confirm_message):
            return
        
        try:
            # Apply changes to database
            asset_id = self.selected_asset_data.id
            self.db.update_asset(asset_id, changes_to_apply)
            
            # Refresh asset display by re-searching for the updated asset
            updated_asset_dict = self.db.get_asset_by_id(asset_id)
            if updated_asset_dict:
                # Convert dictionary to object with attribute access
                updated_asset = type('Asset', (), {})()
                for key, value in updated_asset_dict.items():
                    setattr(updated_asset, key, value)
                
                self._display_asset(updated_asset)
                self.selected_asset_data = updated_asset  # Update our reference
            else:
                # If we can't get the updated asset by ID, try to re-search using current search criteria
                self._refresh_current_search()
            
            messagebox.showinfo("Success", f"Asset updated successfully!\n\nUpdated fields:\n{changes_text}")
            
            # Clear search box and focus on it for next asset
            self.search_value.set("")
            self.search_entry.focus_set()
            
        except Exception as e:
            messagebox.showerror("Update Error", f"Error updating asset: {str(e)}")
            print(f"Update error: {e}")
    
    def _get_display_name(self, db_field_name):
        """Get display name for a database field name."""
        for field in self.db_fields:
            if field['db_name'] == db_field_name:
                return field['display_name']
        return db_field_name.replace('_', ' ').title()
    
    def _refresh_current_search(self):
        """Re-run the current search to refresh the displayed asset."""
        search_field_name = self.search_field.get()
        search_value = self.search_value.get().strip()
        
        if not search_field_name or not search_value:
            # No current search criteria, just clear the display
            self._clear_asset_display()
            return
        
        # Find the database field name
        db_field_name = None
        for field in self.db_fields:
            if field['display_name'] == search_field_name:
                db_field_name = field['db_name']
                break
        
        if not db_field_name:
            self._clear_asset_display()
            return
        
        try:
            # Re-search for assets
            assets = self.db.search_assets_by_field(db_field_name, search_value)
            
            if assets:
                # If the current asset ID is still in the results, show that one
                current_asset = None
                if self.selected_asset_data:
                    for asset in assets:
                        if asset.id == self.selected_asset_data.id:
                            current_asset = asset
                            break
                
                # Display the found asset (either the updated current one or the first match)
                if current_asset:
                    self._display_asset(current_asset)
                else:
                    self._display_asset(assets[0])
            else:
                # No assets found, clear display
                self._clear_asset_display()
                
        except Exception as e:
            print(f"Error refreshing search: {e}")
            self._clear_asset_display()
    
    def _show_add_new_asset_dialog(self, search_field_name: str, search_value: str):
        """Show custom dialog for adding a new asset with optional child relationship."""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Asset Not Found")
        dialog.geometry("500x300")
        dialog.transient(self.window)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 500) // 2
        y = (dialog.winfo_screenheight() - 300) // 2
        dialog.geometry(f"500x300+{x}+{y}")
        
        # Variables for user choices
        add_asset = False
        make_child = False
        
        # Main message
        message_label = ctk.CTkLabel(dialog, 
                                   text=f"No assets found with {search_field_name} = '{search_value}'\n\n"
                                        f"Do you want to add it as a new asset?",
                                   font=ctk.CTkFont(size=14),
                                   wraplength=450)
        message_label.pack(pady=(20, 10))
        
        # Child asset option frame
        child_frame = ctk.CTkFrame(dialog)
        child_frame.pack(fill="x", padx=20, pady=10)
        
        # Checkbox for making it a child asset
        child_var = ctk.BooleanVar()
        
        # Get parent asset info for display
        parent_info = "None"
        if self.selected_asset_data:
            parent_parts = []
            if hasattr(self.selected_asset_data, 'manufacturer') and self.selected_asset_data.manufacturer:
                parent_parts.append(self.selected_asset_data.manufacturer)
            if hasattr(self.selected_asset_data, 'model') and self.selected_asset_data.model:
                parent_parts.append(self.selected_asset_data.model)
            if hasattr(self.selected_asset_data, 'serial_number') and self.selected_asset_data.serial_number:
                parent_parts.append(f"Serial: {self.selected_asset_data.serial_number}")
            
            if parent_parts:
                parent_info = " - ".join(parent_parts)
        
        child_checkbox = ctk.CTkCheckBox(child_frame,
                                       text=f"Make child of: {parent_info}",
                                       variable=child_var,
                                       font=ctk.CTkFont(size=12))
        child_checkbox.pack(pady=10)
        
        # Only show checkbox if we have a parent asset
        if not self.selected_asset_data:
            child_checkbox.configure(state="disabled")
            child_var.set(False)
        
        # Button frame
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=(20, 10))
        
        def on_yes():
            nonlocal add_asset, make_child
            add_asset = True
            make_child = child_var.get()
            dialog.destroy()
        
        def on_no():
            nonlocal add_asset
            add_asset = False
            dialog.destroy()
        
        # Yes and No buttons
        yes_btn = ctk.CTkButton(button_frame, 
                              text="Yes", 
                              command=on_yes,
                              width=100)
        yes_btn.pack(side="left", padx=(0, 10))
        
        no_btn = ctk.CTkButton(button_frame, 
                             text="No", 
                             command=on_no,
                             width=100)
        no_btn.pack(side="left")
        
        # Wait for dialog to close
        dialog.wait_window()
        
        # Process the result
        if add_asset:
            if make_child and self.selected_asset_data:
                # Pass parent asset info for child relationship
                self._open_add_new_asset_with_prefill(search_field_name, search_value, 
                                                    parent_asset=self.selected_asset_data)
            else:
                # Regular new asset
                self._open_add_new_asset_with_prefill(search_field_name, search_value)
        else:
            self._clear_asset_display()
    
    
    def _on_preset_selected(self, preset_name: str):
        """Handle preset selection from dropdown."""
        if preset_name == "Select Preset...":
            return
        elif preset_name == "Save Current as Preset...":
            self._save_current_as_preset()
        else:
            self._load_preset(preset_name)
    
    def _save_current_as_preset(self):
        """Save current field configuration as a new preset."""
        # Get preset name from user
        dialog = ctk.CTkInputDialog(text="Enter preset name:", title="Save Preset")
        preset_name = dialog.get_input()
        
        if not preset_name or preset_name.strip() == "":
            return
        
        preset_name = preset_name.strip()
        
        # Collect current field configurations
        preset_data = []
        date_fields_to_check = []  # Track date fields with today's date for dynamic/static choice
        
        for row_data in self.bulk_change_rows:
            field = row_data['field_var'].get()
            operation = row_data['action_var'].get()
            value = row_data['value_var'].get()
            
            if field and field != "Select Field":
                # Check if this is a date field with today's date
                if "date" in field.lower() and value == _today_audit_date_str():
                    date_fields_to_check.append({
                        "field": field,
                        "operation": operation.lower(),
                        "value": value
                    })
                else:
                    preset_data.append({
                        "field": field,
                        "operation": operation.lower(),
                        "value": value
                    })
        
        # Handle date fields with dynamic/static choice
        for date_field in date_fields_to_check:
            result = messagebox.askyesno("Date Field Options", 
                                       f"The field '{date_field['field']}' is set to today's date ({date_field['value']}).\n\n"
                                       f"Would you like to save this as a dynamic date that always uses the current date?\n\n"
                                       f"• Yes: Always use current date when preset is loaded\n"
                                       f"• No: Use the static date {date_field['value']}")
            
            if result:  # Yes - dynamic date
                preset_data.append({
                    "field": date_field["field"],
                    "operation": date_field["operation"],
                    "value": "current_date"
                })
            else:  # No - static date
                preset_data.append(date_field)
        
        if not preset_data:
            messagebox.showwarning("No Fields", "No fields configured to save as preset.")
            return
        
        # Save to config in new format
        self.config.bulk_update_presets[preset_name] = {
            "type": "user",
            "fields": preset_data
        }
        self.config_manager.save_config()
        
        # Update dropdown options
        preset_options = ["Select Preset...", "Save Current as Preset..."] + list(self.config.bulk_update_presets.keys())
        self.presets_dropdown.configure(values=preset_options)
        
        messagebox.showinfo("Preset Saved", f"Preset '{preset_name}' saved successfully.")
        self.presets_var.set("Select Preset...")
    
    def _load_preset(self, preset_name: str):
        """Load a preset configuration."""
        if preset_name not in self.config.bulk_update_presets:
            messagebox.showerror("Error", f"Preset '{preset_name}' not found.")
            return
        
        preset_data = self.config.bulk_update_presets[preset_name]
        
        # Handle both old and new preset formats
        if isinstance(preset_data, dict) and "fields" in preset_data:
            # New format with type and fields
            fields_data = preset_data["fields"]
        else:
            # Old format - just an array of field configs
            fields_data = preset_data
        
        # Clear existing rows
        for row_data in self.bulk_change_rows:
            row_data['frame'].destroy()
        self.bulk_change_rows.clear()
        
        # Load preset fields
        for field_config in fields_data:
            self._add_change_row()
            row_data = self.bulk_change_rows[-1]
            
            field_name = field_config.get("field", "")
            operation = field_config.get("operation", "replace").title()
            value = field_config.get("value", "")
            
            # Set field
            field_display_names = [field['display_name'] for field in self.db_fields]
            if field_name in field_display_names:
                row_data['field_var'].set(field_name)
                self._on_field_change(len(self.bulk_change_rows) - 1, field_name)
            
            # Set operation
            row_data['action_var'].set(operation)
            
            # Set value
            if value == "current_date":
                # Special handling for current date - use proper format
                current_date = _today_audit_date_str()
                row_data['value_var'].set(current_date)
            else:
                row_data['value_var'].set(value)
        
        # If no rows were added (invalid preset), add default row
        if not self.bulk_change_rows:
            self._add_change_row()
        
        self.presets_var.set("Select Preset...")
        messagebox.showinfo("Preset Loaded", f"Preset '{preset_name}' loaded successfully.")

    def _open_edit_asset(self):
        """Open the Edit Asset window for the currently selected asset."""
        if not self.selected_asset_data:
            messagebox.showwarning("No Asset Selected", "No asset is currently selected to edit.")
            return
        
        try:
            # Open the Edit Asset window with callback to refresh display
            EditAssetWindow(
                parent=self.window,
                asset_id=self.selected_asset_data.id,
                config=self.config,
                on_update_callback=self._on_asset_updated
            )
        except Exception as e:
            messagebox.showerror("Error", f"Error opening edit window: {str(e)}")
            print(f"Edit window error: {e}")

    def _on_asset_updated(self):
        """Callback function called when the asset is updated in the Edit Asset window."""
        if not self.selected_asset_data:
            return
        
        try:
            # Refresh the asset data by fetching updated information from database
            updated_asset_dict = self.db.get_asset_by_id(self.selected_asset_data.id)
            if updated_asset_dict:
                # Convert dictionary to object with attribute access (same as in _apply_changes)
                updated_asset = type('Asset', (), {})()
                for key, value in updated_asset_dict.items():
                    setattr(updated_asset, key, value)
                
                # Refresh the display with updated asset data
                self._display_asset(updated_asset)
                print(f"Asset {self.selected_asset_data.id} display refreshed after edit")
            else:
                # Asset not found, try to refresh using current search
                self._refresh_current_search()
        except Exception as e:
            print(f"Error refreshing asset display after edit: {e}")
            # Try to refresh using current search as fallback
            self._refresh_current_search()

    def _open_add_new_asset_with_prefill(self, search_field_name: str, search_value: str, parent_asset=None):
        """Open the Add New Asset window with pre-filled search value and optional parent relationship."""
        try:
            from add_new_assets import AddNewAssetsWindow
            
            # Create the Add New Assets window
            add_window = AddNewAssetsWindow(parent=self.window, config=self.config)
            
            # Bring the new window to the front
            add_window.window.lift()
            add_window.window.focus_force()
            add_window.window.attributes('-topmost', True)
            add_window.window.after(100, lambda: add_window.window.attributes('-topmost', False))
            
            # Try to pre-fill the field that was searched
            # First try using the exact search field name (display name)
            field_to_fill = search_field_name
            
            # If that doesn't work, try to find the database field name
            if field_to_fill not in add_window.widgets:
                for field in self.db_fields:
                    if field['display_name'] == search_field_name:
                        field_to_fill = field['db_name']
                        break
            
            # Try different possible field name variations
            possible_field_names = [
                field_to_fill,
                search_field_name,
                search_field_name.replace(' ', '_').lower(),
                search_field_name.replace(' ', '').lower(),
                'serial_number' if 'serial' in search_field_name.lower() else None
            ]
            
            widget_filled = False
            for field_name in possible_field_names:
                if field_name and field_name in add_window.widgets:
                    widget = add_window.widgets[field_name]
                    try:
                        # Handle different widget types
                        if hasattr(widget, 'set'):
                            # SearchableDropdown or similar
                            widget.set(search_value)
                        elif hasattr(widget, 'insert'):
                            # CTkEntry
                            widget.delete(0, 'end')
                            widget.insert(0, search_value)
                        elif hasattr(widget, 'configure'):
                            # Try setting text for textbox
                            widget.configure(state="normal")
                            widget.delete("1.0", 'end')
                            widget.insert("1.0", search_value)
                            widget.configure(state="normal")
                        
                        print(f"Pre-filled '{field_name}' with '{search_value}'")
                        widget_filled = True
                        break
                    except Exception as widget_error:
                        print(f"Error setting value for widget '{field_name}': {widget_error}")
                        continue
            
            if not widget_filled:
                print(f"Could not find widget to pre-fill for field: {search_field_name}")
                print(f"Available widgets: {list(add_window.widgets.keys())}")
            
            # Handle parent asset relationship if specified
            if parent_asset:
                # Add a small delay to ensure widgets are fully initialized
                add_window.window.after(200, lambda: self._set_child_asset_fields(add_window, parent_asset))
            
        except Exception as e:
            messagebox.showerror("Error", f"Error opening Add New Asset window: {str(e)}")
            print(f"Error opening add new asset window: {e}")
    
    def _set_child_asset_fields(self, add_window, parent_asset):
        """Set fields for child asset relationship."""
        try:
            # Debug: Print all available widgets to understand the structure
            print(f"Available widgets in Add New Asset window: {list(add_window.widgets.keys())}")
            
            # Helper function to check if a field is configured as a dropdown
            def is_dropdown_field(field_display_name):
                configured_dropdown_headers = set(getattr(self.config, 'dropdown_fields', []) or [])
                return field_display_name in configured_dropdown_headers
            
            # Set Child Asset field to "Y" - using exact name from widget list
            child_field_names = [
                "Child Asset? (Y/N)",  # This exists in the widget list
                "Child Asset",
                "child_asset"
            ]
            
            child_field_set = False
            for field_name in child_field_names:
                if field_name in add_window.widgets:
                    widget = add_window.widgets[field_name]
                    widget_type = type(widget).__name__
                    print(f"Attempting to set {field_name} (type: {widget_type}) to 'Y'")
                    
                    try:
                        # Check if this field is configured as a dropdown
                        if is_dropdown_field(field_name) and widget_type == "SearchableDropdown":
                            # Handle as SearchableDropdown
                            if hasattr(widget, 'variable'):
                                widget.variable.set("Y")
                                print(f"Used variable.set() for dropdown field {field_name}")
                            elif hasattr(widget, 'search_var'):
                                widget.search_var.set("Y")
                                print(f"Used search_var.set() for dropdown field {field_name}")
                        elif hasattr(widget, 'set'):
                            widget.set("Y")
                            print(f"Used .set() method for {field_name}")
                        elif hasattr(widget, 'insert'):
                            widget.delete(0, 'end')
                            widget.insert(0, "Y")
                            print(f"Used insert method for {field_name}")
                        elif hasattr(widget, 'delete') and hasattr(widget, 'insert'):
                            # Handle textbox without using state parameter
                            widget.delete("1.0", 'end')
                            widget.insert("1.0", "Y")
                            print(f"Used textbox methods for {field_name}")
                        
                        print(f"Set '{field_name}' to 'Y'")
                        child_field_set = True
                        break
                    except Exception as widget_error:
                        print(f"Error setting child asset field '{field_name}': {widget_error}")
                        continue
            
            if not child_field_set:
                print(f"Warning: Could not find Child Asset field. Tried: {child_field_names}")
            
            # Set Related Assets field with parent serial number - using exact name from widget list
            parent_serial = ""
            if hasattr(parent_asset, 'serial_number') and parent_asset.serial_number:
                parent_serial = parent_asset.serial_number
            
            if parent_serial:
                related_field_names = [
                    "Related Asset Sync Keys",  # This exists in the widget list
                    "Related Assets",
                    "related_assets",
                    "Related Asset",
                    "Parent Asset",
                    "parent_asset"
                ]
                
                related_field_set = False
                for field_name in related_field_names:
                    if field_name in add_window.widgets:
                        widget = add_window.widgets[field_name]
                        try:
                            if hasattr(widget, 'set'):
                                widget.set(parent_serial)
                            elif hasattr(widget, 'insert'):
                                widget.delete(0, 'end')
                                widget.insert(0, parent_serial)
                            elif hasattr(widget, 'delete') and hasattr(widget, 'insert'):
                                # Handle textbox without using state parameter
                                widget.delete("1.0", 'end')
                                widget.insert("1.0", parent_serial)
                            print(f"Set '{field_name}' to parent serial: '{parent_serial}'")
                            related_field_set = True
                            break
                        except Exception as widget_error:
                            print(f"Error setting related assets field '{field_name}': {widget_error}")
                            continue
                
                if not related_field_set:
                    print(f"Warning: Could not find Related Assets field. Tried: {related_field_names}")
            else:
                print("Warning: Parent asset has no serial number for child relationship")
            
            # Copy location and status fields from parent - using exact names from widget list
            fields_to_copy = {
                "Location": ["Location"],  # Exists in widget list
                "Room": ["Room"],  # Exists in widget list
                "Cubicle": ["Cubicle"],  # Exists in widget list
                "Rack": ["Rack/Elevation", "Rack"],  # Modified to match widget list
                "Status": ["Status"]  # Exists in widget list
            }
            
            for display_name, field_variations in fields_to_copy.items():
                # Get value from parent asset
                parent_value = None
                for field_var in field_variations:
                    if hasattr(parent_asset, field_var.lower()) and getattr(parent_asset, field_var.lower()):
                        parent_value = getattr(parent_asset, field_var.lower())
                        break
                    elif hasattr(parent_asset, field_var) and getattr(parent_asset, field_var):
                        parent_value = getattr(parent_asset, field_var)
                        break
                    # Try without the /Elevation part for rack
                    elif field_var == "Rack/Elevation" and hasattr(parent_asset, 'rack') and getattr(parent_asset, 'rack'):
                        parent_value = getattr(parent_asset, 'rack')
                        break
                
                if parent_value:
                    # Try to set the field in the add window
                    field_set = False
                    for field_name in field_variations:
                        if field_name in add_window.widgets:
                            widget = add_window.widgets[field_name]
                            try:
                                # Add debugging to see widget type
                                widget_type = type(widget).__name__
                                print(f"Attempting to set {field_name} (type: {widget_type}) to '{parent_value}'")
                                
                                success = False
                                error_msg = ""
                                
                                # Try different approaches based on widget type and configuration
                                if widget_type == "SearchableDropdown" or is_dropdown_field(field_name):
                                    # For SearchableDropdown or fields configured as dropdowns, use the variable attribute
                                    try:
                                        if hasattr(widget, 'variable'):
                                            widget.variable.set(str(parent_value))
                                            success = True
                                            print(f"Used variable.set() for dropdown field {field_name}")
                                        elif hasattr(widget, 'search_var'):
                                            widget.search_var.set(str(parent_value))
                                            success = True
                                            print(f"Used search_var.set() for dropdown field {field_name}")
                                        else:
                                            error_msg = "No variable or search_var attribute found on SearchableDropdown"
                                    except Exception as e:
                                        error_msg = f"SearchableDropdown variable setting failed: {e}"
                                elif hasattr(widget, 'set'):
                                    try:
                                        widget.set(str(parent_value))
                                        success = True
                                        print(f"Used .set() method for {field_name}")
                                    except Exception as e:
                                        error_msg = f"set() failed: {e}"
                                elif hasattr(widget, 'insert'):
                                    try:
                                        widget.delete(0, 'end')
                                        widget.insert(0, str(parent_value))
                                        success = True
                                        print(f"Used .insert() method for {field_name}")
                                    except Exception as e:
                                        error_msg = f"insert() failed: {e}"
                                elif hasattr(widget, 'delete') and hasattr(widget, 'insert'):
                                    try:
                                        # Handle textbox without using state parameter
                                        widget.delete("1.0", 'end')
                                        widget.insert("1.0", str(parent_value))
                                        success = True
                                        print(f"Used textbox methods for {field_name}")
                                    except Exception as e:
                                        error_msg = f"textbox methods failed: {e}"
                                else:
                                    error_msg = "No suitable setter method found"
                                
                                if not success:
                                    print(f"Failed to set {field_name}: {error_msg}")
                                
                                # Force a refresh of the widget
                                if hasattr(widget, 'update'):
                                    widget.update()
                                
                                # Force visual refresh for SearchableDropdown widgets
                                if success and widget_type == "SearchableDropdown":
                                    try:
                                        # Force the display_entry to refresh by triggering its update
                                        if hasattr(widget, 'display_entry'):
                                            widget.display_entry.update()
                                            widget.display_entry.update_idletasks()
                                        
                                        # Force the entire widget to refresh
                                        widget.update()
                                        widget.update_idletasks()
                                        
                                        # Schedule a delayed update to ensure the display refreshes
                                        widget.after(50, lambda: widget.display_entry.update() if hasattr(widget, 'display_entry') else None)
                                        
                                        print(f"Triggered display update for {field_name}")
                                    except Exception as e:
                                        print(f"Display update failed for {field_name}: {e}")
                                
                                # Verify the value was set
                                current_value = "unknown"
                                try:
                                    if hasattr(widget, 'search_var') and hasattr(widget.search_var, 'get'):
                                        current_value = widget.search_var.get()
                                    elif hasattr(widget, 'variable') and hasattr(widget.variable, 'get'):
                                        current_value = widget.variable.get()
                                    elif hasattr(widget, 'get'):
                                        current_value = widget.get()
                                except Exception as get_error:
                                    print(f"Could not verify value for {field_name}: {get_error}")
                                
                                if success:
                                    print(f"Successfully set '{field_name}' to '{parent_value}'. Current value: '{current_value}'")
                                field_set = True
                                break
                            except Exception as widget_error:
                                print(f"Error setting {display_name} field '{field_name}': {widget_error}")
                                continue
                    
                    if not field_set:
                        print(f"Warning: Could not find {display_name} field in Add New Asset window")
                else:
                    print(f"No {display_name} value found in parent asset to copy")
                
            print(f"Child asset relationship configured with parent serial: {parent_serial}")
                
        except Exception as e:
            print(f"Error setting child asset fields: {e}")

    def _on_closing(self):
        """Handle window closing."""
        try:
            # Clean up keyboard bindings
            self.window.unbind('<Return>')
            self.window.unbind('<KP_Enter>')
            self.window.unbind('<Control-Return>')
            self.window.unbind('<Control-KP_Enter>')
            self.window.unbind('<Control-BackSpace>')
        except Exception:
            pass  # Ignore errors during cleanup
        
        # Destroy the window
        self.window.destroy()


# Helper function to open the bulk update window
def open_bulk_update_window(parent, config=None):
    """Open the bulk update assets window."""
    BulkUpdateWindow(parent, config)