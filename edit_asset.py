import customtkinter as ctk
from tkinter import messagebox
import os
import csv
from datetime import datetime
from typing import Dict, List, Any, Optional
from asset_database import AssetDatabase, db
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from validation import form_validator, asset_validator
from performance_monitoring import performance_monitor
from ui_components import SearchableDropdown, DatePicker


# Fields that should not be editable (system fields and auto-generated fields)
READONLY_FIELDS = {
    "Asset No.",
    "id",
    "created_date",
    "modified_date", 
    "created_by",
    "modified_by",
    "data_source"
}

# Header name (case-insensitive) to detect audit date field
AUDIT_DATE_HEADER = "audit date"

def _today_audit_date_str() -> str:
    """Return today's date in requested format: MM/D/YYYY (month zero-padded, day without leading zero)."""
    now = datetime.now()
    return f"{now:%m}/{now.day}/{now:%Y}"


class EditAssetWindow:
    """Window for editing existing asset data.
    
    This window reuses the design pattern from AddNewAssetsWindow but:
    - Populates fields with existing asset data
    - Has Cancel and Submit Changes buttons
    - Tracks field changes for audit logging
    - Updates modified_date and audit trail
    """

    def __init__(self, parent, asset_id: int, config=None, on_update_callback=None):
        """Initialize the edit asset window.
        
        Args:
            parent: Parent window
            asset_id: ID of the asset to edit
            config: Configuration object (optional)
            on_update_callback: Callback function to call after successful update
        """
        self.parent = parent
        self.asset_id = asset_id
        self.on_update_callback = on_update_callback
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = config or self.config_manager.get_config()
        
        # Load dynamic required fields from config (like add_new_assets.py does)
        self.required_fields = set(self.config.required_fields)
        self.dropdown_fields = set(self.config.dropdown_fields)
        self.excluded_fields = set(self.config.excluded_fields)
        
        # Create database instance
        self.db = AssetDatabase(self.config.database_path)
        
        # Load the asset data
        self.original_asset = self.db.get_asset_by_id(asset_id)
        if not self.original_asset:
            messagebox.showerror("Error", f"Asset with ID {asset_id} not found.")
            return
        
        # Store original values for change tracking
        self.original_values = dict(self.original_asset)
        
        # Get template path for field structure
        self.template_path = self.config.default_template_path
        
        # Parse template and get field structure
        self.headers = []
        self.unique_values = {}
        self._load_template_structure()
        
        # UI components
        self.window = None
        self.widgets = {}
        self.widget_vars = {}  # Track StringVar instances for dropdowns and date pickers
        self.required_frame = None
        self.additional_frame = None
        
        self._create_window()
        self._create_widgets()
        self._populate_fields()
        self._bind_shortcuts()
        
    def _load_template_structure(self):
        """Load template structure to understand available fields and their types."""
        if not self.template_path or not os.path.exists(self.template_path):
            # Fallback: use asset keys as headers
            self.headers = [key for key in self.original_asset.keys() 
                          if key not in READONLY_FIELDS]
            return
            
        try:
            with open(self.template_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                self.headers = next(reader)
                
                # Load unique values for dropdown fields (like add_new_assets.py does)
                for field in self.dropdown_fields:
                    if field in self.headers:
                        # Get unique values from database for this field
                        db_column = self.db.get_dynamic_column_mapping(self.template_path).get(field)
                        if db_column:
                            unique_values = self.db.get_unique_field_values(db_column)
                            if unique_values:
                                self.unique_values[field] = sorted(unique_values)
                                
        except Exception as e:
            print(f"Warning: Could not load template structure: {e}")
            # Fallback: use asset keys as headers
            self.headers = [key for key in self.original_asset.keys() 
                          if key not in READONLY_FIELDS]

    def _create_window(self):
        """Create the main window."""
        self.window = ctk.CTkToplevel(self.parent)
        asset_no = self.original_asset.get('asset_no', f"ID {self.asset_id}")
        self.window.title(f"Edit Asset - {asset_no}")
        self.window.geometry("1000x700")
        self.window.minsize(900, 600)
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # Handle window closing
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (1500 // 2)
        y = (self.window.winfo_screenheight() // 2) - (850 // 2)
        self.window.geometry(f"1500x850+{x}+{y}")

    def _create_widgets(self):
        """Create all UI widgets."""
        # Scrollable form area (using CTkScrollableFrame) - matches add_new_assets.py
        self.form_inner = ctk.CTkScrollableFrame(self.window)
        self.form_inner.pack(fill="both", expand=True, padx=10, pady=10)

        # Build the form fields
        self._build_form_fields()
        
        # Bottom action frame - matches add_new_assets.py
        action_frame = ctk.CTkFrame(self.window)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        # Buttons - right-aligned like add_new_assets.py
        ctk.CTkButton(action_frame, text="Submit Changes", command=self._submit_changes,
                      fg_color="#1f538d", hover_color="#14375e").pack(side="right", padx=10, pady=10)
        ctk.CTkButton(action_frame, text="Cancel", fg_color="gray", 
                      command=self._cancel_edit).pack(side="right", padx=10, pady=10)

    def _build_form_fields(self):
        """Build form fields using the same pattern as add_new_assets.py."""
        # Configure grid layout for 4 columns (2 pairs of label+widget)
        self.form_inner.grid_columnconfigure(0, weight=0)  # Label 1
        self.form_inner.grid_columnconfigure(1, weight=1)  # Widget 1
        self.form_inner.grid_columnconfigure(2, weight=0)  # Label 2
        self.form_inner.grid_columnconfigure(3, weight=1)  # Widget 2

        current_row = 0
        
        # Title - use grid instead of pack
        title_text = f"Edit Asset: {self.original_asset.get('asset_no', f'ID {self.asset_id}')}"
        title_label = ctk.CTkLabel(self.form_inner, text=title_text, 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(8,4))
        current_row += 1
        
        # Asset info subtitle - use grid instead of pack
        subtitle_parts = []
        if 'manufacturer' in self.original_asset:
            subtitle_parts.append(self.original_asset['manufacturer'])
        if 'model' in self.original_asset:
            subtitle_parts.append(self.original_asset['model'])
        if 'serial_number' in self.original_asset:
            subtitle_parts.append(f"SN: {self.original_asset['serial_number']}")
            
        if subtitle_parts:
            subtitle_label = ctk.CTkLabel(self.form_inner, text=" | ".join(subtitle_parts),
                                        font=ctk.CTkFont(size=14))
            subtitle_label.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(0,15))
            current_row += 1

        # Filter headers to exclude configured excluded fields and readonly fields
        all_excluded = self.excluded_fields | READONLY_FIELDS
        
        # Separate required and additional fields (excluding excluded fields)
        required_headers = [h for h in self.headers if h in self.required_fields and h not in all_excluded]
        additional_headers = [h for h in self.headers if h not in self.required_fields and h not in all_excluded]
        
        # Required fields section
        if required_headers:
            heading_req = ctk.CTkLabel(self.form_inner, text="Required Information", 
                                     font=ctk.CTkFont(size=16, weight="bold"))
            heading_req.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(8,4))
            current_row += 1
            
            # Required fields in two-column layout
            current_row = self._create_field_section(required_headers, current_row)
        
        # Additional fields section
        if additional_headers:
            # Separator
            divider = ctk.CTkFrame(self.form_inner, height=2)
            divider.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=4, pady=(6,10))
            current_row += 1
            
            heading_add = ctk.CTkLabel(self.form_inner, text="Additional Information", 
                                     font=ctk.CTkFont(size=16, weight="bold"))
            heading_add.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(0,4))
            current_row += 1
            
            # Additional fields in two-column layout
            current_row = self._create_field_section(additional_headers, current_row)

    def _create_field_section(self, headers, start_row):
        """Create a two-column (label+input pairs) section for given headers.
        
        Returns the next available row index after placing widgets.
        """
        for idx, header in enumerate(headers):
            col_group = idx % 2  # 0 or 1
            row_offset = idx // 2
            base_col = col_group * 2
            row = start_row + row_offset

            # Create field label - use original header text like Add Assets does
            label = ctk.CTkLabel(self.form_inner, text=header + ":")
            label.grid(row=row, column=base_col, sticky="e", padx=8, pady=4)

            # Create appropriate widget based on field type
            if header in self.dropdown_fields and header in self.unique_values:
                var = ctk.StringVar(value="")
                values = [""] + self.unique_values[header]
                widget = SearchableDropdown(self.form_inner, values=values, variable=var)
                widget.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                self.widget_vars[header] = var
            elif "date" in header.lower():
                # Use DatePicker for any field containing "date"
                var = ctk.StringVar(value="")
                widget = DatePicker(self.form_inner, variable=var)
                widget.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                self.widget_vars[header] = var
                # Pre-populate audit date with today's date if it's the audit date field
                if header.lower() == AUDIT_DATE_HEADER.lower():
                    var.set(_today_audit_date_str())
            elif self.db.should_field_be_multiline(header, self.template_path):
                # Use CTkTextbox for multiline fields (notes, descriptions, etc.)
                widget = ctk.CTkTextbox(self.form_inner, height=80)
                widget.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                # CTkTextbox doesn't use StringVar, so we'll handle it differently
            else:
                # Regular text entry
                widget = ctk.CTkEntry(self.form_inner, placeholder_text=f"Enter {header.lower()}")
                widget.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
            
            # Store widget reference
            self.widgets[header] = widget

        rows_used = (len(headers) + 1) // 2  # number of grid rows consumed
        return start_row + rows_used

    def _populate_fields(self):
        """Populate form fields with existing asset data."""
        # Get column mapping for field name translation
        column_mapping = self.db.get_dynamic_column_mapping(self.template_path)
        reverse_mapping = {v: k for k, v in column_mapping.items()}
        
        for header, widget in self.widgets.items():
            # Get database column name for this header
            db_column = column_mapping.get(header, header.lower().replace(' ', '_'))
            
            # Try to get value from asset data
            value = ""
            if db_column in self.original_asset:
                value = self.original_asset[db_column]
            elif header in self.original_asset:
                value = self.original_asset[header]
            
            # Convert None to empty string
            if value is None:
                value = ""
            
            # Set widget value based on widget type
            try:
                if header in self.widget_vars:
                    # For SearchableDropdown and DatePicker, use the StringVar
                    self.widget_vars[header].set(str(value))
                elif hasattr(widget, 'insert') and hasattr(widget, 'delete'):
                    # For CTkTextbox (multiline) - clear and insert text
                    if isinstance(widget, ctk.CTkTextbox):
                        widget.delete("0.0", "end")
                        widget.insert("0.0", str(value))
                    else:
                        # For CTkEntry widgets
                        widget.delete(0, 'end')
                        widget.insert(0, str(value))
                elif hasattr(widget, 'set'):
                    # For other widgets with set method
                    widget.set(str(value))
            except Exception as e:
                print(f"Warning: Could not populate field {header}: {e}")

    def _get_form_data(self) -> Dict[str, Any]:
        """Extract current form data."""
        form_data = {}
        
        for header, widget in self.widgets.items():
            try:
                if header in self.widget_vars:
                    # For SearchableDropdown and DatePicker, use the StringVar
                    value = self.widget_vars[header].get()
                elif isinstance(widget, ctk.CTkTextbox):
                    # For CTkTextbox (multiline), get text content
                    value = widget.get("0.0", "end-1c")  # Get all text excluding final newline
                elif hasattr(widget, 'get'):
                    # For CTkEntry and other widgets with get method
                    value = widget.get()
                else:
                    value = ""
                
                # Store with both header name and database column name
                form_data[header] = value.strip() if value else ""
                
            except Exception as e:
                print(f"Warning: Could not get value for field {header}: {e}")
                form_data[header] = ""
        
        return form_data

    def _validate_required_fields(self, form_data: Dict[str, Any]) -> List[str]:
        """Validate that required fields are filled."""
        missing_fields = []
        
        for field in self.required_fields:
            value = form_data.get(field, "").strip()
            if not value:
                missing_fields.append(field.replace("*", ""))
        
        return missing_fields

    def _get_changed_fields(self, form_data: Dict[str, Any]) -> Dict[str, Any]:
        """Get only the fields that have changed from original values."""
        changed_fields = {}
        column_mapping = self.db.get_dynamic_column_mapping(self.template_path)
        
        for header, new_value in form_data.items():
            # Get database column name
            db_column = column_mapping.get(header, header.lower().replace(' ', '_'))
            
            # Get original value
            original_value = ""
            if db_column in self.original_asset:
                original_value = self.original_asset[db_column]
            elif header in self.original_asset:
                original_value = self.original_asset[header]
            
            # Convert None to empty string for comparison
            if original_value is None:
                original_value = ""
            
            # Check if value changed
            if str(new_value).strip() != str(original_value).strip():
                changed_fields[db_column] = new_value.strip()
        
        return changed_fields

    def _submit_changes(self):
        """Submit the changes to the database."""
        try:
            # Get form data
            form_data = self._get_form_data()
            
            # Validate required fields
            missing_fields = self._validate_required_fields(form_data)
            if missing_fields:
                messagebox.showerror("Validation Error", 
                                   f"The following required fields are missing:\n\n" + 
                                   "\n".join(f"• {field}" for field in missing_fields))
                self._refocus()
                return
            
            # Get only changed fields
            changed_fields = self._get_changed_fields(form_data)
            
            if not changed_fields:
                messagebox.showinfo("No Changes", "No changes were detected.")
                self._refocus()
                return
            
            # Confirm changes
            change_summary = "\n".join([f"• {field}: {value}" for field, value in changed_fields.items()])
            if not messagebox.askyesno("Confirm Changes", 
                                       f"Are you sure you want to save these changes?\n\n{change_summary}"):
                self._refocus()
                return
            
            # Update the asset in database
            success = self.db.update_asset(self.asset_id, changed_fields)
            
            if success:
                messagebox.showinfo("Success", f"Asset {self.original_asset.get('asset_no', self.asset_id)} updated successfully!")
                
                # Call update callback if provided
                if self.on_update_callback:
                    self.on_update_callback()
                
                # Close the window
                self.window.destroy()
            else:
                messagebox.showerror("Error", "Failed to update asset. Please try again.")
                self._refocus()
                
        except Exception as e:
            messagebox.showerror("Database Error", f"An error occurred while updating the asset:\n{str(e)}")
            self._refocus()

    def _cancel_edit(self):
        """Cancel editing and close the window."""
        # Check if there are any unsaved changes
        form_data = self._get_form_data()
        changed_fields = self._get_changed_fields(form_data)
        
        if changed_fields:
            if not messagebox.askyesno("Unsaved Changes", 
                                       "You have unsaved changes. Are you sure you want to cancel?"):
                self._refocus()
                return
        
        self.window.destroy()

    def _refocus(self):
        """Re-focus and raise this window after messageboxes (Windows quirk)."""
        try:
            self.window.lift()
            self.window.focus_force()
            self.window.grab_set()
        except Exception:
            pass

    def _bind_shortcuts(self):
        """Bind keyboard shortcuts."""
        self.window.bind("<Control-Return>", self._submit_shortcut)
        self.window.bind("<Control-KP_Enter>", self._submit_shortcut)
        self.window.bind("<Escape>", self._cancel_shortcut)

    def _submit_shortcut(self, event=None):
        """Keyboard shortcut for submit."""
        self._submit_changes()
        return "break"

    def _cancel_shortcut(self, event=None):
        """Keyboard shortcut for cancel."""
        self._cancel_edit()
        return "break"

    def _on_closing(self):
        """Handle window closing."""
        self._cancel_edit()


def open_edit_asset_window(parent, asset_id: int, on_update_callback=None):
    """Convenience function to open the edit asset window."""
    return EditAssetWindow(parent, asset_id, on_update_callback=on_update_callback)