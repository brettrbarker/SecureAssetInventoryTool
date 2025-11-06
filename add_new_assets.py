import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
import csv
from datetime import datetime, timedelta
import calendar
from typing import Dict, List
from asset_database import AssetDatabase, db
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from validation import form_validator, asset_validator
from performance_monitoring import performance_monitor
from ui_components import SearchableDropdown, DatePicker
from field_utils import compute_db_fields_from_template, compute_dropdown_fields, compute_date_fields


# AI Initial Prompt:
# In add_new_assets.py, I want tod do the following:
# Read the input file defined in the config settings from main menu. Read that CSV. Parse the header row and create a fillable form
# to fill out all these fields. The window should be 1500x850. 
# If the input form has multiple rows of values, it should parse the unique attributes per field and use those as dropdown selection
# boxes for the new entries being added where the fields mach the static list: "System Name", "*Asset Type", "*Manufacturer",
# "*Model", "Status", "Location", "Room", "Cubicle", "Child Asset? (Y/N)".
# Create a static list of fields not added to the form. I will populate this list later, but it should at least include "Asset No.".
# When pressing "Add Item" button to submt the for, create or append to a csv named with todays date "YYYYMMDD-NewItems.csv".
# The output csv should include the full header row from the input file/template.
# Start with that and I will refine.
#
# Additional AI Prompts:
# I've defined REQUIRED_FIELDS. Make these fields required before the form is submitted. If some are missing, pop up an error message.
# Also, if there is an audit date field, pre-populate it with today's date in the format 08/3/2025
#
# Can the mouse scroll wheel be made to work on this form?
#
# Separate the Required fields into a section at the top with a dividing line and label it Required Fields.
# Additional fields go in a section below. Also, make both section of fields two columns.
#
# Instead of a Reload Template button on the top of the Add Assets screen, make it a Select New Template button.
# This should let a new template be selected and reloaded from. Also add and Output File Box under the template one.
# This should show the default file that will be outputted to with the option to set a custom file name.
#
# Before the Clear button, add a check box for "Add Multiple". If the checkbox is not checked, the form should submit 
# (with confirmation pop-up), clear and the add asset window should close. 
# If the checkbox is checked, after the submission and the added pop-up, the form should clear all the fields except 
# those listed in the "UNIQUE_FIELDS" list that I added to the top of the file. Then the form will be ready to submit 
# a similar item with a different serial number, for example.
#
# After submitting a new asset, the focus goes back to the main menu instead of back to the add asset window. Why?
#
# After refocus on a multi-add item, I want the cursor focus to move to the serial number input box.
# 
# Refactor and change to sqlite database backend.
#
# For the fields that are dropdown menus, is there a way to allow custom input in addition to the dropdown selection options?
# I want to allow typing in the dropdown boxes to search for existing entries, but then if no existing entry is found allow
# adding a manual entry anyway.
#
# It is not actually preserving capitalization. Because it is making it lowercase when it searches,
# when I select add as custom value, it is showing it in lowercase and adds it as lower case.
# Allow conversion to lowercase for searching. But when adding as custom, use the original value typed.

CONFIG_PATH = os.path.join("assets", "config.json")

# Fields that should be rendered as dropdowns (if existing values found)
DROPDOWN_FIELDS = {
    "System Name",
    "*Asset Type",
    "*Manufacturer",
    "*Model",
    "Status",
    "Location",
    "Room",
    "Cubicle",
    "Child Asset? (Y/N)",
}

# Fields to exclude from data entry (will still appear in header / output row as blank)
EXCLUDED_FIELDS = {
    "Asset No.",
    "Version",
    "Client (user names, semicolon delimited)",
    "Client (user names, semicolon delimited",
    "Service Contract? (Y/N)",
    "Service Contract? (Y/N",
    "Contract Expiration Date",
    "Billing Rate Name",
    "Warranty Type",
    "Multi-Install? (Y/N, Child Assets only)",
    "Install Count (Child Assets only)",
    "Reservable? (Y/N)",
    "Discovered Serial Number",
    "Discovery Sync ID",
    "Delete (Y/N)",
    "Delete? (Y/N)",
    "NOTE: * = Field required for new records."
    }

# Required fields that must have data
REQUIRED_FIELDS = {
    "System Name",
    "*Asset Type",
    "*Manufacturer",
    "*Model",
    "Status",
    "Location",
    "Room",
    "Serial Number",
}

# Unique fields that should not be duplicated. These will be values cleared during bulk adds.
UNIQUE_FIELDS = {
    "Serial Number",
    "IP Address",
    "MAC Address",
    "Phone Number",
    "Media Control#",
    "TSCO Control#",
    "Tamper Seal",
    "Network Name"
}

# Header name (case-insensitive) to detect audit date field
AUDIT_DATE_HEADER = "audit date"

def _today_audit_date_str() -> str:
    """Return today's date in requested format: MM/D/YYYY (month zero-padded, day without leading zero)."""
    now = datetime.now()
    return f"{now:%m}/{now.day}/{now:%Y}"


# AI addition from prompt: Is it possible to make the dropdown menus searchable or scrollable?
# Added searchable, scrollable dropdowns via a custom SearchableDropdown widget replacing the previous option menus.
# They support typing to filter and a scrollable list. Only remaining notice is the existing cognitive complexity warning (non-fatal). 
# Let me know if you want multi-select, keyboard navigation refinements, or to refactor for lower complexity.
# NOTE: SearchableDropdown and DatePicker classes are now imported from ui_components module for reusability


class AddNewAssetsWindow:
    """Window for adding new asset rows based on a template CSV header.

    Workflow:
      - Load template CSV from config (default_template_path)
      - Parse header -> build form widgets
      - For dropdown fields, collect unique existing values (if any data rows)
      - On Add Item -> append new row to dated CSV in output directory
    """

    def __init__(self, parent, config = None):
        self.parent = parent
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = config or self.config_manager.get_config()
        
        self.template_path = self.config.default_template_path

        # Create database instance using the configured database path
        self.db = AssetDatabase(self.config.database_path)

        # Load field category sets from config
        self.dropdown_fields = set(self.config.dropdown_fields)
        self.required_fields = set(self.config.required_fields)
        self.excluded_fields = set(self.config.excluded_fields)
        self.unique_fields = set(self.config.unique_fields)

        # Compute template-constrained fields using shared helpers to match other windows
        self.db_fields = compute_db_fields_from_template(self.db, self.config)
        self.dropdown_headers_in_template = set(
            f['display_name'] for f in compute_dropdown_fields(self.db_fields, self.config)
        )



        self.window = ctk.CTkToplevel(parent)
        self.window.title("Add New Assets")
        self.window.geometry("1200x700")
        self.window.minsize(1200, 700)

        # Data structures
        self.headers: List[str] = []
        self.widgets: Dict[str, ctk.CTkBaseClass] = {}
        self.dropdown_value_vars: Dict[str, ctk.StringVar] = {}
        self.unique_values: Dict[str, List[str]] = {}

        # Build UI
        self._build_layout()
        self._load_template_and_build_form()
        self._bind_shortcuts()  # keyboard accelerators
        
        # Set up window close handler
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.window.transient(parent)  # keep on top of parent
        # Removed grab_set() to allow interaction with other windows like Monitor
        # self.window.grab_set()

    # ---------------- UI Construction ---------------- #
    def _build_layout(self):
        # Scrollable form area (using CTkScrollableFrame)
        self.form_inner = ctk.CTkScrollableFrame(self.window)
        self.form_inner.pack(fill="both", expand=True, padx=10, pady=10)

        # Bottom action frame
        action_frame = ctk.CTkFrame(self.window)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        # Instructions
        instructions_label = ctk.CTkLabel(action_frame, 
                                        text="💡 Tips: Ctrl+Enter to add item • Ctrl+Backspace to clear form • Check 'Add Multiple' to keep adding • 'Request Label' auto-requests labels",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(side="left", padx=15, pady=10)
        
        # Add Multiple checkbox (retain non-unique values between submissions)
        self.add_multiple_var = ctk.BooleanVar(value=False)
        add_multi_cb = ctk.CTkCheckBox(action_frame, text="Add Multiple", variable=self.add_multiple_var)
        
        # Request Label checkbox (request label for new asset)
        self.request_label_var = ctk.BooleanVar(value=True)  # Default to checked
        request_label_cb = ctk.CTkCheckBox(action_frame, text="Request Label", variable=self.request_label_var,
                                         fg_color="#2d5a27", hover_color="#1e3f1b")
        request_label_cb.pack(side="right", padx=10, pady=10)
        add_multi_cb.pack(side="right", padx=10, pady=10)
        ctk.CTkButton(action_frame, text="Add Item", command=self._add_item).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(action_frame, text="Clear Form", fg_color="gray", command=self._clear_form).pack(side="right", padx=10, pady=10)

    # --------------- Focus Utility --------------- #
    def _refocus(self):
        """Re-focus and raise this window after messageboxes (Windows quirk)."""
        try:
            self.window.lift()
            self.window.focus_force()
            # Removed grab_set() to keep window non-modal and allow Monitor window to open
            # self.window.grab_set()
        except Exception:
            pass

    def _focus_serial_number(self):
        """Put keyboard focus into the Serial Number field if it exists."""
        widget = self.widgets.get("Serial Number")
        if widget:
            try:
                widget.focus_set()
            except Exception:
                pass

    # --------------- Key Bindings --------------- #
    def _bind_shortcuts(self):
        """Bind global shortcuts for this window.

        Ctrl+Enter     -> submit (Add Item)
        Ctrl+Backspace -> clear entire form
        """
        # Use window-specific bindings instead of bind_all to avoid conflicts with other windows
        self.window.bind("<Control-Return>", self._submit_shortcut)
        self.window.bind("<Control-KP_Enter>", self._submit_shortcut)  # keypad enter
        self.window.bind("<Control-BackSpace>", self._clear_shortcut)

    def _submit_shortcut(self, event=None):  # noqa: D401
        """Keyboard accelerator for submit."""
        self._add_item()
        return "break"

    def _clear_shortcut(self, event=None):  # noqa: D401
        """Keyboard accelerator for clearing the form."""
        self._clear_form()
        self._refocus()
        self._focus_serial_number()
        return "break"

    def _on_closing(self):
        """Handle window closing to clean up resources."""
        try:
            # Clean up any remaining bindings
            self.window.unbind("<Control-Return>")
            self.window.unbind("<Control-KP_Enter>")
            self.window.unbind("<Control-BackSpace>")
        except Exception:
            pass  # Ignore errors during cleanup
        
        # Destroy the window
        self.window.destroy()

    # ---------------- Template & Form ---------------- #


    def _load_template_and_build_form(self):
        """Load template structure and build form using database for dropdown values."""
        if not os.path.exists(self.template_path):
            messagebox.showerror("Template Missing", f"Template file not found:\n{self.template_path}", parent=self.window)
            self._refocus()
            return
        
        # # Update database schema for new template fields
        # try:
        #     schema_updated = self.db.update_schema_for_template(self.template_path)
        #     if schema_updated:
        #         print(f"Database schema updated for template: {self.template_path}")
        # except Exception as e:
        #     print(f"Warning: Could not update database schema: {e}")
        
        try:
            with open(self.template_path, newline='', encoding='utf-8-sig') as f:
                reader = list(csv.reader(f))
        except Exception as e:
            messagebox.showerror("Read Error", f"Failed to read template: {e}", parent=self.window)
            self._refocus()
            return
        if not reader:
            messagebox.showerror("Template Error", "Template CSV is empty.", parent=self.window)
            self._refocus()
            return
        
        self.headers = reader[0]

        # AI Prompt for Modification:
        # Get unique values from database for dropdown fields
        # Instead of getting unique fields from the template for the dropdown lists,
        # the program should get it from unique fields of all items in the database.
        # Change the program to look at the database instead of the template file for unique fields.
        # It should still only build dropdown lists for items specified in the settings menu like it does now.
        # I want it to be dynamic, but I want it to populate the dropdown lists based off the unique values for the database items.
        try:
            # Get dynamic column mapping to know which database columns to query
            column_mapping = self.db.get_dynamic_column_mapping(self.template_path)
            
            # For each dropdown-capable field present in the template, get unique values from the database
            for field in self.dropdown_headers_in_template:
                # Check if this field exists in our template headers first (it should, but keep the guard)
                if field in self.headers:
                    db_column = column_mapping.get(field)
                    if db_column:
                        # Get all unique values from the database for this column
                        values = self.db.get_unique_values(db_column)
                        if values:
                            # Filter out empty/null values and sort
                            filtered_values = [v for v in values if v and v.strip()]
                            if filtered_values:
                                self.unique_values[field] = sorted(filtered_values)
                    else:
                        print(f"Warning: No database column mapping found for dropdown field '{field}'")
        except Exception as e:
            print(f"Warning: Could not load dropdown values from database: {e}")
            # Continue without dropdown values if database access fails
        
        # Split headers into required and additional (preserving original order)
        required_headers = [h for h in self.headers if h in self.required_fields and h not in self.excluded_fields]
        additional_headers = [h for h in self.headers if h not in self.required_fields and h not in self.excluded_fields]

        current_row = 0

        # Section: Required Fields
        if required_headers:
            heading_req = ctk.CTkLabel(self.form_inner, text="Required Fields", anchor="w", font=ctk.CTkFont(size=16, weight="bold"))
            heading_req.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(8,4))
            current_row += 1
            current_row = self._create_field_section(required_headers, start_row=current_row)
            # Divider line
            divider = ctk.CTkFrame(self.form_inner, height=2)
            divider.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=4, pady=(6,10))
            current_row += 1

        # Section: Additional Fields
        if additional_headers:
            heading_add = ctk.CTkLabel(self.form_inner, text="Additional Fields", anchor="w", font=ctk.CTkFont(size=16, weight="bold"))
            heading_add.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(0,4))
            current_row += 1
            current_row = self._create_field_section(additional_headers, start_row=current_row)

        # Configure grid weight for resizing (input columns 1 and 3)
        self.form_inner.columnconfigure(1, weight=1)
        self.form_inner.columnconfigure(3, weight=1)

    def _create_field_section(self, headers: List[str], start_row: int) -> int:
        """Create a two-column (label+input pairs) section for given headers.

        Returns the next available row index after placing widgets.
        """
        for idx, header in enumerate(headers):
            col_group = idx % 2  # 0 or 1
            row_offset = idx // 2
            base_col = col_group * 2
            row = start_row + row_offset

            label = ctk.CTkLabel(self.form_inner, text=header + ":")
            label.grid(row=row, column=base_col, sticky="e", padx=8, pady=4)

            if header in self.dropdown_headers_in_template and header in self.unique_values:
                var = ctk.StringVar(value="")
                values = [""] + self.unique_values[header]
                opt = SearchableDropdown(self.form_inner, values=values, variable=var)
                opt.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                self.dropdown_value_vars[header] = var
                self.widgets[header] = opt
            elif "date" in header.lower():
                # Use DatePicker for any field containing "date"
                var = ctk.StringVar(value="")
                date_picker = DatePicker(self.form_inner, variable=var)
                date_picker.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                self.dropdown_value_vars[header] = var
                self.widgets[header] = date_picker
                # Pre-populate audit date with today's date
                if header.lower() == AUDIT_DATE_HEADER:
                    var.set(_today_audit_date_str())
            elif self.db.should_field_be_multiline(header, self.template_path):
                # Use CTkTextbox for multiline fields (notes, descriptions, etc.)
                textbox = ctk.CTkTextbox(self.form_inner, height=80)
                textbox.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                self.widgets[header] = textbox
            else:
                entry = ctk.CTkEntry(self.form_inner)
                entry.grid(row=row, column=base_col + 1, sticky="we", padx=8, pady=4)
                self.widgets[header] = entry

        rows_used = (len(headers) + 1) // 2  # number of grid rows consumed
        return start_row + rows_used

    def _clear_form_widgets(self):
        for w in self.form_inner.winfo_children():
            w.destroy()
        self.widgets.clear()
        self.dropdown_value_vars.clear()
        self.unique_values.clear()

    def _clear_form(self):
        for header, widget in self.widgets.items():
            if header in self.dropdown_value_vars:
                self.dropdown_value_vars[header].set("")
            elif isinstance(widget, ctk.CTkTextbox):
                # For CTkTextbox, use delete method with "1.0" and "end"
                widget.delete("1.0", "end")
            else:
                # For CTkEntry and other standard widgets
                widget.delete(0, 'end')

    def _clear_for_next_entry(self):
        """Clear only fields that must be unique when adding multiple items.

        NOTE: The project comment above UNIQUE_FIELDS states these are the values
        that should be cleared during bulk adds. We therefore clear ONLY those
        headers in UNIQUE_FIELDS and retain others for faster repeated entry.
        """
        for header, widget in self.widgets.items():
            if header not in self.unique_fields:
                continue
            if header in self.dropdown_value_vars:
                self.dropdown_value_vars[header].set("")
            elif isinstance(widget, ctk.CTkTextbox):
                # For CTkTextbox, use delete method with "1.0" and "end"
                widget.delete("1.0", "end")
            else:
                # For CTkEntry and other standard widgets
                widget.delete(0, 'end')

    # ---------------- Add Item / CSV Write ---------------- #
    
    def _validate_prerequisites(self) -> bool:
        """Check if prerequisites are met for adding an item."""
        if not self.headers:
            messagebox.showerror("Error", "No template loaded.", parent=self.window)
            self._refocus()
            return False
        return True
    
    def _validate_required_fields(self) -> tuple[bool, list[str]]:
        """Validate required fields and return (is_valid, missing_fields)."""
        missing_required = []
        for header in self.headers:
            if header in self.excluded_fields or header not in self.required_fields:
                continue
            widget = self.widgets.get(header)
            if not widget:
                continue
            if header in self.dropdown_value_vars:
                val = self.dropdown_value_vars[header].get().strip()
            elif isinstance(widget, ctk.CTkTextbox):
                # For CTkTextbox (multiline), get text content
                val = widget.get("0.0", "end-1c").strip()
            else:
                val = widget.get().strip()
            if not val:
                missing_required.append(header)
        return len(missing_required) == 0, missing_required
    
    def _extract_form_data(self) -> tuple[list[str], str]:
        """Extract form data and return (row_values, serial_number)."""
        row_values = []
        serial_number = None
        for header in self.headers:
            if header in self.excluded_fields:
                row_values.append("")
                continue
            widget = self.widgets.get(header)
            if not widget:
                row_values.append("")
                continue
            if header in self.dropdown_value_vars:
                val = self.dropdown_value_vars[header].get().strip()
            elif isinstance(widget, ctk.CTkTextbox):
                # For CTkTextbox (multiline), get text content
                val = widget.get("0.0", "end-1c").strip()
            else:
                val = widget.get().strip()
            if header.lower() == "serial number":
                serial_number = val
            row_values.append(val)
        return row_values, serial_number
    
    def _handle_duplicate_overwrite(self, duplicate_idx: int, existing_rows: list, row_values: list, output_path: str) -> bool:
        """Handle duplicate overwrite logic, return success status."""
        result = messagebox.askyesno(
            "Duplicate Serial Number",
            "Item already added. Overwrite?",
            parent=self.window
        )
        if not result:
            self._refocus()
            return False
        
        try:
            # Update the asset in the database
            asset_data = self._convert_row_to_asset_data(row_values)
            success = self.db.update_asset(duplicate_idx, asset_data)
            
            if success:
                # If Request Label checkbox is checked, request a label for the updated asset
                if self.request_label_var.get():
                    try:
                        self.db.request_label(duplicate_idx)
                        print(f"Asset updated in database with ID: {duplicate_idx} - Label requested")
                    except Exception as e:
                        print(f"Warning: Asset updated but label request failed: {e}")
                        # Don't fail the entire operation if label request fails
                else:
                    print(f"Asset updated in database with ID: {duplicate_idx}")
                return True
            else:
                messagebox.showerror("Update Error", "Failed to update asset in database", parent=self.window)
                self._refocus()
                return False
                
        except Exception as e:
            messagebox.showerror("Write Error", f"Failed to overwrite item: {e}", parent=self.window)
            self._refocus()
            return False
    
    def _write_new_item(self, row_values: list, output_path: str) -> bool:
        """Write new item to database, return success status."""
        try:
            # Convert row values to database format
            asset_data = self._convert_row_to_asset_data(row_values)
            
            # Save to database
            asset_id = self.db.add_asset(asset_data)
            
            # If Request Label checkbox is checked, request a label for the new asset
            if self.request_label_var.get():
                try:
                    self.db.request_label(asset_id)
                    print(f"Asset saved to database with ID: {asset_id} - Label requested")
                except Exception as e:
                    print(f"Warning: Asset saved but label request failed: {e}")
                    # Don't fail the entire operation if label request fails
            else:
                print(f"Asset saved to database with ID: {asset_id}")
            
            return True
            
        except Exception as e:
            messagebox.showerror("Write Error", f"Failed to save item: {e}", parent=self.window)
            self._refocus()
            return False
    
    def _convert_row_to_asset_data(self, row_values: list) -> Dict[str, str]:
        """Convert form row values to database asset format."""
        asset_data = {}
        
        # Use dynamic column mapping from database
        column_mapping = self.db.get_dynamic_column_mapping(self.template_path)
        
        for i, header in enumerate(self.headers):
            if i < len(row_values) and row_values[i]:
                db_column = column_mapping.get(header)
                if db_column:
                    value = row_values[i].strip()
                    if value:  # Only add non-empty values
                        asset_data[db_column] = value
        
        return asset_data
    
    def _check_for_duplicate(self, serial_number: str, output_path: str) -> tuple[bool, int, list]:
        """Check for duplicate serial number in database."""
        try:
            existing_asset = self.db.get_asset_by_serial(serial_number)
            if existing_asset:
                # For compatibility with existing code, we'll still return CSV-style data
                # but use database for the actual duplicate check
                return True, existing_asset['id'], []
            return False, -1, []
        except Exception as e:
            print(f"Warning: Could not check for duplicates in database: {e}")
            # Fallback to CSV check if database fails
            return self._check_for_duplicate_csv(serial_number, output_path)
    
    def _check_for_duplicate_csv(self, serial_number: str, output_path: str) -> tuple[bool, int, list]:
        """Fallback CSV duplicate check method - no longer used since CSV is disabled."""
        return False, -1, []  # Always return no duplicates since CSV is disabled
    
    def _format_success_message(self, row_values: list, serial_number: str) -> str:
        """Format success message based on manufacturer/model data."""
        manufacturer = ""
        model = ""
        try:
            manufacturer_idx = self.headers.index("*Manufacturer")
            manufacturer = row_values[manufacturer_idx]
        except ValueError:
            pass
        try:
            model_idx = self.headers.index("*Model")
            model = row_values[model_idx]
        except ValueError:
            pass
            
        if manufacturer and model:
            return f"{manufacturer} {model} added to database\nSN: {serial_number or '(none)'}"
        else:
            return f"Item added to database"
    
    def _handle_validation_error(self, missing_fields: list[str]):
        """Handle validation error display and refocus."""
        messagebox.showerror(
            "Missing Required Fields",
            "Please provide values for required fields:\n" + ", ".join(missing_fields),
            parent=self.window
        )
        self._refocus()
    
    def _handle_success_workflow(self, success_msg: str):
        """Handle post-success UI workflow based on add_multiple mode."""
        if self.add_multiple_var.get():
            messagebox.showinfo("Success", success_msg, parent=self.window)
            self._clear_for_next_entry()
            self._refocus()
            self._focus_serial_number()
        else:
            messagebox.showinfo("Success", success_msg, parent=self.window)
            self._clear_form()
            self.window.destroy()

    @performance_monitor("Add Asset Item")
    def _add_item(self):
        """Main orchestrator method for adding items - delegates to specialized methods."""
        if not self._validate_prerequisites():
            return
            
        # Enhanced validation using our validation framework
        row_values, serial_number = self._extract_form_data()
        
        # Convert row_values list to dictionary for validation
        asset_data = {}
        for i, header in enumerate(self.headers):
            if i < len(row_values):
                asset_data[header] = row_values[i]
        
        # Create template config for validation
        template_config = {
            'required_fields': list(self.required_fields),
            'excluded_fields': list(self.excluded_fields),
            'dropdown_fields': list(self.dropdown_fields),
            'unique_fields': list(self.unique_fields)
        }
        
        # Use centralized validation with template config
        validation_result = asset_validator.validate_asset(asset_data, template_config)
        if not validation_result.is_valid:
            error_messages = validation_result.get_all_messages()
            messagebox.showerror("Validation Error", f"Please fix the following issues:\n\n{error_messages}")
            return
        
        # Show warnings if any
        if validation_result.warnings:
            warning_messages = "\n".join([f"• {warning}" for warning in validation_result.warnings])
            messagebox.showwarning("Validation Warnings", f"Please review:\n\n{warning_messages}")
        
        # Check for unique field conflicts before saving
        if not self._check_unique_field_conflicts(asset_data):
            return  # Conflicts found, user stays on form to make changes
        
        # Handle duplicate check and writing with error handling
        success = safe_execute(
            lambda: self._process_item_write(row_values, serial_number),
            error_handler=error_handler,
            context="adding asset item",
            default_return=False
        )
        
        if not success:
            return
            
        success_msg = self._format_success_message(row_values, serial_number)
        self._handle_success_workflow(success_msg)
    
    def _check_unique_field_conflicts(self, asset_data: Dict[str, str]) -> bool:
        """Check for unique field conflicts with existing assets.
        
        Args:
            asset_data: Dictionary of asset field data from the form
            
        Returns:
            bool: True if no conflicts found, False if conflicts exist
        """
        try:
            # Convert form asset_data to database format for checking
            db_asset_data = self._convert_row_to_asset_data([asset_data.get(header, "") for header in self.headers])
            
            # Get unique fields from config
            unique_fields = list(self.unique_fields)
            
            # Check for conflicts using the database method
            conflicts = self.db.check_unique_field_conflicts(db_asset_data, unique_fields, self.template_path)
            
            if conflicts:
                # Format conflict message
                conflict_messages = []
                for conflict in conflicts:
                    field_name = conflict['field_name']
                    field_value = conflict['field_value']
                    conflicting_asset = conflict['conflicting_asset']
                    
                    # Get asset number and serial number if they exist
                    asset_no = conflicting_asset.get('asset_no', 'N/A')
                    serial_no = conflicting_asset.get('serial_number', 'N/A')
                    
                    conflict_msg = f"• {field_name}: '{field_value}'"
                    conflict_msg += f"\n  Conflicts with Asset No: {asset_no}, Serial Number: {serial_no}"
                    conflict_messages.append(conflict_msg)
                
                error_message = "Unique field conflicts found:\n\n" + "\n\n".join(conflict_messages)
                error_message += "\n\nPlease change the conflicting values and try again."
                
                messagebox.showerror("Unique Field Conflict", error_message, parent=self.window)
                self._refocus()
                return False
            
            return True
            
        except Exception as e:
            print(f"Error checking unique field conflicts: {e}")
            # Allow submission if we can't check (don't block user)
            return True
    
    def _process_item_write(self, row_values: list, serial_number: str) -> bool:
        """Process the item write operation, handling duplicates."""
        try:
            if serial_number:
                # CSV output path commented out for CSV removal - pass empty string for compatibility
                duplicate_found, duplicate_idx, existing_rows = self._check_for_duplicate(serial_number, "")
                if duplicate_found:
                    return self._handle_duplicate_overwrite(duplicate_idx, existing_rows, row_values, "")
                else:
                    return self._write_new_item(row_values, "")
            else:
                return self._write_new_item(row_values, "")
        except Exception:
            return False  # Error already handled in helper methods


# Minimal manual test harness
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    root = ctk.CTk()
    root.geometry("300x100")
    ctk.CTkLabel(root, text="Launcher").pack(pady=10)
    ctk.CTkButton(root, text="Open Add New Assets", command=lambda: AddNewAssetsWindow(root)).pack(pady=10)
    root.mainloop()


