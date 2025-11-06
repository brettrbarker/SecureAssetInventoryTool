import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
import sys
import json
import csv
import shutil
from datetime import datetime
from asset_database import AssetDatabase
from config_manager import ConfigManager
from error_handling import error_handler
from performance_monitoring import performance_monitor
from database_service import DatabaseService
from export_service import export_service


class SettingsWindow:
    def __init__(self, parent, main_menu=None):
        """Initialize the settings window UI and load configuration."""
        self.main_menu = main_menu
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        if main_menu and hasattr(main_menu, 'config'):
            self.config = main_menu.config
        else:
            self.config = self.config_manager.get_config()

        self.config_file = os.path.join("assets", "config.json")
        
        # Get the program's base directory
        # When running as PyInstaller exe, use the executable's directory
        # When running as script, use the script's directory
        if getattr(sys, 'frozen', False):
            # Running as compiled executable - use exe directory
            self.program_dir = os.path.dirname(sys.executable)
        else:
            # Running as script - use script directory
            self.program_dir = os.path.abspath(os.path.dirname(__file__))
        
        # Initialize database service (it gets DB path from config automatically)
        self.database_service = DatabaseService()

        # Cache template headers - load lazily only when needed
        self._cached_headers = None

        # Build window (withdraw first to prevent flash)
        self.window = ctk.CTkToplevel(parent)
        self.window.withdraw()
        self.window.title("Settings")
        self.window.geometry("1180x650")
        self.window.minsize(1000, 580)
        self.window.resizable(True, True)
        self.window.transient(parent)
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)
        
        # Disable automatic updates during widget creation for faster loading
        self.window.update_idletasks()
        
        # Create all widgets in one batch
        self.create_widgets()
        
        # Force a single update after all widgets are created
        self.window.update_idletasks()

        # Position window before showing to avoid repositioning flash
        self._center_window()
        
        # Show window smoothly
        self.window.deiconify()
        self.window.after(10, self._set_modal)
    
    def _convert_to_relative_path_if_appropriate(self, selected_path):
        """Convert absolute path to relative path if it's a subfolder of the program directory.
        
        Args:
            selected_path: The absolute path selected by the user
            
        Returns:
            Relative path if appropriate, otherwise absolute path
        """
        try:
            # Normalize both paths to handle different path separators and resolve symlinks
            selected_abs = os.path.abspath(selected_path)
            program_abs = os.path.abspath(self.program_dir)
            
            # Check if the selected path is under the program directory
            # Use os.path.commonpath to check if they share a common base
            try:
                common_path = os.path.commonpath([selected_abs, program_abs])
                # If the common path is the program directory, selected path is a subfolder
                if os.path.normpath(common_path) == os.path.normpath(program_abs):
                    # Calculate relative path from program directory
                    rel_path = os.path.relpath(selected_abs, program_abs)
                    return rel_path
            except ValueError:
                # Different drives on Windows, cannot create relative path
                pass
            
            # Return absolute path if not a subfolder or on different drives
            return selected_abs
            
        except Exception as e:
            print(f"Error converting path to relative: {e}")
            # If anything goes wrong, return the original path
            return selected_path

    def create_widgets(self):
        # Bottom buttons first - pinned to bottom of window
        self._build_bottom_bar()
        
        # Main scrollable content area that expands but leaves room for bottom bar
        self.main_scroll = ctk.CTkScrollableFrame(self.window)
        self.main_scroll.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        # Title
        title_label = ctk.CTkLabel(self.main_scroll, text="Settings", font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(10, 12))

        # Theme Toggle Section (at top of settings)
        theme_frame = ctk.CTkFrame(self.main_scroll)
        theme_frame.pack(fill="x", padx=10, pady=(0, 8))
        
        theme_inner_frame = ctk.CTkFrame(theme_frame)
        theme_inner_frame.pack(pady=8, padx=10, fill="x")
        
        theme_label = ctk.CTkLabel(theme_inner_frame, text="Appearance Theme:", font=ctk.CTkFont(size=14, weight="bold"))
        theme_label.pack(side="left", padx=(10, 20), pady=10)
        
        # Get current theme and set up theme variable
        current_theme = self.config.get("theme", "dark")
        self.theme_var = ctk.StringVar(value=current_theme.title())  # Convert to title case for display
        
        self.theme_menu = ctk.CTkOptionMenu(theme_inner_frame,
                                           values=["Light", "Dark"],
                                           variable=self.theme_var,
                                           command=self.change_theme,
                                           width=120)
        self.theme_menu.pack(side="right", padx=10, pady=10)

        # Template & output container (top, non-scroll)
        top_frame = ctk.CTkFrame(self.main_scroll)
        top_frame.pack(fill="x", padx=10, pady=(0,8))

        # Template file frame
        template_frame = ctk.CTkFrame(top_frame)
        template_frame.pack(pady=(8, 10), fill="x")
        tmpl_label = ctk.CTkLabel(template_frame, text="Template/Input File:", font=ctk.CTkFont(size=14))
        tmpl_label.pack(anchor="w", padx=10, pady=(8, 4))
        current_path = self.config.get("default_template_path", "assets/default_template.csv")
        self.path_display = ctk.CTkLabel(template_frame, text=current_path, font=ctk.CTkFont(size=12), fg_color=("gray90","gray20"), corner_radius=5, anchor="e")
        self.path_display.pack(fill="x", padx=10, pady=(0,4))
        ctk.CTkButton(template_frame, text="Browse File", width=110, command=self.browse_template_file).pack(anchor="e", padx=10, pady=(0,6))

        # Output directory frame
        output_frame = ctk.CTkFrame(top_frame)
        output_frame.pack(pady=(0, 4), fill="x")
        out_label = ctk.CTkLabel(output_frame, text="Output Directory:", font=ctk.CTkFont(size=14))
        out_label.pack(anchor="w", padx=10, pady=(8, 4))
        current_out_dir = self.config.get("output_directory", "assets/output_files")
        self.out_dir_display = ctk.CTkLabel(output_frame, text=current_out_dir, font=ctk.CTkFont(size=12), fg_color=("gray90","gray20"), corner_radius=5, anchor="e")
        self.out_dir_display.pack(fill="x", padx=10, pady=(0,4))
        ctk.CTkButton(output_frame, text="Browse Folder", width=110, command=self.browse_output_directory).pack(anchor="e", padx=10, pady=(0,6))

        # Database Settings frame
        db_frame = ctk.CTkFrame(top_frame)
        db_frame.pack(pady=(4, 8), fill="x")
        db_label = ctk.CTkLabel(db_frame, text="Database Settings:", font=ctk.CTkFont(size=14, weight="bold"))
        db_label.pack(anchor="w", padx=10, pady=(8, 4))
        
        # Current database file display
        current_db_path = self.config.get("database_path", "assets/asset_database.db")
        self.db_path_display = ctk.CTkLabel(db_frame, text=current_db_path, font=ctk.CTkFont(size=12), fg_color=("gray90","gray20"), corner_radius=5, anchor="e")
        self.db_path_display.pack(fill="x", padx=10, pady=(0,4))
        
        # Database action buttons in a grid layout for dynamic sizing
        db_buttons_frame = ctk.CTkFrame(db_frame, fg_color="transparent")
        db_buttons_frame.pack(fill="x", padx=10, pady=(0,8))
        
        # Configure grid to have 6 columns with equal weight
        for i in range(6):
            db_buttons_frame.grid_columnconfigure(i, weight=1, uniform="db_buttons")
        
        # First row of database buttons (6 buttons)
        ctk.CTkButton(db_buttons_frame, text="Select DB File", command=self._select_database_file).grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        ctk.CTkButton(db_buttons_frame, text="Initialize New DB", command=self._initialize_new_database).grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        ctk.CTkButton(db_buttons_frame, text="Backup Database", command=self._backup_database).grid(row=0, column=2, sticky="ew", padx=2, pady=2)
        ctk.CTkButton(db_buttons_frame, text="Repair/Optimize", command=self._repair_database).grid(row=0, column=3, sticky="ew", padx=2, pady=2)
        ctk.CTkButton(db_buttons_frame, text="Database Info", command=self._show_database_info).grid(row=0, column=4, sticky="ew", padx=2, pady=2)
        ctk.CTkButton(db_buttons_frame, text="Import CSV Data", command=self._import_csv_data).grid(row=0, column=5, sticky="ew", padx=2, pady=2)
        
        # Second row of database buttons (2 remaining buttons)
        ctk.CTkButton(db_buttons_frame, text="Export DB All", command=self._export_database_all).grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        ctk.CTkButton(db_buttons_frame, text="Export DB via Template", command=self._export_database_template).grid(row=1, column=1, sticky="ew", padx=2, pady=2)

        # Field category columns container - this expands within the main scroll area
        cols_frame = ctk.CTkFrame(self.main_scroll)
        cols_frame.pack(padx=10, pady=(0,10), fill="both", expand=True)
        for i in range(4):
            cols_frame.grid_columnconfigure(i, weight=1, uniform="fields")
        cols_frame.grid_rowconfigure(0, weight=1)

        # Load headers once if not already cached
        headers = self._load_template_headers()
        
        categories = [
            ("Dropdown Fields", "dropdown_fields"),
            ("Required Fields", "required_fields"),
            ("Excluded Fields", "excluded_fields"),
            ("Unique Fields (Reset With Multi-Add)", "unique_fields"),
        ]
        
        # Create all columns in batch to minimize redraws
        for col_idx, (label, key) in enumerate(categories):
            self._build_field_column(cols_frame, col_idx, label, key, headers, self.config.get(key))

        # Monitor Asset Details Settings frame - lazy loaded for performance
        monitor_frame = ctk.CTkFrame(self.main_scroll)
        monitor_frame.pack(pady=(8, 8), fill="x")
        monitor_label = ctk.CTkLabel(monitor_frame, text="Monitor Asset Details:", font=ctk.CTkFont(size=14, weight="bold"))
        monitor_label.pack(anchor="w", padx=10, pady=(8, 4))
        
        # Monitor columns container - create placeholder that will be populated on demand
        monitor_cols_frame = ctk.CTkFrame(monitor_frame)
        monitor_cols_frame.pack(padx=10, pady=(0, 8), fill="x")
        
        # Add loading placeholder
        self._monitor_placeholder = ctk.CTkLabel(monitor_cols_frame, 
                                                 text="Loading monitor fields...",
                                                 font=ctk.CTkFont(size=12),
                                                 text_color="gray60")
        self._monitor_placeholder.pack(pady=20)
        
        # Store data for lazy loading
        self._monitor_cols_frame = monitor_cols_frame
        self._monitor_loaded = False
        
        # Schedule lazy loading after window is visible
        self.window.after(100, self._lazy_load_monitor_fields)

        # Bulk Update Presets Management section
        self.presets_frame = ctk.CTkFrame(self.main_scroll)
        self.presets_frame.pack(pady=(8, 8), fill="x")
        presets_label = ctk.CTkLabel(self.presets_frame, text="Bulk Update Presets:", font=ctk.CTkFont(size=14, weight="bold"))
        presets_label.pack(anchor="w", padx=10, pady=(8, 4))
        
        # Presets list and management
        self._build_presets_section(self.presets_frame)

        # Report Fields section
        report_fields_frame = ctk.CTkFrame(self.main_scroll)
        report_fields_frame.pack(pady=(8, 8), fill="x")
        report_fields_label = ctk.CTkLabel(report_fields_frame, text="Report Fields:", font=ctk.CTkFont(size=14, weight="bold"))
        report_fields_label.pack(anchor="w", padx=10, pady=(8, 4))
        
        # Report fields columns container
        report_cols_frame = ctk.CTkFrame(report_fields_frame)
        report_cols_frame.pack(padx=10, pady=(0, 8), fill="x")
        
        # Configure grid for three equal columns
        for i in range(3):
            report_cols_frame.grid_columnconfigure(i, weight=1, uniform="report_fields")
        report_cols_frame.grid_rowconfigure(0, weight=1)

        # Create the report field columns
        report_categories = [
            ("Label Output Fields", "label_output_fields"),
            ("HMR Fields (Not Yet Implemented)", "hmr_fields"), 
            ("Destruction Report Fields (Not Yet Implemented)", "destruction_report_fields"),
        ]
        
        for col_idx, (label, key) in enumerate(report_categories):
            self._build_report_field_column(report_cols_frame, col_idx, label, key, headers, self.config.get(key))

    def _build_presets_section(self, parent):
        """Build the preset management section - optimized for faster loading."""
        # Description label
        desc_label = ctk.CTkLabel(parent, text="Manage saved bulk update presets", 
                                 font=ctk.CTkFont(size=12), text_color="gray60")
        desc_label.pack(anchor="w", padx=10, pady=(0, 8))
        
        # Content frame
        content_frame = ctk.CTkFrame(parent)
        content_frame.pack(fill="x", padx=10, pady=(0, 8))
        
        # Get current presets
        presets = self.config.get("bulk_update_presets", {})
        
        if not presets:
            # No presets message
            no_presets_label = ctk.CTkLabel(content_frame, text="No presets saved yet", 
                                           font=ctk.CTkFont(size=12), text_color="gray50")
            no_presets_label.pack(pady=20)
        else:
            # Show count instead of full list if many presets
            preset_count = len(presets)
            
            if preset_count > 10:
                # Just show summary for many presets to load faster
                summary_label = ctk.CTkLabel(content_frame, 
                                            text=f"{preset_count} presets saved (click 'Expand' to view all)",
                                            font=ctk.CTkFont(size=12))
                summary_label.pack(pady=10)
                
                expand_btn = ctk.CTkButton(content_frame, text="Expand Presets", width=120,
                                          command=lambda: self._expand_presets(content_frame, presets))
                expand_btn.pack(pady=(0, 10))
            else:
                # Show all presets directly for smaller counts
                self._show_all_presets(content_frame, presets)
    
    def _expand_presets(self, parent, presets):
        """Expand and show all presets when user clicks expand."""
        # Clear current content
        for widget in parent.winfo_children():
            widget.destroy()
        
        # Show all presets
        self._show_all_presets(parent, presets)
    
    def _show_all_presets(self, parent, presets):
        """Display all preset items."""
        # Presets list with scrollable frame
        presets_scroll = ctk.CTkScrollableFrame(parent, height=150)
        presets_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Batch create preset items
        for preset_name, preset_data in presets.items():
            self._create_preset_item(presets_scroll, preset_name, preset_data)

    def _create_preset_item(self, parent, preset_name, preset_data):
        """Create a single preset item with details and delete button."""
        # Preset item frame
        item_frame = ctk.CTkFrame(parent, fg_color=("gray90", "gray20"))
        item_frame.pack(fill="x", pady=2, padx=2)
        
        # Left side - preset info
        info_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        info_frame.pack(side="left", fill="both", expand=True, padx=10, pady=8)
        
        # Preset name
        name_label = ctk.CTkLabel(info_frame, text=preset_name, 
                                 font=ctk.CTkFont(size=14, weight="bold"))
        name_label.pack(anchor="w")
        
        # Get preset type and fields (handle both old and new format)
        if isinstance(preset_data, dict) and "fields" in preset_data:
            # New format with type
            preset_type = preset_data.get("type", "user")
            fields_data = preset_data["fields"]
        else:
            # Old format - assume user type
            preset_type = "user"
            fields_data = preset_data
        
        # Preset details - show field count and first few fields with values
        field_count = len(fields_data)
        
        if field_count == 1:
            # Single field - show full details
            item = fields_data[0]
            field = item.get("field", "Unknown")
            operation = item.get("operation", "replace")
            value = item.get("value", "")
            
            # Handle special values
            if value == "current_date":
                display_value = "current date"
            else:
                display_value = value if value else "(empty)"
            
            details_text = f"1 field: {field} → {operation.title()} with \"{display_value}\""
        else:
            # Multiple fields - show summary with first 2 detailed, then count
            details_parts = []
            for i, item in enumerate(fields_data[:2]):
                field = item.get("field", "Unknown")
                operation = item.get("operation", "replace")
                value = item.get("value", "")
                
                # Handle special values
                if value == "current_date":
                    display_value = "current date"
                else:
                    display_value = value if value else "(empty)"
                
                details_parts.append(f"{field} → {display_value}")
            
            details_text = f"{field_count} fields: {', '.join(details_parts)}"
            if field_count > 2:
                details_text += f" +{field_count - 2} more"
        
        details_label = ctk.CTkLabel(info_frame, text=details_text, 
                                    font=ctk.CTkFont(size=11), text_color="gray60")
        details_label.pack(anchor="w")
        
        # Right side - delete button (disabled for system presets)
        if preset_type == "system":
            # System preset - show disabled delete button
            delete_btn = ctk.CTkButton(item_frame, text="System", width=70, height=28,
                                      fg_color="gray", hover_color="gray", 
                                      state="disabled", text_color="darkgray")
        else:
            # User preset - normal delete button
            delete_btn = ctk.CTkButton(item_frame, text="Delete", width=70, height=28,
                                      fg_color="red", hover_color="darkred",
                                      command=lambda: self._delete_preset(preset_name))
        delete_btn.pack(side="right", padx=10, pady=8)

    def _delete_preset(self, preset_name):
        """Delete a preset after confirmation."""
        # Check if preset exists and get its type
        presets = self.config.get("bulk_update_presets", {})
        if preset_name not in presets:
            messagebox.showerror("Error", f"Preset '{preset_name}' not found.")
            return
        
        preset_data = presets[preset_name]
        
        # Get preset type (handle both old and new format)
        if isinstance(preset_data, dict) and "type" in preset_data:
            preset_type = preset_data["type"]
        else:
            preset_type = "user"  # Assume user type for old format
        
        # Prevent deletion of system presets
        if preset_type == "system":
            messagebox.showwarning("Cannot Delete", 
                                 f"The '{preset_name}' preset is a system preset and cannot be deleted.")
            return
        
        # Confirmation dialog
        result = messagebox.askyesno("Delete Preset", 
                                   f"Are you sure you want to delete the preset '{preset_name}'?\n\nThis action cannot be undone.")
        
        if result:
            # Remove from config
            if "bulk_update_presets" in self.config and preset_name in self.config["bulk_update_presets"]:
                del self.config["bulk_update_presets"][preset_name]
                
                # Save config
                self.config_manager.save_config(self.config)
                
                # Refresh the presets section first so user sees the updated list
                self._refresh_presets_section()
                
                # Then show success message
                messagebox.showinfo("Preset Deleted", f"Preset '{preset_name}' has been deleted.")

    def _refresh_presets_section(self):
        """Refresh the presets section to reflect changes."""
        # Reload config to ensure we have latest data
        self.config = self.config_manager.get_config()
        
        # Destroy all child widgets in the presets frame except the title label
        for widget in self.presets_frame.winfo_children():
            # Keep the title label "Bulk Update Presets:"
            if isinstance(widget, ctk.CTkLabel) and "Bulk Update Presets:" in str(widget.cget("text")):
                continue
            widget.destroy()
        
        # Rebuild the presets section content
        self._build_presets_section(self.presets_frame)
        
        # Force window to update display
        self.window.update_idletasks()

    def _build_bottom_bar(self):
        # Bottom bar anchored to window bottom, not packed after content
        bottom_bar = ctk.CTkFrame(self.window)
        bottom_bar.pack(side="bottom", fill="x", padx=10, pady=(0, 10))
        
        # Left group for action buttons
        left_actions = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        left_actions.pack(side="left")
        ctk.CTkButton(left_actions, text="Reload From Template", width=170, command=self._reload_from_template).pack(side="left", padx=(0,12))
        ctk.CTkButton(left_actions, text="Reset to Defaults", width=150, command=self._reset_to_defaults).pack(side="left")
        
        # Spacer to push Close button to the right
        spacer = ctk.CTkLabel(bottom_bar, text="")
        spacer.pack(side="left", expand=True)
        
        # Close button on the right
        ctk.CTkButton(bottom_bar, text="Close", width=90, command=self.close_window).pack(side="right")
        

    def close_window(self):
        # Always save configuration on close, no prompts
        self.save_config()
        self.window.grab_release()
        self.window.destroy()
    
    def change_theme(self, theme):
        """Save theme setting to config and prompt user to restart."""
        try:
            # Update and save config
            self.config_manager.update_config(theme=theme.lower())
            self.config_manager.save_config()  # Explicitly save to disk
            
            # Refresh our local config reference
            self.config = self.config_manager.get_config()
            
            # Update main menu config if it exists
            if self.main_menu:
                self.main_menu.config = self.config_manager.get_config()
            
            # Inform user that restart is required
            messagebox.showinfo("Theme Setting Saved", 
                              f"Theme changed to {theme}.\n\n"
                              f"Please restart the application for the theme change to take effect.")
                
        except Exception as e:
            print(f"Error saving theme setting: {e}")
            messagebox.showerror("Theme Error", f"Failed to save theme setting: {str(e)}")
    
    def _center_window(self):
        """Center the window on the screen (monitor)."""
        self.window.update_idletasks()
        
        # Get window dimensions
        width = self.window.winfo_reqwidth()
        height = self.window.winfo_reqheight()
        
        # Always center on screen (primary monitor) for consistent positioning
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calculate center position on screen
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.window.geometry(f"+{x}+{y}")

    def _set_modal(self):
        try:
            self.window.grab_set()
            self.window.focus()
        except Exception:
            self.window.focus()
    
    # AI Prompted Addition
    # Instead of a box to type in, this should just be a text field label with the path from the config and a selection button to choose another file if desired.
    def browse_template_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Template File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.config.get("default_template_path", ""))
        )
        if file_path:
            # Convert to relative path if appropriate
            file_path = self._convert_to_relative_path_if_appropriate(file_path)
            
            if self._validate_template_file(file_path):
                self._process_template_change(file_path)
    
    def _validate_template_file(self, file_path):
        """Validate that the template file exists and is readable."""
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "Selected template file does not exist.")
            return False
            
        try:
            with open(file_path, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                headers = next(reader, [])
                if not headers:
                    messagebox.showerror("Error", "Template file is empty or invalid.")
                    return False
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read template file: {str(e)}")
            return False
        return True
    
    def _process_template_change(self, file_path):
        """Process the template change including database schema updates."""
        try:
            current_db = self.config.get("database_path", "assets/asset_database.db")
            db = AssetDatabase(current_db)
            compatibility = db.verify_template_compatibility(file_path)
            
            # Check for errors in compatibility check
            if 'error' in compatibility:
                messagebox.showerror("Error", f"Template compatibility check failed: {compatibility['error']}")
                return
            
            # Get list of new fields that will be created
            new_fields = [field for field in compatibility.get('field_details', []) if field.get('status') == 'will_create']
            
            if new_fields and not self._confirm_schema_changes(new_fields):
                return
            
            if new_fields:
                if not self._update_database_schema(db, file_path, new_fields):
                    return
            
            self._apply_template_change(file_path)
            
        except Exception as e:
            messagebox.showerror("Database Error", f"Error updating database schema: {str(e)}")
    
    def _confirm_schema_changes(self, new_fields):
        """Show confirmation dialog for schema changes."""
        new_cols_str = ", ".join([field['header'] for field in new_fields])
        message = f"This template will add {len(new_fields)} new columns to the database:\n\n{new_cols_str}\n\nDo you want to continue?"
        return messagebox.askyesno("Database Schema Update", message)
    
    def _update_database_schema(self, db, file_path, new_fields):
        """Update the database schema and show result."""
        schema_updated = db.update_schema_for_template(file_path)
        if schema_updated:
            messagebox.showinfo("Success", f"Database schema updated successfully.\nAdded {len(new_fields)} new columns.")
            return True
        else:
            messagebox.showerror("Error", "Failed to update database schema.")
            return False
    
    def _apply_template_change(self, file_path):
        """Apply the template change to UI and config."""
        self.path_display.configure(text=file_path)
        self.config["default_template_path"] = file_path

        # Clear cached headers before reloading
        self._cached_headers = None

        # Try to update the database schema to add any missing columns from the new template
        db_path = self.config.get("database_path", "")
        if db_path and os.path.exists(db_path):
            try:
                db = AssetDatabase(db_path)
                schema_result = db.update_schema_for_template(file_path)
                if schema_result:
                    # Reload headers and UI after schema update
                    self._cached_headers = self._load_template_headers()
                    self._reload_from_template()
                    messagebox.showinfo("Template Updated", "Template file changed successfully.\nDatabase schema was checked and updated if needed.\nField categories have been refreshed.")
                    return
                else:
                    # Even if schema update reported failure, continue to apply template in UI
                    self._cached_headers = self._load_template_headers()
                    self._reload_from_template()
                    messagebox.showwarning("Template Updated", "Template file changed, but updating the database schema failed.\nPlease check the application log for details.")
                    return
            except Exception as e:
                # Log and continue to apply the template in the UI
                try:
                    error_handler.logger.error(f"Error updating schema on template change: {e}")
                except Exception:
                    pass

        # If no valid database configured or schema update not attempted, still apply template to UI
        self._cached_headers = self._load_template_headers()
        self._reload_from_template()
        messagebox.showinfo("Template Updated", "Template file changed successfully.\nDatabase schema has not been modified (no valid database configured).\nField categories have been refreshed.")
            
    def browse_output_directory(self):
        dir_path = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.config.get("output_directory", "assets/output_files")
        )
        if dir_path:
            # Convert to relative path if appropriate
            dir_path = self._convert_to_relative_path_if_appropriate(dir_path)
            
            self.out_dir_display.configure(text=dir_path)
            self.config["output_directory"] = dir_path

    def _lazy_load_monitor_fields(self):
        """Lazy load monitor fields after window is visible to improve startup time."""
        if self._monitor_loaded:
            return
        
        # Remove placeholder
        if hasattr(self, '_monitor_placeholder') and self._monitor_placeholder:
            self._monitor_placeholder.destroy()
        
        # Configure grid for three equal columns
        for i in range(3):
            self._monitor_cols_frame.grid_columnconfigure(i, weight=1, uniform="monitor_fields")
        self._monitor_cols_frame.grid_rowconfigure(0, weight=1)
        
        # Get headers
        headers = self._load_template_headers()
        
        # Create monitor field columns
        monitor_categories = [
            ("Primary (Row 0)", "monitor_primary_fields"),
            ("Secondary (Row 1)", "monitor_secondary_fields"), 
            ("Tertiary (Row 2)", "monitor_tertiary_fields"),
        ]
        
        for col_idx, (label, key) in enumerate(monitor_categories):
            self._build_monitor_field_column(self._monitor_cols_frame, col_idx, label, key, headers, self.config.get(key))
        
        self._monitor_loaded = True

    # -------- Field Category Helpers -------- #
    def _load_template_headers(self):
        """Load template headers with caching."""
        if self._cached_headers is not None:
            return self._cached_headers
            
        path = self.config.get("default_template_path", "assets/default_template.csv")
        if not os.path.exists(path):
            self._cached_headers = []
            return []
        try:
            with open(path, newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                self._cached_headers = next(reader, [])
                return self._cached_headers
        except Exception:
            self._cached_headers = []
            return []

    def _build_field_column(self, parent, column, title, key, headers, default=None):
        """Build a field selection column optimized for performance."""
        # Outer column frame
        col = ctk.CTkFrame(parent, fg_color=("gray92","gray18"))
        col.grid(row=0, column=column, sticky="nsew", padx=6, pady=6)
        
        # Title label
        title_label = ctk.CTkLabel(col, text=title, font=ctk.CTkFont(size=14, weight="bold"))
        title_label.pack(anchor="w", padx=8, pady=(8,4))
        
        # Calculate minimum height based on line count (ensure at least 4 lines visible)
        min_lines = 4
        line_height = 25
        min_height = min_lines * line_height + 20
        
        # Use larger of minimum height or responsive height based on headers count
        header_count = len(headers) if headers else 4
        responsive_height = min(max(min_height, header_count * line_height + 20), 220)
        
        # Scrollable area for checkboxes with calculated height
        scroll = ctk.CTkScrollableFrame(col, width=250, height=responsive_height)
        scroll.pack(fill="both", expand=True, padx=6, pady=(0,8))
        
        # Determine selection efficiently
        if default is None:
            selection = set(self._get_builtin_defaults().get(key, []))
        else:
            selection = set(default)
            
        if key not in self.config:
            self.config[key] = list(selection)
        
        # Pre-create all variables and widgets in one pass
        vars_list = []
        for h in headers:
            var = ctk.IntVar(value=1 if h in selection else 0)
            chk = ctk.CTkCheckBox(scroll, text=h, variable=var, 
                                 command=lambda k=key: self._update_multiselect_config(k))
            chk.pack(anchor="w")
            vars_list.append((h, var))
        
        # Store reference for updates
        setattr(self, f"_ms_{key}", {
            "frame": col, 
            "vars": vars_list, 
            "key": key, 
            "scroll": scroll
        })
        
        return col

    def _build_monitor_field_column(self, parent, column, title, key, headers, default=None):
        """Build a column for monitor field selection with drag & drop ordering"""
        # Outer column frame
        col = ctk.CTkFrame(parent, fg_color=("gray92","gray18"))
        col.grid(row=0, column=column, sticky="nsew", padx=6, pady=6)
        
        # Title label
        title_label = ctk.CTkLabel(col, text=title, font=ctk.CTkFont(size=14, weight="bold"))
        title_label.pack(anchor="w", padx=8, pady=(8,4))
        
        # Instructions label
        instr_label = ctk.CTkLabel(col, text="Drag to reorder • Check to include", 
                                  font=ctk.CTkFont(size=10), text_color="gray60")
        instr_label.pack(anchor="w", padx=8, pady=(0,4))
        
        # Calculate height for monitor columns
        min_lines = 4
        line_height = 30  # Slightly taller for drag handles
        min_height = min_lines * line_height + 40
        
        # Use responsive height for monitor columns
        header_count = len(headers) if headers else 4
        responsive_height = min(max(min_height, header_count * line_height + 40), 200)
        
        # Scrollable area for drag & drop list
        scroll = ctk.CTkScrollableFrame(col, width=220, height=responsive_height)
        scroll.pack(fill="both", expand=True, padx=6, pady=(0,8))
        
        # Get current configuration
        current_config = self.config.get(key, [])
        if not current_config and default:
            current_config = list(default)
            self.config[key] = current_config
        
        # Create drag & drop list
        self._create_drag_drop_list(scroll, key, headers, current_config)
        
        # Store reference for updates
        setattr(self, f"_monitor_{key}", {
            "frame": col,
            "scroll": scroll,
            "key": key
        })
        
        return col

    def _build_report_field_column(self, parent, column, title, key, headers, default=None):
        """Build a column for report field selection (multi-select checkboxes)"""
        # Outer column frame
        col = ctk.CTkFrame(parent, fg_color=("gray92","gray18"))
        col.grid(row=0, column=column, sticky="nsew", padx=6, pady=6)
        
        # Title label
        title_label = ctk.CTkLabel(col, text=title, font=ctk.CTkFont(size=14, weight="bold"))
        title_label.pack(anchor="w", padx=8, pady=(8,4))
        
        # Instructions label
        instr_label = ctk.CTkLabel(col, text="Select fields to include in reports", 
                                  font=ctk.CTkFont(size=10), text_color="gray60")
        instr_label.pack(anchor="w", padx=8, pady=(0,4))
        
        # Calculate height for report columns
        min_lines = 6
        line_height = 25
        min_height = min_lines * line_height + 40
        
        # Use responsive height for report columns
        header_count = len(headers) if headers else 6
        responsive_height = min(max(min_height, header_count * line_height + 40), 250)
        
        # Scrollable area for checkbox list
        scroll = ctk.CTkScrollableFrame(col, width=220, height=responsive_height)
        scroll.pack(fill="both", expand=True, padx=6, pady=(0,8))
        
        # Get current configuration - initialize with empty list if None
        current_config = self.config.get(key, [])
        if current_config is None:
            current_config = []
            self.config[key] = current_config
        
        # Create checkbox list
        self._create_report_field_checkboxes(scroll, key, headers, current_config)
        
        # Store reference for updates
        setattr(self, f"_report_{key}", {
            "frame": col,
            "scroll": scroll,
            "key": key
        })
        
        return col

    def _create_report_field_checkboxes(self, parent, key, available_fields, current_config):
        """Create checkbox list for report field selection"""
        if not available_fields:
            no_fields_label = ctk.CTkLabel(parent, text="No fields available", 
                                         font=ctk.CTkFont(size=12), text_color="gray50")
            no_fields_label.pack(pady=10)
            return
        
        # Store checkbox variables
        checkbox_vars = {}
        
        # Create checkboxes for each available field - match format of other columns
        for field in available_fields:
            var = ctk.BooleanVar(value=field in current_config)
            checkbox_vars[field] = var
            
            # Checkbox - simple pack like Dropdown Fields section
            checkbox = ctk.CTkCheckBox(
                parent, 
                text=field, 
                variable=var,
                command=lambda f=field, v=var: self._update_report_field_selection(key, f, v.get())
            )
            checkbox.pack(anchor="w", padx=4, pady=2)
        
        # Store references for cleanup
        setattr(self, f"_checkboxes_{key}", checkbox_vars)

    def _update_report_field_selection(self, key, field, selected):
        """Update the report field selection when checkbox changes"""
        current_config = self.config.get(key, [])
        if current_config is None:
            current_config = []
            
        if selected and field not in current_config:
            current_config.append(field)
        elif not selected and field in current_config:
            current_config.remove(field)
            
        self.config[key] = current_config

    def _create_drag_drop_list(self, parent, key, available_fields, current_config):
        """Create a drag & drop list for field ordering - optimized for batch creation"""
        # Defer updates until all widgets are created
        parent.update_idletasks()
        
        # Create ordered list: first show configured fields in order, then remaining fields
        configured_fields = [f for f in current_config if f in available_fields]
        remaining_fields = [f for f in available_fields if f not in current_config]
        ordered_fields = configured_fields + remaining_fields
        
        # Store the list items and their state
        list_items = []
        
        # Batch create all widgets before packing to reduce redraws
        for i, field_name in enumerate(ordered_fields):
            # Create frame for each list item
            item_frame = ctk.CTkFrame(parent, height=28, fg_color=("gray85", "gray25"))
            
            # Drag handle (⋮⋮)
            drag_label = ctk.CTkLabel(item_frame, text="⋮⋮", width=20, 
                                     font=ctk.CTkFont(size=12), text_color="gray50")
            
            # Checkbox for inclusion
            is_selected = field_name in current_config
            var = ctk.IntVar(value=1 if is_selected else 0)
            checkbox = ctk.CTkCheckBox(item_frame, text="", variable=var, width=20,
                                      command=lambda k=key: self._update_monitor_drag_config(k))
            
            # Field name label
            field_label = ctk.CTkLabel(item_frame, text=field_name, anchor="w",
                                      font=ctk.CTkFont(size=11))
            
            # Store item data
            item_data = {
                'frame': item_frame,
                'field_name': field_name,
                'checkbox': checkbox,
                'var': var,
                'label': field_label,
                'drag_label': drag_label,
                'index': i
            }
            list_items.append(item_data)
            
            # Bind drag events
            self._bind_drag_events(item_data, key)
        
        # Now pack everything at once - much faster than pack-as-you-go
        for item_data in list_items:
            item_data['frame'].pack(fill="x", padx=2, pady=1)
            item_data['frame'].pack_propagate(False)
            item_data['drag_label'].pack(side="left", padx=(4,2))
            item_data['checkbox'].pack(side="left", padx=(2,4))
            item_data['label'].pack(side="left", fill="x", expand=True, padx=(0,4))
        
        # Store list items reference
        setattr(self, f"_drag_items_{key}", list_items)

    def _bind_drag_events(self, item_data, key):
        """Bind drag and drop events to list item - optimized with fewer bindings"""
        frame = item_data['frame']
        drag_label = item_data['drag_label']
        
        # Use single lambda for each event type instead of separate ones per widget
        drag_start = lambda e: self._start_drag(e, item_data, key)
        drag_move = lambda e: self._drag_motion(e, item_data, key)
        drag_end = lambda e: self._end_drag(e, item_data, key)
        
        # Bind events to drag handle and frame - reuse lambda functions
        for widget in [drag_label, frame]:
            widget.bind("<Button-1>", drag_start)
            widget.bind("<B1-Motion>", drag_move)
            widget.bind("<ButtonRelease-1>", drag_end)
            
        # Simplified hover feedback - single function pair instead of closures
        frame.bind("<Enter>", lambda e: self._on_item_enter(frame, drag_label))
        frame.bind("<Leave>", lambda e: self._on_item_leave(frame, drag_label))
        drag_label.bind("<Enter>", lambda e: self._on_item_enter(frame, drag_label))
        drag_label.bind("<Leave>", lambda e: self._on_item_leave(frame, drag_label))
    
    def _on_item_enter(self, frame, drag_label):
        """Hover enter effect"""
        frame.configure(fg_color=("gray80", "gray30"))
        drag_label.configure(text_color="gray70")
    
    def _on_item_leave(self, frame, drag_label):
        """Hover leave effect"""
        frame.configure(fg_color=("gray85", "gray25"))
        drag_label.configure(text_color="gray50")

    def _start_drag(self, event, item_data, key):
        """Start dragging a list item"""
        self._drag_data = {
            'item': item_data,
            'key': key,
            'start_y': event.y_root,
            'original_bg': item_data['frame'].cget('fg_color')
        }
        # Visual feedback for dragging
        item_data['frame'].configure(fg_color=("gray75", "gray35"))

    def _drag_motion(self, event, item_data, key):
        """Handle drag motion"""
        if not hasattr(self, '_drag_data'):
            return
            
        # Calculate relative position change
        y_diff = event.y_root - self._drag_data['start_y']
        
        # Determine if we should swap positions
        if abs(y_diff) > 30:  # Threshold for swapping
            list_items = getattr(self, f"_drag_items_{key}", [])
            dragged_item = self._drag_data['item']
            current_index = dragged_item['index']
            
            if y_diff > 0 and current_index < len(list_items) - 1:
                # Moving down
                self._swap_items(list_items, current_index, current_index + 1, key)
                self._drag_data['start_y'] = event.y_root
            elif y_diff < 0 and current_index > 0:
                # Moving up  
                self._swap_items(list_items, current_index, current_index - 1, key)
                self._drag_data['start_y'] = event.y_root

    def _end_drag(self, event, item_data, key):
        """End dragging and restore visual state"""
        if hasattr(self, '_drag_data'):
            # Restore original background
            item_data['frame'].configure(fg_color=self._drag_data['original_bg'])
            delattr(self, '_drag_data')
            
        # Update configuration after drag
        self._update_monitor_drag_config(key)

    def _swap_items(self, list_items, index1, index2, key):
        """Swap two items in the drag & drop list"""
        if index1 == index2 or index1 < 0 or index2 < 0:
            return
        if index1 >= len(list_items) or index2 >= len(list_items):
            return
            
        # Swap items in list
        list_items[index1], list_items[index2] = list_items[index2], list_items[index1]
        
        # Update indices
        list_items[index1]['index'] = index1
        list_items[index2]['index'] = index2
        
        # Repack frames in new order
        for i, item in enumerate(list_items):
            item['frame'].pack_forget()
        
        for i, item in enumerate(list_items):
            item['frame'].pack(fill="x", padx=2, pady=1)

    def _update_monitor_drag_config(self, key):
        """Update monitor field configuration from drag & drop list"""
        list_items = getattr(self, f"_drag_items_{key}", [])
        if not list_items:
            return
            
        # Get selected fields in their current order
        selected_fields = []
        for item in list_items:
            if item['var'].get() == 1:  # If checkbox is checked
                selected_fields.append(item['field_name'])
        
        # Update configuration
        self.config[key] = selected_fields

    def _update_multiselect_config(self, key):
        data = getattr(self, f"_ms_{key}", None)
        if not data or not isinstance(data, dict) or "vars" not in data:
            return
        selected = [h for h, var in data["vars"] if var.get() == 1]
        self.config[key] = selected

    def _reload_from_template(self):
        # Clear cached headers to force reload from file
        self._cached_headers = None
        headers = self._load_template_headers()
        for key in ["dropdown_fields","required_fields","excluded_fields","unique_fields"]:
            data = getattr(self, f"_ms_{key}", None)
            if not data or not isinstance(data, dict) or "scroll" not in data:
                continue
            selection = self.config.get(key, [])
            
            # Get the scrollable frame reference
            scroll_frame = data["scroll"]
            
            # Clear all checkboxes from scrollable frame
            for widget in scroll_frame.winfo_children():
                widget.destroy()
            
            # Recreate checkboxes directly in scrollable frame
            vars_list = []
            for h in headers:
                var = ctk.IntVar(value=1 if h in selection else 0)
                chk = ctk.CTkCheckBox(scroll_frame, text=h, variable=var, 
                                     command=lambda k=key: self._update_multiselect_config(k))
                chk.pack(anchor="w")
                vars_list.append((h, var))
            
            # Update stored reference
            if isinstance(data, dict):
                data["vars"] = vars_list
            self._update_multiselect_config(key)
        
        # Also reload monitor field columns with the new headers
        for key in ["monitor_primary_fields", "monitor_secondary_fields", "monitor_tertiary_fields"]:
            data = getattr(self, f"_monitor_{key}", None)
            if not data or not isinstance(data, dict) or "scroll" not in data:
                continue
            
            current_config = self.config.get(key, [])
            scroll_frame = data["scroll"]
            
            # Clear all widgets from scrollable frame
            for widget in scroll_frame.winfo_children():
                widget.destroy()
            
            # Recreate drag & drop list with new headers
            self._create_drag_drop_list(scroll_frame, key, headers, current_config)
        
        # Also reload report field columns with the new headers
        for key in ["label_output_fields", "hmr_fields", "destruction_report_fields"]:
            data = getattr(self, f"_report_{key}", None)
            if not data or not isinstance(data, dict) or "scroll" not in data:
                continue
            
            current_config = self.config.get(key, [])
            scroll_frame = data["scroll"]
            
            # Clear all widgets from scrollable frame
            for widget in scroll_frame.winfo_children():
                widget.destroy()
            
            # Recreate checkbox list with new headers
            self._create_report_field_checkboxes(scroll_frame, key, headers, current_config)

    def _reset_to_defaults(self):
        defaults = self._get_builtin_defaults()
        for key, selection in defaults.items():
            self.config[key] = selection[:]
            data = getattr(self, f"_ms_{key}", None)
            if not data or not isinstance(data, dict) or "vars" not in data:
                continue
            # Update checkbox states
            for h, var in data["vars"]:
                var.set(1 if h in selection else 0)
            self._update_multiselect_config(key)
        
        # Also reset monitor field defaults
        monitor_defaults = {
            'monitor_primary_fields': ["Serial Number", "Asset No."],
            'monitor_secondary_fields': ["*Manufacturer", "*Model"],
            'monitor_tertiary_fields': ["Room", "Cubicle", "System Name"]
        }
        
        for key, selection in monitor_defaults.items():
            self.config[key] = selection[:]
            data = getattr(self, f"_monitor_{key}", None)
            if not data or not isinstance(data, dict) or "vars" not in data:
                continue
            # Update checkbox states  
            for h, var in data["vars"]:
                var.set(1 if h in selection else 0)
            self._update_monitor_config(key)
        
        # Also reset report field defaults
        report_defaults = {
            'label_output_fields': ["Asset No.", "Serial Number", "*Manufacturer", "*Model"],
            'hmr_fields': ["Asset No.", "Serial Number", "*Manufacturer", "*Model", "Location", "Room"],
            'destruction_report_fields': ["Asset No.", "Serial Number", "*Manufacturer", "*Model", "Status", "Location"]
        }
        
        for key, selection in report_defaults.items():
            self.config[key] = selection[:]
            data = getattr(self, f"_report_{key}", None)
            if not data or not isinstance(data, dict):
                continue
            
            # Update checkbox states if available
            checkbox_vars = getattr(self, f"_checkboxes_{key}", {})
            for field, var in checkbox_vars.items():
                var.set(field in selection)

    def _get_builtin_defaults(self):
        return {
            'dropdown_fields': ["System Name","*Asset Type","*Manufacturer","*Model","Status","Location","Room","Cubicle","Child Asset? (Y/N)"],
            'required_fields': ["System Name","*Asset Type","*Manufacturer","*Model","Status","Location","Room","Serial Number"],
            'excluded_fields': ["Asset No.","Version","Client (user names, semicolon delimited)","Service Contract? (Y/N)","Contract Expiration Date","Billing Rate Name","Warranty Type","Multi-Install? (Y/N, Child Assets only)","Install Count (Child Assets only)","Reservable? (Y/N)","Discovered Serial Number","Discovery Sync ID","Delete? (Y/N)","Delete (Y/N)","NOTE: * = Field required for new records."],
            'unique_fields': ["Serial Number","IP Address","MAC Address","Phone Number","Media Control#","TSCO Control#","Tamper Seal","Network Name"],
            'monitor_primary_fields': ["Serial Number", "Asset No."],
            'monitor_secondary_fields': ["*Manufacturer", "*Model"],
            'monitor_tertiary_fields': ["Room", "Cubicle", "System Name"],
            'label_output_fields': ["Asset No.", "Serial Number", "*Manufacturer", "*Model"],
            'hmr_fields': ["Asset No.", "Serial Number", "*Manufacturer", "*Model", "Location", "Room"],
            'destruction_report_fields': ["Asset No.", "Serial Number", "*Manufacturer", "*Model", "Status", "Location"],
        }

    # ---------------- Configuration Helpers ---------------- #
    def load_config(self):
        """Load configuration from JSON file (used only if main_menu not provided)."""
        cfg_path = os.path.join("assets", "config.json")
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError):
                return self.default_config()
        return self.default_config()

    def default_config(self):
        d = {
            "theme": "dark",
            "default_template_path": "assets/default_template.csv",
            "output_directory": "assets/output_files",
            "database_path": "assets/asset_database.db",
        }
        d.update(self._get_builtin_defaults())
        return d

    def save_config(self):
        """Persist current config to disk and update main menu if present."""
        self._normalize_field_sets()
        
        # Use ConfigManager to save the configuration
        try:
            # Save using ConfigManager (it will save the current _config instance)
            success = self.config_manager.save_config(self.config)
            
            if not success:
                messagebox.showerror("Error", "Failed to save configuration to file.")
                return
            
            # Update main menu config if available
            if self.main_menu:
                self.main_menu.config = self.config_manager.get_config()
                
        except Exception as e:
            error_handler.handle_exception(e, "saving configuration")
            
        # Ensure output directory exists if specified
        out_dir = self.config.get("output_directory")
        if out_dir:
            try:
                os.makedirs(out_dir, exist_ok=True)
            except OSError:
                pass

    def _normalize_field_sets(self):
        # Ensure lists, remove duplicates while preserving order, and resolve conflicts
        def dedupe(seq):
            seen = set()
            out = []
            for x in seq:
                if x not in seen:
                    seen.add(x)
                    out.append(x)
            return out
        keys = ["dropdown_fields","required_fields","excluded_fields","unique_fields"]
        for k in keys:
            if k in self.config and isinstance(self.config[k], list):
                self.config[k] = dedupe(self.config[k])
        # Remove any required fields that ended up in excluded list
        req = set(self.config.get("required_fields", []))
        excl = self.config.get("excluded_fields", [])
        self.config["excluded_fields"] = [x for x in excl if x not in req]

    # -------- Database Management Methods -------- #
    
    def _select_database_file(self):
        """Allow user to select a different SQLite database file."""
        file_path = filedialog.askopenfilename(
            title="Select Database File",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.config.get("database_path", "assets/"))
        )
        if file_path:
            try:
                # Convert to relative path if appropriate
                file_path = self._convert_to_relative_path_if_appropriate(file_path)
                
                # Test if the file is a valid SQLite database
                test_db = AssetDatabase(file_path)
                # Try to get table info to verify it's a valid asset database
                test_db.get_table_columns()
                
                # Update config and UI
                self.config["database_path"] = file_path
                self.db_path_display.configure(text=file_path)
                messagebox.showinfo("Success", f"Database file changed to:\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Invalid database file or connection failed:\n{str(e)}")
    
    def _initialize_new_database(self):
        """Create a new empty database and initialize it with the current template."""
        # Get save location for new database
        file_path = filedialog.asksaveasfilename(
            title="Create New Database",
            defaultextension=".db",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.config.get("database_path", "assets/"))
        )
        if not file_path:
            return
        
        # Convert to relative path if appropriate
        file_path = self._convert_to_relative_path_if_appropriate(file_path)
            
        try:
            # Backup existing database if it exists
            current_db = self.config.get("database_path", "assets/asset_database.db")
            if os.path.exists(current_db):
                backup_name = self._create_backup_filename(current_db)
                shutil.copy2(current_db, backup_name)
                messagebox.showinfo("Backup Created", f"Existing database backed up to:\n{backup_name}")
            
            # Remove new file if it exists
            if os.path.exists(file_path):
                os.remove(file_path)
            
            # Create new database and initialize with current template
            new_db = AssetDatabase(file_path)
            template_path = self.config.get("default_template_path", "assets/default_template.csv")
            
            if os.path.exists(template_path):
                new_db.update_schema_for_template(template_path)
                messagebox.showinfo("Success", f"New database created and initialized with template:\n{file_path}")
            else:
                messagebox.showinfo("Success", f"New database created (no template found for initialization):\n{file_path}")
            
            # Update config and UI
            self.config["database_path"] = file_path
            self.db_path_display.configure(text=file_path)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create new database:\n{str(e)}")
    
    def _backup_database(self):
        """Create a backup copy of the current database."""
        try:
            current_db = self.config.get("database_path", "assets/asset_database.db")
            if not os.path.exists(current_db):
                messagebox.showerror("Error", "Current database file not found.")
                return
            
            backup_name = self._create_backup_filename(current_db)
            shutil.copy2(current_db, backup_name)
            
            # Get file size for user info
            size_mb = os.path.getsize(backup_name) / (1024 * 1024)
            messagebox.showinfo("Backup Complete", f"Database backed up successfully:\n{backup_name}\n\nSize: {size_mb:.2f} MB")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to backup database:\n{str(e)}")
    
    def _repair_database(self):
        """Run database repair and optimization operations."""
        try:
            current_db = self.config.get("database_path", "assets/asset_database.db")
            db = AssetDatabase(current_db)
            
            # Show confirmation dialog
            if not messagebox.askyesno("Database Repair", 
                                     "This will run VACUUM and integrity check operations.\n\n"
                                     "This may take a few moments for large databases.\n\n"
                                     "Continue?"):
                return
            
            # Run repair operations
            with db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Check integrity
                cursor.execute("PRAGMA integrity_check")
                integrity_result = cursor.fetchone()[0]
                
                # Run VACUUM to optimize database
                cursor.execute("VACUUM")
                
                # Get database size after optimization
                cursor.execute("PRAGMA page_count")
                page_count = cursor.fetchone()[0]
                cursor.execute("PRAGMA page_size")
                page_size = cursor.fetchone()[0]
                size_mb = (page_count * page_size) / (1024 * 1024)
            
            status = "✓ OK" if integrity_result == "ok" else f"⚠ {integrity_result}"
            messagebox.showinfo("Repair Complete", 
                              f"Database repair completed:\n\n"
                              f"Integrity Check: {status}\n"
                              f"Database optimized (VACUUM completed)\n"
                              f"Current size: {size_mb:.2f} MB")
            
        except Exception as e:
            messagebox.showerror("Error", f"Database repair failed:\n{str(e)}")
    
    def _show_database_info(self):
        """Display detailed information about the current database."""
        try:
            current_db = self.config.get("database_path", "assets/asset_database.db")
            db = AssetDatabase(current_db)  # Use the configured database path
            
            # Get database statistics
            with db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get asset count
                cursor.execute("SELECT COUNT(*) FROM assets")
                asset_count = cursor.fetchone()[0]
                
                # Get audit log count (check if table exists first)
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='asset_audit_log'")
                audit_table_exists = cursor.fetchone() is not None
                
                if audit_table_exists:
                    cursor.execute("SELECT COUNT(*) FROM asset_audit_log")
                    audit_count = cursor.fetchone()[0]
                else:
                    audit_count = 0
                
                # Get database size and page info
                cursor.execute("PRAGMA page_count")
                page_count = cursor.fetchone()[0]
                cursor.execute("PRAGMA page_size")
                page_size = cursor.fetchone()[0]
                size_bytes = page_count * page_size
                size_mb = size_bytes / (1024 * 1024)
            
            # Get table column count
            columns = db.get_table_columns()
            column_count = len(columns)
            
            # Get all database tables
            tables = db.get_database_tables()
            table_list = "\n".join([f"  • {table}" for table in tables])
            
            # Get file modification date
            mod_time = datetime.fromtimestamp(os.path.getmtime(current_db)).strftime("%Y-%m-%d %H:%M:%S")
            
            # Display information
            info_text = (f"Database Information:\n"
                        f"{'='*40}\n\n"
                        f"File: {current_db}\n"
                        f"Last Modified: {mod_time}\n"
                        f"Size: {size_mb:.2f} MB ({size_bytes:,} bytes)\n\n"
                        f"Data Statistics:\n"
                        f"Assets: {asset_count:,}\n"
                        f"Audit Entries: {audit_count:,}\n"
                        f"Table Columns: {column_count}\n\n"
                        f"Database Tables:\n{table_list}\n\n"
                        f"Database Structure:\n"
                        f"Page Size: {page_size} bytes\n"
                        f"Page Count: {page_count:,}")
            
            messagebox.showinfo("Database Information", info_text)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get database information:\n{str(e)}")
    
    def _import_csv_data(self):
        """Import data from a CSV file into the database."""
        csv_file = None
        
        # Check if template file is configured and has data
        template_path = self.config.get("default_template_path", "")
        if template_path and os.path.exists(template_path):
            try:
                # Check if template has data rows (not just headers)
                with open(template_path, 'r', newline='', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    headers = next(reader, [])
                    first_data_row = next(reader, None)
                    
                    if first_data_row:  # Template has data to import
                        # Count total rows
                        f.seek(0)
                        next(csv.reader(f))  # Skip header
                        row_count = sum(1 for _ in csv.reader(f))
                        
                        # Ask user if they want to use the template file
                        template_name = os.path.basename(template_path)
                        choice_msg = (
                            f"A Template/Import file is already selected:\n\n"
                            f"File: {template_name}\n"
                            f"Rows: {row_count}\n\n"
                            f"Do you want to import from this file?\n\n"
                            f"• Yes - Import from this file\n"
                            f"• No - Select a different file\n"
                            f"• Cancel - Cancel import"
                        )
                        
                        result = messagebox.askyesnocancel("Import from Template?", choice_msg)
                        
                        if result is None:  # Cancel
                            return
                        elif result:  # Yes - use template file
                            csv_file = template_path
                        # else: No - will prompt for file below
            except Exception as e:
                # If there's an error reading the template, just continue to file selection
                print(f"Error checking template file: {e}")
        
        # If no file selected yet, prompt user to select one
        if not csv_file:
            csv_file = filedialog.askopenfilename(
                title="Select CSV File to Import",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialdir=self.config.get("output_directory", "assets/output_files")
            )
            if not csv_file:
                return
            
        try:
            current_db = self.config.get("database_path", "assets/asset_database.db")
            db = AssetDatabase(current_db)
            
            # Show confirmation dialog with preview
            with open(csv_file, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                headers = next(reader, [])
                # Count remaining rows for preview
                f.seek(0)
                next(csv.reader(f))  # Skip header
                row_count = sum(1 for _ in csv.reader(f))
            
            preview_text = f"Import Preview:\n\nFile: {os.path.basename(csv_file)}\nColumns: {len(headers)}\nRows to import: {row_count}\nSample Headers: {', '.join(headers[:5])}{'...' if len(headers) > 5 else ''}\n\nThis will add the data to your existing database.\nDuplicates will be detected and you'll be asked for confirmation.\n\nContinue with import?"
            
            if not messagebox.askyesno("Confirm Import", preview_text):
                return
            
            # Create duplicate handling callback
            def handle_duplicate(duplicate_type, duplicate_value, existing_data, new_data):
                return self._show_duplicate_dialog(duplicate_type, duplicate_value, existing_data, new_data)
            
            # Perform the import with duplicate handling
            imported_count = db.import_from_csv(csv_file, handle_duplicate)
            messagebox.showinfo("Import Complete", f"Successfully imported {imported_count} records from:\n{csv_file}")
            
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import CSV data:\n{str(e)}")
    
    def _show_duplicate_dialog(self, duplicate_type, duplicate_value, existing_data, new_data):
        """Show dialog for handling duplicate assets during import."""
        # AI Prompted Addition
        # Fix the buttons at the bottom of _show_duplicate_dialog to always show no matter the window size.
        # Do it the same way that you did for the "Reload From Template" and other buttons on the bottom of the settings page.
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Duplicate Asset Found")
        dialog.geometry("600x500")
        dialog.transient(self.window)
        dialog.grab_set()
        dialog.minsize(600, 500)
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"+{x}+{y}")
        
        result = {"action": "skip"}
        
        # Define button actions first
        def on_overwrite():
            result["action"] = "overwrite"
            dialog.destroy()
        
        def on_skip():
            result["action"] = "skip"
            dialog.destroy()
        
        def on_overwrite_all():
            result["action"] = "overwrite_all"
            dialog.destroy()
        
        def on_skip_all():
            result["action"] = "skip_all"
            dialog.destroy()
        
        # Bottom buttons first - pinned to bottom of dialog
        bottom_bar = ctk.CTkFrame(dialog)
        bottom_bar.pack(side="bottom", fill="x", padx=20, pady=(0, 20))
        
        # Button layout in bottom bar
        button_row1 = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        button_row1.pack(pady=(0, 5))
        
        ctk.CTkButton(button_row1, text="Overwrite", width=100, command=on_overwrite).pack(side="left", padx=(0, 10))
        ctk.CTkButton(button_row1, text="Skip", width=100, command=on_skip).pack(side="left")
        
        button_row2 = ctk.CTkFrame(bottom_bar, fg_color="transparent")
        button_row2.pack()
        
        ctk.CTkButton(button_row2, text="Overwrite All", width=100, command=on_overwrite_all).pack(side="left", padx=(0, 10))
        ctk.CTkButton(button_row2, text="Skip All", width=100, command=on_skip_all).pack(side="left")
        
        # Main content area that expands but leaves room for bottom bar
        content_frame = ctk.CTkFrame(dialog)
        content_frame.pack(fill="both", expand=True, padx=20, pady=(20, 0))
        
        # Title
        title_label = ctk.CTkLabel(content_frame, text="Duplicate Asset Detected", font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=(20, 10))
        
        # Duplicate info
        info_text = f"Found duplicate {duplicate_type}: {duplicate_value}"
        info_label = ctk.CTkLabel(content_frame, text=info_text, font=ctk.CTkFont(size=12))
        info_label.pack(pady=(0, 15))
        
        # Comparison frame that expands to fill remaining space
        comparison_frame = ctk.CTkFrame(content_frame)
        comparison_frame.pack(fill="both", expand=True, padx=10, pady=(0, 20))
        
        # Existing data column
        existing_frame = ctk.CTkFrame(comparison_frame)
        existing_frame.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)
        
        existing_label = ctk.CTkLabel(existing_frame, text="Existing Database Record:", font=ctk.CTkFont(size=12, weight="bold"))
        existing_label.pack(pady=(10, 5))
        
        existing_text = ctk.CTkTextbox(existing_frame, height=250)
        existing_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Format existing data for display
        existing_display = []
        for key, value in existing_data.items():
            if value and key not in ['id', 'created_date', 'modified_date']:
                existing_display.append(f"{key}: {value}")
        existing_text.insert("1.0", "\n".join(existing_display))
        existing_text.configure(state="disabled")
        
        # New data column
        new_frame = ctk.CTkFrame(comparison_frame)
        new_frame.pack(side="right", fill="both", expand=True, padx=(5, 10), pady=10)
        
        new_label = ctk.CTkLabel(new_frame, text="New CSV Record:", font=ctk.CTkFont(size=12, weight="bold"))
        new_label.pack(pady=(10, 5))
        
        new_text = ctk.CTkTextbox(new_frame, height=250)
        new_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Format new data for display
        new_display = []
        for key, value in new_data.items():
            if value:
                new_display.append(f"{key}: {value}")
        new_text.insert("1.0", "\n".join(new_display))
        new_text.configure(state="disabled")
        
        # Wait for dialog to close
        dialog.wait_window()
        return result["action"]
    
    def _export_database_all(self):
        """Export all database fields and all assets to CSV."""
        try:
            # Get save location for export file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"database_export_all_{timestamp}.csv"
            
            file_path = filedialog.asksaveasfilename(
                title="Export Database - All Fields",
                defaultextension=".csv",
                initialfile=default_filename,
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialdir=self.config.get("output_directory", "assets/output_files")
            )
            if not file_path:
                return
            
            current_db = self.config.get("database_path", "assets/asset_database.db")
            db = AssetDatabase(current_db)
            
            # Get all assets
            assets = db.search_assets(limit=999999)  # Get all assets with very high limit
            
            if not assets:
                messagebox.showinfo("No Data", "No assets found in database.")
                return
            
            # Get all column names from the database
            all_columns = db.get_table_columns()
            
            # Write to CSV
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=all_columns)
                writer.writeheader()
                
                for asset in assets:
                    # Ensure all columns are present in the row
                    row = {col: asset.get(col, '') for col in all_columns}
                    writer.writerow(row)
            
            messagebox.showinfo("Export Complete", 
                              f"Database exported successfully:\n{file_path}\n\n"
                              f"Records exported: {len(assets)}\n"
                              f"Fields exported: {len(all_columns)}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export database:\n{str(e)}")
    
    @performance_monitor("Export Database Template")
    def _export_database_template(self):
        """Export database using template fields with filtering options."""
        export_service.export_database_template(self.window)
    
    def _create_backup_filename(self, original_path):
        """Create a backup filename with timestamp."""
        base_name = os.path.splitext(original_path)[0]
        extension = os.path.splitext(original_path)[1]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base_name}_backup_{timestamp}{extension}"

