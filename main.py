import customtkinter as ctk
from tkinter import messagebox
import os
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from performance_monitoring import performance_monitor

# Version Information
VERSION = "1.0.251114"  # Format: Major.Minor.YYMMDD

# Import settings and other modules with error handling for invalid config paths
try:
    from settings_menu import SettingsWindow
    from add_new_assets import AddNewAssetsWindow
    from export_service import export_service
    from monitor_window import MonitorWindow
    from reports_analysis import open_reports_analysis_window
    MODULES_LOADED = True
except Exception as e:
    print(f"Error loading modules (likely due to invalid config paths): {e}")
    MODULES_LOADED = False
    # We'll handle this in the main class

# AI Prompts:
# Position the main menu in the top-left corner of the screen.
#
# Add an "Export Assets via Template" Button to the main menu that
# calls the exact same "_export_database_template" function that is
# called when this button is clicked in settings.

class MainMenu:
    def __init__(self, root):
        self.root = root
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = self.config_manager.get_config()
        
        # Ensure all required directories exist at startup
        self.config_manager.ensure_directories()
        
        # Check if modules loaded successfully, if not validate and fix config
        if not MODULES_LOADED:
            self._handle_module_load_failure()
            return
        
        # Validate configuration paths at startup
        self._validate_config_paths()
        
        # Create automatic backup of database if it exists
        if MODULES_LOADED:
            from database_service import database_service
            database_service.create_automatic_backup()
        
        # Set theme from config
        ctk.set_appearance_mode(self.config.theme)
        ctk.set_default_color_theme("dark-blue")

        self.root.geometry("500x680")
        self.root.minsize(500, 680)
        self.root.resizable(False, False)
        self.root.title("Secure Asset Inventory Tool")

        # Position the main window in top-left corner
        self._center_window()

        # Create title section with improved styling
        title_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        title_frame.pack(pady=(20, 10))
        
        # Main title - clean and professional
        self.title_label = ctk.CTkLabel(title_frame, 
                                       text="Secure Asset Inventory Tool", 
                                       font=ctk.CTkFont(size=32, weight="bold"))
        self.title_label.pack()
        
        # Subtitle for context
        self.subtitle_label = ctk.CTkLabel(title_frame, 
                                          text="Professional Asset Management System", 
                                          font=ctk.CTkFont(size=14),
                                          text_color=("gray50", "gray70"))
        self.subtitle_label.pack(pady=(5, 0))
        
        # Elegant divider with gradient effect
        divider_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        divider_frame.pack(fill="x", padx=40, pady=(15, 25))
        
        divider = ctk.CTkFrame(divider_frame, height=2, fg_color=("gray70", "gray30"))
        divider.pack(fill="x")

        # Create main buttons frame for two-column layout
        buttons_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        buttons_frame.pack(fill="both", expand=True, padx=25, pady=(0, 10))
        
        # Configure grid for two equal columns
        buttons_frame.grid_columnconfigure(0, weight=1)
        buttons_frame.grid_columnconfigure(1, weight=1)
        
        # Enhanced button dimensions and styling
        button_width = 190
        button_height = 95
        
        # Row 0: Add New Assets, Browse Assets
        self.buttonNewAssets = ctk.CTkButton(buttons_frame, text="📝 Add New Assets", 
                                           font=ctk.CTkFont(size=16, weight="bold"), 
                                           command=self.open_add_new_assets,
                                           width=button_width, height=button_height,
                                           corner_radius=12)
        self.buttonNewAssets.grid(row=0, column=0, padx=8, pady=8, sticky="ew")

        self.buttonBrowseAssets = ctk.CTkButton(buttons_frame, text="🔍 Browse Assets", 
                                              font=ctk.CTkFont(size=16, weight="bold"), 
                                              command=self.open_browse_assets,
                                              width=button_width, height=button_height,
                                              corner_radius=12)
        self.buttonBrowseAssets.grid(row=0, column=1, padx=8, pady=8, sticky="ew")

        # Row 1: Export Assets (left) and Bulk Update (right)
        self.buttonExportAssets = ctk.CTkButton(buttons_frame, text="📤 Export Assets\nvia Template", 
                                              font=ctk.CTkFont(size=16, weight="bold"), 
                                              command=self.export_assets_via_template,
                                              width=button_width, height=button_height,
                                              corner_radius=12)
        self.buttonExportAssets.grid(row=1, column=0, padx=8, pady=8, sticky="ew")

        self.buttonBulkUpdate = ctk.CTkButton(buttons_frame, text="⚙️ Search/Change\nAssets", 
                                            font=ctk.CTkFont(size=16, weight="bold"), 
                                            command=self.open_bulk_update_assets,
                                            width=button_width, height=button_height,
                                            corner_radius=12)
        self.buttonBulkUpdate.grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        # Row 2: Reports & Analysis (left) and Monitor Changes (right)
        self.buttonReports = ctk.CTkButton(buttons_frame, text="📊 Reports\nand Analysis", 
                                          font=ctk.CTkFont(size=16, weight="bold"), 
                                          command=self.open_reports_analysis,
                                          width=button_width, height=button_height,
                                          corner_radius=12)
        self.buttonReports.grid(row=2, column=0, padx=8, pady=8, sticky="ew")

        self.buttonMonitor = ctk.CTkButton(buttons_frame, text="👁️ Monitor Changes", 
                                         font=ctk.CTkFont(size=16, weight="bold"), 
                                         command=self.open_monitor,
                                         width=button_width, height=button_height,
                                         corner_radius=12)
        self.buttonMonitor.grid(row=2, column=1, padx=8, pady=8, sticky="ew")

        # Row 3: Settings (bottom row, spanning both columns)
        self.buttonSettings = ctk.CTkButton(buttons_frame, text="⚙️ Settings", 
                                          font=ctk.CTkFont(size=18, weight="bold"), 
                                          command=self.open_settings,
                                          width=button_width*2, height=button_height,
                                          corner_radius=12,
                                          fg_color=("gray50", "gray30"), 
                                          hover_color=("gray60", "gray40"))
        self.buttonSettings.grid(row=3, column=0, columnspan=2, padx=8, pady=8, sticky="ew")
        
        # Footer with version info
        footer_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        footer_frame.pack(side="bottom", pady=(10, 15))
        
        version_label = ctk.CTkLabel(footer_frame, 
                                   text=f"v{VERSION} • BRB", 
                                   font=ctk.CTkFont(size=11),
                                   text_color=("gray40", "gray60"))
        version_label.pack()

    def _center_window(self):
        """Position the main window in the top-left corner of screen."""
        self.root.update_idletasks()
        
        # Get screen dimensions - use primary screen for reliable positioning
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Position in top-left corner with small margins
        left_margin = 20  # Small margin from left edge
        top_margin = 20   # Small margin from top edge
        
        x = left_margin
        y = top_margin
        
        # Set window position in top-left corner
        self.root.geometry(f"500x680+{x}+{y}")
    
    def _handle_module_load_failure(self):
        """Handle the case where modules couldn't load due to invalid config paths."""
        # Reset problematic config paths
        self.config_manager.update_config(
            default_template_path="",
            output_directory="", 
            database_path=""
        )
        self.config = self.config_manager.get_config()
        
        # Set basic theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        
        # Create minimal UI
        self.root.geometry("500x300")
        self.root.minsize(500, 300)
        self.root.title("Configuration Required")
        self._center_window_simple()
        
        # Show error message and instructions
        error_frame = ctk.CTkFrame(self.root)
        error_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        title_label = ctk.CTkLabel(error_frame, 
                                  text="Configuration Error", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(20, 10))
        
        message_label = ctk.CTkLabel(error_frame, 
                                    text="Invalid file paths detected in configuration.\nPlease configure valid file and folder paths.",
                                    font=ctk.CTkFont(size=14))
        message_label.pack(pady=10)
        
        restart_button = ctk.CTkButton(error_frame, 
                                      text="Restart Application", 
                                      command=self._restart_application,
                                      font=ctk.CTkFont(size=16))
        restart_button.pack(pady=20)
    
    def _center_window_simple(self):
        """Simple window centering for error state."""
        self.root.update_idletasks()
        width = 500
        height = 300
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def _restart_application(self):
        """Restart the application to reload modules with corrected config."""
        import sys
        import subprocess
        self.root.destroy()
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit()
    
    def _validate_config_paths(self):
        """Validate configuration paths and reset invalid ones."""
        try:
            # Special handling for template file - check if it exists
            template_path = self.config.default_template_path
            template_missing = not template_path or not os.path.exists(template_path)
            
            if template_missing:
                # Schedule template prompt after UI is created
                self.root.after(100, self._prompt_for_template)
                return
            
            paths_to_check = {
                'default_template_path': self.config.default_template_path,
                'output_directory': self.config.output_directory,
                'database_path': self.config.database_path
            }
            
            invalid_paths = []
            config_updates = {}
            
            for path_name, path_value in paths_to_check.items():
                path_valid = True
                
                if not path_value:
                    path_valid = False
                else:
                    try:
                        # For database_path, test if the directory can be created
                        if path_name == 'database_path':
                            db_dir = os.path.dirname(path_value)
                            if db_dir:
                                # Try to create the directory to test permissions/validity
                                os.makedirs(db_dir, exist_ok=True)
                            # Check if file exists or directory is writable
                            if not os.path.exists(path_value):
                                # Test if we can create the file
                                try:
                                    with open(path_value, 'a'):
                                        pass
                                    # Clean up test file if it was just created
                                    if not os.path.getsize(path_value):
                                        os.remove(path_value)
                                except (PermissionError, OSError):
                                    path_valid = False
                        else:
                            # For other paths, just check existence
                            if not os.path.exists(path_value):
                                path_valid = False
                    except (PermissionError, OSError, TypeError, ValueError) as e:
                        print(f"Error validating {path_name}: {e}")
                        path_valid = False
                
                if not path_valid:
                    invalid_paths.append(path_name)
                    config_updates[path_name] = ""
            
            # If any paths are invalid, update config and show warning
            if invalid_paths:
                # Update config with empty strings for invalid paths
                self.config_manager.update_config(**config_updates)
                # Refresh our local config reference
                self.config = self.config_manager.get_config()
                
                # Schedule the popup and settings menu to open after UI is created
                self.root.after(100, self._show_config_warning)
                
        except Exception as e:
            print(f"Error during path validation: {e}")
            # If validation itself fails, reset all paths and show warning
            self.config_manager.update_config(
                default_template_path="",
                output_directory="", 
                database_path=""
            )
            self.config = self.config_manager.get_config()
            self.root.after(100, self._show_config_warning)
    
    def _show_config_warning(self):
        """Show configuration warning and open settings menu."""
        messagebox.showwarning(
            "Configuration Required", 
            "Please Configure File and Folder Paths",
            parent=self.root
        )
        # Open settings menu
        self.open_settings()
    
    def _prompt_for_template(self):
        """Prompt user to use default template or select their own."""
        result = messagebox.askyesno(
            "No Template File Loaded",
            "No template file is currently loaded.\n\n"
            "Would you like to use the default template?\n\n"
            "• Yes - Load the default template (assets/templates/default_template.csv)\n"
            "• No - Go to Settings to select a template file",
            parent=self.root
        )
        
        if result:
            # User wants to use default template
            default_template = "assets/templates/default_template.csv"
            
            # Check if default template exists
            if os.path.exists(default_template):
                try:
                    # Update database schema with template columns
                    from asset_database import AssetDatabase
                    db_path = self.config.database_path
                    
                    if db_path:
                        db = AssetDatabase(db_path)
                        schema_updated = db.update_schema_for_template(default_template)
                        
                        if schema_updated:
                            # Update config with default template
                            self.config_manager.update_config(default_template_path=default_template)
                            self.config_manager.save_config()
                            self.config = self.config_manager.get_config()
                            
                            messagebox.showinfo(
                                "Template Loaded",
                                f"Default template loaded successfully:\n{default_template}\n\n"
                                "Database schema has been updated with template columns.",
                                parent=self.root
                            )
                        else:
                            # Schema update failed but still save the template path
                            self.config_manager.update_config(default_template_path=default_template)
                            self.config_manager.save_config()
                            self.config = self.config_manager.get_config()
                            
                            messagebox.showwarning(
                                "Template Loaded with Warning",
                                f"Default template path set to:\n{default_template}\n\n"
                                "However, database schema update encountered issues.\n"
                                "You may need to check Settings.",
                                parent=self.root
                            )
                    else:
                        # No database path configured
                        self.config_manager.update_config(default_template_path=default_template)
                        self.config_manager.save_config()
                        self.config = self.config_manager.get_config()
                        
                        messagebox.showinfo(
                            "Template Loaded",
                            f"Default template loaded successfully:\n{default_template}",
                            parent=self.root
                        )
                        
                except Exception as e:
                    print(f"Error loading default template: {e}")
                    messagebox.showerror(
                        "Template Load Error",
                        f"Failed to load default template:\n{str(e)}\n\n"
                        "Opening Settings to configure manually.",
                        parent=self.root
                    )
                    self.open_settings()
            else:
                # Default template doesn't exist
                messagebox.showerror(
                    "Template Not Found",
                    f"Default template file not found:\n{default_template}\n\n"
                    "Opening Settings to select a template file.",
                    parent=self.root
                )
                self.open_settings()
        else:
            # User wants to go to settings
            self.open_settings()

    def button_notimplemented(self):
        messagebox.showinfo("Not Implemented", "Feature Not Implemented Yet.")
        print("Feature Not Implemented Yet.")

    @performance_monitor("Open Add New Assets")
    def open_add_new_assets(self):
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        # Pass current config so window uses latest settings
        AddNewAssetsWindow(self.root, self.config)

    @performance_monitor("Open Browse Assets")
    def open_browse_assets(self):
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        # Import here to avoid circular imports
        from browse_assets import BrowseAssetsWindow
        BrowseAssetsWindow(self.root, self.config)

    @performance_monitor("Open Bulk Update Assets")
    def open_bulk_update_assets(self):
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        # Import here to avoid circular imports
        from bulk_update_assets import BulkUpdateWindow
        BulkUpdateWindow(self.root, self.config)

    @performance_monitor("Open Monitor Window")
    def open_monitor(self):
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        # Monitor window can run alongside other windows
        MonitorWindow(self.root)

    @performance_monitor("Open Reports and Analysis")
    def open_reports_analysis(self):
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        # Reports window can run alongside other windows
        open_reports_analysis_window(self.root)

    def open_settings(self):
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        SettingsWindow(self.root, self)  # Pass self (MainMenu instance)

    @performance_monitor("Export Assets via Template")
    def export_assets_via_template(self):
        """Export assets using template formatting - uses centralized export service."""
        if not MODULES_LOADED:
            messagebox.showerror("Error", "Modules not loaded. Please restart application.", parent=self.root)
            return
        export_service.export_database_template(self.root)

    def change_theme(self, theme):
        """Change application theme and save to config."""
        ctk.set_appearance_mode(theme.lower())
        # Update config through manager
        self.config_manager.update_config(theme=theme.lower())
        # Refresh our local config reference
        self.config = self.config_manager.get_config()

if __name__ == "__main__":
    # Set customtkinter font directory
    font_dir = os.path.join(os.path.dirname(__file__), "assets", "fonts")
    ctk.CTkFont.fallback_font_paths = [font_dir]

    root = ctk.CTk()
    MainMenu(root)
    root.mainloop()