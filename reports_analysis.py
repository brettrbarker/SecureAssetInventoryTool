"""
Reports and Analysis window for asset management system.
Provides various analytical reports and data insights.
"""

import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import os
import re
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import seaborn as sns

# Configure matplotlib for tkinter integration
plt.switch_backend('Agg')  # Use non-interactive backend

from asset_database import AssetDatabase
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from performance_monitoring import performance_monitor
from ui_components import SearchableDropdown, AssetDetailWindow, MultiAssetViewer, DatePicker
from field_utils import compute_db_fields_from_template, compute_dropdown_fields


class ReportsAnalysisWindow:
    """Reports and Analysis window for generating various asset reports."""
    
    def __init__(self, parent=None):
        self.parent = parent
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = self.config_manager.get_config()
        
        # Create database instance
        self.db = AssetDatabase(self.config.database_path)
        
        # Get database fields for dropdowns
        self.db_fields = compute_db_fields_from_template(self.db, self.config)
        self.dropdown_fields = compute_dropdown_fields(self.db_fields, self.config)
        
        # Create window
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Reports and Analysis")
        self.window.geometry("1000x700")
        self.window.minsize(900, 600)
        self.window.transient(parent)
        
        # Center the window
        self._center_window()
        
        # Build UI
        self._create_widgets()
        
        # Set up window close handler
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Focus the window
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
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """Create all window widgets."""
        # Main container
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ctk.CTkLabel(main_frame, 
                                  text="Reports and Analysis", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(10, 20))
        
        # Create notebook for different report types
        self.notebook = ctk.CTkTabview(main_frame)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Add tabs for different report types
        self._create_overview_tab()
        self._create_audit_report_tab()
        self._create_labels_report_tab()
        self._create_duplicate_report_tab()
        self._create_cubicle_analysis_tab()
        self._create_hmr_tab()
        self._create_destruction_cert_tab()
    
    def _create_overview_tab(self):
        """Create overview dashboard tab with charts and statistics."""
        tab = self.notebook.add("Overview")
        
        # Create main scroll frame for the dashboard
        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Title
        title_label = ctk.CTkLabel(scroll_frame, 
                                  text="Asset Database Overview", 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(10, 20))
        
        # Refresh button
        refresh_btn = ctk.CTkButton(scroll_frame, text="🔄 Refresh Dashboard", 
                                   command=self._refresh_overview,
                                   width=150, height=35)
        refresh_btn.pack(pady=(0, 20))
        
        # Statistics summary frame
        self.stats_frame = ctk.CTkFrame(scroll_frame)
        self.stats_frame.pack(fill="x", padx=10, pady=(0, 20))
        
        # Charts container frame
        self.charts_frame = ctk.CTkFrame(scroll_frame)
        self.charts_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Load initial dashboard
        self._refresh_overview()
    
    def _create_audit_report_tab(self):
        """Create tab for audit reporting (items not audited in X days)."""
        tab = self.notebook.add("Audit Report")
        
        # Description
        desc_label = ctk.CTkLabel(tab, 
                                 text="Find assets that have not been audited within the specified timeframe",
                                 font=ctk.CTkFont(size=14))
        desc_label.pack(pady=(10, 20))
        
        # Parameters frame
        params_frame = ctk.CTkFrame(tab)
        params_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Days parameter
        days_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
        days_frame.pack(pady=20)
        
        ctk.CTkLabel(days_frame, text="Days since last audit:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        
        self.audit_days_var = ctk.StringVar(value="365")
        self.audit_days_entry = ctk.CTkEntry(days_frame, textvariable=self.audit_days_var, width=100)
        self.audit_days_entry.pack(side="left", padx=(0, 10))
        
        # Run button
        audit_run_btn = ctk.CTkButton(days_frame, text="Generate Report", 
                                     command=self._generate_audit_report,
                                     width=150, height=35)
        audit_run_btn.pack(side="left", padx=20)
        
        # Export button (initially disabled)
        self.audit_export_btn = ctk.CTkButton(days_frame, text="Export to Excel", 
                                            command=self._export_audit_report,
                                            width=150, height=35,
                                            state="disabled",
                                            fg_color="gray")
        self.audit_export_btn.pack(side="left", padx=10)
        
        # Store data for export
        self.audit_data = None
        
        # Instructions
        instructions_label = ctk.CTkLabel(tab, 
                                        text="💡 Tip: Double-click on any asset in the results table to view details",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(pady=(0, 10))
        
        # Results area
        self.audit_results_frame = ctk.CTkFrame(tab)
        self.audit_results_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
    
    def _create_labels_report_tab(self):
        """Create tab for label reporting (label requests by date)."""
        tab = self.notebook.add("Labels")
        
        # Description
        desc_label = ctk.CTkLabel(tab, 
                                 text="Find assets with label requests based on date criteria",
                                 font=ctk.CTkFont(size=14))
        desc_label.pack(pady=(10, 20))
        
        # Parameters frame
        params_frame = ctk.CTkFrame(tab)
        params_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Criteria and date selection frame
        criteria_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
        criteria_frame.pack(pady=20)
        
        # Criteria dropdown
        ctk.CTkLabel(criteria_frame, text="Search criteria:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        
        self.labels_criteria_var = ctk.StringVar(value="On")
        criteria_options = ["On", "On or After", "On or Before", "All"]
        self.labels_criteria_dropdown = ctk.CTkComboBox(criteria_frame, 
                                                       variable=self.labels_criteria_var,
                                                       values=criteria_options,
                                                       width=150,
                                                       command=self._on_labels_criteria_change)
        self.labels_criteria_dropdown.pack(side="left", padx=(0, 20))
        
        # Date picker
        self.labels_date_label = ctk.CTkLabel(criteria_frame, text="Date:", font=ctk.CTkFont(size=14))
        self.labels_date_label.pack(side="left", padx=(0, 10))
        
        self.labels_date_var = ctk.StringVar(value=datetime.now().strftime("%m/%d/%Y"))
        self.labels_date_picker = DatePicker(criteria_frame, variable=self.labels_date_var, width=150)
        self.labels_date_picker.pack(side="left", padx=(0, 20))
        
        # Run button
        self.labels_run_btn = ctk.CTkButton(criteria_frame, text="Generate Report", 
                                           command=self._generate_labels_report,
                                           width=150, height=35)
        self.labels_run_btn.pack(side="left", padx=20)
        
        # Export button (initially disabled)
        self.labels_export_btn = ctk.CTkButton(criteria_frame, text="Export to Excel", 
                                              command=self._export_labels_report,
                                              width=150, height=35,
                                              state="disabled",
                                              fg_color="gray")
        self.labels_export_btn.pack(side="left", padx=10)
        
        # Store data for export
        self.labels_data = None
        
        # Set initial state for date picker based on default criteria
        self._on_labels_criteria_change(self.labels_criteria_var.get())
        
        # Instructions
        instructions_label = ctk.CTkLabel(tab, 
                                        text="💡 Tip: Use 'All' criteria to see all assets with label requests regardless of date",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(pady=(0, 10))
        
        # Separator line
        separator = ctk.CTkFrame(tab, height=2, fg_color="gray")
        separator.pack(fill="x", padx=20, pady=10)
        
        # Barcode Generation Section - Condensed
        barcode_section_label = ctk.CTkLabel(tab, 
                                           text="Generate Barcode Labels", 
                                           font=ctk.CTkFont(size=14, weight="bold"))
        barcode_section_label.pack(pady=(5, 0))
        
        # Barcode configuration frame - more compact
        barcode_config_frame = ctk.CTkFrame(tab)
        barcode_config_frame.pack(fill="x", padx=20, pady=(5, 10))
        
        # Single row for all field selections
        field_selection_frame = ctk.CTkFrame(barcode_config_frame, fg_color="transparent")
        field_selection_frame.pack(pady=10, fill="x")
        
        # Get label output fields from config for dropdown options
        label_fields = self.config.get("label_output_fields", [])
        
        # Primary field (barcode value) - compact
        ctk.CTkLabel(field_selection_frame, text="Primary:", 
                    font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 5))
        
        self.barcode_primary_var = ctk.StringVar(value="Serial Number" if "Serial Number" in label_fields else (label_fields[0] if label_fields else "Asset No."))
        self.barcode_primary_dropdown = ctk.CTkComboBox(field_selection_frame,
                                                       variable=self.barcode_primary_var,
                                                       values=label_fields,
                                                       width=140)
        self.barcode_primary_dropdown.pack(side="left", padx=(0, 15))
        
        # Secondary field 1 - compact
        ctk.CTkLabel(field_selection_frame, text="Secondary 1:", 
                    font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 5))
        
        self.barcode_secondary1_var = ctk.StringVar(value="*Manufacturer" if "*Manufacturer" in label_fields else (label_fields[1] if len(label_fields) > 1 else "*Asset Type"))
        self.barcode_secondary1_dropdown = ctk.CTkComboBox(field_selection_frame,
                                                          variable=self.barcode_secondary1_var,
                                                          values=label_fields,
                                                          width=140)
        self.barcode_secondary1_dropdown.pack(side="left", padx=(0, 15))
        
        # Secondary field 2 - compact
        ctk.CTkLabel(field_selection_frame, text="Secondary 2:", 
                    font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 5))
        
        self.barcode_secondary2_var = ctk.StringVar(value="*Model" if "*Model" in label_fields else (label_fields[2] if len(label_fields) > 2 else "*Manufacturer"))
        self.barcode_secondary2_dropdown = ctk.CTkComboBox(field_selection_frame,
                                                          variable=self.barcode_secondary2_var,
                                                          values=label_fields,
                                                          width=140)
        self.barcode_secondary2_dropdown.pack(side="left", padx=(0, 15))
        
        # Generate barcode button - smaller and inline
        self.generate_barcode_btn = ctk.CTkButton(field_selection_frame, 
                                                 text="Generate Barcode Labels", 
                                                 command=self._generate_barcode_labels,
                                                 width=150, height=30,
                                                 state="disabled",
                                                 fg_color="gray")
        self.generate_barcode_btn.pack(side="left", padx=(10, 0))
        
        # Results area
        self.labels_results_frame = ctk.CTkFrame(tab)
        self.labels_results_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
    
    def _on_labels_criteria_change(self, value):
        """Handle change in labels criteria dropdown to enable/disable date picker."""
        if value == "All":
            # Disable the date picker for 'All' criteria since date is not needed
            self.labels_date_picker.display_entry.configure(state="disabled")
        else:
            # Enable the date picker for other criteria
            self.labels_date_picker.display_entry.configure(state="normal")

    def _create_duplicate_report_tab(self):
        """Create tab for duplicate detection reporting."""
        tab = self.notebook.add("Duplicate Detection")
        
        # Description
        desc_label = ctk.CTkLabel(tab, 
                                 text="Find duplicate values in selected fields",
                                 font=ctk.CTkFont(size=14))
        desc_label.pack(pady=(10, 20))
        
        # Parameters frame
        params_frame = ctk.CTkFrame(tab)
        params_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Field selection
        field_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
        field_frame.pack(pady=20)
        
        ctk.CTkLabel(field_frame, text="Field to check for duplicates:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        
        # Get field names for dropdown
        field_names = [field['display_name'] for field in self.db_fields]
        self.duplicate_field_var = ctk.StringVar()
        self.duplicate_field_dropdown = SearchableDropdown(field_frame,
                                                          values=field_names,
                                                          variable=self.duplicate_field_var,
                                                          width=250)
        self.duplicate_field_dropdown.pack(side="left", padx=(0, 10))
        
        # Run button
        duplicate_run_btn = ctk.CTkButton(field_frame, text="Generate Report", 
                                        command=self._generate_duplicate_report,
                                        width=150, height=35)
        duplicate_run_btn.pack(side="left", padx=20)
        
        # Export button (initially disabled)
        self.duplicate_export_btn = ctk.CTkButton(field_frame, text="Export to Excel", 
                                                command=self._export_duplicate_report,
                                                width=150, height=35,
                                                state="disabled",
                                                fg_color="gray")
        self.duplicate_export_btn.pack(side="left", padx=10)
        
        # Duplicate value selection frame (initially hidden)
        self.duplicate_value_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
        # Don't pack yet - will be shown after report generation
        
        ctk.CTkLabel(self.duplicate_value_frame, text="Select duplicate value to view assets:", 
                    font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        
        # Dropdown for specific duplicate values (populated after report generation)
        self.duplicate_value_var = ctk.StringVar()
        self.duplicate_value_dropdown = SearchableDropdown(self.duplicate_value_frame,
                                                          values=[],
                                                          variable=self.duplicate_value_var,
                                                          width=250)
        self.duplicate_value_dropdown.pack(side="left", padx=(0, 10))
        
        # View Assets button
        self.view_assets_btn = ctk.CTkButton(self.duplicate_value_frame, text="View Assets", 
                                           command=self._view_duplicate_assets,
                                           width=120, height=35,
                                           fg_color="#1f538d", hover_color="#14375e")
        self.view_assets_btn.pack(side="left", padx=10)
        
        # Store duplicate results for asset viewing
        self.duplicate_results = None
        
        # Store data for export
        self.duplicate_data = None
        
        # Instructions
        instructions_label = ctk.CTkLabel(tab, 
                                        text="💡 Tip: Double-click any asset in results to view details • Use dropdown to view all assets with a specific duplicate value",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(pady=(0, 10))
        
        # Results area
        self.duplicate_results_frame = ctk.CTkFrame(tab)
        self.duplicate_results_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
    
    def _create_cubicle_analysis_tab(self):
        """Create tab for cubicle analysis (anomaly detection)."""
        tab = self.notebook.add("Cubicle Analysis")
        
        # Description
        desc_label = ctk.CTkLabel(tab, 
                                 text="Find cubicles that don't match specific asset type requirements",
                                 font=ctk.CTkFont(size=14))
        desc_label.pack(pady=(10, 20))
        
        # Parameters frame
        params_frame = ctk.CTkFrame(tab)
        params_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Configuration frame
        config_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
        config_frame.pack(pady=20, fill="x")
        
        # Asset type selection
        asset_type_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        asset_type_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(asset_type_frame, text="Asset Type:", 
                    font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        
        # Get unique asset types from database
        try:
            with self.db.get_connection() as conn:
                cursor = conn.execute("SELECT DISTINCT asset_type FROM assets WHERE asset_type IS NOT NULL AND asset_type != '' ORDER BY asset_type")
                asset_types = [row[0] for row in cursor.fetchall()]
        except Exception:
            asset_types = ["Monitor", "Computer", "Keyboard", "Mouse"]  # Default fallback
        
        self.asset_type_var = ctk.StringVar()
        self.asset_type_dropdown = ctk.CTkComboBox(asset_type_frame, 
                                                  variable=self.asset_type_var,
                                                  values=asset_types,
                                                  width=200)
        self.asset_type_dropdown.pack(side="left", padx=(0, 20))
        if asset_types:
            self.asset_type_dropdown.set(asset_types[0])
        
        # Quantity comparison frame
        quantity_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        quantity_frame.pack(pady=10, fill="x")
        
        ctk.CTkLabel(quantity_frame, text="Find cubicles with:", 
                    font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        
        # Comparison operator
        self.comparison_var = ctk.StringVar(value="exactly")
        self.comparison_dropdown = ctk.CTkComboBox(quantity_frame,
                                                  variable=self.comparison_var,
                                                  values=["exactly", "less than", "greater than", "not equal to"],
                                                  width=120)
        self.comparison_dropdown.pack(side="left", padx=(0, 10))
        
        # Quantity input
        self.quantity_var = ctk.StringVar(value="2")
        self.quantity_entry = ctk.CTkEntry(quantity_frame, textvariable=self.quantity_var, width=80)
        self.quantity_entry.pack(side="left", padx=(0, 10))
        
        ctk.CTkLabel(quantity_frame, text="items", 
                    font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 20))
        
        # Run button
        cubicle_run_btn = ctk.CTkButton(quantity_frame, text="Generate Analysis", 
                                       command=self._generate_cubicle_analysis,
                                       width=150, height=35)
        cubicle_run_btn.pack(side="left", padx=20)
        
        # Export button (initially disabled)
        self.cubicle_export_btn = ctk.CTkButton(quantity_frame, text="Export to Excel", 
                                              command=self._export_cubicle_analysis,
                                              width=150, height=35,
                                              state="disabled",
                                              fg_color="gray")
        self.cubicle_export_btn.pack(side="left", padx=10)
        
        # Store data for export
        self.cubicle_data = None
        
        # Instructions
        instructions_label = ctk.CTkLabel(tab, 
                                        text="💡 Tip: Double-click any row to view asset details (multi-asset viewer for rows with multiple assets)",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(pady=(0, 10))
        
        # Results area
        self.cubicle_results_frame = ctk.CTkFrame(tab)
        self.cubicle_results_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
    
    def _create_hmr_tab(self):
        """Create tab for Hardware Movement Request (HMR) form generation."""
        tab = self.notebook.add("HMR")
        
        # Create a centered container
        container = ctk.CTkFrame(tab, fg_color="transparent")
        container.pack(expand=True, fill="both")
        
        # Icon and message
        icon_label = ctk.CTkLabel(container, 
                                 text="🚧",
                                 font=ctk.CTkFont(size=72))
        icon_label.pack(pady=(50, 20))
        
        title_label = ctk.CTkLabel(container, 
                                  text="Hardware Movement Request", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(0, 20))
        
        message_label = ctk.CTkLabel(container, 
                                    text="Feature not implemented yet",
                                    font=ctk.CTkFont(size=16),
                                    text_color="gray")
        message_label.pack(pady=(0, 10))
        
        description_label = ctk.CTkLabel(container, 
                                        text="This feature will allow you to:\n"
                                             "• Generate HMR forms from asset data\n"
                                             "• Fill PDF templates with asset information\n"
                                             "• Track hardware movement requests",
                                        font=ctk.CTkFont(size=14),
                                        text_color="gray",
                                        justify="left")
        description_label.pack(pady=(0, 30))
    
    def _create_destruction_cert_tab(self):
        """Create tab for Destruction Certification form generation."""
        tab = self.notebook.add("Destruction Cert")
        
        # Create a centered container
        container = ctk.CTkFrame(tab, fg_color="transparent")
        container.pack(expand=True, fill="both")
        
        # Icon and message
        icon_label = ctk.CTkLabel(container, 
                                 text="🚧",
                                 font=ctk.CTkFont(size=72))
        icon_label.pack(pady=(50, 20))
        
        title_label = ctk.CTkLabel(container, 
                                  text="Destruction Certification", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(0, 20))
        
        message_label = ctk.CTkLabel(container, 
                                    text="Feature not implemented yet",
                                    font=ctk.CTkFont(size=16),
                                    text_color="gray")
        message_label.pack(pady=(0, 10))
        
        description_label = ctk.CTkLabel(container, 
                                        text="This feature will allow you to:\n"
                                             "• Generate destruction certification forms\n"
                                             "• Document asset disposal and destruction\n"
                                             "• Maintain destruction records",
                                        font=ctk.CTkFont(size=14),
                                        text_color="gray",
                                        justify="left")
        description_label.pack(pady=(0, 30))
    
    @performance_monitor("Generate Audit Report")
    def _generate_audit_report(self):
        """Generate report of items not audited in X days."""
        try:
            # Get parameters
            days = int(self.audit_days_var.get())
            
            # Calculate cutoff date in ISO format for proper date comparison
            cutoff_date = datetime.now() - timedelta(days=days)
            cutoff_str = cutoff_date.strftime("%Y-%m-%d")
            
            # Clear previous results
            for widget in self.audit_results_frame.winfo_children():
                widget.destroy()
            
            # Query database for assets not audited since cutoff date
            # Need to handle multiple date formats properly by converting to ISO format for comparison
            query = """
                SELECT asset_no, manufacturer, model, serial_number, location, room, cubicle, 
                       audit_date, status, asset_type
                FROM assets 
                WHERE is_deleted = 0 
                AND (
                    audit_date IS NULL 
                    OR audit_date = ''
                    OR (
                        CASE 
                            WHEN audit_date LIKE '%/%/____' THEN 
                                date(
                                    substr(audit_date, length(audit_date) - 3, 4) || '-' || 
                                    substr('0' || substr(audit_date, 1, instr(audit_date, '/') - 1), -2) || '-' || 
                                    substr('0' || substr(
                                        substr(audit_date, instr(audit_date, '/') + 1), 
                                        1, 
                                        instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1
                                    ), -2)
                                )
                            WHEN audit_date LIKE '____-__-__' THEN 
                                date(audit_date)
                            ELSE 
                                '1900-01-01'
                        END
                    ) < date(?)
                )
                ORDER BY location, room, cubicle, manufacturer, model
            """
            
            with self.db.get_connection() as conn:
                cursor = conn.execute(query, (cutoff_str,))
                results = cursor.fetchall()
            
            # Convert to pandas DataFrame for better processing
            columns = ["Asset No", "Manufacturer", "Model", "Serial Number", 
                      "Location", "Room", "Cubicle", "Last Audit", "Status", "Asset Type"]
            
            if results:
                self.audit_data = pd.DataFrame(results, columns=columns)
                
                # Add calculated fields
                self.audit_data['Days Since Audit'] = self.audit_data['Last Audit'].apply(
                    lambda x: self._calculate_days_since_audit(x, cutoff_date)
                )
                
                # Enable export button
                self.audit_export_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            else:
                self.audit_data = None
                self.audit_export_btn.configure(state="disabled", fg_color="gray")
            
            # Display results
            if not results:
                no_results_label = ctk.CTkLabel(self.audit_results_frame, 
                                              text=f"✅ All assets have been audited within the last {days} days!",
                                              font=ctk.CTkFont(size=16, weight="bold"),
                                              text_color="green")
                no_results_label.pack(pady=20)
                return
            
            # Results header with statistics
            stats_text = f"Assets not audited in the last {days} days: {len(results)} found\n"
            if self.audit_data is not None:
                # Add some statistics
                by_location = self.audit_data.groupby('Location').size()
                stats_text += f"Locations affected: {len(by_location)}\n"
                stats_text += f"Most affected location: {by_location.idxmax()} ({by_location.max()} assets)"
            
            header_label = ctk.CTkLabel(self.audit_results_frame, 
                                      text=stats_text,
                                      font=ctk.CTkFont(size=14, weight="bold"))
            header_label.pack(pady=(10, 20))
            
            # Create results table
            self._create_results_table(self.audit_results_frame, results, columns)
            
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number of days.")
        except Exception as e:
            error_handler.handle_exception(e, context="generating audit report",
                                     parent_window=self.window)
    
    def _calculate_days_since_audit(self, audit_date_str, cutoff_date):
        """Calculate days since last audit."""
        if not audit_date_str or audit_date_str.strip() == '':
            return "Never"
        try:
            audit_date = datetime.strptime(audit_date_str, "%m/%d/%Y")
            days_diff = (datetime.now() - audit_date).days
            return str(days_diff)
        except:
            return "Invalid Date"
    
    @performance_monitor("Generate Labels Report")
    def _generate_labels_report(self):
        """Generate report of label requests based on date criteria."""
        try:
            # Get parameters
            criteria = self.labels_criteria_var.get()
            
            # Clear previous results
            for widget in self.labels_results_frame.winfo_children():
                widget.destroy()
            
            # Get label output fields from config
            label_fields = self.config.get("label_output_fields", [])
            
            # Get column mapping to convert display names to database column names
            column_mapping = self.db.get_dynamic_column_mapping(self.config.default_template_path)
            
            # Build list of database columns to select
            db_columns = []
            display_columns = []
            
            # Add configured fields
            for field in label_fields:
                # Get the database column name for this display field
                db_column = column_mapping.get(field)
                if db_column:
                    db_columns.append(db_column)
                    display_columns.append(field)
                else:
                    # Try to generate safe column name if not in mapping
                    safe_column = self.db._generate_safe_column_name(field)
                    db_columns.append(safe_column)
                    display_columns.append(field)
            
            # Always add label_requested_date (hardcoded requirement)
            db_columns.append("label_requested_date")
            display_columns.append("Label Request Date")
            
            # Build SELECT clause
            select_clause = ", ".join(db_columns)
            
            # Build query based on criteria
            if criteria == "All":
                # Show all assets with label requests regardless of date
                query = f"""
                    SELECT {select_clause}
                    FROM assets 
                    WHERE is_deleted = 0 
                    AND label_requested_date IS NOT NULL 
                    AND label_requested_date != ''
                    AND label_requested_date != '1901-01-01 00:00:00'
                    ORDER BY label_requested_date DESC
                """
                params = ()
            else:
                # Get date parameter and parse it
                date_str = self.labels_date_var.get()
                try:
                    selected_date = datetime.strptime(date_str, "%m/%d/%Y")
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid date in MM/DD/YYYY format.")
                    return
                
                # Convert to date only (ignore time)
                target_date = selected_date.strftime("%Y-%m-%d")
                
                if criteria == "On":
                    # Exact date match (ignore time component)
                    query = f"""
                        SELECT {select_clause}
                        FROM assets 
                        WHERE is_deleted = 0 
                        AND label_requested_date IS NOT NULL 
                        AND label_requested_date != ''
                        AND label_requested_date != '1901-01-01 00:00:00'
                        AND date(label_requested_date) = date(?)
                        ORDER BY label_requested_date DESC
                    """
                    params = (target_date,)
                elif criteria == "On or After":
                    # On or after the selected date
                    query = f"""
                        SELECT {select_clause}
                        FROM assets 
                        WHERE is_deleted = 0 
                        AND label_requested_date IS NOT NULL 
                        AND label_requested_date != ''
                        AND label_requested_date != '1901-01-01 00:00:00'
                        AND date(label_requested_date) >= date(?)
                        ORDER BY label_requested_date DESC
                    """
                    params = (target_date,)
                elif criteria == "On or Before":
                    # On or before the selected date
                    query = f"""
                        SELECT {select_clause}
                        FROM assets 
                        WHERE is_deleted = 0 
                        AND label_requested_date IS NOT NULL 
                        AND label_requested_date != ''
                        AND label_requested_date != '1901-01-01 00:00:00'
                        AND date(label_requested_date) <= date(?)
                        ORDER BY label_requested_date DESC
                    """
                    params = (target_date,)
            
            # Execute query
            with self.db.get_connection() as conn:
                cursor = conn.execute(query, params)
                results = cursor.fetchall()
            
            if results:
                # Create DataFrame with dynamic columns based on config
                self.labels_data = pd.DataFrame(results, columns=display_columns)
                
                # Format the label requested date for display
                if "Label Request Date" in self.labels_data.columns:
                    self.labels_data['Label Request Date'] = self.labels_data['Label Request Date'].apply(
                        lambda x: self._format_label_date_for_display(x)
                    )
                
                # Enable export button
                self.labels_export_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
                
                # Enable barcode generation button
                self.generate_barcode_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            else:
                self.labels_data = None
                self.labels_export_btn.configure(state="disabled", fg_color="gray")
                
                # Disable barcode generation button
                self.generate_barcode_btn.configure(state="disabled", fg_color="gray")
            
            # Display results
            if not results:
                if criteria == "All":
                    no_results_text = "No assets have label requests in the database."
                else:
                    no_results_text = f"No assets have label requests {criteria.lower()} {date_str}."
                
                no_results_label = ctk.CTkLabel(self.labels_results_frame, 
                                              text=f"ℹ️ {no_results_text}",
                                              font=ctk.CTkFont(size=16, weight="bold"),
                                              text_color="orange")
                no_results_label.pack(pady=20)
                return
            
            # Results header with statistics
            if criteria == "All":
                stats_text = f"Total assets with label requests: {len(results)} found"
            else:
                stats_text = f"Assets with label requests {criteria.lower()} {date_str}: {len(results)} found"
            
            # Add field info to stats
            stats_text += f"\nShowing {len(display_columns)} fields: {', '.join(display_columns)}"
            
            header_label = ctk.CTkLabel(self.labels_results_frame, 
                                      text=stats_text,
                                      font=ctk.CTkFont(size=14, weight="bold"))
            header_label.pack(pady=(10, 20))
            
            # Create results table with dynamic columns
            self._create_results_table(self.labels_results_frame, results, display_columns)
            
        except Exception as e:
            error_handler.handle_exception(e, context="generating labels report",
                                     parent_window=self.window)
    
    def _format_label_date_for_display(self, date_str):
        """Format label requested date for display in results."""
        if not date_str or date_str.strip() == '' or date_str == '1901-01-01 00:00:00':
            return "No Date"
        try:
            # Try to parse as ISO datetime format first
            date_obj = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            return date_obj.strftime("%m/%d/%Y %I:%M %p")
        except:
            try:
                # Try to parse as other common formats
                date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                return date_obj.strftime("%m/%d/%Y %I:%M %p")
            except:
                return str(date_str)  # Return as-is if can't parse

    @performance_monitor("Generate Barcode Labels")
    def _generate_barcode_labels(self):
        """Generate barcode labels from current label report results."""
        try:
            # Check if we have label data
            if self.labels_data is None or self.labels_data.empty:
                messagebox.showwarning("No Data", "No label report data available. Please generate a label report first.")
                return
            
            # Get selected fields
            primary_field = self.barcode_primary_var.get()
            secondary1_field = self.barcode_secondary1_var.get()
            secondary2_field = self.barcode_secondary2_var.get()
            
            if not all([primary_field, secondary1_field, secondary2_field]):
                messagebox.showwarning("Missing Fields", "Please select all three fields for barcode generation.")
                return
            
            # Check if selected fields exist in the data
            missing_fields = []
            for field in [primary_field, secondary1_field, secondary2_field]:
                if field not in self.labels_data.columns:
                    missing_fields.append(field)
            
            if missing_fields:
                messagebox.showerror("Field Error", 
                                   f"Selected fields not found in report data: {', '.join(missing_fields)}\n"
                                   f"Available fields: {', '.join(self.labels_data.columns)}")
                return
            
            # Prepare barcode data
            barcode_data = []
            for _, row in self.labels_data.iterrows():
                primary_value = str(row[primary_field]) if pd.notna(row[primary_field]) else ""
                secondary1_value = str(row[secondary1_field]) if pd.notna(row[secondary1_field]) else ""
                secondary2_value = str(row[secondary2_field]) if pd.notna(row[secondary2_field]) else ""
                
                # Skip empty primary values
                if primary_value.strip():
                    barcode_data.append((primary_value, secondary1_value, secondary2_value))
            
            if not barcode_data:
                messagebox.showwarning("No Data", "No valid barcode data found. Primary field values cannot be empty.")
                return
            
            # Ask user for save location using configured output directory
            output_dir = self.config.output_directory
            labels_dir = os.path.join(output_dir, "labels")
            os.makedirs(labels_dir, exist_ok=True)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
                title="Save Barcode Labels",
                initialfile=f"Barcode_Labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                initialdir=labels_dir
            )
            
            if filename:
                # Import and use the barcode generator
                try:
                    from generate_barcodes_pdf import BarcodeGenerator
                    
                    generator = BarcodeGenerator()
                    generator.generate_pdf(barcode_data, filename)
                    
                    messagebox.showinfo("Success", 
                                      f"Barcode labels generated successfully!\n"
                                      f"File: {filename}\n"
                                      f"Labels: {len(barcode_data)}")
                    
                except ImportError:
                    messagebox.showerror("Import Error", 
                                       "Could not import barcode generator. Please ensure generate_barcodes_pdf.py is available.")
                except Exception as e:
                    error_handler.handle_exception(e, context="generating barcode labels",
                                             parent_window=self.window)
                    
        except Exception as e:
            error_handler.handle_exception(e, context="preparing barcode label generation",
                                     parent_window=self.window)

    @performance_monitor("Generate Duplicate Report")
    def _generate_duplicate_report(self):
        """Generate report of duplicate values in selected field."""
        try:
            # Get parameters
            field_display_name = self.duplicate_field_var.get()
            
            if not field_display_name:
                messagebox.showwarning("Warning", "Please select a field to check for duplicates.")
                return
            
            # Find database field name
            db_field_name = None
            for field in self.db_fields:
                if field['display_name'] == field_display_name:
                    db_field_name = field['db_name']
                    break
            
            if not db_field_name:
                messagebox.showerror("Error", "Invalid field selected.")
                return
            
            # Clear previous results
            for widget in self.duplicate_results_frame.winfo_children():
                widget.destroy()
            
            # Query database for duplicates
            query = f"""
                SELECT {db_field_name}, COUNT(*) as count,
                       GROUP_CONCAT(asset_no) as asset_numbers,
                       GROUP_CONCAT(manufacturer || ' ' || model) as assets
                FROM assets 
                WHERE is_deleted = 0 
                AND {db_field_name} IS NOT NULL 
                AND {db_field_name} != ''
                GROUP BY {db_field_name}
                HAVING COUNT(*) > 1
                ORDER BY count DESC, {db_field_name}
            """
            
            with self.db.get_connection() as conn:
                cursor = conn.execute(query)
                results = cursor.fetchall()
            
            # Convert to pandas DataFrame for export
            if results:
                # Create detailed duplicate data for export
                export_data = []
                for row in results:
                    value, count, asset_nos, assets = row
                    asset_list = assets.split(',')
                    asset_no_list = asset_nos.split(',')
                    
                    for asset_no, asset_desc in zip(asset_no_list, asset_list):
                        export_data.append({
                            'Field': field_display_name,
                            'Duplicate Value': value,
                            'Occurrences': count,
                            'Asset No': asset_no.strip(),
                            'Asset Description': asset_desc.strip()
                        })
                
                self.duplicate_data = pd.DataFrame(export_data)
                
                # Enable export button
                self.duplicate_export_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            else:
                self.duplicate_data = None
                self.duplicate_export_btn.configure(state="disabled", fg_color="gray")
            
            # Display results
            if not results:
                # Hide duplicate value selection frame when no results
                self.duplicate_value_frame.pack_forget()
                self.duplicate_results = None
                
                no_results_label = ctk.CTkLabel(self.duplicate_results_frame, 
                                              text=f"✅ No duplicate values found in {field_display_name}!",
                                              font=ctk.CTkFont(size=16, weight="bold"),
                                              text_color="green")
                no_results_label.pack(pady=20)
                return
            
            # Results header with statistics
            stats_text = f"Duplicate values in {field_display_name}: {len(results)} groups found\n"
            total_duplicates = sum(row[1] for row in results)  # Sum of all counts
            stats_text += f"Total affected assets: {total_duplicates}"
            
            header_label = ctk.CTkLabel(self.duplicate_results_frame, 
                                      text=stats_text,
                                      font=ctk.CTkFont(size=14, weight="bold"))
            header_label.pack(pady=(10, 20))
            
            # Create results table - flatten duplicate data for table display
            table_data = []
            table_headers = ["Field", "Duplicate Value", "Count", "Asset No", "Asset Description"]
            
            for row in results:
                value, count, asset_nos, assets = row
                asset_list = assets.split(',')
                asset_no_list = asset_nos.split(',')
                
                for asset_no, asset_desc in zip(asset_no_list, asset_list):
                    table_data.append([
                        field_display_name,
                        value,
                        count,
                        asset_no.strip(),
                        asset_desc.strip()
                    ])
            
            # Store results for duplicate value viewing and populate dropdown
            self.duplicate_results = results
            duplicate_values = [str(row[0]) for row in results]  # Extract duplicate values
            
            # Update dropdown with duplicate values and show the selection frame
            try:
                # Update the SearchableDropdown values
                self.duplicate_value_dropdown.values_all = [""] + duplicate_values
                if hasattr(self.duplicate_value_dropdown, 'values'):
                    self.duplicate_value_dropdown.values = [""] + duplicate_values
                self.duplicate_value_var.set("")  # Clear selection
                self.duplicate_value_frame.pack(pady=(10, 20))  # Show the frame
            except Exception as dropdown_error:
                print(f"Warning: Could not update duplicate value dropdown: {dropdown_error}")
                # Still show the frame even if dropdown update fails
                self.duplicate_value_frame.pack(pady=(10, 20))
            
            # Create results table
            self._create_results_table(self.duplicate_results_frame, table_data, table_headers)
            
        except Exception as e:
            error_handler.handle_exception(e, context="generating duplicate report",
                                     parent_window=self.window)
    
    def _view_duplicate_assets(self):
        """Open MultiAssetViewer for all assets with the selected duplicate value."""
        try:
            # Get selected duplicate value
            selected_value = self.duplicate_value_var.get().strip()
            
            if not selected_value:
                messagebox.showwarning("Warning", "Please select a duplicate value to view its assets.")
                return
            
            if not self.duplicate_results:
                messagebox.showerror("Error", "No duplicate results available. Please generate a report first.")
                return
            
            # Find the selected duplicate value in the results
            selected_row = None
            for row in self.duplicate_results:
                if str(row[0]) == selected_value:
                    selected_row = row
                    break
            
            if not selected_row:
                messagebox.showerror("Error", "Selected duplicate value not found in results.")
                return
            
            # Extract asset numbers from the selected row
            value, count, asset_nos, assets = selected_row
            asset_no_list = [asset_no.strip() for asset_no in asset_nos.split(',')]
            
            # Look up all assets by asset numbers
            assets_data = []
            for asset_no in asset_no_list:
                with self.db.get_connection() as conn:
                    cursor = conn.execute("SELECT * FROM assets WHERE asset_no = ? AND is_deleted = 0", (asset_no,))
                    result = cursor.fetchone()
                    if result:
                        assets_data.append(dict(result))
            
            if not assets_data:
                messagebox.showwarning("Assets Not Found", "None of the assets with this duplicate value could be found.")
                return
            
            # Open multi-asset viewer window
            MultiAssetViewer(self.window, assets_data)
            
        except Exception as e:
            error_handler.handle_exception(e, context=f"viewing duplicate assets for value '{selected_value}'",
                                     parent_window=self.window)
    
    @performance_monitor("Generate Cubicle Analysis")
    def _generate_cubicle_analysis(self):
        """Generate cubicle analysis based on asset type and quantity criteria."""
        try:
            # Get parameters
            asset_type = self.asset_type_var.get().strip()
            comparison = self.comparison_var.get()
            quantity = int(self.quantity_var.get())
            
            if not asset_type:
                messagebox.showerror("Error", "Please select an asset type.")
                return
            
            # Clear previous results
            for widget in self.cubicle_results_frame.winfo_children():
                widget.destroy()
            
            # Build query based on comparison operator
            if comparison == "exactly":
                condition = "= ?"
                description = f"exactly {quantity}"
            elif comparison == "less than":
                condition = "< ?"
                description = f"less than {quantity}"
            elif comparison == "greater than":
                condition = "> ?"
                description = f"greater than {quantity}"
            elif comparison == "not equal to":
                condition = "!= ?"
                description = f"not equal to {quantity}"
            else:
                condition = "= ?"
                description = f"exactly {quantity}"
            
            # Query to get cubicle counts for the specified asset type
            query = f"""
                WITH cubicle_counts AS (
                    SELECT location, room, cubicle, 
                           COUNT(*) as actual_count,
                           GROUP_CONCAT(asset_no || ': ' || manufacturer || ' ' || model) as assets
                    FROM assets 
                    WHERE is_deleted = 0 
                    AND location IS NOT NULL AND location != ''
                    AND room IS NOT NULL AND room != ''
                    AND cubicle IS NOT NULL AND cubicle != ''
                    AND asset_type = ?
                    GROUP BY location, room, cubicle
                )
                SELECT location, room, cubicle, actual_count, assets
                FROM cubicle_counts
                WHERE actual_count {condition}
                
                UNION
                
                SELECT DISTINCT a.location, a.room, a.cubicle, 0 as actual_count, '' as assets
                FROM assets a
                WHERE a.is_deleted = 0 
                AND a.location IS NOT NULL AND a.location != ''
                AND a.room IS NOT NULL AND a.room != ''
                AND a.cubicle IS NOT NULL AND a.cubicle != ''
                AND NOT EXISTS (
                    SELECT 1 FROM assets a2 
                    WHERE a2.location = a.location 
                    AND a2.room = a.room 
                    AND a2.cubicle = a.cubicle 
                    AND a2.asset_type = ?
                    AND a2.is_deleted = 0
                )
                AND 0 {condition}
                
                ORDER BY location, room, cubicle
            """
            
            with self.db.get_connection() as conn:
                cursor = conn.execute(query, (asset_type, quantity, asset_type, quantity))
                results = cursor.fetchall()
            
            if not results:
                no_results_label = ctk.CTkLabel(self.cubicle_results_frame, 
                                              text=f"No cubicles found with {description} {asset_type}(s).",
                                              font=ctk.CTkFont(size=16))
                no_results_label.pack(pady=20)
                return
            
            # Display header with statistics
            header_text = f"Cubicles with {description} {asset_type}(s): {len(results)} found"
            header_label = ctk.CTkLabel(self.cubicle_results_frame, 
                                      text=header_text,
                                      font=ctk.CTkFont(size=16, weight="bold"))
            header_label.pack(anchor="w", pady=(0, 15))
            
            # Create table data for display
            table_data = []
            table_headers = ["Location", "Room", "Cubicle", "Asset Type", "Expected", "Actual Count", "Asset Numbers"]
            export_data = []
            
            for row in results:
                location, room, cubicle, actual_count, assets = row
                
                # Extract asset numbers for display (only show asset numbers in table for performance)
                asset_numbers = ""
                if assets and assets.strip():
                    asset_items = assets.split(',')
                    asset_nos = []
                    for asset_item in asset_items:
                        if ':' in asset_item:
                            asset_no = asset_item.split(':')[0].strip()
                            asset_nos.append(asset_no)
                    asset_numbers = ", ".join(asset_nos)
                
                # Add to table data
                table_data.append([
                    location,
                    room,
                    cubicle,
                    asset_type,
                    f"{comparison.title()} {quantity}",
                    actual_count,
                    asset_numbers if asset_numbers else "None"
                ])
                
                # Prepare export data (keep full asset details for export)
                export_data.append({
                    'Location': location,
                    'Room': room,
                    'Cubicle': cubicle,
                    'Asset Type': asset_type,
                    'Expected': f"{comparison.title()} {quantity}",
                    'Actual Count': actual_count,
                    'Assets': assets if assets else 'None'
                })
            
            # Create results table
            self._create_results_table(self.cubicle_results_frame, table_data, table_headers)
            
            # Store data for export and enable export button
            if export_data:
                self.cubicle_data = export_data
                self.cubicle_export_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
            
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid quantity number.")
        except Exception as e:
            error_handler.handle_exception(e, context="generating cubicle analysis",
                                     parent_window=self.window)
    
    def _display_anomaly(self, anomaly):
        """Display details of a cubicle anomaly."""
        anomaly_frame = ctk.CTkFrame(self.cubicle_results_frame)
        anomaly_frame.pack(fill="x", padx=20, pady=5)
        
        # Header
        header_text = (f"{anomaly['location']} - {anomaly['room']} - {anomaly['cubicle']} "
                      f"({anomaly['total_items']} total items)")
        header_label = ctk.CTkLabel(anomaly_frame, text=header_text,
                                  font=ctk.CTkFont(size=14, weight="bold"))
        header_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Standard type(s)
        standard_text = f"Standard: {', '.join(anomaly['dominant_types'])} ({anomaly['dominant_count']} items)"
        standard_label = ctk.CTkLabel(anomaly_frame, text=standard_text,
                                    font=ctk.CTkFont(size=12),
                                    text_color="green")
        standard_label.pack(anchor="w", padx=20, pady=2)
        
        # Anomalies
        for asset_type, type_data in anomaly['anomalies'].items():
            anomaly_text = f"Anomaly: {asset_type} ({type_data['count']} items)"
            anomaly_label = ctk.CTkLabel(anomaly_frame, text=anomaly_text,
                                       font=ctk.CTkFont(size=12),
                                       text_color="orange")
            anomaly_label.pack(anchor="w", padx=20, pady=2)
            
            # List specific assets
            for asset in type_data['assets']:
                asset_label = ctk.CTkLabel(anomaly_frame, text=f"    • {asset}",
                                         font=ctk.CTkFont(size=11))
                asset_label.pack(anchor="w", padx=30, pady=1)
    
    def _create_results_table(self, parent, data, headers):
        """Create a results table for displaying report data."""
        # Create frame for table
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create treeview for table display with improved styling
        tree = ttk.Treeview(table_frame, columns=headers, show="headings", height=15)
        
        # Configure improved styling for the treeview
        style = ttk.Style()
        
        # Configure heading style (bold headers, black text on light background)
        style.configure("Treeview.Heading", 
                       font=("Segoe UI", 11, "bold"),
                       background="#e0e0e0",
                       foreground="black",
                       relief="raised",
                       borderwidth=1)
        
        # Configure row style (larger font, centered text)
        style.configure("Treeview", 
                       font=("Segoe UI", 10),
                       rowheight=25,
                       background="white",
                       foreground="black",
                       fieldbackground="white")
        
        # Configure selected row style
        style.map("Treeview",
                 background=[('selected', '#0078d4')],
                 foreground=[('selected', 'white')])
        
        # Configure heading hover effect
        style.map("Treeview.Heading",
                 background=[('active', '#d0d0d0')])
        
        # Store original data for sorting
        self._table_data = data
        self._table_headers = headers
        self._sort_direction = {}  # Track sort direction for each column
        
        # Configure columns with improved formatting
        for i, header in enumerate(headers):
            tree.heading(header, text=header, command=lambda h=header: self._sort_table(tree, h))
            tree.column(header, width=120, minwidth=80, anchor="center")
        
        # Add data and store asset numbers for double-click functionality
        asset_data = {}  # Map item_id to asset_no
        
        # Find the Asset No column index
        asset_no_col_index = None
        for i, header in enumerate(headers):
            if "Asset No" in header:
                asset_no_col_index = i
                break
        
        # Find Asset Numbers column for cubicle analysis
        asset_numbers_col_index = None
        for i, header in enumerate(headers):
            if "Asset Numbers" in header:
                asset_numbers_col_index = i
                break
        
        for row in data:
            # Handle None values and convert to strings
            display_row = []
            asset_nos = []
            
            for i, item in enumerate(row):
                if item is None:
                    display_row.append("")
                else:
                    display_row.append(str(item))
                    
                    # Capture asset number from Asset No column
                    if i == asset_no_col_index:
                        asset_nos.append(str(item))
                    
                    # For cubicle analysis, extract asset numbers from Asset Numbers column
                    elif i == asset_numbers_col_index and str(item) not in ["None", ""]:
                        # Parse comma-separated asset numbers
                        numbers = [num.strip() for num in str(item).split(',') if num.strip()]
                        asset_nos.extend(numbers)
            
            item_id = tree.insert("", "end", values=display_row)
            if asset_nos:
                # Store all asset numbers for multi-asset handling
                asset_data[item_id] = asset_nos
        
        # Store asset data for sorting preservation
        self._asset_data = asset_data
        
        # Add double-click functionality to open asset details
        def on_double_click(event):
            item = tree.selection()[0] if tree.selection() else None
            if item and item in asset_data:
                asset_nos = asset_data[item]
                if len(asset_nos) == 1:
                    # Single asset - open directly
                    self._open_asset_details(asset_nos[0])
                else:
                    # Multiple assets - open multi-asset viewer
                    self._open_multi_asset_viewer(asset_nos)
        
        tree.bind("<Double-1>", on_double_click)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack elements
        tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        
        # Store tree reference for sorting
        self._current_tree = tree
    
    def _sort_table(self, tree, column):
        """Sort table by clicked column header."""
        try:
            # Get column index
            col_index = self._table_headers.index(column)
            
            # Toggle sort direction
            if column not in self._sort_direction:
                self._sort_direction[column] = "asc"
            else:
                self._sort_direction[column] = "desc" if self._sort_direction[column] == "asc" else "asc"
            
            # Get current data from tree
            items = [(tree.set(child, column), child) for child in tree.get_children('')]
            
            # Smart sorting function
            def smart_sort_key(value):
                """Smart sorting that handles numbers, dates, and text appropriately."""
                if not value or value == "":
                    return ("", "")  # Empty values sort first
                
                value_str = str(value).strip()
                
                # Check if it's a pure number
                if value_str.replace('.', '').replace('-', '').replace('+', '').isdigit():
                    try:
                        return (0, float(value_str))  # Numbers sort with priority 0
                    except ValueError:
                        pass
                
                # Check if it's a date (various formats)
                date_patterns = [
                    r'^\d{1,2}/\d{1,2}/\d{4}$',  # MM/DD/YYYY or M/D/YYYY
                    r'^\d{4}-\d{2}-\d{2}$',      # YYYY-MM-DD
                    r'^\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s+(AM|PM)$'  # MM/DD/YYYY HH:MM AM/PM
                ]
                
                for pattern in date_patterns:
                    if re.match(pattern, value_str):
                        try:
                            if '/' in value_str and ':' in value_str:
                                # Handle MM/DD/YYYY HH:MM AM/PM format
                                date_part = value_str.split()[0]
                                date_obj = datetime.strptime(date_part, "%m/%d/%Y")
                            elif '/' in value_str:
                                # Handle MM/DD/YYYY format
                                date_obj = datetime.strptime(value_str, "%m/%d/%Y")
                            elif '-' in value_str:
                                # Handle YYYY-MM-DD format
                                date_obj = datetime.strptime(value_str, "%Y-%m-%d")
                            else:
                                continue
                            return (1, date_obj.timestamp())  # Dates sort with priority 1
                        except ValueError:
                            continue
                
                # Default to string sorting
                return (2, value_str.lower())  # Text sorts with priority 2
            
            # Sort data
            reverse_order = self._sort_direction[column] == "desc"
            items.sort(key=lambda x: smart_sort_key(x[0]), reverse=reverse_order)
            
            # Rearrange items in tree
            for index, (val, child) in enumerate(items):
                tree.move(child, '', index)
            
            # Update column header to show sort direction
            for header in self._table_headers:
                if header == column:
                    direction_indicator = " ↑" if self._sort_direction[column] == "asc" else " ↓"
                    tree.heading(header, text=f"{header}{direction_indicator}")
                else:
                    tree.heading(header, text=header)
                    
        except Exception as e:
            print(f"Error sorting table: {e}")
            # Reset header if sorting fails
            tree.heading(column, text=column)
    
    def _export_audit_report(self):
        """Export audit report to Excel."""
        if self.audit_data is None:
            messagebox.showwarning("No Data", "No audit report data to export. Please generate a report first.")
            return
        
        try:
            # Ask user for save location using configured output directory with reports subdirectory
            reports_dir = os.path.join(self.config.output_directory, "reports")
            os.makedirs(reports_dir, exist_ok=True)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Audit Report",
                initialfile=f"Audit_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                initialdir=reports_dir
            )
            
            if filename:
                # Export to Excel with formatting
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    self.audit_data.to_excel(writer, sheet_name='Audit Report', index=False)
                    
                    # Get workbook and worksheet for formatting
                    workbook = writer.book
                    worksheet = writer.sheets['Audit Report']
                    
                    # Auto-adjust column widths
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                messagebox.showinfo("Export Complete", f"Audit report exported successfully to:\n{filename}")
                
        except Exception as e:
            error_handler.handle_exception(e, context="exporting audit report",
                                     parent_window=self.window)
    
    def _export_labels_report(self):
        """Export labels report to Excel."""
        if self.labels_data is None:
            messagebox.showwarning("No Data", "No labels report data to export. Please generate a report first.")
            return
        
        try:
            # Ask user for save location using configured output directory with labels subdirectory
            labels_dir = os.path.join(self.config.output_directory, "labels")
            os.makedirs(labels_dir, exist_ok=True)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Labels Report",
                initialfile=f"Labels_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                initialdir=labels_dir
            )
            
            if filename:
                # Export to Excel with formatting
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    self.labels_data.to_excel(writer, sheet_name='Labels Report', index=False)
                    
                    # Get workbook and worksheet for formatting
                    workbook = writer.book
                    worksheet = writer.sheets['Labels Report']
                    
                    # Auto-adjust column widths
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                messagebox.showinfo("Export Complete", f"Labels report exported successfully to:\n{filename}")
                
        except Exception as e:
            error_handler.handle_exception(e, context="exporting labels report",
                                     parent_window=self.window)

    def _export_duplicate_report(self):
        """Export duplicate report to Excel."""
        if self.duplicate_data is None:
            messagebox.showwarning("No Data", "No duplicate report data to export. Please generate a report first.")
            return
        
        try:
            # Ask user for save location using configured output directory with reports subdirectory
            reports_dir = os.path.join(self.config.output_directory, "reports")
            os.makedirs(reports_dir, exist_ok=True)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Duplicate Report",
                initialfile=f"Duplicate_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                initialdir=reports_dir
            )
            
            if filename:
                # Export to Excel
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    self.duplicate_data.to_excel(writer, sheet_name='Duplicate Report', index=False)
                    
                    # Auto-adjust column widths
                    workbook = writer.book
                    worksheet = writer.sheets['Duplicate Report']
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                messagebox.showinfo("Export Complete", f"Duplicate report exported successfully to:\n{filename}")
                
        except Exception as e:
            error_handler.handle_exception(e, context="exporting duplicate report",
                                     parent_window=self.window)
    
    def _export_cubicle_analysis(self):
        """Export cubicle analysis to Excel."""
        if self.cubicle_data is None:
            messagebox.showwarning("No Data", "No cubicle analysis data to export. Please generate an analysis first.")
            return
        
        try:
            # Ask user for save location using configured output directory with reports subdirectory
            reports_dir = os.path.join(self.config.output_directory, "reports")
            os.makedirs(reports_dir, exist_ok=True)
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Cubicle Analysis",
                initialfile=f"Cubicle_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                initialdir=reports_dir
            )
            
            if filename:
                # Export to Excel
                df = pd.DataFrame(self.cubicle_data)
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Cubicle Analysis', index=False)
                    
                    # Auto-adjust column widths
                    worksheet = writer.sheets['Cubicle Analysis']
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                messagebox.showinfo("Export Complete", f"Cubicle analysis exported successfully to:\n{filename}")
                
        except Exception as e:
            error_handler.handle_exception(e, context="exporting cubicle analysis",
                                     parent_window=self.window)
    
    def _open_asset_details(self, asset_no):
        """Open the asset details window for the specified asset number."""
        try:
            # Look up asset by asset number
            with self.db.get_connection() as conn:
                cursor = conn.execute("SELECT * FROM assets WHERE asset_no = ? AND is_deleted = 0", (asset_no,))
                result = cursor.fetchone()
            
            if result:
                # Convert Row object to dictionary
                asset_dict = dict(result)
                # Open the asset detail window (view-only with Details and History tabs)
                AssetDetailWindow(self.window, asset_dict, on_edit_callback=None)
            else:
                messagebox.showwarning("Asset Not Found", f"Asset {asset_no} could not be found.")
                
        except Exception as e:
            error_handler.handle_exception(e, context=f"opening asset details for {asset_no}",
                                     parent_window=self.window)
    
    def _open_multi_asset_viewer(self, asset_nos):
        """Open a multi-asset viewer popup for cycling through multiple assets."""
        try:
            # Look up all assets by asset numbers
            assets = []
            for asset_no in asset_nos:
                with self.db.get_connection() as conn:
                    cursor = conn.execute("SELECT * FROM assets WHERE asset_no = ? AND is_deleted = 0", (asset_no,))
                    result = cursor.fetchone()
                    if result:
                        assets.append(dict(result))
            
            if not assets:
                messagebox.showwarning("Assets Not Found", "None of the specified assets could be found.")
                return
            
            # Open multi-asset viewer window
            MultiAssetViewer(self.window, assets)
                
        except Exception as e:
            error_handler.handle_exception(e, context=f"opening multi-asset viewer for {asset_nos}",
                                     parent_window=self.window)
    
    def _refresh_overview(self):
        """Refresh the overview dashboard with current data."""
        try:
            # Clear existing widgets
            for widget in self.stats_frame.winfo_children():
                widget.destroy()
            for widget in self.charts_frame.winfo_children():
                widget.destroy()
            
            # Get database statistics
            stats = self._get_database_statistics()
            
            # Create statistics summary
            self._create_statistics_summary(stats)
            
            # Create charts
            self._create_dashboard_charts(stats)
            
        except Exception as e:
            error_handler.handle_exception(e, context="refreshing overview dashboard",
                                     parent_window=self.window)
    
    def _get_database_statistics(self):
        """Get comprehensive database statistics for the dashboard."""
        stats = {}
        
        try:
            with self.db.get_connection() as conn:
                # Total assets
                cursor = conn.execute("SELECT COUNT(*) FROM assets WHERE is_deleted = 0")
                stats['total_assets'] = cursor.fetchone()[0]
                
                # Assets by type
                cursor = conn.execute("""
                    SELECT asset_type, COUNT(*) as count 
                    FROM assets 
                    WHERE is_deleted = 0 AND asset_type IS NOT NULL AND asset_type != ''
                    GROUP BY asset_type 
                    ORDER BY count DESC
                """)
                stats['by_type'] = dict(cursor.fetchall())
                
                # Assets by manufacturer
                cursor = conn.execute("""
                    SELECT manufacturer, COUNT(*) as count 
                    FROM assets 
                    WHERE is_deleted = 0 AND manufacturer IS NOT NULL AND manufacturer != ''
                    GROUP BY manufacturer 
                    ORDER BY count DESC
                    LIMIT 10
                """)
                stats['by_manufacturer'] = dict(cursor.fetchall())
                
                # Assets by system
                cursor = conn.execute("""
                    SELECT system_name, COUNT(*) as count 
                    FROM assets 
                    WHERE is_deleted = 0 AND system_name IS NOT NULL AND system_name != ''
                    GROUP BY system_name 
                    ORDER BY count DESC
                    LIMIT 10
                """)
                stats['by_system'] = dict(cursor.fetchall())
                
                # Assets by location/room
                cursor = conn.execute("""
                    SELECT room, COUNT(*) as count 
                    FROM assets 
                    WHERE is_deleted = 0 AND room IS NOT NULL AND room != ''
                    GROUP BY room 
                    ORDER BY count DESC
                    LIMIT 10
                """)
                stats['by_room'] = dict(cursor.fetchall())
                
                # Audit compliance stats
                one_year_ago = (datetime.now() - timedelta(days=365)).isoformat()
                cursor = conn.execute("""
                    SELECT 
                        SUM(CASE WHEN 
                            CASE 
                                WHEN audit_date LIKE '%/%/%' THEN
                                    CASE 
                                        WHEN length(substr(audit_date, 1, instr(audit_date, '/') - 1)) = 1 THEN
                                            substr(audit_date, length(audit_date) - 3, 4) || '-0' || 
                                            substr(audit_date, 1, instr(audit_date, '/') - 1) || '-' ||
                                            CASE 
                                                WHEN length(substr(audit_date, instr(audit_date, '/') + 1, instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1)) = 1 
                                                THEN '0' || substr(audit_date, instr(audit_date, '/') + 1, instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1)
                                                ELSE substr(audit_date, instr(audit_date, '/') + 1, instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1)
                                            END
                                        ELSE
                                            substr(audit_date, length(audit_date) - 3, 4) || '-' || 
                                            substr(audit_date, 1, instr(audit_date, '/') - 1) || '-' ||
                                            CASE 
                                                WHEN length(substr(audit_date, instr(audit_date, '/') + 1, instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1)) = 1 
                                                THEN '0' || substr(audit_date, instr(audit_date, '/') + 1, instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1)
                                                ELSE substr(audit_date, instr(audit_date, '/') + 1, instr(substr(audit_date, instr(audit_date, '/') + 1), '/') - 1)
                                            END
                                    END
                                ELSE audit_date
                            END >= ? THEN 1 ELSE 0 END) as audited_recently,
                        COUNT(*) as total
                    FROM assets 
                    WHERE is_deleted = 0 AND audit_date IS NOT NULL AND audit_date != ''
                """, (one_year_ago,))
                audit_result = cursor.fetchone()
                stats['audit_recent'] = audit_result[0] if audit_result[0] else 0
                stats['audit_total'] = audit_result[1] if audit_result[1] else 0
                stats['audit_overdue'] = stats['audit_total'] - stats['audit_recent']
                
                # Assets without audit dates
                cursor = conn.execute("""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE is_deleted = 0 AND (audit_date IS NULL OR audit_date = '')
                """)
                stats['audit_never'] = cursor.fetchone()[0]
                
                # Recent additions (last 30 days)
                cursor = conn.execute("""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE is_deleted = 0 AND created_date >= ?
                """, (one_year_ago,))
                stats['recent_additions'] = cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting database statistics: {e}")
            # Return empty stats on error
            stats = {
                'total_assets': 0,
                'by_type': {},
                'by_manufacturer': {},
                'by_system': {},
                'by_room': {},
                'audit_recent': 0,
                'audit_total': 0,
                'audit_overdue': 0,
                'audit_never': 0,
                'recent_additions': 0
            }
        
        return stats
    
    def _create_statistics_summary(self, stats):
        """Create the statistics summary section."""
        # Title
        title_label = ctk.CTkLabel(self.stats_frame, 
                                  text="📊 Database Summary", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=(10, 15))
        
        # Create grid of stat cards
        cards_frame = ctk.CTkFrame(self.stats_frame, fg_color="transparent")
        cards_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Configure grid
        cards_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        # Stat cards data
        cards_data = [
            ("📦 Total Assets", f"{stats['total_assets']:,}", "blue"),
            ("✅ Recently Audited", f"{stats['audit_recent']:,}", "green"),
            ("⚠️ Audit Overdue", f"{stats['audit_overdue']:,}", "orange"),
            ("➕ Added (30 days)", f"{stats['recent_additions']:,}", "purple")
        ]
        
        # Create stat cards
        for i, (title, value, color) in enumerate(cards_data):
            card = ctk.CTkFrame(cards_frame)
            card.grid(row=0, column=i, padx=10, pady=10, sticky="ew")
            
            title_label = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=12))
            title_label.pack(pady=(10, 5))
            
            value_label = ctk.CTkLabel(card, text=value, 
                                     font=ctk.CTkFont(size=18, weight="bold"))
            value_label.pack(pady=(0, 10))
    
    def _create_dashboard_charts(self, stats):
        """Create charts for the dashboard."""
        # Set up matplotlib for dark theme
        plt.style.use('default')
        
        # Create main charts frame with grid
        charts_container = ctk.CTkFrame(self.charts_frame, fg_color="transparent")
        charts_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configure grid
        charts_container.grid_columnconfigure((0, 1), weight=1)
        charts_container.grid_rowconfigure((0, 1, 2), weight=1)
        
        # Chart 1: Asset Types Pie Chart
        if stats['by_type']:
            self._create_pie_chart(charts_container, stats['by_type'], 
                                 "Asset Types Distribution", 0, 0)
        
        # Chart 2: Top Manufacturers Bar Chart
        if stats['by_manufacturer']:
            self._create_bar_chart(charts_container, stats['by_manufacturer'], 
                                 "Top 10 Manufacturers", 0, 1)
        
        # Chart 3: Audit Status Pie Chart
        audit_data = {
            'Recently Audited': stats['audit_recent'],
            'Overdue for Audit': stats['audit_overdue'],
            'Never Audited': stats['audit_never']
        }
        if any(audit_data.values()):
            self._create_pie_chart(charts_container, audit_data, 
                                 "Audit Compliance Status", 1, 0)
        
        # Chart 4: Top Systems Bar Chart
        if stats['by_system']:
            self._create_bar_chart(charts_container, stats['by_system'], 
                                 "Top 10 Systems", 1, 1)
        
        # Chart 5: Room Distribution (spans both columns)
        if stats['by_room']:
            self._create_bar_chart(charts_container, stats['by_room'], 
                                 "Top 10 Locations", 2, 0, columnspan=2)
    
    def _create_pie_chart(self, parent, data, title, row, col, columnspan=1):
        """Create a pie chart."""
        # Create frame for the chart
        chart_frame = ctk.CTkFrame(parent)
        chart_frame.grid(row=row, column=col, columnspan=columnspan, 
                        padx=10, pady=10, sticky="nsew")
        
        # Create figure
        fig = Figure(figsize=(6, 4), dpi=100)
        fig.patch.set_facecolor('#2b2b2b')  # Dark background
        ax = fig.add_subplot(111)
        
        # Prepare data (limit to top 8 items for readability)
        if len(data) > 8:
            sorted_items = sorted(data.items(), key=lambda x: x[1], reverse=True)
            top_items = dict(sorted_items[:7])
            other_sum = sum(v for k, v in sorted_items[7:])
            if other_sum > 0:
                top_items['Others'] = other_sum
            data = top_items
        
        labels = list(data.keys())
        sizes = list(data.values())
        
        # Create pie chart with nice colors
        colors = plt.cm.Set3(range(len(labels)))
        wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct='%1.1f%%', 
                                         colors=colors, startangle=90)
        
        # Style the chart
        ax.set_title(title, color='white', fontsize=12, weight='bold', pad=20)
        
        # Style text
        for text in texts:
            text.set_color('white')
            text.set_fontsize(9)
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(8)
            autotext.set_weight('bold')
        
        # Create canvas and pack
        canvas = FigureCanvasTkAgg(fig, chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)
    
    def _create_bar_chart(self, parent, data, title, row, col, columnspan=1):
        """Create a horizontal bar chart."""
        # Create frame for the chart
        chart_frame = ctk.CTkFrame(parent)
        chart_frame.grid(row=row, column=col, columnspan=columnspan, 
                        padx=10, pady=10, sticky="nsew")
        
        # Create figure
        fig = Figure(figsize=(8, 4), dpi=100)
        fig.patch.set_facecolor('#2b2b2b')  # Dark background
        ax = fig.add_subplot(111)
        ax.set_facecolor('#2b2b2b')
        
        # Prepare data (reverse for horizontal bar chart)
        labels = list(data.keys())
        values = list(data.values())
        
        # Limit to top 10 and reverse order for better display
        if len(labels) > 10:
            sorted_pairs = sorted(zip(labels, values), key=lambda x: x[1], reverse=True)
            labels, values = zip(*sorted_pairs[:10])
        
        # Reverse order so highest values appear at top
        labels = labels[::-1]
        values = values[::-1]
        
        # Create horizontal bar chart
        bars = ax.barh(range(len(labels)), values, color=plt.cm.viridis(range(len(labels))))
        
        # Style the chart
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels, color='white', fontsize=9)
        ax.set_xlabel('Count', color='white', fontsize=10)
        ax.set_title(title, color='white', fontsize=12, weight='bold', pad=20)
        
        # Style axes
        ax.tick_params(colors='white')
        ax.spines['bottom'].set_color('white')
        ax.spines['left'].set_color('white')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        # Add value labels on bars
        for i, (bar, value) in enumerate(zip(bars, values)):
            ax.text(value + max(values) * 0.01, bar.get_y() + bar.get_height()/2, 
                   str(value), va='center', color='white', fontsize=8, weight='bold')
        
        # Adjust layout
        fig.tight_layout()
        
        # Create canvas and pack
        canvas = FigureCanvasTkAgg(fig, chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)

    def _on_closing(self):
        """Handle window closing."""
        try:
            if hasattr(self.db, 'close'):
                self.db.close()
        except Exception:
            pass  # Ignore errors during cleanup
        
        self.window.destroy()


def open_reports_analysis_window(parent):
    """Helper function to open the Reports and Analysis window."""
    ReportsAnalysisWindow(parent)