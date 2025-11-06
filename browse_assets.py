"""
Browse and search assets window for viewing database contents.
Enhanced with advanced search capabilities, filtering, and asset management.
"""

import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
import tkinter as tk
import os
from typing import Dict, List, Any, Optional
from asset_database import AssetDatabase
from datetime import datetime, timedelta
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from performance_monitoring import performance_monitor
from database_service import database_service
from ui_components import AssetDetailWindow, SearchableDropdown, DatePicker
import re


class BrowseAssetsWindow:
    """Enhanced window for browsing, searching, and managing assets from the database."""
    
    def __init__(self, parent, config = None):
        """Initialize the browse assets window."""
        self.parent = parent
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = config or self.config_manager.get_config()
        
        # Create database instance using the configured database path
        self.db = AssetDatabase(self.config.database_path)
        
        # Use global database service instance
        self.db_service = database_service
        
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Browse & Search Assets")
        self.window.geometry("1000x600")
        self.window.minsize(800, 500)
        self.window.transient(parent)
        
        # Center the window
        self._center_window()
        
        # Variables for search and filtering
        self.filter_rows = []  # Will store filter row widgets and variables
        self.filter_logic_var = tk.StringVar(value="AND")  # AND or OR logic for combining filters
        self.sort_field = tk.StringVar()
        self.sort_direction = tk.StringVar(value="asc")
        self.items_per_page = tk.IntVar(value=100)
        self.current_page = tk.IntVar(value=1)
        
        # Current data
        self.current_assets = []
        self.total_count = 0
        self.filtered_count = 0
        self.selected_asset = None
        self._search_after_id = None
        
        # Get database fields and unique values for dropdowns
        self.db_fields = self._get_database_fields()
        self.unique_values = self._get_unique_field_values()
        
        # Create a set of database field names that should use dropdowns
        # Convert config display names to database field names
        self.dropdown_db_fields = set()
        if hasattr(self.config, 'dropdown_fields'):
            for display_name in self.config.dropdown_fields:
                db_field = self.db._generate_safe_column_name(display_name)
                self.dropdown_db_fields.add(db_field)
        
        self._create_widgets()
        self._initialize_empty_state()
        
        # Load saved searches into listbox
        self._refresh_saved_searches_list()
        
        # Focus handling
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.window.focus_force()
        
        # Bind keyboard shortcuts
        self._setup_keyboard_shortcuts()
    
    def _get_database_fields(self):
        """Get database fields for column headers and search dropdown - includes ALL fields."""
        try:
            # Get all table columns
            all_columns = self.db.get_table_columns()
            
            # Only exclude 'id' field - include everything else (including system fields)
            exclude_columns = {'id'}
            display_columns = [col for col in all_columns if col not in exclude_columns]
            
            # Convert column names to human-readable format with priority ordering
            priority_fields = [
                'asset_no', 'asset_type', 'manufacturer', 'model', 'serial_number', 
                'status', 'location', 'room', 'system_name', 'notes'
            ]
            
            # System/metadata fields to show at the end
            system_fields = [
                'data_source', 'created_by', 'modified_by', 'created_date', 
                'modified_date', 'is_deleted'
            ]
            
            readable_fields = []
            
            # Add priority fields first
            for priority_col in priority_fields:
                if priority_col in display_columns:
                    readable = self.db._column_to_header(priority_col)
                    readable_fields.append({
                        'db_name': priority_col, 
                        'display_name': readable,
                        'is_searchable': True,
                        'is_filterable': priority_col in ['asset_type', 'manufacturer', 'status', 'location']
                    })
            
            # Add remaining regular fields (excluding system fields)
            for col in display_columns:
                if col not in priority_fields and col not in system_fields:
                    readable = self.db._column_to_header(col)
                    readable_fields.append({
                        'db_name': col, 
                        'display_name': readable,
                        'is_searchable': True,
                        'is_filterable': False
                    })
            
            # Add system/metadata fields at the end
            for col in system_fields:
                if col in display_columns:
                    readable = self.db._column_to_header(col)
                    readable_fields.append({
                        'db_name': col, 
                        'display_name': readable,
                        'is_searchable': True,
                        'is_filterable': col in ['data_source', 'created_by', 'modified_by']
                    })
            
            return readable_fields
        except Exception as e:
            print(f"Error getting database fields: {e}")
            # Fallback
            return [
                {'db_name': 'asset_no', 'display_name': 'Asset No.', 'is_searchable': True, 'is_filterable': False},
                {'db_name': 'manufacturer', 'display_name': 'Manufacturer', 'is_searchable': True, 'is_filterable': True},
                {'db_name': 'model', 'display_name': 'Model', 'is_searchable': True, 'is_filterable': False},
            ]
    
    def _get_unique_field_values(self):
        """Get unique values for dropdown fields from config."""
        unique_vals = {}
        
        # Get dropdown fields from config (these are display names)
        dropdown_fields = self.config.dropdown_fields if hasattr(self.config, 'dropdown_fields') else []
        
        # Convert display names to database field names and get unique values
        for display_name in dropdown_fields:
            try:
                # Convert display name to database field name
                db_field = self.db._generate_safe_column_name(display_name)
                
                # Get unique values using the database field name
                values = self.db.get_unique_field_values(db_field)
                
                # Store using the database field name as key
                unique_vals[db_field] = sorted([v for v in values if v and v.strip()])
            except Exception as e:
                print(f"Error getting unique values for {display_name} (db field: {db_field}): {e}")
                unique_vals[db_field] = []
        
        return unique_vals
    
    def _setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts for improved usability."""
        # Focus on first filter value entry when Ctrl+F is pressed
        self.window.bind("<Control-f>", lambda e: self._focus_first_filter())
        self.window.bind("<F5>", lambda e: self._load_initial_data())
        self.window.bind("<Control-r>", lambda e: self._load_initial_data())
        self.window.bind("<Escape>", lambda e: self._clear_all_filters())
        self.window.bind("<Delete>", lambda e: self._delete_asset() if self.selected_asset else None)
        self.window.bind("<Return>", lambda e: self._view_details() if self.selected_asset else None)
        self.window.bind("<Control-Return>", lambda e: self._do_search())
    
    def _focus_first_filter(self):
        """Focus on the value entry of the first filter row."""
        if self.filter_rows:
            first_row = self.filter_rows[0]
            if 'value_entry' in first_row:
                first_row['value_entry'].focus_set()
    
    def _center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """Create the enhanced UI widgets."""
        # Main container with sidebar and content
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configure main frame grid
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
        
        # Left sidebar for filters and tools
        self._create_sidebar(main_frame)
        
        # Right content area for search results and details
        self._create_content_area(main_frame)
        
        # Bottom status bar
        self._create_enhanced_status_bar()
    
    def _create_sidebar(self, parent):
        """Create the left sidebar with search controls and filters."""
        sidebar = ctk.CTkScrollableFrame(parent, width=350)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Title
        title_label = ctk.CTkLabel(sidebar, text="Asset Search & Filter", 
                                  font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=(10, 20))
        
        # Filter Builder Section (replaces Quick Search, Advanced Filters, and Date Range)
        self._create_filter_builder_section(sidebar)
        
        # Saved Searches Section
        self._create_saved_searches_section(sidebar)
        
        # Action Buttons Section
        self._create_sidebar_actions(sidebar)
    
    def _create_filter_builder_section(self, parent):
        """Create the new advanced filter builder with AND/OR grouping."""
        filter_frame = ctk.CTkFrame(parent)
        filter_frame.pack(fill="x", pady=(0, 15))
        
        # Section title with clear button
        title_frame = ctk.CTkFrame(filter_frame)
        title_frame.pack(fill="x", pady=(10, 10))
        
        ctk.CTkLabel(title_frame, text="Filters", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=10)
        
        clear_btn = ctk.CTkButton(title_frame, text="Clear All", width=70, height=25,
                                 command=self._clear_all_filters)
        clear_btn.pack(side="right", padx=10)
        
        # Container for filter rows
        self.filters_container = ctk.CTkScrollableFrame(filter_frame, height=300)
        self.filters_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Track filter rows and groups
        self.filter_rows = []
        self.filter_groups = []  # For nested grouping structure
        self.group_counter = 0  # Counter for unique group IDs
        
        # Create root group (always exists)
        self._create_root_group()
        
        # Root group logic selector
        root_logic_frame = ctk.CTkFrame(filter_frame)
        root_logic_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(root_logic_frame, text="Root Logic:", 
                    font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(5, 10))
        
        ctk.CTkLabel(root_logic_frame, text="Combine with:", 
                    font=ctk.CTkFont(size=10)).pack(side="left", padx=(0, 5))
        
        ctk.CTkRadioButton(root_logic_frame, text="AND", 
                         variable=self.root_group['logic_var'], value="AND",
                         width=50, height=20).pack(side="left", padx=2)
        ctk.CTkRadioButton(root_logic_frame, text="OR", 
                         variable=self.root_group['logic_var'], value="OR",
                         width=50, height=20).pack(side="left", padx=2)
        
        # Control buttons frame
        controls_frame = ctk.CTkFrame(filter_frame)
        controls_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        add_filter_btn = ctk.CTkButton(controls_frame, text="+ Add Filter", 
                                      command=lambda: self._add_filter_row(self.root_group),
                                      height=28, width=120)
        add_filter_btn.pack(side="left", padx=(0, 5))
        
        add_group_btn = ctk.CTkButton(controls_frame, text="+ Add Group", 
                                     command=lambda: self._add_group(self.root_group),
                                     height=28, width=120,
                                     fg_color="gray40", hover_color="gray30")
        add_group_btn.pack(side="left", padx=(0, 5))
        
        # Help text
        help_frame = ctk.CTkFrame(filter_frame)
        help_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        help_text = ctk.CTkLabel(help_frame, 
                                text="💡 Tip: Use groups to create complex logic like (F1 AND F2) OR F3\nEach group can have its own AND/OR logic.", 
                                font=ctk.CTkFont(size=10),
                                text_color="gray60",
                                justify="left")
        help_text.pack(anchor="w", padx=5, pady=5)
    
    def _create_root_group(self):
        """Create the root group that contains all filters."""
        self.root_group = {
            'id': 'root',
            'parent': None,
            'frame': self.filters_container,
            'content_frame': self.filters_container,
            'logic_var': tk.StringVar(value="AND"),
            'items': [],  # List of filters and sub-groups
            'is_group': True,
            'depth': 0
        }
        self.filter_groups.append(self.root_group)
    
    def _add_group(self, parent_group, logic="AND"):
        """Add a new filter group with its own AND/OR logic."""
        self.group_counter += 1
        group_id = f"group_{self.group_counter}"
        depth = parent_group['depth'] + 1
        
        # Create group frame with visual nesting
        group_outer_frame = ctk.CTkFrame(parent_group['content_frame'], 
                                        fg_color=("gray85", "gray20"),
                                        border_width=2,
                                        border_color=("gray60", "gray40"))
        group_outer_frame.pack(fill="x", pady=5, padx=(depth * 20, 0))
        
        # Group header with logic selector and controls (using two rows)
        header_frame = ctk.CTkFrame(group_outer_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=5, pady=5)
        
        # First row: Group label, logic selector, and delete button
        top_row = ctk.CTkFrame(header_frame, fg_color="transparent")
        top_row.pack(fill="x", pady=(0, 3))
        
        # Group label
        ctk.CTkLabel(top_row, text=f"Group {self.group_counter}:", 
                    font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(0, 10))
        
        # Logic selector for this group
        logic_var = tk.StringVar(value=logic)
        ctk.CTkLabel(top_row, text="Combine with:", 
                    font=ctk.CTkFont(size=10)).pack(side="left", padx=(0, 5))
        ctk.CTkRadioButton(top_row, text="AND", variable=logic_var, value="AND",
                         width=50, height=20).pack(side="left", padx=2)
        ctk.CTkRadioButton(top_row, text="OR", variable=logic_var, value="OR",
                         width=50, height=20).pack(side="left", padx=2)
        
        # Delete button on right side
        delete_btn = ctk.CTkButton(top_row, text="✕", width=35, height=24,
                                  fg_color="red", hover_color="darkred",
                                  command=lambda: self._remove_group(group_id))
        delete_btn.pack(side="right")
        
        # Second row: Add Filter and Add Group buttons
        bottom_row = ctk.CTkFrame(header_frame, fg_color="transparent")
        bottom_row.pack(fill="x")
        
        # Group controls
        ctk.CTkButton(bottom_row, text="+ Filter", width=80, height=24,
                     command=lambda g=None: self._add_filter_row(g or group_data)).pack(side="left", padx=(0, 5))
        ctk.CTkButton(bottom_row, text="+ Group", width=80, height=24,
                     fg_color="gray40", hover_color="gray30",
                     command=lambda g=None: self._add_group(g or group_data)).pack(side="left", padx=2)
        
        # Content frame for filters and sub-groups
        content_frame = ctk.CTkFrame(group_outer_frame, fg_color="transparent")
        content_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        # Store group data
        group_data = {
            'id': group_id,
            'parent': parent_group,
            'frame': group_outer_frame,
            'content_frame': content_frame,
            'logic_var': logic_var,
            'items': [],
            'is_group': True,
            'depth': depth
        }
        
        parent_group['items'].append(group_data)
        self.filter_groups.append(group_data)
        
        return group_data

    
    def _add_filter_row(self, parent_group, field=None, operator=None, value=None):
        """Add a new filter criteria row to a specific group."""
        depth = parent_group['depth'] + 1
        
        row_frame = ctk.CTkFrame(parent_group['content_frame'])
        row_frame.pack(fill="x", pady=3, padx=(depth * 20, 0))
        
        # Row controls
        row_controls = ctk.CTkFrame(row_frame)
        row_controls.pack(fill="x", padx=5, pady=2)
        
        row_num = len(self.filter_rows) + 1
        ctk.CTkLabel(row_controls, text=f"Filter {row_num}:", 
                    font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(0, 10))
        
        delete_btn = ctk.CTkButton(row_controls, text="✕", width=25, height=25,
                                  fg_color="red", hover_color="darkred",
                                  command=lambda: self._remove_filter_row(filter_data))
        delete_btn.pack(side="right")
        
        # Field and Operator on same line
        field_operator_frame = ctk.CTkFrame(row_frame)
        field_operator_frame.pack(fill="x", padx=5, pady=2)
        
        # Field dropdown - use SearchableDropdown with alphabetized list
        field_var = tk.StringVar(value=field if field else "")
        field_names = sorted([f['display_name'] for f in self.db_fields if f['is_searchable']])
        field_dropdown = SearchableDropdown(field_operator_frame, values=field_names, 
                                           variable=field_var, width=140, height=28)
        field_dropdown.pack(side="left", padx=(0, 5))
        
        # Operator dropdown - use SearchableDropdown
        operator_var = tk.StringVar(value=operator if operator else "contains")
        operators = ["equals", "contains", "does not equal", "does not contain", "starts with", "ends with"]
        operator_dropdown = SearchableDropdown(field_operator_frame, values=operators,
                                              variable=operator_var, width=140, height=28)
        operator_dropdown.pack(side="left")
        
        # Value entry on second line (will be replaced with DatePicker for date fields)
        value_frame = ctk.CTkFrame(row_frame)
        value_frame.pack(fill="x", padx=5, pady=2)
        
        value_var = tk.StringVar(value=value if value else "")
        value_entry = ctk.CTkEntry(value_frame, textvariable=value_var,
                                  placeholder_text="Enter value...")
        value_entry.pack(fill="x")
        
        # Store filter row data
        filter_data = {
            'frame': row_frame,
            'parent_group': parent_group,
            'field_var': field_var,
            'operator_var': operator_var,
            'value_var': value_var,
            'field_dropdown': field_dropdown,
            'operator_dropdown': operator_dropdown,
            'value_entry': value_entry,
            'value_frame': value_frame,
            'is_date_field': False,
            'is_dropdown_field': False,
            'is_group': False
        }
        self.filter_rows.append(filter_data)
        parent_group['items'].append(filter_data)
        
        # Set up field change callback to switch between entry and date picker
        field_var.trace_add('write', lambda *args: self._on_filter_field_change(filter_data))
        
        # Initialize the correct widget for the current field
        if field:
            self._on_filter_field_change(filter_data)
        
        return filter_data
    
    def _on_filter_field_change(self, filter_data):
        """Handle field change to switch between text entry, date picker, or dropdown."""
        field_display = filter_data['field_var'].get()
        
        # Get the database field name
        db_field = next((f['db_name'] for f in self.db_fields 
                        if f['display_name'] == field_display), None)
        
        # Check if this is a date field (any field containing "date")
        is_date = db_field and 'date' in db_field.lower()
        
        # Check if this is a dropdown field from config (using db field name)
        is_dropdown = db_field and db_field in self.dropdown_db_fields
        
        # Determine current field type
        current_type = 'date' if filter_data.get('is_date_field', False) else \
                      ('dropdown' if filter_data.get('is_dropdown_field', False) else 'text')
        new_type = 'date' if is_date else ('dropdown' if is_dropdown else 'text')
        
        # If field type changed OR if staying dropdown but field changed, update the value widget
        # For dropdowns, we need to update even if type is same (dropdown->dropdown) but field changed
        needs_update = (current_type != new_type) or \
                      (new_type == 'dropdown' and filter_data.get('last_db_field') != db_field)
        
        # Track the current db_field for dropdown comparisons
        if new_type == 'dropdown':
            filter_data['last_db_field'] = db_field
        
        if needs_update:
            filter_data['is_date_field'] = is_date
            filter_data['is_dropdown_field'] = is_dropdown
            
            # Clear the value FIRST to prevent trace callbacks on destroyed widgets
            try:
                filter_data['value_var'].set("")
            except tk.TclError:
                pass
            
            # Destroy old value widget
            if 'value_entry' in filter_data and filter_data['value_entry'].winfo_exists():
                filter_data['value_entry'].destroy()
            if 'date_picker' in filter_data and hasattr(filter_data['date_picker'], 'destroy'):
                filter_data['date_picker'].destroy()
            
            value_frame = filter_data['value_frame']
            value_var = filter_data['value_var']
            
            if is_date:
                # Create DatePicker for date fields
                date_picker = DatePicker(value_frame, variable=value_var, width=285)
                date_picker.pack(fill="x")
                filter_data['date_picker'] = date_picker
                filter_data['value_entry'] = date_picker  # For focus purposes
                
                # Update operators for date fields
                date_operators = ["equals", "before", "after", "between"]
                filter_data['operator_dropdown'].values_all = date_operators[:]
                
                # Set default operator for dates if current operator not applicable
                current_op = filter_data['operator_var'].get()
                if current_op not in date_operators:
                    filter_data['operator_var'].set("equals")
                
                # Add operator change callback to show/hide second date picker for "between"
                filter_data['operator_var'].trace_add('write', 
                    lambda *args: self._on_date_operator_change(filter_data))
            
            elif is_dropdown:
                # Get new dropdown values for this field
                dropdown_values = [""] + self.unique_values.get(db_field, [])
                
                # If value_entry exists and is already a SearchableDropdown, update its values
                # Otherwise, create a new SearchableDropdown
                if ('value_entry' in filter_data and 
                    hasattr(filter_data['value_entry'], 'values_all') and
                    filter_data['value_entry'].winfo_exists()):
                    # Update existing SearchableDropdown values
                    filter_data['value_entry'].values_all = dropdown_values[:]
                else:
                    # Create new SearchableDropdown
                    dropdown = SearchableDropdown(value_frame, values=dropdown_values, 
                                                variable=value_var, width=285, height=28)
                    dropdown.pack(fill="x")
                    filter_data['value_entry'] = dropdown
                
                # Update operators for dropdown fields (same as text)
                text_operators = ["equals", "contains", "does not equal", "does not contain", 
                                "starts with", "ends with"]
                filter_data['operator_dropdown'].values_all = text_operators[:]
                
                # Set default operator if current operator not applicable
                current_op = filter_data['operator_var'].get()
                if current_op not in text_operators:
                    filter_data['operator_var'].set("equals")
            
            else:
                # Create text entry for regular fields
                value_entry = ctk.CTkEntry(value_frame, textvariable=value_var,
                                          placeholder_text="Enter value...")
                value_entry.pack(fill="x")
                filter_data['value_entry'] = value_entry
                
                # Update operators for text fields
                text_operators = ["equals", "contains", "does not equal", "does not contain", 
                                "starts with", "ends with"]
                filter_data['operator_dropdown'].values_all = text_operators[:]
                
                # Set default operator for text if current operator not applicable
                current_op = filter_data['operator_var'].get()
                if current_op not in text_operators:
                    filter_data['operator_var'].set("contains")
    
    def _on_date_operator_change(self, filter_data):
        """Handle date operator change to show/hide second date picker for 'between'."""
        operator = filter_data['operator_var'].get()
        
        if operator == "between":
            # Show second date picker if not already shown
            if 'date_picker_end' not in filter_data or not filter_data['date_picker_end'].winfo_exists():
                value_frame = filter_data['value_frame']
                
                # Add "to" label
                to_label = ctk.CTkLabel(value_frame, text="to", font=ctk.CTkFont(size=11))
                to_label.pack(pady=2)
                filter_data['to_label'] = to_label
                
                # Add second date picker for end date
                end_date_var = tk.StringVar(value="")
                date_picker_end = DatePicker(value_frame, variable=end_date_var, width=285)
                date_picker_end.pack(fill="x")
                filter_data['date_picker_end'] = date_picker_end
                filter_data['value_var_end'] = end_date_var
        else:
            # Hide/remove second date picker if it exists
            if 'date_picker_end' in filter_data and filter_data['date_picker_end'].winfo_exists():
                filter_data['date_picker_end'].destroy()
                del filter_data['date_picker_end']
            if 'to_label' in filter_data and filter_data['to_label'].winfo_exists():
                filter_data['to_label'].destroy()
                del filter_data['to_label']
            if 'value_var_end' in filter_data:
                del filter_data['value_var_end']
    
    def _remove_filter_row(self, filter_data):
        """Remove a filter row."""
        if filter_data in self.filter_rows:
            # Remove trace callbacks FIRST to prevent errors
            try:
                trace_info = filter_data['field_var'].trace_info()
                for trace_id in trace_info:
                    filter_data['field_var'].trace_remove('write', trace_id[1])
            except (KeyError, tk.TclError, AttributeError):
                pass
            
            try:
                trace_info = filter_data['operator_var'].trace_info()
                for trace_id in trace_info:
                    filter_data['operator_var'].trace_remove('write', trace_id[1])
            except (KeyError, tk.TclError, AttributeError):
                pass
            
            # Remove from parent group's items list BEFORE destroying widgets
            if 'parent_group' in filter_data and filter_data['parent_group']:
                parent_group = filter_data['parent_group']
                if filter_data in parent_group['items']:
                    parent_group['items'].remove(filter_data)
            
            # Remove from filter_rows list BEFORE destroying
            self.filter_rows.remove(filter_data)
            
            # Destroy the frame to destroy all widgets
            # DO NOT clear StringVars after this - it triggers callbacks on destroyed widgets
            try:
                filter_data['frame'].destroy()
            except (KeyError, tk.TclError):
                pass
    
    def _remove_group(self, group_id):
        """Remove a filter group and all its contents."""
        group = next((g for g in self.filter_groups if g['id'] == group_id), None)
        if not group or group['id'] == 'root':
            return
        
        # Remove all filters and sub-groups in this group
        items_copy = group['items'][:]  # Copy to avoid modification during iteration
        for item in items_copy:
            if item.get('is_group'):
                self._remove_group(item['id'])
            else:
                self._remove_filter_row(item)
        
        # Remove from parent group's items list
        if group['parent'] and group in group['parent']['items']:
            group['parent']['items'].remove(group)
        
        # Destroy the frame
        group['frame'].destroy()
        
        # Remove from filter_groups list
        if group in self.filter_groups:
            self.filter_groups.remove(group)
    
    def _clear_all_filters(self):
        """Clear all filters and groups, recreate root group."""
        # Remove all filter rows
        for filter_data in self.filter_rows[:]:
            self._remove_filter_row(filter_data)
        
        # Remove all groups except root
        for group in self.filter_groups[:]:
            if group['id'] != 'root':
                self._remove_group(group['id'])
        
        # Clear root group items
        self.root_group['items'] = []
        
        # Clear the results area and show empty state
        self._initialize_empty_state()
        
        # Update status
        self.status_label.configure(text="All filters cleared - click Search to find assets")
        
        # Reset pagination
        self.current_page.set(1)
    
    def _create_saved_searches_section(self, parent):
        """Create saved searches management."""
        saved_frame = ctk.CTkFrame(parent)
        saved_frame.pack(fill="x", pady=(0, 15))
        
        # Section title
        title_frame = ctk.CTkFrame(saved_frame)
        title_frame.pack(fill="x", pady=(10, 5))
        
        ctk.CTkLabel(title_frame, text="Saved Searches", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        
        save_btn = ctk.CTkButton(title_frame, text="Save Current", width=80, height=25,
                                command=self._save_current_search)
        save_btn.pack(side="right", padx=5)
        
        # Saved searches list with larger font
        self.saved_searches_listbox = tk.Listbox(saved_frame, height=4, 
                                                 font=("TkDefaultFont", 12))
        self.saved_searches_listbox.pack(fill="x", padx=10, pady=(0, 5))
        
        # Saved search actions
        saved_actions = ctk.CTkFrame(saved_frame)
        saved_actions.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkButton(saved_actions, text="Load", width=60, height=25,
                     command=self._load_saved_search).pack(side="left", padx=(0, 5))
        ctk.CTkButton(saved_actions, text="Delete", width=60, height=25,
                     command=self._delete_saved_search).pack(side="left")
    
    def _create_sidebar_actions(self, parent):
        """Create action buttons in sidebar."""
        actions_frame = ctk.CTkFrame(parent)
        actions_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(actions_frame, text="Actions", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(10, 5))
        
        actions = [
            ("🔍 Search", self._do_search, "#1f538d"),
            ("🗑️ Clear All", self._clear_all_filters, "#666666"),
            ("🏷️ Request Label", self._request_labels_for_filtered, "#2d5a27"),
            ("💾 Export CSV", self._export_current_results, "#8f6f2d")
        ]
        
        for text, command, color in actions:
            btn = ctk.CTkButton(actions_frame, text=text, command=command,
                               fg_color=color, height=32)
            btn.pack(fill="x", padx=10, pady=2)
    
    def _create_content_area(self, parent):
        """Create the main content area with results table and details."""
        self.content_frame = ctk.CTkFrame(parent)
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        
        # Configure content frame grid
        self.content_frame.grid_rowconfigure(1, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)
        
        # Results header with sorting and pagination
        self._create_results_header(self.content_frame)
        
        # Results table
        self._create_enhanced_table(self.content_frame)
        
        # Asset details panel (initially hidden)
        self._create_details_panel(self.content_frame)
        
        # Pagination controls
        self._create_pagination_controls(self.content_frame)
    def _create_results_header(self, parent):
        """Create results header with sorting and display options."""
        header_frame = ctk.CTkFrame(parent)
        header_frame.grid(row=0, column=0, sticky="ew", pady=(10, 5))
        header_frame.grid_columnconfigure(1, weight=1)
        
        # Results count and info
        self.results_info_label = ctk.CTkLabel(header_frame, text="Loading...", 
                                              font=ctk.CTkFont(size=12))
        self.results_info_label.grid(row=0, column=0, padx=10, sticky="w")
        
        # Sorting controls
        sort_frame = ctk.CTkFrame(header_frame)
        sort_frame.grid(row=0, column=2, padx=10, sticky="e")
        
        ctk.CTkLabel(sort_frame, text="Sort by:").pack(side="left", padx=(5, 2))
        
        sort_fields = [field['display_name'] for field in self.db_fields[:8]]  # Top fields only
        sort_dropdown = ctk.CTkOptionMenu(sort_frame, variable=self.sort_field, values=sort_fields,
                                         command=self._on_sort_change)
        sort_dropdown.pack(side="left", padx=2)
        if sort_fields:
            self.sort_field.set(sort_fields[0])
        
        direction_dropdown = ctk.CTkOptionMenu(sort_frame, variable=self.sort_direction, 
                                              values=["Ascending", "Descending"],
                                              command=self._on_sort_change)
        direction_dropdown.pack(side="left", padx=2)
        
        # Items per page
        ctk.CTkLabel(sort_frame, text="Show:").pack(side="left", padx=(10, 2))
        items_dropdown = ctk.CTkOptionMenu(sort_frame, variable=self.items_per_page,
                                          values=["25", "50", "100", "250", "500"],
                                          command=self._on_page_size_change)
        items_dropdown.pack(side="left", padx=2)
    
    def _create_enhanced_table(self, parent):
        """Create enhanced table with better styling and functionality."""
        table_frame = ctk.CTkFrame(parent)
        table_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        
        # Create Treeview with custom styling
        tree_container = tk.Frame(table_frame, bg=table_frame._fg_color[1])
        tree_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Define columns with priorities for display
        self.priority_columns = ['asset_no', 'asset_type', 'manufacturer', 'model', 'serial_number', 'status', 'location']
        self.all_columns = ['id'] + [field['db_name'] for field in self.db_fields]
        
        # Create treeview
        self.tree = ttk.Treeview(tree_container, columns=self.all_columns, show="headings", height=20)
        
        # Configure column headers and widths
        for col in self.all_columns:
            if col == "id":
                self.tree.heading(col, text="ID")
                self.tree.column(col, width=50, minwidth=30)
            else:
                display_name = next((f['display_name'] for f in self.db_fields if f['db_name'] == col), col.title())
                self.tree.heading(col, text=display_name, command=lambda c=col: self._sort_by_column(c))
                
                # Set intelligent column widths
                if col in ['asset_no', 'room', 'status']:
                    self.tree.column(col, width=80, minwidth=50)
                elif col in ['manufacturer', 'model', 'location', 'system_name']:
                    self.tree.column(col, width=120, minwidth=80)
                elif col == 'serial_number':
                    self.tree.column(col, width=130, minwidth=90)
                elif col == 'notes':
                    self.tree.column(col, width=200, minwidth=100)
                else:
                    self.tree.column(col, width=100, minwidth=60)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Event bindings
        self.tree.bind("<<TreeviewSelect>>", self._on_item_select)
        self.tree.bind("<Double-1>", self._on_item_double_click)
        self.tree.bind("<Button-3>", self._show_context_menu)  # Right-click context menu
        
        # Create context menu
        self._create_context_menu()
    
    def _create_details_panel(self, parent):
        """Create collapsible asset details panel."""
        self.details_frame = ctk.CTkFrame(parent)
        # Initially hidden - will be shown when asset is selected
        
        # Details header
        details_header = ctk.CTkFrame(self.details_frame)
        details_header.pack(fill="x", padx=10, pady=(10, 5))
        
        self.details_title = ctk.CTkLabel(details_header, text="Asset Details", 
                                         font=ctk.CTkFont(size=16, weight="bold"))
        self.details_title.pack(side="left")
        
        close_details_btn = ctk.CTkButton(details_header, text="✕", width=30, height=30,
                                         command=self._hide_details_panel)
        close_details_btn.pack(side="right")
        
        # Quick action buttons
        quick_actions = ctk.CTkFrame(details_header)
        quick_actions.pack(side="right", padx=(0, 10))
        
        ctk.CTkButton(quick_actions, text="View Full", width=80, height=30,
                     command=self._view_details).pack(side="left", padx=2)
        ctk.CTkButton(quick_actions, text="Edit", width=60, height=30,
                     command=self._edit_asset).pack(side="left", padx=2)
        
        # Details content
        self.details_content = ctk.CTkScrollableFrame(self.details_frame, height=200)
        self.details_content.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def _create_pagination_controls(self, parent):
        """Create pagination controls."""
        pagination_frame = ctk.CTkFrame(parent)
        pagination_frame.grid(row=3, column=0, sticky="ew", pady=(5, 10))
        
        # Previous page
        self.prev_btn = ctk.CTkButton(pagination_frame, text="← Previous", width=100,
                                     command=self._prev_page)
        self.prev_btn.pack(side="left", padx=10)
        
        # Page info
        self.page_info_label = ctk.CTkLabel(pagination_frame, text="Page 1 of 1")
        self.page_info_label.pack(side="left", expand=True)
        
        # Next page  
        self.next_btn = ctk.CTkButton(pagination_frame, text="Next →", width=100,
                                     command=self._next_page)
        self.next_btn.pack(side="right", padx=10)
    
    def _create_context_menu(self):
        """Create right-click context menu for table."""
        self.context_menu = tk.Menu(self.window, tearoff=0)
        self.context_menu.add_command(label="View Details", command=self._view_details)
        self.context_menu.add_command(label="Edit Asset", command=self._edit_asset)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Copy Asset No.", command=self._copy_asset_no)
        self.context_menu.add_command(label="Copy Serial Number", command=self._copy_serial)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Delete Asset", command=self._delete_asset)
    
    def _create_enhanced_status_bar(self):
        """Create enhanced status bar with more information."""
        self.status_frame = ctk.CTkFrame(self.window, height=40)
        self.status_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        # Left side - search/filter status
        self.status_label = ctk.CTkLabel(self.status_frame, text="Ready", anchor="w")
        self.status_label.pack(side="left", padx=10, pady=8)
        
        # Center - selection info
        self.selection_label = ctk.CTkLabel(self.status_frame, text="", anchor="center")
        self.selection_label.pack(side="left", expand=True, padx=10, pady=8)
        
        # Right side - database stats
        self.stats_label = ctk.CTkLabel(self.status_frame, text="", anchor="e")
        self.stats_label.pack(side="right", padx=10, pady=8)
        
        # Bottom tip label
        tip_frame = ctk.CTkFrame(self.window)
        tip_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        tip_label = ctk.CTkLabel(tip_frame, 
                                text="💡 Tip: Press Ctrl+Enter to Search",
                                font=ctk.CTkFont(size=10),
                                text_color="gray60")
        tip_label.pack(pady=5)
    
    # Event Handlers and Search Methods
    def _on_sort_change(self, value=None):
        """Handle sort field or direction change."""
        self._perform_search()
    
    def _on_page_size_change(self, value=None):
        """Handle items per page change."""
        self.current_page.set(1)  # Reset to first page
        self._perform_search()
    
    def _sort_by_column(self, column):
        """Sort by clicking column header."""
        # Find display name for column
        display_name = next((f['display_name'] for f in self.db_fields if f['db_name'] == column), column.title())
        
        # Toggle direction if same column, otherwise set ascending
        if self.sort_field.get() == display_name:
            current_dir = self.sort_direction.get()
            self.sort_direction.set("Descending" if current_dir == "Ascending" else "Ascending")
        else:
            self.sort_field.set(display_name)
            self.sort_direction.set("Ascending")
        
        self._perform_search()
    
    @performance_monitor("Enhanced Asset Search")
    def _perform_search(self):
        """Perform enhanced search with the new filter builder."""
        try:
            filters = self._build_search_filters()
            
            # If no filters, get all assets
            if not filters or 'conditions' not in filters:
                all_results = self.db.search_assets({}, limit=100000)
            else:
                # Apply filters with AND/OR logic
                all_results = self._apply_custom_filters(filters)
            
            self.total_count = len(all_results)
            
            # Apply sorting
            if self.sort_field.get() and self.sort_field.get() != "Select field":
                sort_column = next((f['db_name'] for f in self.db_fields 
                                  if f['display_name'] == self.sort_field.get()), None)
                if sort_column:
                    reverse = self.sort_direction.get() == "Descending"
                    all_results.sort(key=lambda x: str(x.get(sort_column, '')).lower(), reverse=reverse)
            
            # Apply pagination
            page_size = int(self.items_per_page.get())
            current_page = self.current_page.get()
            start_idx = (current_page - 1) * page_size
            end_idx = start_idx + page_size
            
            # Get current page results
            self.current_assets = all_results[start_idx:end_idx]
            self.filtered_count = len(all_results)
            
            # Update display
            self._populate_enhanced_table(self.current_assets)
            self._update_pagination_info()
            self._update_results_info()
            
            # Update status
            if filters and 'conditions' in filters:
                self.status_label.configure(text=f"Found {self.filtered_count} matching assets")
            else:
                self.status_label.configure(text="Showing all assets")
            
        except Exception as e:
            messagebox.showerror("Search Error", f"Search failed: {e}")
            self.status_label.configure(text="Search error - check logs")
            print(f"Search error: {e}")
            import traceback
            traceback.print_exc()
    
    def _apply_custom_filters(self, filters):
        """Apply custom filters with nested group logic to all assets."""
        all_assets = self.db.search_assets({}, limit=100000)
        
        if not filters or 'root' not in filters:
            return all_assets
        
        root_structure = filters['root']
        
        if not root_structure:
            return all_assets
        
        filtered_assets = []
        
        for asset in all_assets:
            # Recursively evaluate the filter structure
            if self._evaluate_group(asset, root_structure):
                filtered_assets.append(asset)
        
        return filtered_assets
    
    def _evaluate_group(self, asset, group_structure):
        """Recursively evaluate a group structure against an asset."""
        if not group_structure or not group_structure.get('conditions'):
            return True
        
        logic = group_structure.get('logic', 'AND')
        conditions = group_structure['conditions']
        
        if not conditions:
            return True
        
        # Evaluate first condition (could be a sub-group or a condition)
        first_cond = conditions[0]
        if first_cond['type'] == 'group':
            result = self._evaluate_group(asset, first_cond)
        else:
            result = self._test_condition(asset, first_cond)
        
        # Apply logic operator with remaining conditions
        for i in range(1, len(conditions)):
            cond = conditions[i]
            
            if cond['type'] == 'group':
                cond_result = self._evaluate_group(asset, cond)
            else:
                cond_result = self._test_condition(asset, cond)
            
            if logic == "AND":
                result = result and cond_result
            else:  # OR
                result = result or cond_result
        
        return result
    
    def _test_condition(self, asset, condition):
        """Test if an asset matches a single filter condition."""
        field = condition['field']
        operator = condition['operator']
        value = condition['value']
        
        # Get the raw asset value
        asset_value_raw = asset.get(field, '')
        
        # Check if this is a date comparison
        if operator in ["before", "after", "between"]:
            return self._test_date_condition(asset_value_raw, operator, value)
        
        # For text comparisons, convert to lowercase
        asset_value = str(asset_value_raw).lower()
        value = value.lower()
        
        if operator == "equals":
            return asset_value == value
        elif operator == "contains":
            return value in asset_value
        elif operator == "does not equal":
            return asset_value != value
        elif operator == "does not contain":
            return value not in asset_value
        elif operator == "starts with":
            return asset_value.startswith(value)
        elif operator == "ends with":
            return asset_value.endswith(value)
        else:
            return False
    
    def _test_date_condition(self, asset_date_str, operator, filter_value):
        """Test date-specific conditions."""
        if not asset_date_str or not filter_value:
            return False
        
        try:
            # Parse asset date
            asset_date = self._parse_date(asset_date_str)
            if not asset_date:
                return False
            
            if operator == "equals":
                filter_date = self._parse_date(filter_value)
                if not filter_date:
                    return False
                # Compare just the date part (ignore time)
                return asset_date.date() == filter_date.date()
                
            elif operator == "before":
                filter_date = self._parse_date(filter_value)
                if not filter_date:
                    return False
                return asset_date.date() < filter_date.date()
                
            elif operator == "after":
                filter_date = self._parse_date(filter_value)
                if not filter_date:
                    return False
                return asset_date.date() > filter_date.date()
                
            elif operator == "between":
                # Format expected: "MM/DD/YYYY - MM/DD/YYYY" or similar
                if ' - ' in filter_value:
                    parts = filter_value.split(' - ')
                    if len(parts) == 2:
                        start_date = self._parse_date(parts[0].strip())
                        end_date = self._parse_date(parts[1].strip())
                        if start_date and end_date:
                            return start_date.date() <= asset_date.date() <= end_date.date()
                return False
            
            return False
            
        except Exception as e:
            print(f"Error testing date condition: {e}")
            return False
    
    def _parse_date(self, date_str):
        """Parse a date string in various formats."""
        if not date_str:
            return None
        
        # Try different date formats
        formats = [
            "%m/%d/%Y",      # MM/DD/YYYY (DatePicker format)
            "%m/%#d/%Y",     # MM/D/YYYY (Windows format without leading zero)
            "%Y-%m-%d",      # ISO format
            "%Y-%m-%dT%H:%M:%S",  # ISO with time
            "%Y-%m-%d %H:%M:%S",  # ISO with space
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(str(date_str).strip(), fmt)
            except ValueError:
                continue
        
        # Try parsing ISO format with timezone
        try:
            return datetime.fromisoformat(str(date_str).replace('Z', '+00:00'))
        except (ValueError, AttributeError):
            pass
        
        return None
    
    def _build_search_filters(self):
        """Build comprehensive filter dictionary from the nested group structure."""
        # Build the filter structure recursively from the root group
        root_structure = self._build_group_structure(self.root_group)
        
        if not root_structure:
            return {}
        
        return {'root': root_structure}
    
    def _build_group_structure(self, group):
        """Recursively build the filter structure for a group."""
        if not group['items']:
            return None
        
        structure = {
            'type': 'group',
            'logic': group['logic_var'].get(),
            'conditions': []
        }
        
        for item in group['items']:
            if item.get('is_group'):
                # Recursively build sub-group
                sub_structure = self._build_group_structure(item)
                if sub_structure:
                    structure['conditions'].append(sub_structure)
            else:
                # Build filter condition
                field_display = item['field_var'].get()
                operator = item['operator_var'].get()
                value = item['value_var'].get().strip()
                
                # For "between" operator, combine start and end dates
                if operator == "between" and 'value_var_end' in item:
                    end_value = item['value_var_end'].get().strip()
                    if value and end_value:
                        value = f"{value} - {end_value}"
                    elif not value:
                        continue
                
                # Skip if no value or field not selected
                if not value or field_display == "Select Field..." or not field_display:
                    continue
                
                # Convert display name to database field name
                db_field = next((f['db_name'] for f in self.db_fields 
                               if f['display_name'] == field_display), None)
                
                if db_field:
                    structure['conditions'].append({
                        'type': 'condition',
                        'field': db_field,
                        'operator': operator,
                        'value': value
                    })
        
        # Return None if no valid conditions
        if not structure['conditions']:
            return None
        
        return structure
    
    def _do_search(self):
        """Perform search with filters."""
        # Get custom filters from the filter builder
        custom_filters = self._build_search_filters()
        
        print(f"DEBUG: custom_filters = {custom_filters}")  # Debug output
        
        try:
            if custom_filters:
                # Apply custom filters
                filtered_assets = self._apply_custom_filters(custom_filters)
                print(f"DEBUG: Filtered {len(filtered_assets)} assets from total")  # Debug output
                
                # Store filtered results
                self.current_assets = filtered_assets
                self._populate_enhanced_table(filtered_assets)
                self.filtered_count = len(filtered_assets)
                self.status_label.configure(text=f"Found {self.filtered_count} matching assets")
            else:
                # No filters, show all
                all_assets = self.db.search_assets({}, limit=100000)
                # Store all assets
                self.current_assets = all_assets
                self._populate_enhanced_table(all_assets)
                self.filtered_count = len(all_assets)
                self.status_label.configure(text=f"Showing all {self.filtered_count} assets")
            
        except Exception as e:
            messagebox.showerror("Search Error", f"Search failed: {e}")
            self.status_label.configure(text="Search error - check logs")
            print(f"Search error: {e}")
            import traceback
            traceback.print_exc()
        
    def _populate_enhanced_table(self, assets):
        """Populate table with enhanced display and formatting."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # If no assets, show helpful message
        if not assets:
            if hasattr(self, 'tree') and hasattr(self, 'all_columns'):
                # Insert a help message as a single row
                help_values = [""] * len(self.all_columns)
                if len(help_values) > 1:
                    help_values[1] = "Enter search criteria and click the Search button to find assets"
                self.tree.insert("", "end", values=help_values)
            return
        
        # Add new items with enhanced formatting
        for asset in assets:
            values = []
            for col_name in self.all_columns:
                value = asset.get(col_name, "")
                if value is None:
                    value = ""
                
                # Enhanced formatting for specific fields
                if col_name == 'status':
                    value = str(value).title()
                elif col_name in ['asset_no', 'serial_number'] and value:
                    value = str(value).upper()
                elif col_name == 'notes' and isinstance(value, str) and len(value) > 100:
                    value = value[:97] + "..."
                
                values.append(str(value))
            
            # Insert with item ID for easy retrieval
            item_id = self.tree.insert("", "end", values=values)
            self.tree.set(item_id, "id", asset.get("id", ""))
    
    def _update_results_info(self):
        """Update results information display."""
        if hasattr(self, 'filtered_count') and hasattr(self, 'total_db_count'):
            if self.filtered_count == self.total_db_count:
                info_text = f"Showing {len(self.current_assets)} of {self.total_db_count} total assets"
            else:
                info_text = f"Showing {len(self.current_assets)} of {self.filtered_count} filtered assets ({self.total_db_count} total)"
        else:
            info_text = f"Showing {len(self.current_assets)} assets"
        
        self.results_info_label.configure(text=info_text)
    
    def _update_pagination_info(self):
        """Update pagination controls and information."""
        page_size = int(self.items_per_page.get())
        total_pages = max(1, (self.filtered_count + page_size - 1) // page_size)
        current_page = self.current_page.get()
        
        # Update page info
        self.page_info_label.configure(text=f"Page {current_page} of {total_pages}")
        
        # Enable/disable navigation buttons
        self.prev_btn.configure(state="normal" if current_page > 1 else "disabled")
        self.next_btn.configure(state="normal" if current_page < total_pages else "disabled")
    
    def _prev_page(self):
        """Go to previous page."""
        if self.current_page.get() > 1:
            self.current_page.set(self.current_page.get() - 1)
            self._perform_search()
    
    def _next_page(self):
        """Go to next page."""
        page_size = int(self.items_per_page.get())
        total_pages = max(1, (self.filtered_count + page_size - 1) // page_size)
        if self.current_page.get() < total_pages:
            self.current_page.set(self.current_page.get() + 1)
            self._perform_search()
    
    def _show_details_panel(self, asset):
        """Show asset details in side panel."""
        if not self.details_frame.winfo_ismapped():
            self.details_frame.grid(row=0, column=1, rowspan=4, sticky="nsew", padx=(5, 0))
            self.content_frame.grid_columnconfigure(1, weight=0, minsize=300)
            
            # Force UI update to ensure the panel is fully rendered
            self.details_frame.update_idletasks()
        
        # Clear previous details
        for widget in self.details_content.winfo_children():
            widget.destroy()
        
        # Ensure the details_content is properly updated after clearing
        self.details_content.update_idletasks()
        
        # Update title
        asset_id = asset.get('asset_no', 'Unknown')
        self.details_title.configure(text=f"Asset: {asset_id}")
        
        # Add key details
        key_fields = ['asset_type', 'manufacturer', 'model', 'serial_number', 'status', 'location', 'notes']
        
        for field in key_fields:
            if field in asset and asset[field]:
                # Find display name
                display_name = next((f['display_name'] for f in self.db_fields if f['db_name'] == field), field.title())
                
                detail_frame = ctk.CTkFrame(self.details_content)
                detail_frame.pack(fill="x", pady=2)
                
                label = ctk.CTkLabel(detail_frame, text=f"{display_name}:", 
                                   font=ctk.CTkFont(weight="bold"), anchor="nw")
                label.pack(anchor="nw", padx=5, pady=2)
                
                value = str(asset[field])
                
                # Determine if this is a multiline field (like Notes)
                is_multiline = field == 'notes' or '\n' in value or len(value) > 200
                
                if is_multiline:
                    # Use textbox for multiline/long content
                    line_count = value.count('\n') + 1
                    textbox_height = min(max(40, line_count * 20), 120)
                    
                    value_textbox = ctk.CTkTextbox(detail_frame, 
                                                  height=textbox_height,
                                                  wrap="word",
                                                  activate_scrollbars=True,
                                                  fg_color=("gray90", "gray20"),
                                                  corner_radius=6)
                    value_textbox.insert("0.0", value)
                    value_textbox.configure(state="disabled")  # Read-only but selectable
                    value_textbox.pack(fill="x", padx=5, pady=(0, 5))
                else:
                    # Use textbox for single-line content (selectable)
                    value_textbox = ctk.CTkTextbox(detail_frame, 
                                                  height=30,
                                                  wrap="word",
                                                  activate_scrollbars=False,
                                                  fg_color=("gray90", "gray20"),
                                                  corner_radius=6)
                    value_textbox.insert("0.0", value)
                    value_textbox.configure(state="disabled")  # Read-only but selectable
                    value_textbox.pack(fill="x", padx=5, pady=(0, 5))
    
    def _hide_details_panel(self):
        """Hide asset details panel."""
        self.details_frame.grid_forget()
        self.content_frame.grid_columnconfigure(1, weight=0, minsize=0)
    
    def _initialize_empty_state(self):
        """Initialize the interface with empty state - no data loaded."""
        try:
            # Get basic database stats without loading all assets
            sample_assets = self.db.search_assets({}, limit=1)
            if sample_assets:
                # Count total assets efficiently
                all_assets = self.db.search_assets({}, limit=10000)
                self.total_db_count = len(all_assets)
            else:
                self.total_db_count = 0
            
            # Initialize empty state
            self.current_assets = []
            self.filtered_count = 0
            
            # Populate filter dropdowns
            self._populate_filter_dropdowns()
            
            # Clear table and show welcome message
            self._populate_enhanced_table([])
            
            # Update status
            self.status_label.configure(text="Enter search criteria and click Search to find assets")
            self.results_info_label.configure(text="Click Search to load assets")
            self.stats_label.configure(text=f"Total: {self.total_db_count} assets in database")
            
            # Reset pagination
            self.current_page.set(1)
            self._update_pagination_info()
            
        except Exception as e:
            error_msg = f"Failed to initialize: {e}"
            messagebox.showerror("Initialization Error", error_msg)
            self.status_label.configure(text="Initialization error - check logs")
    
    def _load_initial_data(self):
        """Load initial data and setup interface."""
        try:
            # Get database field information
            if hasattr(self.db, 'get_table_fields'):
                self.db_fields = self.db.get_table_fields()
            else:
                # Fallback to getting fields from a sample asset
                sample_assets = self.db.search_assets({}, limit=1)
                if sample_assets:
                    sample_asset = sample_assets[0]
                    self.db_fields = [
                        {'db_name': key, 'display_name': key.replace('_', ' ').title()}
                        for key in sample_asset.keys()
                        if key != 'id'
                    ]
                else:
                    self.db_fields = []
            
            # Prioritize important fields
            priority_fields = ['asset_no', 'asset_type', 'manufacturer', 'model', 'serial_number', 'status', 'location']
            prioritized_fields = []
            remaining_fields = []
            
            for field in self.db_fields:
                if field['db_name'] in priority_fields:
                    prioritized_fields.append(field)
                else:
                    remaining_fields.append(field)
            
            # Sort priority fields by the order in priority_fields
            prioritized_fields.sort(key=lambda x: priority_fields.index(x['db_name']) 
                                  if x['db_name'] in priority_fields else len(priority_fields))
            
            self.db_fields = prioritized_fields + remaining_fields
            
            # Get total count for statistics
            all_assets = self.db.search_assets({}, limit=10000)
            self.total_db_count = len(all_assets)
            
            # Initial search (load all assets)
            self.current_assets = all_assets[:int(self.items_per_page.get())]
            self.filtered_count = len(all_assets)
            
            # Populate interface elements
            self._populate_filter_dropdowns()
            self._populate_enhanced_table(self.current_assets)
            self._update_results_info()
            self._update_pagination_info()
            self._update_database_stats()
            
            self.status_label.configure(text="Ready")
            
        except Exception as e:
            error_msg = f"Failed to load data: {e}"
            messagebox.showerror("Load Error", error_msg)
            self.status_label.configure(text="Load error - check logs")
    
    def _populate_filter_dropdowns(self):
        """Populate quick filter dropdowns with unique values."""
        try:
            # Note: SearchableDropdown values are set during creation in _create_quick_filters_section
            # This method is kept for compatibility but doesn't need to update the dropdowns
            # since they get their values from self.unique_values during initialization
            pass
                
        except Exception as e:
            print(f"Warning: Could not populate filter dropdowns: {e}")
    
    def _update_database_stats(self):
        """Update database statistics in status bar."""
        try:
            stats_text = f"Total: {self.total_db_count} assets"
            if hasattr(self, 'filtered_count') and self.filtered_count != self.total_db_count:
                stats_text += f" | Filtered: {self.filtered_count}"
            self.stats_label.configure(text=stats_text)
        except Exception:
            self.stats_label.configure(text="Stats unavailable")
    
    def _save_current_search(self):
        """Save current search parameters to config."""
        try:
            # Build current filter structure
            filter_structure = self._build_search_filters()
            
            if not filter_structure or not filter_structure.get('root'):
                messagebox.showwarning("No Filters", "Please add at least one filter before saving.")
                return
            
            # Prompt for search name
            from tkinter import simpledialog
            search_name = simpledialog.askstring(
                "Save Search", 
                "Enter a name for this saved search:",
                parent=self.window
            )
            
            if not search_name:
                return  # User cancelled
            
            # Get current saved searches from config
            saved_searches = self.config.saved_searches if hasattr(self.config, 'saved_searches') else {}
            if saved_searches is None:
                saved_searches = {}
            
            # Save the filter structure
            saved_searches[search_name] = {
                'filter_structure': filter_structure,
                'description': f"Saved on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            }
            
            # Update config
            self.config.saved_searches = saved_searches
            self.config_manager.save_config(self.config)
            
            # Refresh the saved searches list
            self._refresh_saved_searches_list()
            
            messagebox.showinfo("Success", f"Search '{search_name}' saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save search: {e}")
            print(f"Error saving search: {e}")
            import traceback
            traceback.print_exc()
    
    def _load_saved_search(self):
        """Load a saved search from config."""
        try:
            # Get selected search
            selection = self.saved_searches_listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a saved search to load.")
                return
            
            search_name = self.saved_searches_listbox.get(selection[0])
            
            # Get saved searches from config
            saved_searches = self.config.saved_searches if hasattr(self.config, 'saved_searches') else {}
            
            if search_name not in saved_searches:
                messagebox.showerror("Error", f"Saved search '{search_name}' not found.")
                return
            
            search_data = saved_searches[search_name]
            filter_structure = search_data.get('filter_structure')
            
            if not filter_structure:
                messagebox.showerror("Error", "Invalid saved search data.")
                return
            
            # Clear current filters
            self._clear_all_filters()
            
            # Rebuild filters from saved structure
            self._rebuild_filters_from_structure(filter_structure)
            
            messagebox.showinfo("Success", f"Loaded search '{search_name}'!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load search: {e}")
            print(f"Error loading search: {e}")
            import traceback
            traceback.print_exc()
    
    def _delete_saved_search(self):
        """Delete a saved search from config."""
        try:
            # Get selected search
            selection = self.saved_searches_listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a saved search to delete.")
                return
            
            search_name = self.saved_searches_listbox.get(selection[0])
            
            # Confirm deletion
            from tkinter import messagebox as mb
            if not mb.askyesno("Confirm Delete", f"Are you sure you want to delete '{search_name}'?"):
                return
            
            # Get saved searches from config
            saved_searches = self.config.saved_searches if hasattr(self.config, 'saved_searches') else {}
            
            if search_name in saved_searches:
                del saved_searches[search_name]
                
                # Update config
                self.config.saved_searches = saved_searches
                self.config_manager.save_config(self.config)
                
                # Refresh the list
                self._refresh_saved_searches_list()
                
                messagebox.showinfo("Success", f"Deleted search '{search_name}'!")
            else:
                messagebox.showerror("Error", f"Saved search '{search_name}' not found.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete search: {e}")
            print(f"Error deleting search: {e}")
            import traceback
            traceback.print_exc()
    
    def _refresh_saved_searches_list(self):
        """Refresh the saved searches listbox with current saved searches."""
        try:
            self.saved_searches_listbox.delete(0, tk.END)
            
            saved_searches = self.config.saved_searches if hasattr(self.config, 'saved_searches') else {}
            if saved_searches:
                for search_name in sorted(saved_searches.keys()):
                    self.saved_searches_listbox.insert(tk.END, search_name)
                    
        except Exception as e:
            print(f"Error refreshing saved searches list: {e}")
    
    def _rebuild_filters_from_structure(self, filter_structure):
        """Rebuild filter UI from saved structure."""
        try:
            if 'root' not in filter_structure:
                return
            
            root_structure = filter_structure['root']
            
            # Set root logic
            if 'logic' in root_structure:
                self.root_group['logic_var'].set(root_structure['logic'])
            
            # Recursively rebuild groups and filters
            self._rebuild_group_items(self.root_group, root_structure.get('conditions', []))
            
        except Exception as e:
            print(f"Error rebuilding filters: {e}")
            import traceback
            traceback.print_exc()
    
    def _rebuild_group_items(self, parent_group, conditions):
        """Recursively rebuild groups and filters from conditions list."""
        for condition in conditions:
            if condition['type'] == 'group':
                # Create a new group
                new_group = self._add_group(parent_group, logic=condition.get('logic', 'AND'))
                
                # Recursively add items to this group
                self._rebuild_group_items(new_group, condition.get('conditions', []))
                
            else:  # condition
                # Convert db field name back to display name
                db_field = condition['field']
                display_name = next((f['display_name'] for f in self.db_fields 
                                   if f['db_name'] == db_field), db_field)
                
                operator = condition.get('operator', 'contains')
                value = condition.get('value', '')
                
                # Add filter WITHOUT value initially - this allows field change to complete
                filter_data = self._add_filter_row(
                    parent_group,
                    field=display_name,
                    operator=operator,
                    value=None  # Don't set value yet
                )
                
                # Now set the value AFTER the field change has processed
                # This ensures the correct widget type (entry, datepicker, dropdown) exists
                self.window.after(100, lambda fd=filter_data, v=value: self._set_filter_value(fd, v))
    
    def _set_filter_value(self, filter_data, value):
        """Set the value for a filter after the field change has been processed."""
        try:
            if not value:
                return
                
            # For "between" operator with date fields, handle dual values
            if (filter_data.get('is_date_field') and 
                filter_data['operator_var'].get() == 'between' and 
                ' - ' in value):
                # Split the date range
                parts = value.split(' - ')
                if len(parts) == 2:
                    filter_data['value_var'].set(parts[0].strip())
                    if 'value_var_end' in filter_data:
                        filter_data['value_var_end'].set(parts[1].strip())
            else:
                # Normal single value
                filter_data['value_var'].set(value)
                
        except Exception as e:
            print(f"Error setting filter value: {e}")

    
    def _export_current_results(self):
        """Export current filtered results."""
        try:
            if not hasattr(self, 'current_assets') or not self.current_assets:
                messagebox.showwarning("No Data", "No assets to export.")
                return
            
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Current Results"
            )
            
            if filename:
                import csv
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    if self.current_assets:
                        fieldnames = self.current_assets[0].keys()
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(self.current_assets)
                
                messagebox.showinfo("Export Complete", f"Exported {len(self.current_assets)} assets to {filename}")
                self.status_label.configure(text=f"Exported {len(self.current_assets)} assets")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data: {e}")
    
    def _request_labels_for_filtered(self):
        """Request labels for all assets in the current filtered results."""
        try:
            if not hasattr(self, 'current_assets') or not self.current_assets:
                messagebox.showwarning("No Assets", "No assets to request labels for.\n\nPlease perform a search first.")
                return
            
            asset_count = len(self.current_assets)
            
            # Confirm with user
            confirm = messagebox.askyesno(
                "Confirm Label Request",
                f"Request labels for {asset_count} asset{'s' if asset_count != 1 else ''}?\n\n"
                f"This will update the 'Label Requested Date' field for all assets in the current filtered results."
            )
            
            if not confirm:
                return
            
            # Process each asset
            success_count = 0
            failed_count = 0
            
            for asset in self.current_assets:
                asset_id = asset.get('id')
                if asset_id:
                    try:
                        if self.db.request_label(asset_id):
                            success_count += 1
                        else:
                            failed_count += 1
                    except Exception as e:
                        error_handler.logger.warning(f"Failed to request label for asset {asset_id}: {e}")
                        failed_count += 1
                else:
                    failed_count += 1
            
            # Show results
            if failed_count == 0:
                messagebox.showinfo(
                    "Success",
                    f"Successfully requested labels for all {success_count} asset{'s' if success_count != 1 else ''}!"
                )
                self.status_label.configure(text=f"Requested labels for {success_count} assets")
            else:
                messagebox.showwarning(
                    "Partial Success",
                    f"Requested labels for {success_count} asset{'s' if success_count != 1 else ''}.\n"
                    f"Failed for {failed_count} asset{'s' if failed_count != 1 else ''}."
                )
                self.status_label.configure(text=f"Requested labels: {success_count} succeeded, {failed_count} failed")
            
            # Refresh the display to show updated dates
            self._do_search()
            
        except Exception as e:
            error_handler.handle_exception(e, "Failed to request labels for filtered assets")
            messagebox.showerror("Error", f"Failed to request labels: {e}")
    
    def _export_all_assets(self):
        """Export all assets in database."""
        try:
            # Get all assets without filters
            all_assets = self.db.search_assets({}, limit=10000)
            
            if not all_assets:
                messagebox.showwarning("No Data", "No assets found in database.")
                return
            
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export All Assets"
            )
            
            if filename:
                import csv
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    fieldnames = all_assets[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(all_assets)
                
                messagebox.showinfo("Export Complete", f"Exported {len(all_assets)} total assets to {filename}")
                self.status_label.configure(text=f"Exported {len(all_assets)} total assets")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export all assets: {e}")
    
    
    # Event Handlers
    def _on_item_select(self, event=None):
        """Handle item selection in the table."""
        selected_items = self.tree.selection()
        if selected_items:
            # Update selection info
            count = len(selected_items)
            self.selection_label.configure(text=f"{count} asset{'s' if count != 1 else ''} selected")
            
            # Show details for single selection
            if count == 1:
                selected_asset = self._get_selected_asset()
                if selected_asset:
                    self._show_details_panel(selected_asset)
        else:
            self.selection_label.configure(text="")
            self._hide_details_panel()
    
    def _on_item_double_click(self, event=None):
        """Handle double-click on table item."""
        self._view_details()
    
    def _get_selected_asset(self):
        """Get the currently selected asset data."""
        selection = self.tree.selection()
        if not selection:
            return None
        
        item = selection[0]
        values = self.tree.item(item)['values']
        
        if not values:
            return None
        
        # Build asset dictionary from values
        asset = {}
        for i, col_name in enumerate(self.all_columns):
            if i < len(values):
                asset[col_name] = values[i]
        
        return asset
    
    def _copy_asset_no(self):
        """Copy selected asset number to clipboard."""
        selected = self._get_selected_asset()
        if selected:
            asset_no = selected.get('asset_no', '')
            if asset_no:
                self.window.clipboard_clear()
                self.window.clipboard_append(str(asset_no))
                self.status_label.configure(text=f"Copied asset number: {asset_no}")
    
    def _copy_serial(self):
        """Copy selected asset serial number to clipboard."""
        selected = self._get_selected_asset()
        if selected:
            serial = selected.get('serial_number', '')
            if serial:
                self.window.clipboard_clear()
                self.window.clipboard_append(str(serial))
                self.status_label.configure(text=f"Copied serial number: {serial}")
    
    def _show_context_menu(self, event):
        """Show context menu on right-click."""
        # Select the item under the cursor
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def _export_filtered_results(self):
        """Export current filtered results to CSV."""
        try:
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export Filtered Results"
            )
            
            if filename:
                # Use existing export functionality with current filtered results
                if hasattr(self, 'export_service'):
                    self.export_service.export_to_csv(self.current_assets, filename)
                else:
                    # Fallback implementation
                    import csv
                    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                        if self.current_assets:
                            fieldnames = self.current_assets[0].keys()
                            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                            writer.writeheader()
                            writer.writerows(self.current_assets)
                
                messagebox.showinfo("Export Complete", f"Exported {len(self.current_assets)} assets to {filename}")
                self.status_label.configure(text=f"Exported {len(self.current_assets)} assets")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data: {e}")
    
    def _view_details(self):
        """View detailed information about selected asset."""
        selected_asset = self._get_selected_asset()
        if not selected_asset:
            messagebox.showwarning("No Selection", "Please select an asset to view details.")
            return
        
        try:
            from ui_components import AssetDetailWindow
            
            # Define callback to refresh data after asset edit
            def on_asset_edited():
                self._load_initial_data()  # Refresh the asset list
            
            AssetDetailWindow(self.window, selected_asset, on_edit_callback=on_asset_edited)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open asset details: {e}")
    
    def _edit_asset(self):
        """Edit selected asset."""
        selected_asset = self._get_selected_asset()
        if not selected_asset:
            messagebox.showwarning("No Selection", "Please select an asset to edit.")
            return
        
        try:
            from edit_asset import open_edit_asset_window
            
            asset_id = selected_asset.get('id')
            if not asset_id:
                messagebox.showerror("Error", "Cannot edit asset: Asset ID not found.")
                return
            
            # Define callback to refresh data after edit
            def on_asset_updated():
                self._load_initial_data()  # Refresh the asset list
            
            # Open edit window
            open_edit_asset_window(self.window, asset_id, on_update_callback=on_asset_updated)
            
        except ImportError as e:
            messagebox.showerror("Error", f"Could not open edit window: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    
    def _delete_asset(self):
        """Delete selected asset."""
        selected_asset = self._get_selected_asset()
        if not selected_asset:
            messagebox.showwarning("No Selection", "Please select an asset to delete.")
            return
        
        result = messagebox.askyesno("Confirm Delete", 
                                   f"Are you sure you want to delete asset {selected_asset.get('asset_no', 'Unknown')}?\n"
                                   f"This action cannot be undone.", parent=self.window)
        
        if result:
            try:
                success = self.db.delete_asset(selected_asset['id'])
                if success:
                    messagebox.showinfo("Success", "Asset deleted successfully.")
                    self._load_initial_data()  # Refresh the list
                else:
                    messagebox.showerror("Error", "Failed to delete asset.")
            except Exception as e:
                messagebox.showerror("Delete Error", f"Failed to delete asset: {e}")
    
    def _on_closing(self):
        """Handle window closing."""
        self.window.destroy()


# Keyboard shortcuts handlers
def _setup_keyboard_shortcuts(self):
    """Setup keyboard shortcuts for enhanced productivity."""
    self.window.bind("<Control-f>", lambda e: self.search_entry.focus())
    self.window.bind("<F5>", lambda e: self._load_initial_data())
    self.window.bind("<Control-r>", lambda e: self._load_initial_data())
    self.window.bind("<Return>", lambda e: self._view_details())
    self.window.bind("<Control-e>", lambda e: self._edit_asset())
    self.window.bind("<Escape>", lambda e: self._clear_all_filters())


if __name__ == "__main__":
    # Test the enhanced browse window
    root = ctk.CTk()
    ctk.set_appearance_mode("dark")
    BrowseAssetsWindow(root)
    root.mainloop()
