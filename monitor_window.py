"""
Real-time monitor window for viewing recent asset database changes.
Shows a live feed of recently added or modified assets.
"""

import customtkinter as ctk
from tkinter import ttk
import threading
import time
import os
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Union
from asset_database import AssetDatabase
from config_manager import ConfigManager
from error_handling import error_handler, safe_execute
from performance_monitoring import performance_monitor
from ui_components import AssetDetailWindow


class MonitorWindow:
    """Real-time monitor window for database changes.
    
    Features:
    - Shows recent asset additions and modifications
    - Auto-refreshes every few seconds
    - Can run alongside other windows
    - Tall, narrow layout optimized for monitoring
    - Configurable number of items to display
    """
    
    def __init__(self, parent=None):
        self.parent = parent
        
        # Use centralized configuration manager
        self.config_manager = ConfigManager()
        self.config = self.config_manager.get_config()
        
        # Create database instance
        self.db = AssetDatabase(self.config.database_path)
        
        # Monitor settings
        self.max_items = 10  # Maximum number of items to display
        self.refresh_interval = 5  # Seconds between auto-refresh
        self.auto_refresh_enabled = True
        self.days_filter = 0.5  # Default to 1 day (today) for statistics filtering
        
        # Cache for reducing flicker
        self.last_assets_data = []
        self.asset_widgets = []  # Keep track of created widgets
        
        # Threading for auto-refresh
        self.refresh_thread = None
        self.stop_refresh = threading.Event()
        
        # Window setup
        self.window = ctk.CTkToplevel(parent) if parent else ctk.CTk()
        self.window.title("Asset Monitor")
        
        # Position window in upper right corner
        self.window.update_idletasks()  # Ensure window exists for positioning
        
        # Get screen dimensions - always use primary screen for reliable positioning
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # # Get virtual screen info for debugging
        # try:
        #     virtual_width = self.window.winfo_vrootwidth()
        #     virtual_height = self.window.winfo_vrootheight()
        #     print(f"Primary screen: {screen_width}x{screen_height}")
        #     print(f"Virtual screen (all monitors): {virtual_width}x{virtual_height}")
        # except Exception as e:
        #     print(f"Could not get virtual screen info: {e}")
        #     virtual_width = screen_width
        
        window_width = 400
        # Make window span full height of screen with small margins
        top_margin = 20  # Small margin from top
        bottom_margin = 100  # Small margin from bottom (for taskbar)
        window_height = screen_height - top_margin - bottom_margin
        
        # Always position on primary screen to ensure visibility
        # Position flush against the right edge of the PRIMARY screen
        pos_x = screen_width - window_width
        pos_y = top_margin  # Start at top margin
        
        # print(f"Positioning on primary screen: x={pos_x}, y={pos_y}")
        # print(f"Window dimensions: {window_width}x{window_height}")
        # print(f"Window will span x={pos_x} to x={pos_x + window_width} on primary screen (width: {screen_width})")
        # print(f"Window will span y={pos_y} to y={pos_y + window_height} on primary screen (height: {screen_height})")
        
        self.window.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
        self.window.minsize(400, 700)  # Keep minimum size requirement
        
        # Build UI
        self._build_layout()
        self._load_statistics()
        self._load_recent_changes()
        self._start_auto_refresh()
        
        # Set up window close handler
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Ensure non-modal behavior - allow other windows to remain active
        # self.window.grab_set()  # Commented out to allow simultaneous windows
        
        # Ensure window independence from parent for true non-modal behavior
        if parent:
            self.window.transient()  # Remove transient relationship to parent
    
    def _build_layout(self):
        """Build the monitor window layout."""
        # Statistics section at the top
        stats_frame = ctk.CTkFrame(self.window)
        stats_frame.pack(fill="x", padx=10, pady=10)
        
        # Header row with Statistics title and Days filter
        header_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(5, 2))
        header_frame.columnconfigure(0, weight=1)  # Title takes available space
        header_frame.columnconfigure(1, weight=0)  # Filter stays right
        
        # Statistics title (left side)
        stats_title = ctk.CTkLabel(header_frame, text="Statistics", 
                                  font=ctk.CTkFont(size=17, weight="bold"))
        stats_title.grid(row=0, column=0, sticky="w")
        
        # Days filter (right side)
        filter_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        filter_frame.grid(row=0, column=1, sticky="e")
        
        filter_label = ctk.CTkLabel(filter_frame, text="Within", 
                                   font=ctk.CTkFont(size=12))
        filter_label.grid(row=0, column=0, padx=(0, 2))
        
        self.days_filter_combo = ctk.CTkComboBox(filter_frame, 
                                               values=["12 hours", "1 day", "1 week", "2 weeks", "1 month", "2 months", "3 months", "1 year", "All"],
                                               width=80,
                                               height=26,
                                               font=ctk.CTkFont(size=12),
                                               command=self._on_days_filter_changed)
        self.days_filter_combo.set("12 hours")  # Default to 1 day (today)
        self.days_filter_combo.grid(row=0, column=1, padx=2)
        
        # Remove the "days" label since the dropdown now includes units
        # days_label = ctk.CTkLabel(filter_frame, text="days", 
        #                          font=ctk.CTkFont(size=12))
        # days_label.grid(row=0, column=2, padx=(2, 0))
        
        # Thin horizontal line under Statistics header
        stats_divider = ctk.CTkFrame(stats_frame, height=1, fg_color=("gray70", "gray30"))
        stats_divider.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=(2, 8))
        
        # Configure grid for card layout - 3 columns for Room, Cube, Overall
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.columnconfigure(1, weight=1)
        stats_frame.columnconfigure(2, weight=1)
        
        # Room card (left)
        room_card = ctk.CTkFrame(stats_frame)
        room_card.grid(row=2, column=0, padx=3, pady=5, sticky="ew")
        
        self.room_header_label = ctk.CTkLabel(room_card, text="Room: Loading...", 
                                             font=ctk.CTkFont(size=12, weight="bold"))
        self.room_header_label.pack(pady=(5, 2))
        
        # Horizontal line under Room header
        room_divider = ctk.CTkFrame(room_card, height=1, fg_color=("gray70", "gray30"))
        room_divider.pack(fill="x", padx=8, pady=(0, 3))
        
        self.room_added_label = ctk.CTkLabel(room_card, text="Added: 0", 
                                            font=ctk.CTkFont(size=11))
        self.room_added_label.pack(pady=1)
        
        self.room_modified_label = ctk.CTkLabel(room_card, text="Modified: 0", 
                                               font=ctk.CTkFont(size=11))
        self.room_modified_label.pack(pady=1)
        
        self.room_total_label = ctk.CTkLabel(room_card, text="Total: 0", 
                                            font=ctk.CTkFont(size=11, weight="bold"))
        self.room_total_label.pack(pady=(1, 5))
        
        # Cube card (center)
        cube_card = ctk.CTkFrame(stats_frame)
        cube_card.grid(row=2, column=1, padx=3, pady=5, sticky="ew")
        
        self.cube_header_label = ctk.CTkLabel(cube_card, text="Cube: Loading...", 
                                             font=ctk.CTkFont(size=12, weight="bold"))
        self.cube_header_label.pack(pady=(5, 2))
        
        # Horizontal line under Cube header
        cube_divider = ctk.CTkFrame(cube_card, height=1, fg_color=("gray70", "gray30"))
        cube_divider.pack(fill="x", padx=8, pady=(0, 3))
        
        self.cube_added_label = ctk.CTkLabel(cube_card, text="Added: 0", 
                                            font=ctk.CTkFont(size=11))
        self.cube_added_label.pack(pady=1)
        
        self.cube_modified_label = ctk.CTkLabel(cube_card, text="Modified: 0", 
                                               font=ctk.CTkFont(size=11))
        self.cube_modified_label.pack(pady=1)
        
        self.cube_total_label = ctk.CTkLabel(cube_card, text="Total: 0", 
                                            font=ctk.CTkFont(size=11, weight="bold"))
        self.cube_total_label.pack(pady=(1, 5))
        
        # Overall card (right)
        overall_card = ctk.CTkFrame(stats_frame)
        overall_card.grid(row=2, column=2, padx=3, pady=5, sticky="ew")
        
        overall_header = ctk.CTkLabel(overall_card, text="Overall", 
                                     font=ctk.CTkFont(size=12, weight="bold"))
        overall_header.pack(pady=(5, 2))
        
        # Horizontal line under Overall header
        overall_divider = ctk.CTkFrame(overall_card, height=1, fg_color=("gray70", "gray30"))
        overall_divider.pack(fill="x", padx=8, pady=(0, 3))
        
        self.overall_added_label = ctk.CTkLabel(overall_card, text="Added: 0", 
                                               font=ctk.CTkFont(size=11))
        self.overall_added_label.pack(pady=1)
        
        self.overall_modified_label = ctk.CTkLabel(overall_card, text="Modified: 0", 
                                                  font=ctk.CTkFont(size=11))
        self.overall_modified_label.pack(pady=1)
        
        self.overall_total_label = ctk.CTkLabel(overall_card, text="Total: 0", 
                                               font=ctk.CTkFont(size=11, weight="bold"))
        self.overall_total_label.pack(pady=(1, 5))
        

        
        # Title frame - centered "Recent Asset Changes"
        title_frame = ctk.CTkFrame(self.window)
        title_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        title_label = ctk.CTkLabel(title_frame, text="Recent Asset Changes", 
                                  font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(pady=8)
        
        # Control buttons frame - second row
        controls_frame = ctk.CTkFrame(self.window)
        controls_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        # Create a centered container for the controls
        center_container = ctk.CTkFrame(controls_frame, fg_color="transparent")
        center_container.pack(expand=True)
        
        # Manual refresh button
        refresh_btn = ctk.CTkButton(center_container, text="Refresh", width=80,
                                   command=self._refresh_data)
        refresh_btn.pack(side="left", padx=5, pady=5)
        
        # Database reload button (for when database config changes)
        reload_db_btn = ctk.CTkButton(center_container, text="Reload DB", width=90,
                                     command=self.reload_configuration,
                                     fg_color="orange", hover_color="darkorange")
        reload_db_btn.pack(side="left", padx=5, pady=5)
        
        # Auto-refresh toggle
        self.auto_refresh_var = ctk.BooleanVar(value=self.auto_refresh_enabled)
        auto_refresh_cb = ctk.CTkCheckBox(center_container, text="Auto-refresh", 
                                         variable=self.auto_refresh_var,
                                         command=self._toggle_auto_refresh)
        auto_refresh_cb.pack(side="left", padx=5, pady=5)
        
        # Smooth refresh toggle to reduce flicker
        self.smooth_refresh_var = ctk.BooleanVar(value=True)
        smooth_refresh_cb = ctk.CTkCheckBox(center_container, text="Smooth", 
                                           variable=self.smooth_refresh_var)
        smooth_refresh_cb.pack(side="left", padx=5, pady=5)
        
        # Settings frame
        settings_frame = ctk.CTkFrame(self.window)
        settings_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        # Max items selector
        ctk.CTkLabel(settings_frame, text="Show:").pack(side="left", padx=5, pady=5)
        self.max_items_var = ctk.StringVar(value=str(self.max_items))
        max_items_dropdown = ctk.CTkComboBox(settings_frame, values=["10", "25", "50", "100"],
                                           variable=self.max_items_var, width=80,
                                           command=self._on_max_items_change)
        max_items_dropdown.pack(side="left", padx=5, pady=5)
        ctk.CTkLabel(settings_frame, text="items").pack(side="left", padx=5, pady=5)
        
        # Refresh interval selector
        ctk.CTkLabel(settings_frame, text="Refresh every:").pack(side="left", padx=(20, 5), pady=5)
        self.refresh_interval_var = ctk.StringVar(value=str(self.refresh_interval))
        interval_dropdown = ctk.CTkComboBox(settings_frame, values=["2", "5", "10", "30"],
                                          variable=self.refresh_interval_var, width=80,
                                          command=self._on_interval_change)
        interval_dropdown.pack(side="left", padx=5, pady=5)
        ctk.CTkLabel(settings_frame, text="sec").pack(side="left", padx=5, pady=5)
        
        # Data source filter
        ctk.CTkLabel(settings_frame, text="Source:").pack(side="left", padx=(20, 5), pady=5)
        self.data_source_var = ctk.StringVar(value="all")
        source_dropdown = ctk.CTkComboBox(settings_frame, values=["all", "manual", "import"],
                                        variable=self.data_source_var, width=80,
                                        command=self._on_source_filter_change)
        source_dropdown.pack(side="left", padx=5, pady=5)
        
        # Status frame
        status_frame = ctk.CTkFrame(self.window)
        status_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.status_label = ctk.CTkLabel(status_frame, text="Loading...", 
                                        font=ctk.CTkFont(size=13))
        self.status_label.pack(padx=5, pady=5)
        
        # Main content area - scrollable list
        self.content_frame = ctk.CTkScrollableFrame(self.window)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        
        # Configure grid for content
        self.content_frame.columnconfigure(0, weight=1)
        
        # Database status label at the bottom
        self.db_status_label = ctk.CTkLabel(self.window, 
                                           text=f"DB: {os.path.basename(self.config.database_path)}", 
                                           font=ctk.CTkFont(size=10),
                                           text_color="gray60")
        self.db_status_label.pack(pady=(0, 5))
    
    def _toggle_auto_refresh(self):
        """Toggle auto-refresh on/off."""
        self.auto_refresh_enabled = self.auto_refresh_var.get()
        if self.auto_refresh_enabled:
            self._start_auto_refresh()
        else:
            self._stop_auto_refresh()
    
    def _on_max_items_change(self, value):
        """Handle max items dropdown change."""
        try:
            self.max_items = int(value)
            self._refresh_data()
        except ValueError:
            pass
    
    def _on_interval_change(self, value):
        """Handle refresh interval dropdown change."""
        try:
            self.refresh_interval = int(value)
            # Restart auto-refresh with new interval
            if self.auto_refresh_enabled:
                self._stop_auto_refresh()
                self._start_auto_refresh()
        except ValueError:
            pass
    
    def _on_source_filter_change(self, value):
        """Handle data source filter change."""
        self._refresh_data()
    
    def _start_auto_refresh(self):
        """Start the auto-refresh background thread."""
        if not self.auto_refresh_enabled:
            return
            
        self._stop_auto_refresh()  # Stop any existing thread
        self.stop_refresh.clear()
        
        def refresh_loop():
            while not self.stop_refresh.wait(self.refresh_interval):
                if self.auto_refresh_enabled:
                    try:
                        # Schedule UI update in main thread
                        self.window.after(0, self._refresh_data)
                    except Exception:
                        break  # Window might be destroyed
                else:
                    break
        
        self.refresh_thread = threading.Thread(target=refresh_loop, daemon=True)
        self.refresh_thread.start()
    
    def _stop_auto_refresh(self):
        """Stop the auto-refresh background thread."""
        self.stop_refresh.set()
        if self.refresh_thread and self.refresh_thread.is_alive():
            self.refresh_thread.join(timeout=1.0)
    
    @performance_monitor("Monitor Window Refresh")
    def _refresh_data(self):
        """Refresh the displayed data."""
        safe_execute(
            self._load_statistics,
            error_handler=error_handler,
            context="refreshing monitor statistics"
        )
        safe_execute(
            self._load_recent_changes,
            error_handler=error_handler,
            context="refreshing monitor data"
        )
    
    def reload_configuration(self):
        """Reload configuration and database connection.
        
        This method should be called when the database configuration changes
        to ensure the monitor window points to the correct database.
        """
        try:
            # Reload configuration
            self.config_manager = ConfigManager()
            self.config = self.config_manager.get_config()
            
            # Close existing database connection if it exists
            if hasattr(self.db, 'close'):
                try:
                    self.db.close()
                except Exception:
                    pass  # Ignore errors when closing
            
            # Create new database connection with updated path
            self.db = AssetDatabase(self.config.database_path)
            
            # Update the database status label
            if hasattr(self, 'db_status_label'):
                self.db_status_label.configure(text=f"DB: {os.path.basename(self.config.database_path)}")
            
            print(f"Monitor window reloaded - now using database: {self.config.database_path}")
            
            # Refresh data with new database connection
            self._refresh_data()
            
        except Exception as e:
            error_handler.handle_error(
                e, 
                context="reloading monitor configuration",
                user_message=f"Error reloading monitor configuration: {str(e)}"
            )
    
    def _on_days_filter_changed(self, value):
        """Handle days filter change."""
        try:
            if value == "All":
                self.days_filter = None
            else:
                # Parse descriptive filter values
                filter_mapping = {
                    "12 hours": 0.5,
                    "1 day": 1,
                    "1 week": 7,
                    "2 weeks": 14,
                    "1 month": 30,
                    "2 months": 60,
                    "3 months": 90,
                    "1 year": 365
                }
                self.days_filter = filter_mapping.get(value, 0.5)  # Default to 1 day if unknown
            # Refresh statistics with new filter
            self._load_statistics()
        except Exception:
            # If any error occurs, default to 1 day
            self.days_filter = 0.5
            self._load_statistics()
    
    def _load_statistics(self):
        """Load and display statistics."""
        try:
            # Get current room and cube
            current_room, current_cube = self._get_current_room_cube()
            
            # Update room card
            room_display = current_room or "n/a"
            self.room_header_label.configure(text=f"Room: {room_display}")
            
            if current_room:
                room_added = self._get_room_added_count(current_room, self.days_filter)
                room_modified = self._get_room_modified_count(current_room, self.days_filter)
                room_total = room_added + room_modified
                
                self.room_added_label.configure(text=f"Added: {room_added}")
                self.room_modified_label.configure(text=f"Modified: {room_modified}")
                self.room_total_label.configure(text=f"Total: {room_total}")
            else:
                self.room_added_label.configure(text="Added: 0")
                self.room_modified_label.configure(text="Modified: 0")
                self.room_total_label.configure(text="Total: 0")
            
            # Update cube/rack card - check if we should show rack instead of cube
            is_using_rack = self._is_using_rack_field()
            
            if is_using_rack:
                # Show rack statistics
                rack_display = current_cube or "n/a"  # current_cube contains the rack number when rack field is found
                self.cube_header_label.configure(text=f"Rack: {rack_display}")
                
                if current_cube:  # current_cube is actually the rack number in this case
                    rack_added = self._get_rack_added_count(current_cube, self.days_filter)
                    rack_modified = self._get_rack_modified_count(current_cube, self.days_filter)
                    rack_total = rack_added + rack_modified
                    
                    self.cube_added_label.configure(text=f"Added: {rack_added}")
                    self.cube_modified_label.configure(text=f"Modified: {rack_modified}")
                    self.cube_total_label.configure(text=f"Total: {rack_total}")
                else:
                    self.cube_added_label.configure(text="Added: 0")
                    self.cube_modified_label.configure(text="Modified: 0")
                    self.cube_total_label.configure(text="Total: 0")
            else:
                # Show cube statistics (original behavior)
                cube_display = current_cube or "n/a"
                self.cube_header_label.configure(text=f"Cube: {cube_display}")
                
                if current_cube:
                    cube_added = self._get_cube_added_count(current_cube, self.days_filter)
                    cube_modified = self._get_cube_modified_count(current_cube, self.days_filter)
                    cube_total = cube_added + cube_modified
                    
                    self.cube_added_label.configure(text=f"Added: {cube_added}")
                    self.cube_modified_label.configure(text=f"Modified: {cube_modified}")
                    self.cube_total_label.configure(text=f"Total: {cube_total}")
                else:
                    self.cube_added_label.configure(text="Added: 0")
                    self.cube_modified_label.configure(text="Modified: 0")
                    self.cube_total_label.configure(text="Total: 0")
            
            # Update overall card
            overall_added = self._get_overall_added_count(self.days_filter)
            overall_modified = self._get_overall_modified_count(self.days_filter)
            overall_total = overall_added + overall_modified
            
            self.overall_added_label.configure(text=f"Added: {overall_added}")
            self.overall_modified_label.configure(text=f"Modified: {overall_modified}")
            self.overall_total_label.configure(text=f"Total: {overall_total}")
            
        except Exception as e:
            # Error state
            self.room_header_label.configure(text="Room: Error")
            self.room_added_label.configure(text="Added: --")
            self.room_modified_label.configure(text="Modified: --")
            self.room_total_label.configure(text="Total: --")
            
            # Set generic label for cube/rack error (will show Cube or Rack based on current state)
            try:
                is_using_rack = self._is_using_rack_field()
                header_text = "Rack: Error" if is_using_rack else "Cube: Error"
            except:
                header_text = "Cube: Error"
            
            self.cube_header_label.configure(text=header_text)
            self.cube_added_label.configure(text="Added: --")
            self.cube_modified_label.configure(text="Modified: --")
            self.cube_total_label.configure(text="Total: --")
            
            self.overall_added_label.configure(text="Added: --")
            self.overall_modified_label.configure(text="Modified: --")
            self.overall_total_label.configure(text="Total: --")
            
            print(f"Error loading statistics: {e}")
    
    def _get_today_total(self, days_filter: Optional[int] = None) -> int:
        """Get the total number of assets added or modified within the specified days.
        
        Excludes imported items that haven't been changed.
        Only counts manually added items or any item that has been modified.
        
        Args:
            days_filter: Number of days to look back (None for all time, 0.5 for 12 hours)
        """
        try:
            now = datetime.now()
            
            if days_filter is None:
                # All time
                date_condition = "1 = 1"  # Always true
                date_params = []
            else:
                # Within specified days
                cutoff_date = now - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Count assets created within timeframe that are manual OR have been actually modified
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {date_condition}
                    AND is_deleted = 0
                    AND (
                        data_source = 'manual' 
                        OR modified_date > '1901-01-02'
                    )
                """
                cursor.execute(query, date_params)
                created_count = cursor.fetchone()[0]
                
                # Count assets modified within timeframe (but not created in timeframe)
                # These are automatically included regardless of data_source
                if days_filter is None:
                    modify_date_condition = "modified_date > '1901-01-02'"
                    modify_params = []
                else:
                    modify_date_condition = "modified_date >= ? AND modified_date > '1901-01-02'"
                    modify_params = [cutoff_date.isoformat()]
                
                modify_query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {modify_date_condition}
                    AND modified_date != created_date
                    AND is_deleted = 0
                    AND NOT ({date_condition})
                """
                cursor.execute(modify_query, modify_params + date_params)
                modified_count = cursor.fetchone()[0]
                
                return created_count + modified_count
                
        except Exception as e:
            print(f"Error getting today's total: {e}")
            return 0
    
    def _get_current_room_cube(self) -> tuple[str, str]:
        """Get the current room and cube/cubicle/rack from the most recent asset."""
        try:
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get available columns to check what fields exist
                available_columns = self.db.get_table_columns()
                
                # Check for room field
                room_column = None
                for col in ['room', 'Room', 'ROOM']:
                    if col in available_columns:
                        room_column = col
                        break
                
                # Check for both rack and cube fields
                rack_column = None
                for col in available_columns:
                    if 'rack' in col.lower():
                        rack_column = col
                        break
                
                cube_column = None
                for col in ['cube', 'Cube', 'CUBE', 'cubicle', 'Cubicle', 'CUBICLE']:
                    if col in available_columns:
                        cube_column = col
                        break
                
                # Need at least one location column
                if not room_column and not rack_column and not cube_column:
                    return None, None
                
                # Build query to get the most recent room/rack/cube values
                select_parts = []
                if room_column:
                    select_parts.append(room_column)
                if rack_column:
                    select_parts.append(rack_column)
                if cube_column:
                    select_parts.append(cube_column)
                
                if not select_parts:
                    return None, None
                
                select_sql = ', '.join(select_parts)
                
                cursor.execute(f"""
                    SELECT {select_sql}
                    FROM assets 
                    WHERE is_deleted = 0
                    ORDER BY 
                        CASE 
                            WHEN modified_date > '1901-01-02' AND modified_date != created_date 
                            THEN modified_date 
                            ELSE created_date 
                        END DESC
                    LIMIT 1
                """)
                
                result = cursor.fetchone()
                if result:
                    room_value = None
                    rack_value = None
                    cube_value = None
                    
                    # Parse result based on which columns we selected
                    idx = 0
                    if room_column:
                        room_value = result[idx] if len(result) > idx else None
                        idx += 1
                    if rack_column:
                        rack_value = result[idx] if len(result) > idx else None
                        idx += 1
                    if cube_column:
                        cube_value = result[idx] if len(result) > idx else None
                    
                    # Clean up None or empty values
                    room_value = room_value if room_value and str(room_value).strip() and room_value != 'None' else None
                    rack_value = rack_value if rack_value and str(rack_value).strip() and rack_value != 'None' else None
                    cube_value = cube_value if cube_value and str(cube_value).strip() and cube_value != 'None' else None
                    
                    # Determine which location value to use
                    # Priority: rack (if it has extractable data) > cube
                    location_value = None
                    
                    if rack_value:
                        # Try to extract rack number
                        rack_number = self._extract_rack_number(rack_value)
                        if rack_number:  # Only use rack if we can extract a number
                            location_value = rack_number
                        elif cube_value:  # Fall back to cube if rack extraction fails
                            location_value = cube_value
                    elif cube_value:
                        location_value = cube_value
                    
                    return room_value, location_value
                
                return None, None
                
        except Exception as e:
            print(f"Error getting current room/cube: {e}")
            return None, None
    
    def _extract_rack_number(self, rack_field_value: str) -> str:
        """Extract rack number from a rack field value.
        
        Parses text to get the first number until end of field or '/' is reached.
        Example: "3/23" returns "3", "Rack15" returns "15"
        """
        try:
            if not rack_field_value:
                return None
            
            # Convert to string and strip whitespace
            value = str(rack_field_value).strip()
            
            # Find first digit in the string
            rack_number = ""
            found_digit = False
            
            for char in value:
                if char.isdigit():
                    rack_number += char
                    found_digit = True
                elif char == '/' and found_digit:
                    # Stop at '/' if we've already found digits
                    break
                elif found_digit and not char.isdigit():
                    # Stop if we found digits but hit a non-digit (other than '/')
                    break
            
            return rack_number if rack_number else None
            
        except Exception as e:
            print(f"Error extracting rack number from '{rack_field_value}': {e}")
            return None
    
    def _is_using_rack_field(self) -> bool:
        """Check if the most recent asset uses a rack field instead of cube."""
        try:
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get available columns
                available_columns = self.db.get_table_columns()
                
                # Check for rack fields
                rack_column = None
                for col in available_columns:
                    if 'rack' in col.lower():
                        rack_column = col
                        break
                
                if not rack_column:
                    return False
                
                # Get the most recent asset (same logic as _get_current_room_cube)
                # Check if THIS specific asset has rack data
                cursor.execute(f"""
                    SELECT {rack_column}
                    FROM assets 
                    WHERE is_deleted = 0
                    ORDER BY 
                        CASE 
                            WHEN modified_date > '1901-01-02' AND modified_date != created_date 
                            THEN modified_date 
                            ELSE created_date 
                        END DESC
                    LIMIT 1
                """)
                
                result = cursor.fetchone()
                if result and result[0] is not None and str(result[0]).strip() != '':
                    # The most recent asset has rack data
                    return True
                
                return False
                
        except Exception as e:
            print(f"Error checking for rack field usage: {e}")
            return False
    
    def _get_room_total(self, current_room: str, days_filter: Optional[int] = None) -> int:
        """Get the total number of assets in the current room within the specified days.
        
        Excludes imported items that haven't been changed.
        Only counts manually added items or any item that has been modified.
        
        Args:
            current_room: The room to filter by
            days_filter: Number of days to look back (None for all time)
        """
        try:
            if not current_room:
                return 0
            
            # Build date filter condition
            if days_filter is None:
                date_condition = "1 = 1"  # Always true
                date_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get available columns
                available_columns = self.db.get_table_columns()
                
                # Find room column
                room_column = None
                for col in ['room', 'Room', 'ROOM']:
                    if col in available_columns:
                        room_column = col
                        break
                
                if not room_column:
                    return 0
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {room_column} = ?
                    AND {date_condition}
                    AND is_deleted = 0
                    AND (
                        data_source = 'manual' 
                        OR modified_date > '1901-01-02'
                    )
                """
                cursor.execute(query, [current_room] + date_params)
                
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting room total: {e}")
            return 0
    
    def _get_room_added_count(self, current_room: str, days_filter: Optional[int] = None) -> int:
        """Get count of assets added to the current room within the specified days.
        
        Counts manual assets created within the period as "Added".
        Import assets are only counted if they were also modified (handled separately).
        """
        try:
            if not current_room:
                return 0
            
            # Build date filter condition for creation
            if days_filter is None:
                date_condition = "1 = 1"
                date_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                available_columns = self.db.get_table_columns()
                
                # Find room column
                room_column = None
                for col in ['room', 'Room', 'ROOM']:
                    if col in available_columns:
                        room_column = col
                        break
                
                if not room_column:
                    return 0
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {room_column} = ?
                    AND {date_condition}
                    AND is_deleted = 0
                    AND data_source = 'manual'
                """
                cursor.execute(query, [current_room] + date_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting room added count: {e}")
            return 0
    
    def _get_room_modified_count(self, current_room: str, days_filter: Optional[int] = None) -> int:
        """Get count of assets modified in the current room within the specified days.
        
        Includes both manual and import assets that were modified.
        For assets created within the same period: manual = Added, import+modified = Modified.
        Excludes manual assets created within same period to avoid double-counting.
        """
        try:
            if not current_room:
                return 0
            
            # Build date filter condition for modification
            if days_filter is None:
                date_condition = "modified_date > '1901-01-02'"
                date_params = []
                exclude_new_manual_condition = ""
                exclude_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                cutoff_iso = cutoff_date.isoformat()
                
                date_condition = "modified_date >= ? AND modified_date > '1901-01-02'"
                date_params = [cutoff_iso]
                
                # Exclude manual assets created within the same period (they count as Added)
                # But include import assets created within period if they were also modified
                exclude_new_manual_condition = "AND (created_date < ? OR (created_date >= ? AND data_source = 'import'))"
                exclude_params = [cutoff_iso, cutoff_iso]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                available_columns = self.db.get_table_columns()
                
                # Find room column
                room_column = None
                for col in ['room', 'Room', 'ROOM']:
                    if col in available_columns:
                        room_column = col
                        break
                
                if not room_column:
                    return 0
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {room_column} = ?
                    AND {date_condition}
                    AND modified_date != created_date
                    AND is_deleted = 0
                    {exclude_new_manual_condition}
                """
                cursor.execute(query, [current_room] + date_params + exclude_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting room modified count: {e}")
            return 0
    
    def _get_cube_added_count(self, current_cube: str, days_filter: Optional[int] = None) -> int:
        """Get count of assets added to the current cube within the specified days."""
        try:
            if not current_cube:
                return 0
            
            # Build date filter condition for creation
            if days_filter is None:
                date_condition = "1 = 1"
                date_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                available_columns = self.db.get_table_columns()
                
                # Find cube column
                cube_column = None
                for col in ['cube', 'Cube', 'CUBE', 'cubicle', 'Cubicle', 'CUBICLE']:
                    if col in available_columns:
                        cube_column = col
                        break
                
                if not cube_column:
                    return 0
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {cube_column} = ?
                    AND {date_condition}
                    AND is_deleted = 0
                    AND data_source = 'manual'
                """
                cursor.execute(query, [current_cube] + date_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting cube added count: {e}")
            return 0
    
    def _get_cube_modified_count(self, current_cube: str, days_filter: Optional[int] = None) -> int:
        """Get count of assets modified in the current cube within the specified days.
        
        Includes both manual and import assets that were modified.
        For assets created within the same period: manual = Added, import+modified = Modified.
        Excludes manual assets created within same period to avoid double-counting.
        """
        try:
            if not current_cube:
                return 0
            
            # Build date filter condition for modification
            if days_filter is None:
                date_condition = "modified_date > '1901-01-02'"
                date_params = []
                exclude_new_manual_condition = ""
                exclude_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                cutoff_iso = cutoff_date.isoformat()
                
                date_condition = "modified_date >= ? AND modified_date > '1901-01-02'"
                date_params = [cutoff_iso]
                
                # Exclude manual assets created within the same period (they count as Added)
                # But include import assets created within period if they were also modified
                exclude_new_manual_condition = "AND (created_date < ? OR (created_date >= ? AND data_source = 'import'))"
                exclude_params = [cutoff_iso, cutoff_iso]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                available_columns = self.db.get_table_columns()
                
                # Find cube column
                cube_column = None
                for col in ['cube', 'Cube', 'CUBE', 'cubicle', 'Cubicle', 'CUBICLE']:
                    if col in available_columns:
                        cube_column = col
                        break
                
                if not cube_column:
                    return 0
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {cube_column} = ?
                    AND {date_condition}
                    AND modified_date != created_date
                    AND is_deleted = 0
                    {exclude_new_manual_condition}
                """
                cursor.execute(query, [current_cube] + date_params + exclude_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting cube modified count: {e}")
            return 0
    
    def _get_rack_added_count(self, current_rack: str, days_filter: Optional[int] = None) -> int:
        """Get count of assets added to the current rack within the specified days."""
        try:
            if not current_rack:
                return 0
            
            # Build date filter condition for creation
            if days_filter is None:
                date_condition = "1 = 1"
                date_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                available_columns = self.db.get_table_columns()
                
                # Find rack column
                rack_column = None
                for col in available_columns:
                    if 'rack' in col.lower():
                        rack_column = col
                        break
                
                if not rack_column:
                    return 0
                
                # Need to handle rack field differently since we extracted the number
                # We'll look for rack fields that contain our extracted rack number
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {rack_column} LIKE ?
                    AND {date_condition}
                    AND is_deleted = 0
                    AND data_source = 'manual'
                """
                # Use LIKE pattern to match rack number at start of field
                rack_pattern = f"{current_rack}%"
                cursor.execute(query, [rack_pattern] + date_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting rack added count: {e}")
            return 0
    
    def _get_rack_modified_count(self, current_rack: str, days_filter: Optional[int] = None) -> int:
        """Get count of assets modified in the current rack within the specified days.
        
        Includes both manual and import assets that were modified.
        For assets created within the same period: manual = Added, import+modified = Modified.
        Excludes manual assets created within same period to avoid double-counting.
        """
        try:
            if not current_rack:
                return 0
            
            # Build date filter condition for modification
            if days_filter is None:
                date_condition = "modified_date > '1901-01-02'"
                date_params = []
                exclude_new_manual_condition = ""
                exclude_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                cutoff_iso = cutoff_date.isoformat()
                
                date_condition = "modified_date >= ? AND modified_date > '1901-01-02'"
                date_params = [cutoff_iso]
                
                # Exclude manual assets created within the same period (they count as Added)
                # But include import assets created within period if they were also modified
                exclude_new_manual_condition = "AND (created_date < ? OR (created_date >= ? AND data_source = 'import'))"
                exclude_params = [cutoff_iso, cutoff_iso]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                available_columns = self.db.get_table_columns()
                
                # Find rack column
                rack_column = None
                for col in available_columns:
                    if 'rack' in col.lower():
                        rack_column = col
                        break
                
                if not rack_column:
                    return 0
                
                # Need to handle rack field differently since we extracted the number
                # We'll look for rack fields that contain our extracted rack number
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {rack_column} LIKE ?
                    AND {date_condition}
                    AND modified_date != created_date
                    AND is_deleted = 0
                    {exclude_new_manual_condition}
                """
                # Use LIKE pattern to match rack number at start of field
                rack_pattern = f"{current_rack}%"
                cursor.execute(query, [rack_pattern] + date_params + exclude_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting rack modified count: {e}")
            return 0
    
    def _get_manual_total(self, days_filter: Optional[int] = None) -> int:
        """Get the total number of manually added assets within the specified days."""
        try:
            now = datetime.now()
            
            if days_filter is None:
                date_condition = "1 = 1"
                date_params = []
            else:
                cutoff_date = now - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {date_condition}
                    AND is_deleted = 0
                    AND data_source = 'manual'
                """
                cursor.execute(query, date_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting manual total: {e}")
            return 0
    
    def _get_import_total(self, days_filter: Optional[int] = None) -> int:
        """Get the total number of imported assets within the specified days."""
        try:
            now = datetime.now()
            
            if days_filter is None:
                date_condition = "1 = 1"
                date_params = []
            else:
                cutoff_date = now - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {date_condition}
                    AND is_deleted = 0
                    AND data_source = 'import'
                """
                cursor.execute(query, date_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting import total: {e}")
            return 0
    
    def _get_overall_added_count(self, days_filter: Optional[int] = None) -> int:
        """Get the total number of assets added within the specified days.
        
        This counts assets that were created within the time period, regardless of 
        whether they were also modified after creation.
        
        Args:
            days_filter: Number of days to look back (None for all time, 0.5 for 12 hours)
        """
        try:
            now = datetime.now()
            
            if days_filter is None:
                date_condition = "1 = 1"
                date_params = []
            else:
                cutoff_date = now - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {date_condition}
                    AND is_deleted = 0
                    AND data_source = 'manual'
                """
                cursor.execute(query, date_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting overall added count: {e}")
            return 0
    
    def _get_overall_modified_count(self, days_filter: Optional[int] = None) -> int:
        """Get the total number of assets modified within the specified days.
        
        Includes both manual and import assets that were modified.
        For assets created within the same period: manual = Added, import+modified = Modified.
        Excludes manual assets created within same period to avoid double-counting.
        """
        try:
            if days_filter is None:
                # For "all time", just count all modified assets
                date_condition = "modified_date > '1901-01-02'"
                date_params = []
                exclude_new_manual_condition = ""
                exclude_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                cutoff_iso = cutoff_date.isoformat()
                
                # Assets modified within the time period
                date_condition = "modified_date >= ? AND modified_date > '1901-01-02'"
                date_params = [cutoff_iso]
                
                # Exclude manual assets created within the same period (they count as Added)
                # But include import assets created within period if they were also modified
                exclude_new_manual_condition = "AND (created_date < ? OR (created_date >= ? AND data_source = 'import'))"
                exclude_params = [cutoff_iso, cutoff_iso]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {date_condition}
                    AND modified_date != created_date
                    AND is_deleted = 0
                    {exclude_new_manual_condition}
                """
                cursor.execute(query, date_params + exclude_params)
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting overall modified count: {e}")
            return 0

    def _get_cube_total(self, current_cube: str, days_filter: Optional[int] = None) -> int:
        """Get the total number of assets in the current cube/cubicle within the specified days.
        
        Excludes imported items that haven't been changed.
        Only counts manually added items or any item that has been modified.
        
        Args:
            current_cube: The cube/cubicle to filter by
            days_filter: Number of days to look back (None for all time)
        """
        try:
            if not current_cube:
                return 0
            
            # Build date filter condition
            if days_filter is None:
                date_condition = "1 = 1"  # Always true
                date_params = []
            else:
                cutoff_date = datetime.now() - timedelta(days=days_filter)
                date_condition = "created_date >= ?"
                date_params = [cutoff_date.isoformat()]
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get available columns
                available_columns = self.db.get_table_columns()
                
                # Find cube column
                cube_column = None
                for col in ['cube', 'Cube', 'CUBE', 'cubicle', 'Cubicle', 'CUBICLE']:
                    if col in available_columns:
                        cube_column = col
                        break
                
                if not cube_column:
                    return 0
                
                query = f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {cube_column} = ?
                    AND {date_condition}
                    AND is_deleted = 0
                    AND (
                        data_source = 'manual' 
                        OR modified_date > '1901-01-02'
                    )
                """
                cursor.execute(query, [current_cube] + date_params)
                
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting cube total: {e}")
            return 0

    def _get_room_cube_total(self, current_room: str, current_cube: str) -> int:
        """Get the total number of assets in the current room and cube.
        
        Excludes imported items that haven't been changed.
        Only counts manually added items or any item that has been modified.
        """
        try:
            if not current_room and not current_cube:
                return 0
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get available columns
                available_columns = self.db.get_table_columns()
                
                # Find room and cube columns
                room_column = None
                cube_column = None
                
                for col in ['room', 'Room', 'ROOM']:
                    if col in available_columns:
                        room_column = col
                        break
                
                for col in ['cube', 'Cube', 'CUBE', 'cubicle', 'Cubicle', 'CUBICLE']:
                    if col in available_columns:
                        cube_column = col
                        break
                
                # Build WHERE conditions
                where_conditions = []
                params = []
                
                if current_room and room_column:
                    where_conditions.append(f"{room_column} = ?")
                    params.append(current_room)
                
                if current_cube and cube_column:
                    where_conditions.append(f"{cube_column} = ?")
                    params.append(current_cube)
                
                if not where_conditions:
                    return 0
                
                where_sql = ' AND '.join(where_conditions)
                
                cursor.execute(f"""
                    SELECT COUNT(*) 
                    FROM assets 
                    WHERE {where_sql}
                    AND is_deleted = 0
                    AND (
                        data_source = 'manual' 
                        OR modified_date > '1901-01-02'
                    )
                """, params)
                
                return cursor.fetchone()[0]
                
        except Exception as e:
            print(f"Error getting room/cube total: {e}")
            return 0

    def _load_recent_changes(self):
        """Load and display recent asset changes with minimal flicker."""
        try:
            # Get recent changes from database
            recent_assets = self._get_recent_changes()
            
            # Check if smooth refresh is enabled and data actually changed
            if self.smooth_refresh_var.get() and self._assets_data_unchanged(recent_assets):
                # Just update the status timestamp without rebuilding UI
                count = len(recent_assets)
                last_update = datetime.now().strftime("%H:%M:%S")
                self._update_status(f"Showing {count} items (Last updated: {last_update})")
                return
            
            # Store new data
            self.last_assets_data = recent_assets.copy() if recent_assets else []
            
            # Use update_idletasks to reduce visual flicker during rebuild
            if self.smooth_refresh_var.get():
                self.window.update_idletasks()
            
            # Clear existing content only when data actually changed
            for widget in self.content_frame.winfo_children():
                widget.destroy()
            self.asset_widgets.clear()
            
            if not recent_assets:
                # No recent changes
                no_data_label = ctk.CTkLabel(self.content_frame, 
                                           text="No recent changes found",
                                           font=ctk.CTkFont(size=14),
                                           text_color="gray50")
                no_data_label.grid(row=0, column=0, pady=20)
                self._update_status("No recent changes")
                return
            
            # Temporarily disable window updates during bulk widget creation
            if self.smooth_refresh_var.get():
                self.content_frame.update_idletasks()
            
            # Display each asset change
            for i, asset in enumerate(recent_assets):
                widget = self._create_asset_item(asset, i)
                self.asset_widgets.append(widget)
            
            # Re-enable updates and force a refresh
            if self.smooth_refresh_var.get():
                self.content_frame.update_idletasks()
            
            # Update status
            count = len(recent_assets)
            last_update = datetime.now().strftime("%H:%M:%S")
            self._update_status(f"Showing {count} items (Last updated: {last_update})")
            
        except Exception as e:
            error_label = ctk.CTkLabel(self.content_frame, 
                                     text=f"Error loading data: {str(e)[:50]}...",
                                     font=ctk.CTkFont(size=12),
                                     text_color="red")
            error_label.grid(row=0, column=0, pady=10)
            self._update_status("Error loading data")
    
    def _assets_data_unchanged(self, new_assets: List[Dict[str, Any]]) -> bool:
        """Check if the asset data has actually changed since last refresh."""
        if len(new_assets) != len(self.last_assets_data):
            return False
        
        # Quick check - compare key fields of first few items
        for i, (new_asset, old_asset) in enumerate(zip(new_assets[:5], self.last_assets_data[:5])):
            if (new_asset.get('id') != old_asset.get('id') or 
                new_asset.get('modified_date') != old_asset.get('modified_date') or
                new_asset.get('created_date') != old_asset.get('created_date')):
                return False
        
        return True
    
    def _get_recent_changes(self) -> List[Dict[str, Any]]:
        """Get recent asset changes from the database."""
        try:
            # Get assets modified within the same time period as statistics
            if self.days_filter is None:
                # All time - use a reasonable cutoff for display
                cutoff_date = datetime.now() - timedelta(days=365)
            else:
                # Use the same filter as statistics
                cutoff_date = datetime.now() - timedelta(days=self.days_filter)
            
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                
                # Get the template path for column mapping
                template_path = self.config.default_template_path
                
                # Get dynamic column mapping from template to database columns
                try:
                    column_mapping = self.db.get_dynamic_column_mapping(template_path)
                except Exception as e:
                    print(f"Warning: Could not get column mapping: {e}")
                    column_mapping = {}
                
                # Get required fields from config - these should exist in the database
                required_fields = self.config.required_fields or []
                
                # Also get monitor fields to ensure they're included in the query
                monitor_fields = set()
                monitor_fields.update(self.config.get('monitor_primary_fields', []))
                monitor_fields.update(self.config.get('monitor_secondary_fields', []))
                monitor_fields.update(self.config.get('monitor_tertiary_fields', []))
                
                # Combine required fields with monitor fields
                all_display_fields = set(required_fields) | monitor_fields
                
                # Use SELECT * to get all columns including notes
                # This ensures AssetDetailWindow has access to all asset data
                columns_sql = '*'
                
                # Build WHERE clause to match refined statistics logic
                # Show assets that are either:
                # 1. Created within the time period (Added - manual only, or Modified - import that was modified)
                # 2. Created before the period but modified within it (Modified)
                cutoff_iso = cutoff_date.isoformat()
                where_conditions = [
                    "((created_date >= ? AND data_source = 'manual') OR (created_date >= ? AND data_source = 'import' AND modified_date >= ? AND modified_date != created_date AND modified_date > '1901-01-02') OR (created_date < ? AND modified_date >= ? AND modified_date != created_date AND modified_date > '1901-01-02'))",
                    "is_deleted = 0"
                ]
                params = [cutoff_iso, cutoff_iso, cutoff_iso, cutoff_iso, cutoff_iso]
                
                # Add data source filter if not "all"
                data_source_filter = self.data_source_var.get()
                if data_source_filter != "all":
                    where_conditions.append("data_source = ?")
                    params.append(data_source_filter)
                
                where_sql = " AND ".join(where_conditions)
                
                # Query for recent changes with refined change type detection
                # Logic: 
                # - Modified AND added manually within period = Added
                # - Modified AND added by import within period = Modified 
                # - Added manually within period = Added
                # - Added by import within period = Imported (filtered out later)
                cursor.execute(f"""
                    SELECT 
                        {columns_sql},
                        CASE 
                            WHEN created_date >= ? AND data_source = 'manual' THEN 'Added'
                            WHEN created_date >= ? AND data_source = 'import' AND modified_date >= ? AND modified_date != created_date AND modified_date > '1901-01-02' THEN 'Modified'
                            WHEN created_date >= ? AND data_source = 'import' THEN 'Imported'
                            WHEN created_date < ? AND modified_date >= ? AND modified_date != created_date AND modified_date > '1901-01-02' THEN 'Modified'
                            ELSE 'Added'
                        END as change_type
                    FROM assets 
                    WHERE {where_sql}
                    ORDER BY 
                        CASE 
                            WHEN modified_date > '1901-01-02' AND modified_date != created_date 
                            THEN modified_date 
                            ELSE created_date 
                        END DESC
                    LIMIT ?
                """, [cutoff_iso, cutoff_iso, cutoff_iso, cutoff_iso, cutoff_iso, cutoff_iso] + params + [self.max_items])
                
                rows = cursor.fetchall()
                
                # Convert to dictionaries with template field names for easier display
                results = []
                for row in rows:
                    row_dict = dict(row)
                    
                    # Add reverse mappings from database columns back to template field names
                    for field_name in all_display_fields:
                        db_column = column_mapping.get(field_name)
                        if db_column and db_column in row_dict:
                            # Add the value under the template field name for easier access
                            row_dict[field_name] = row_dict[db_column]
                    
                    results.append(row_dict)
                
                return results
                
        except Exception as e:
            print(f"Error getting recent changes: {e}")
            # Try a simpler fallback query with just basic columns
            try:
                if self.days_filter is None:
                    cutoff_date = datetime.now() - timedelta(days=365)
                else:
                    cutoff_date = datetime.now() - timedelta(days=self.days_filter)
                with self.db.get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        SELECT 
                            id,
                            created_date,
                            modified_date,
                            'Added' as change_type
                        FROM assets 
                        WHERE created_date >= ?
                        AND is_deleted = 0
                        ORDER BY created_date DESC
                        LIMIT ?
                    """, (cutoff_date.isoformat(), min(self.max_items, 10)))
                    
                    rows = cursor.fetchall()
                    return [dict(row) for row in rows]
            except Exception as fallback_error:
                print(f"Fallback query also failed: {fallback_error}")
                return []
    
    def _create_asset_item(self, asset: Dict[str, Any], row_index: int):
        """Create a display item for a single asset."""
        # Main item frame
        item_frame = ctk.CTkFrame(self.content_frame)
        item_frame.grid(row=row_index, column=0, sticky="ew", padx=5, pady=2)
        item_frame.columnconfigure(1, weight=1)  # Asset identification column expands
        # Configure other columns to have fixed width
        item_frame.columnconfigure(0, weight=0)  # Change type
        item_frame.columnconfigure(2, weight=0)  # Data source
        item_frame.columnconfigure(3, weight=0)  # Timestamp
        item_frame.columnconfigure(4, weight=0)  # Details button
        
        # Change type indicator (Added/Modified)
        change_type = asset.get('change_type', 'Unknown')
        type_color = "#4CAF50" if change_type == "Added" else "#FF9800"  # Green for added, orange for modified
        
        type_label = ctk.CTkLabel(item_frame, text=change_type, 
                                 font=ctk.CTkFont(size=12, weight="bold"),
                                 text_color=type_color, width=60)
        type_label.grid(row=0, column=0, padx=5, pady=2, sticky="w")
        
        # Data source indicator
        data_source = asset.get('data_source', 'unknown')
        source_color = "#2196F3" if data_source == "manual" else "#9C27B0"  # Blue for manual, purple for import
        source_symbol = "✋" if data_source == "manual" else "📄"  # Hand for manual, document for import
        
        source_label = ctk.CTkLabel(item_frame, text=f"{source_symbol}", 
                                   font=ctk.CTkFont(size=14),
                                   text_color=source_color, width=20)
        source_label.grid(row=0, column=2, padx=2, pady=2, sticky="w")
        
        # Dynamic asset display based on configuration
        current_row = 0
        
        # Row 0 (Primary) - Main line with configured primary fields
        primary_fields = self.config.get('monitor_primary_fields', ["Serial Number", "Asset No."])
        primary_text = self._build_field_display_text(asset, primary_fields)
        
        if primary_text:
            primary_label = ctk.CTkLabel(item_frame, text=primary_text,
                                       font=ctk.CTkFont(size=14, weight="bold"),
                                       anchor="w")
            primary_label.grid(row=current_row, column=1, padx=5, pady=2, sticky="ew")
        
        # Row 1 (Secondary) - Secondary line with configured secondary fields  
        secondary_fields = self.config.get('monitor_secondary_fields', ["*Manufacturer", "*Model"])
        secondary_text = self._build_field_display_text(asset, secondary_fields)
        
        if secondary_text:
            current_row += 1
            secondary_label = ctk.CTkLabel(item_frame, text=secondary_text,
                                         font=ctk.CTkFont(size=12),
                                         anchor="w", text_color="gray70")
            secondary_label.grid(row=current_row, column=1, padx=5, pady=(0, 2), sticky="ew")
        
        # Row 2 (Tertiary) - Third line with configured tertiary fields
        tertiary_fields = self.config.get('monitor_tertiary_fields', ["Room", "Cubicle", "System Name"])
        tertiary_text = self._build_field_display_text(asset, tertiary_fields)
        
        if tertiary_text:
            current_row += 1
            tertiary_label = ctk.CTkLabel(item_frame, text=tertiary_text,
                                        font=ctk.CTkFont(size=11),
                                        anchor="w", text_color="gray60")
            tertiary_label.grid(row=current_row, column=1, padx=5, pady=(0, 2), sticky="ew")
        
        # Timestamp
        if change_type == "Added":
            timestamp_str = asset.get('created_date', '')
        else:
            timestamp_str = asset.get('modified_date', '')
        
        if timestamp_str:
            try:
                # Parse the timestamp and format it nicely
                if 'T' in timestamp_str:
                    dt = datetime.fromisoformat(timestamp_str)
                else:
                    dt = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
                
                # Show relative time for recent items
                now = datetime.now()
                diff = now - dt
                
                if diff.total_seconds() < 60:
                    time_text = "Just now"
                elif diff.total_seconds() < 3600:
                    minutes = int(diff.total_seconds() / 60)
                    time_text = f"{minutes}m ago"
                elif diff.total_seconds() < 86400:
                    hours = int(diff.total_seconds() / 3600)
                    time_text = f"{hours}h ago"
                else:
                    time_text = dt.strftime("%m/%d %H:%M")
                
            except Exception:
                time_text = timestamp_str[:16]  # Fallback to truncated string
        else:
            time_text = "Unknown time"
        
        time_label = ctk.CTkLabel(item_frame, text=time_text, 
                                 font=ctk.CTkFont(size=11),
                                 text_color="gray50")
        time_label.grid(row=0, column=3, padx=5, pady=2, sticky="e")
        
        # Details button - small button to show asset details
        details_btn = ctk.CTkButton(item_frame, text="📋", width=30, height=24,
                                   font=ctk.CTkFont(size=14),
                                   command=lambda: self._show_asset_details(asset))
        details_btn.grid(row=0, column=4, padx=2, pady=2, sticky="e")
    
    def _build_field_display_text(self, asset: Dict[str, Any], field_names: list) -> str:
        """Build display text from asset data using configured field names"""
        if not field_names:
            return ""
        
        field_values = []
        for field_name in field_names:
            # Get field value, trying multiple possible keys for compatibility
            value = (asset.get(field_name) or 
                    asset.get(field_name.lower()) or 
                    asset.get(field_name.replace('*', '')) or 
                    asset.get(field_name.replace('*', '').lower()) or '')
            
            # Clean up value
            if value and str(value).strip() and str(value).strip().lower() not in ['none', 'n/a']:
                # Special formatting for specific fields
                if field_name in ['Asset No.', 'asset_no']:
                    if not value or str(value).strip() in ['None', 'N/A', None]:
                        value = f"ID:{asset.get('id', 'Unknown')}"
                    field_values.append(f"Asset: {value}")
                elif field_name in ['Serial Number', 'serial_number']:
                    field_values.append(f"SN: {value}")
                elif field_name in ['Room', 'room']:
                    field_values.append(f"Room: {value}")
                elif field_name in ['Cubicle', 'cubicle', 'Cube', 'cube']:
                    field_values.append(f"Cube: {value}")
                elif field_name in ['System Name', 'system_name']:
                    field_values.append(f"System: {value}")
                elif field_name in ['Location', 'location']:
                    field_values.append(f"Location: {value}")
                elif field_name in ['Status', 'status']:
                    field_values.append(f"Status: {value}")
                else:
                    # For other fields, just show the value
                    field_values.append(str(value))
        
        return " | ".join(field_values) if field_values else ""

    def _show_asset_details(self, asset: Dict[str, Any]):
        """Show detailed asset information in a popup window."""
        try:
            # Define callback to refresh data after asset edit
            def on_asset_edited():
                self._refresh_data()  # Refresh the monitor data
            
            AssetDetailWindow(self.window, asset, on_edit_callback=on_asset_edited)
        except Exception as e:
            print(f"Error showing asset details: {e}")
    
    def _update_status(self, message: str):
        """Update the status label."""
        self.status_label.configure(text=message)
    
    def _on_closing(self):
        """Handle window closing to clean up resources."""
        try:
            self._stop_auto_refresh()
        except Exception:
            pass
        
        self.window.destroy()


def main():
    """Standalone test for the monitor window."""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    
    # Create root window for testing
    root = ctk.CTk()
    root.geometry("300x200")
    root.title("Monitor Test")
    
    # Button to open monitor
    def open_monitor():
        MonitorWindow(root)
    
    ctk.CTkLabel(root, text="Monitor Window Test").pack(pady=20)
    ctk.CTkButton(root, text="Open Monitor", command=open_monitor).pack(pady=10)
    
    root.mainloop()


if __name__ == "__main__":
    main()
