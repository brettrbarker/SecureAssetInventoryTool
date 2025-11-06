"""
Reusable UI components for the Secure Asset Inventory Tool.
Contains custom widgets and common UI patterns.
"""

import customtkinter as ctk
from tkinter import messagebox
import calendar
import csv
import os
from datetime import datetime, timedelta
from typing import List, Callable, Optional, Dict, Any
from config_manager import ConfigManager
from asset_database import AssetDatabase

class SearchableDropdown(ctk.CTkFrame):
    """Custom searchable + scrollable dropdown with custom input support.

    Features:
    - Type directly in the entry field for custom values
    - Click dropdown button to see/select from existing values
    - Supports filtering existing values by typing
    - Allows custom entries not in the predefined list

    Usage:
        var = ctk.StringVar()
        w = SearchableDropdown(parent, values=[...], variable=var)
    """
    def __init__(self, master, values: List[str], variable: ctk.StringVar | None = None, width: int = 220, height: int = 32):
        super().__init__(master)
        self.values_all = values[:]  # master list
        self.variable = variable or ctk.StringVar(value="")
        self.width = width
        self.height = height
        self.popup = None
        self.search_var = ctk.StringVar()
        self.allow_custom = True  # Allow custom input

        # Main display: entry (editable) + button
        self.display_entry = ctk.CTkEntry(self, textvariable=self.variable, width=width-30)
        self.display_entry.grid(row=0, column=0, sticky="we")
        # Allow direct typing in the entry field
        self.display_entry.bind("<KeyRelease>", self._on_entry_change)
        self.display_entry.bind("<Button-1>", self._on_entry_click)
        toggle_btn = ctk.CTkButton(self, text="▼", width=28, command=self.open_popup)
        toggle_btn.grid(row=0, column=1, padx=(4,0))
        self.columnconfigure(0, weight=1)
        
        # Track if popup should auto-open for filtering
        self.auto_popup = False

    def _on_entry_click(self, event=None):
        """Handle click in entry field - focus but don't auto-open popup."""
        self.display_entry.focus_set()
        return "break"

    def _on_entry_change(self, event=None):
        """Handle typing in the entry field - optionally show filtered popup."""
        current_text = self.variable.get()
        
        # If the user is typing and there are matching values, show filtered popup
        if current_text and self.values_all:
            filtered = [v for v in self.values_all if v and current_text.lower() in v.lower()]
            if filtered and not (self.popup and self.popup.winfo_exists()):
                self.auto_popup = True
                self.open_popup()
                self.search_var.set(current_text)

    def open_popup(self):
        if self.popup and self.popup.winfo_exists():
            return
        self.popup = ctk.CTkToplevel(self)
        self.popup.transient(self.winfo_toplevel())
        self.popup.title("Select or type custom value")
        self.popup.geometry(self._popup_geometry())
        self.popup.attributes("-topmost", True)
        self.popup.grab_set()  # modal for selection
        self.popup.bind("<Escape>", lambda e: self.close_popup())

        # Instructions
        instructions = ctk.CTkLabel(self.popup, text="Select from list or type custom value in main field", 
                                  font=ctk.CTkFont(size=11), text_color="gray60")
        instructions.pack(fill="x", padx=8, pady=(8,2))

        # Search box for filtering
        search_entry = ctk.CTkEntry(self.popup, placeholder_text="Type to filter existing options", textvariable=self.search_var)
        search_entry.pack(fill="x", padx=8, pady=(2,4))
        
        # Set initial search value if auto-opened from typing
        if self.auto_popup:
            search_entry.insert(0, self.variable.get())
            self.auto_popup = False
        
        search_entry.focus_set()
        
        # Add trace callback with error handling
        def safe_rebuild(*args):
            try:
                self._rebuild_list()
            except Exception:
                pass  # Ignore errors from destroyed widgets
        
        self.search_var.trace_add("write", safe_rebuild)

        # Scrollable list
        self.list_frame = ctk.CTkScrollableFrame(self.popup, width=self.width+40, height=240)
        self.list_frame.pack(fill="both", expand=True, padx=8, pady=(0,4))

        # Action buttons frame
        btn_frame = ctk.CTkFrame(self.popup)
        btn_frame.pack(fill="x", padx=8, pady=(0,8))
        
        # Use current entry value button (for custom entries)
        current_val = self.variable.get().strip()
        if current_val and current_val not in self.values_all:
            use_custom_btn = ctk.CTkButton(btn_frame, text=f'Use "{current_val}"', 
                                         command=lambda: self._select_custom(current_val),
                                         fg_color="green", hover_color="darkgreen")
            use_custom_btn.pack(side="left", padx=(0,4), pady=4)
        
        # Clear selection button
        clear_btn = ctk.CTkButton(btn_frame, text="Clear", command=lambda: self._select(""),
                                fg_color="gray", hover_color="darkgray")
        clear_btn.pack(side="right", padx=(4,0), pady=4)

        self._rebuild_list()

    def _popup_geometry(self) -> str:
        # Position just under widget
        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        return f"{self.width+80}x380+{x}+{y}"  # Slightly taller for new elements

    def _rebuild_list(self):
        # Safety check: don't rebuild if popup is closed/destroyed
        if not self.popup or not self.popup.winfo_exists():
            return
            
        # Clear existing buttons
        try:
            for child in self.list_frame.winfo_children():
                child.destroy()
        except Exception:
            # Widget may have been destroyed, skip silently
            return
        
        original_search = self.search_var.get().strip()  # Preserve original capitalization
        pattern = original_search.lower()  # Use lowercase only for filtering
        if pattern:
            filtered = [v for v in self.values_all if v and pattern in v.lower()]
        else:
            filtered = self.values_all
        
        if not filtered:
            if pattern:
                # Show option to add the search term as custom value
                try:
                    no_match_label = ctk.CTkLabel(self.list_frame, text=f'No matches found for "{original_search}"')
                    no_match_label.pack(fill="x", pady=4)
                    if self.allow_custom:
                        add_custom_btn = ctk.CTkButton(self.list_frame, text=f'Add "{original_search}" as custom value',
                                                     command=lambda: self._select_custom(original_search),
                                                     fg_color="orange", hover_color="darkorange")
                        add_custom_btn.pack(fill="x", padx=2, pady=2)
                except Exception:
                    # Widget may have been destroyed, skip silently
                    return
            else:
                try:
                    ctk.CTkLabel(self.list_frame, text="No existing options").pack(fill="x", pady=4)
                except Exception:
                    return
            return
        
        # Show existing options
        try:
            for v in filtered:
                b = ctk.CTkButton(self.list_frame, text=v if v else "(blank)", anchor="w",
                                  command=lambda val=v: self._select(val))
                b.pack(fill="x", padx=2, pady=2)
        except Exception:
            # Widget may have been destroyed, skip silently
            return

    def _select(self, value: str):
        """Select a value from the existing options."""
        self.variable.set(value)
        self.close_popup()

    def _select_custom(self, value: str):
        """Select a custom value (not in predefined list)."""
        self.variable.set(value)
        self.close_popup()

    def close_popup(self):
        if self.popup and self.popup.winfo_exists():
            # Clean up any trace callbacks to prevent errors
            try:
                # Remove all trace callbacks from search_var
                for trace_id in self.search_var.trace_info():
                    self.search_var.trace_remove(*trace_id)
            except Exception:
                pass
            
            self.popup.grab_release()
            self.popup.destroy()
        self.popup = None
        self.search_var.set("")  # Clear search when closing

class DatePicker(ctk.CTkFrame):
    """Custom date picker widget with calendar and Today button.
    
    Features:
    - Click calendar button to open date picker popup
    - Calendar grid for date selection
    - Today button to quickly select current date
    - Manual entry field for typing dates
    - Uses MM/D/YYYY format to match audit date format
    
    Usage:
        var = ctk.StringVar()
        picker = DatePicker(parent, variable=var)
    """
    
    def __init__(self, master, variable: ctk.StringVar | None = None, width: int = 220, height: int = 32):
        super().__init__(master)
        self.variable = variable or ctk.StringVar(value="")
        self.width = width
        self.height = height
        self.popup = None
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year
        
        # Main display: entry (editable) + calendar button
        self.display_entry = ctk.CTkEntry(self, textvariable=self.variable, width=width-30)
        self.display_entry.grid(row=0, column=0, sticky="we")
        
        # Calendar button
        calendar_btn = ctk.CTkButton(self, text="📅", width=28, command=self.open_calendar)
        calendar_btn.grid(row=0, column=1, padx=(4,0))
        self.columnconfigure(0, weight=1)
    
    def open_calendar(self):
        if self.popup and self.popup.winfo_exists():
            return
            
        self.popup = ctk.CTkToplevel(self)
        self.popup.transient(self.winfo_toplevel())
        self.popup.title("Select Date")
        self.popup.geometry(self._popup_geometry())
        self.popup.attributes("-topmost", True)
        self.popup.grab_set()
        self.popup.bind("<Escape>", lambda e: self.close_calendar())
        
        # Parse current date if valid
        current_date = self._parse_current_date()
        if current_date:
            self.current_month = current_date.month
            self.current_year = current_date.year
        
        # Header with month/year navigation
        header_frame = ctk.CTkFrame(self.popup)
        header_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        # Previous month button
        prev_btn = ctk.CTkButton(header_frame, text="◀", width=30, command=self._prev_month)
        prev_btn.pack(side="left")
        
        # Month/Year label
        self.month_year_label = ctk.CTkLabel(header_frame, 
                                           text=f"{calendar.month_name[self.current_month]} {self.current_year}",
                                           font=ctk.CTkFont(size=14, weight="bold"))
        self.month_year_label.pack(side="left", expand=True)
        
        # Next month button
        next_btn = ctk.CTkButton(header_frame, text="▶", width=30, command=self._next_month)
        next_btn.pack(side="right")
        
        # Calendar grid frame
        self.calendar_frame = ctk.CTkFrame(self.popup)
        self.calendar_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Action buttons frame
        action_frame = ctk.CTkFrame(self.popup)
        action_frame.pack(fill="x", padx=10, pady=(5, 10))
        
        # Today button
        today_btn = ctk.CTkButton(action_frame, text="Today", command=self._select_today)
        today_btn.pack(side="left", padx=(0, 10))
        
        # Clear button
        clear_btn = ctk.CTkButton(action_frame, text="Clear", command=self._clear_date,
                                fg_color="gray", hover_color="darkgray")
        clear_btn.pack(side="right")
        
        self._build_calendar()
    
    def _popup_geometry(self) -> str:
        # Position just under widget
        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        return f"280x320+{x}+{y}"
    
    def _parse_current_date(self):
        """Parse the current date value in the entry field."""
        try:
            date_str = self.variable.get().strip()
            if not date_str:
                return None
            # Parse MM/D/YYYY or MM/DD/YYYY format
            parts = date_str.split('/')
            if len(parts) == 3:
                month, day, year = int(parts[0]), int(parts[1]), int(parts[2])
                return datetime(year, month, day)
        except (ValueError, IndexError):
            pass
        return None
    
    def _prev_month(self):
        """Navigate to previous month."""
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self._update_calendar()
    
    def _next_month(self):
        """Navigate to next month."""
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self._update_calendar()
    
    def _update_calendar(self):
        """Update the month/year label and rebuild calendar."""
        self.month_year_label.configure(text=f"{calendar.month_name[self.current_month]} {self.current_year}")
        self._build_calendar()
    
    def _build_calendar(self):
        """Build the calendar grid for the current month."""
        # Clear existing calendar
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        # Day headers (Sunday first)
        days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        for col, day in enumerate(days):
            label = ctk.CTkLabel(self.calendar_frame, text=day, font=ctk.CTkFont(size=12, weight="bold"))
            label.grid(row=0, column=col, padx=2, pady=2, sticky="nsew")
        
        # Get calendar data for current month
        # Set first day of week to Sunday (6) to match our headers
        cal = calendar.Calendar(firstweekday=6)
        month_days = cal.monthdayscalendar(self.current_year, self.current_month)
        
        # Current date for highlighting
        today = datetime.now()
        current_date = self._parse_current_date()
        
        # Create date buttons
        for week_num, week in enumerate(month_days, start=1):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cell for days from other months
                    empty = ctk.CTkLabel(self.calendar_frame, text="")
                    empty.grid(row=week_num, column=day_num, padx=2, pady=2, sticky="nsew")
                else:
                    # Determine button color
                    fg_color = None
                    hover_color = None
                    
                    # Highlight today
                    if (self.current_year == today.year and 
                        self.current_month == today.month and 
                        day == today.day):
                        fg_color = "orange"
                        hover_color = "darkorange"
                    
                    # Highlight selected date
                    elif (current_date and
                          self.current_year == current_date.year and
                          self.current_month == current_date.month and
                          day == current_date.day):
                        fg_color = "green"
                        hover_color = "darkgreen"
                    
                    btn = ctk.CTkButton(self.calendar_frame, text=str(day), width=30, height=30,
                                      command=lambda d=day: self._select_date(d),
                                      fg_color=fg_color, hover_color=hover_color)
                    btn.grid(row=week_num, column=day_num, padx=2, pady=2, sticky="nsew")
        
        # Configure grid weights for even spacing
        for i in range(7):
            self.calendar_frame.grid_columnconfigure(i, weight=1)
    
    def _select_date(self, day):
        """Select a specific date."""
        selected_date = datetime(self.current_year, self.current_month, day)
        formatted_date = f"{selected_date:%m}/{selected_date.day}/{selected_date:%Y}"
        self.variable.set(formatted_date)
        self.close_calendar()
    
    def _select_today(self):
        """Select today's date."""
        today = datetime.now()
        formatted_date = f"{today:%m}/{today.day}/{today:%Y}"
        self.variable.set(formatted_date)
        self.close_calendar()
    
    def _clear_date(self):
        """Clear the selected date."""
        self.variable.set("")
        self.close_calendar()
    
    def close_calendar(self):
        """Close the calendar popup."""
        if self.popup and self.popup.winfo_exists():
            self.popup.grab_release()
            self.popup.destroy()
        self.popup = None

class FilterDialog(ctk.CTkToplevel):
    """Reusable filter dialog for exports and searches."""
    
    def __init__(self, parent, title: str = "Filter Options", filters: List[dict] = None):
        super().__init__(parent)
        self.parent = parent
        self.title(title)
        self.geometry("500x400")
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.filters = filters or []
        self._build_dialog()
        self._center_window()
    
    def _build_dialog(self):
        """Build the filter dialog interface."""
        # Title
        title_label = ctk.CTkLabel(self, text="Select Filter Options", 
                                  font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(pady=(20, 15))
        
        # Dynamic filter creation based on provided filters
        # Implementation for flexible filter creation...
    
    def _center_window(self):
        """Center the dialog window."""
        WindowManager.center_window(self)
    
    def get_result(self):
        """Get the filter result after dialog closes."""
        self.wait_window()
        return self.result

class DatabaseOperationDialog(ctk.CTkToplevel):
    """Reusable dialog for database operations with progress indication."""
    
    def __init__(self, parent, operation_name: str, operation_func: Callable):
        super().__init__(parent)
        self.parent = parent
        self.operation_name = operation_name
        self.operation_func = operation_func
        self.result = None
        self._build_dialog()
        self._center_window()
    
    def _build_dialog(self):
        """Build the operation dialog."""
        # Progress bar, operation details, etc.
        pass
    
    def _center_window(self):
        """Center the dialog window."""
        WindowManager.center_window(self)

class WindowManager:
    """Utility class for common window operations."""
    
    @staticmethod
    def center_window(window, width: int = None, height: int = None):
        """Center a window on screen or parent."""
        window.update_idletasks()
        
        if width and height:
            window.geometry(f"{width}x{height}")
        
        w = window.winfo_reqwidth()
        h = window.winfo_reqheight()
        
        try:
            # Try to center on parent
            parent_x = window.master.winfo_rootx()
            parent_y = window.master.winfo_rooty()
            parent_w = window.master.winfo_width()
            parent_h = window.master.winfo_height()
            
            x = parent_x + (parent_w - w) // 2
            y = parent_y + (parent_h - h) // 2
        except:
            # Center on screen
            screen_w = window.winfo_screenwidth()
            screen_h = window.winfo_screenheight()
            x = (screen_w - w) // 2
            y = (screen_h - h) // 2
        
        window.geometry(f"+{x}+{y}")
    
    @staticmethod
    def create_action_frame(parent, buttons: List[dict], pack_side: str = "bottom"):
        """Create a standardized action button frame."""
        frame = ctk.CTkFrame(parent)
        frame.pack(side=pack_side, fill="x", padx=10, pady=10)
        
        for button_config in buttons:
            btn = ctk.CTkButton(frame, **button_config)
            btn.pack(side="right" if pack_side == "bottom" else "left", 
                    padx=5, pady=5)
        
        return frame


class EmbeddedAssetDetail:
    """Embeddable asset detail component for use within other windows.
    
    This provides the same functionality as AssetDetailWindow but embeds
    into an existing frame rather than creating its own window.
    """
    
    def __init__(self, parent_frame, asset: Dict[str, Any], on_edit_callback=None, show_edit_button=True):
        """Initialize the embedded asset detail component.
        
        Args:
            parent_frame: Parent frame to embed the component into
            asset: Dictionary containing asset data
            on_edit_callback: Optional callback function to call when asset is edited
            show_edit_button: Whether to show the edit button (default True)
        """
        self.parent_frame = parent_frame
        self.asset = asset
        self.on_edit_callback = on_edit_callback
        self.show_edit_button = show_edit_button
        
        # Load configuration for field structure
        self.config_manager = ConfigManager()
        self.config = self.config_manager.get_config()
        
        # Create database instance to get field structure
        self.db = AssetDatabase(self.config.database_path)
        
        # Get template structure
        self.template_path = self.config.default_template_path
        self.headers = []
        self._load_template_structure()
        
        # Load field configurations
        self.required_fields = set(self.config.required_fields)
        self.excluded_fields = set(self.config.excluded_fields)
        
        # Create the embedded content
        self._create_embedded_widgets()
    
    def _load_template_structure(self):
        """Load template structure to get field order and names."""
        try:
            if os.path.exists(self.template_path):
                with open(self.template_path, 'r', newline='', encoding='utf-8') as file:
                    reader = csv.reader(file)
                    self.headers = next(reader, [])
            else:
                # Fallback headers if template not found
                self.headers = [
                    "*Asset Type", "*Manufacturer", "*Model", "Status", "Serial Number", 
                    "Location", "Room", "System Name", "IP Address", "MAC Address",
                    "Purchase Date", "Vendor", "Cost", "Notes"
                ]
        except Exception as e:
            print(f"Error loading template: {e}")
            # Use basic fallback headers
            self.headers = list(self.asset.keys())
    
    def refresh(self, updated_asset: Dict[str, Any] = None):
        """Refresh the display with updated asset data.
        
        Args:
            updated_asset: Updated asset dictionary. If None, uses existing self.asset
        """
        if updated_asset:
            self.asset = updated_asset
        
        # Clear all existing widgets
        for widget in self.parent_frame.winfo_children():
            widget.destroy()
        
        # Rebuild the display
        self._create_embedded_widgets()
    
    def _create_embedded_widgets(self):
        """Create the embedded detail view widgets."""
        # Pack directly into parent frame without creating a container
        # Title frame
        title_frame = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        title_frame.pack(fill="x", pady=(8, 5), padx=15)  # Reduced from pady=(10, 10), padx=20
        
        # Title
        asset_no = self.asset.get('asset_no') or f"ID {self.asset.get('id', 'Unknown')}"
        title_text = f"Asset Details: {asset_no}"
        title_label = ctk.CTkLabel(title_frame, text=title_text, 
                                  font=ctk.CTkFont(size=18, weight="bold"))  # Reduced from size=20
        title_label.pack(side="left")
        
        # Edit button (if enabled)
        if self.show_edit_button:
            # Button frame for multiple buttons
            button_frame = ctk.CTkFrame(title_frame, fg_color="transparent")
            button_frame.pack(side="right", padx=(10, 0))
            
            # Request Label button
            label_btn = ctk.CTkButton(button_frame, text="🏷️ Request Label", 
                                    command=self._request_label,
                                    width=120, height=28,
                                    fg_color="#2d5a27", hover_color="#1e3f1b")
            label_btn.pack(side="left", padx=(0, 5))
            
            # Edit button
            edit_btn = ctk.CTkButton(button_frame, text="✏️ Edit", command=self._edit_asset,
                                   width=75, height=28,
                                   fg_color="#1f538d", hover_color="#14375e")
            edit_btn.pack(side="left")
        
        # Form area (using scrollable frame - scrollbar appears automatically when needed)
        scrollable_frame = ctk.CTkScrollableFrame(self.parent_frame, fg_color="transparent")
        scrollable_frame.pack(fill="both", expand=True, padx=15, pady=(0, 5))
        
        # Inner frame for the form fields
        self.form_inner = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        self.form_inner.pack(fill="both", expand=True)
        
        # Build the form fields (read-only)
        self._build_form_fields()
    
    def _build_form_fields(self):
        """Build form fields using the same pattern as AssetDetailWindow."""
        # Configure grid layout for 4 columns (2 pairs of label+widget)
        self.form_inner.grid_columnconfigure(0, weight=0)  # Label 1
        self.form_inner.grid_columnconfigure(1, weight=1)  # Widget 1
        self.form_inner.grid_columnconfigure(2, weight=0)  # Label 2
        self.form_inner.grid_columnconfigure(3, weight=1)  # Widget 2

        current_row = 0
        
        # Asset info subtitle
        subtitle_parts = []
        if self.asset.get('manufacturer'):
            subtitle_parts.append(self.asset['manufacturer'])
        if self.asset.get('model'):
            subtitle_parts.append(self.asset['model'])
        if self.asset.get('serial_number'):
            subtitle_parts.append(f"SN: {self.asset['serial_number']}")
            
        if subtitle_parts:
            subtitle_label = ctk.CTkLabel(self.form_inner, text=" | ".join(subtitle_parts),
                                        font=ctk.CTkFont(size=14))
            subtitle_label.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=8, pady=(0,8))  # Reduced from pady=(0,15)
            current_row += 1

        # Filter headers to exclude configured excluded fields and system fields
        readonly_fields = {"Asset No.", "id", "created_date", "modified_date", "created_by", "modified_by", "data_source"}
        all_excluded = self.excluded_fields | readonly_fields
        
        # Get column mapping for field name translation
        column_mapping = self.db.get_dynamic_column_mapping(self.template_path)
        
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
            current_row = self._create_field_section(required_headers, current_row, column_mapping)
        
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
            current_row = self._create_field_section(additional_headers, current_row, column_mapping)
        
        # System information section - simple centered text
        divider = ctk.CTkFrame(self.form_inner, height=2)
        divider.grid(row=current_row, column=0, columnspan=4, sticky="we", padx=4, pady=(10,10))
        current_row += 1
        
        # Build system info text
        created_by = self.asset.get('created_by', 'Unknown')
        created_date_raw = self.asset.get('created_date', '')
        created_date = self._format_date(created_date_raw)
        data_source = self.asset.get('data_source', 'Unknown')
        modified_by = self.asset.get('modified_by', 'Unknown')
        modified_date_raw = self.asset.get('modified_date', '')
        modified_date = self._format_date(modified_date_raw)
        
        # Create centered frame for system info
        system_info_frame = ctk.CTkFrame(self.form_inner, fg_color="transparent")
        system_info_frame.grid(row=current_row, column=0, columnspan=4, pady=(5,10))
        
        # System info text
        system_text = f"Created by {created_by} on {created_date} via {data_source}"
        
        # Check if modified date is different from created date (compare raw timestamps)
        # The default modified_date is '1901-01-01 00:00:00' for unmodified assets
        if modified_date_raw and modified_date_raw != created_date_raw and modified_date_raw != '1901-01-01 00:00:00':
            system_text += f"\nLast Modified by {modified_by} on {modified_date}"
        
        system_label = ctk.CTkLabel(system_info_frame, 
                                    text=system_text,
                                    font=ctk.CTkFont(size=11),
                                    text_color="gray60",
                                    justify="center")
        system_label.pack(pady=5)
    
    def _create_field_section(self, headers, start_row, column_mapping):
        """Create a two-column section for given headers with read-only values."""
        # Use the same implementation as AssetDetailWindow
        fields_with_data = []
        
        for header in headers:
            # Get database column name for this header
            db_column = column_mapping.get(header)
            if not db_column:
                # Fallback heuristic
                db_column = header.lower().replace("*", "").replace(" ", "_")
                db_column = db_column.replace("(", "").replace(")", "").replace("?", "")
                db_column = db_column.replace(",", "").replace("#", "").replace("/", "_").replace("-", "_")
                while "__" in db_column:
                    db_column = db_column.replace("__", "_")
            
            # Get the value for this field
            value = ""
            
            # Handle both dictionary and object formats
            if hasattr(self.asset, 'get'):
                # Dictionary format
                if db_column in self.asset:
                    value = self.asset[db_column]
                elif header in self.asset:
                    value = self.asset[header]
                elif header.lower() in self.asset:
                    value = self.asset[header.lower()]
                else:
                    # Try some common variations
                    header_clean = header.replace("*", "").strip()
                    if header_clean.lower() in self.asset:
                        value = self.asset[header_clean.lower()]
                    elif header_clean.replace(" ", "_").lower() in self.asset:
                        value = self.asset[header_clean.replace(" ", "_").lower()]
            else:
                # Object format (with attributes)
                if hasattr(self.asset, db_column):
                    value = getattr(self.asset, db_column)
                elif hasattr(self.asset, header):
                    value = getattr(self.asset, header)
                elif hasattr(self.asset, header.lower()):
                    value = getattr(self.asset, header.lower())
                else:
                    # Try some common variations
                    header_clean = header.replace("*", "").strip()
                    if hasattr(self.asset, header_clean.lower()):
                        value = getattr(self.asset, header_clean.lower())
                    elif hasattr(self.asset, header_clean.replace(" ", "_").lower()):
                        value = getattr(self.asset, header_clean.replace(" ", "_").lower())
            
            # Convert None to empty string
            if value is None:
                value = ""
            
            display_value = str(value) if value else ""
            
            # Format date fields
            if "date" in header.lower() and display_value:
                display_value = self._format_date(display_value)
            
            # Only include fields with data
            if display_value:
                fields_with_data.append((header, display_value))
        
        # Create widgets only for fields with data
        for idx, (header, display_value) in enumerate(fields_with_data):
            col_group = idx % 2  # 0 or 1
            row_offset = idx // 2
            base_col = col_group * 2
            row = start_row + row_offset
            
            # Create field label
            label = ctk.CTkLabel(self.form_inner, text=header + ":")
            label.grid(row=row, column=base_col, sticky="e", padx=8, pady=2)  # Reduced from pady=4
            
            # Check if this is a Related Asset field
            is_related_asset = "related" in header.lower() and "asset" in header.lower()
            has_related_assets = is_related_asset and display_value
            
            if has_related_assets:
                # Parse multiple related assets (comma or semicolon separated)
                separators = [',', ';']
                related_assets = [display_value]  # Default to single value
                
                for separator in separators:
                    if separator in display_value:
                        related_assets = [asset.strip() for asset in display_value.split(separator) if asset.strip()]
                        break
                
                # Format display text with make/model information
                formatted_display = []
                for asset in related_assets:
                    formatted_display.append(self._format_related_asset_display(asset))
                
                # Join formatted displays (preserve separator style with line breaks)
                separator_used = ',' if ',' in display_value else ';'
                display_text = f"{separator_used}\n".join(formatted_display)
                
                # Create frame for value and buttons
                value_frame = ctk.CTkFrame(self.form_inner, fg_color="transparent")
                value_frame.grid(row=row, column=base_col + 1, sticky="ew", padx=8, pady=2)
                value_frame.grid_columnconfigure(0, weight=1)
                
                if len(related_assets) == 1:
                    # Single related asset - calculate height based on display text length
                    line_count = display_text.count('\n') + 1
                    text_height = min(max(30, line_count * 20), 80)
                    
                    value_textbox = ctk.CTkTextbox(value_frame, 
                                                 height=text_height,
                                                 wrap="word",
                                                 activate_scrollbars=False,
                                                 fg_color=("gray90", "gray20"),
                                                 corner_radius=6)
                    value_textbox.insert("0.0", display_text)
                    value_textbox.configure(state="disabled")  # Read-only but selectable
                    value_textbox.grid(row=0, column=0, sticky="ew", padx=(0, 4))
                    
                    view_btn = ctk.CTkButton(value_frame, text="👁", width=28, height=28,
                                           command=lambda val=related_assets[0]: self._view_related_asset(val),
                                           fg_color="#1f538d", hover_color="#14375e")
                    view_btn.grid(row=0, column=1, padx=(4, 0))
                else:
                    # Multiple related assets - calculate height based on number and content
                    line_count = display_text.count('\n') + 1
                    text_height = min(max(60, line_count * 20), 120)
                    enable_scrollbar = line_count > 6
                    
                    value_textbox = ctk.CTkTextbox(value_frame, 
                                                 height=text_height,
                                                 wrap="word",
                                                 activate_scrollbars=enable_scrollbar,
                                                 fg_color=("gray90", "gray20"),
                                                 corner_radius=6)
                    value_textbox.insert("0.0", display_text)
                    value_textbox.configure(state="disabled")  # Read-only but selectable
                    value_textbox.grid(row=0, column=0, sticky="ew", padx=(0, 4))
                    
                    if len(related_assets) <= 3:
                        # Few assets - show buttons horizontally
                        for i, asset in enumerate(related_assets):
                            view_btn = ctk.CTkButton(value_frame, text=f"👁{i+1}", width=32, height=28,
                                                   command=lambda val=asset: self._view_related_asset(val),
                                                   fg_color="#1f538d", hover_color="#14375e")
                            view_btn.grid(row=0, column=i+1, padx=(2, 0))
                    else:
                        # Many assets - create dropdown-style button
                        menu_btn = ctk.CTkButton(value_frame, text="👁▼", width=40, height=28,
                                               command=lambda assets=related_assets: self._show_related_assets_menu(assets, value_frame),
                                               fg_color="#1f538d", hover_color="#14375e")
                        menu_btn.grid(row=0, column=1, padx=(4, 0))
            else:
                # Regular value textbox (selectable)
                # Check if this is a multiline field (like Notes) that needs more height
                is_multiline = self.db.should_field_be_multiline(header, self.template_path)
                
                # Calculate appropriate height based on content and field type
                if is_multiline and display_value:
                    # Count lines in the content
                    line_count = display_value.count('\n') + 1
                    # Set height based on line count, with min 60 and max 150
                    textbox_height = min(max(60, line_count * 20), 150)
                    enable_scrollbar = line_count > 7  # Enable scrollbar if more than 7 lines
                else:
                    textbox_height = 30
                    enable_scrollbar = False
                
                value_textbox = ctk.CTkTextbox(self.form_inner, 
                                             height=textbox_height,
                                             wrap="word",
                                             activate_scrollbars=enable_scrollbar,
                                             fg_color=("gray90", "gray20"),
                                             corner_radius=6)
                value_textbox.insert("0.0", display_value)
                value_textbox.configure(state="disabled")  # Read-only but selectable
                value_textbox.grid(row=row, column=base_col + 1, sticky="ew", padx=8, pady=2)

        rows_used = (len(fields_with_data) + 1) // 2
        return start_row + rows_used
    
    def _format_date(self, date_str: str) -> str:
        """Format date string for display."""
        if not date_str:
            return ""
        try:
            dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            return dt.strftime("%m/%d/%Y")
        except (ValueError, TypeError):
            return date_str
    
    def _show_related_assets_menu(self, related_assets: List[str], parent_frame):
        """Show a popup menu with buttons for each related asset."""
        try:
            # Create popup menu (need to find the toplevel parent)
            toplevel = parent_frame.winfo_toplevel()
            popup = ctk.CTkToplevel(toplevel)
            popup.title("Select Related Asset")
            popup.transient(toplevel)
            popup.grab_set()
            popup.attributes("-topmost", True)
            
            # Position popup near the parent frame
            parent_frame.update_idletasks()
            x = parent_frame.winfo_rootx()
            y = parent_frame.winfo_rooty() + parent_frame.winfo_height()
            popup.geometry(f"450x{min(400, len(related_assets) * 45 + 80)}+{x}+{y}")  # Wider for make/model
            
            # Title
            title_label = ctk.CTkLabel(popup, text="Select Related Asset to View:", 
                                     font=ctk.CTkFont(size=14, weight="bold"))
            title_label.pack(pady=(10, 5))
            
            # Scrollable frame for asset buttons
            scroll_frame = ctk.CTkScrollableFrame(popup, width=420, height=min(350, len(related_assets) * 45))
            scroll_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
            
            # Create button for each related asset with make/model info
            for i, asset in enumerate(related_assets):
                # Format the button text with make/model
                display_text = f"{i+1}. {self._format_related_asset_display(asset)}"
                
                asset_btn = ctk.CTkButton(scroll_frame, text=display_text, 
                                        anchor="w", height=40,
                                        command=lambda val=asset, p=popup: self._select_related_asset(val, p),
                                        fg_color="#1f538d", hover_color="#14375e")
                asset_btn.pack(fill="x", padx=5, pady=2)
            
            # Close button
            close_btn = ctk.CTkButton(popup, text="Close", fg_color="gray", 
                                    command=popup.destroy)
            close_btn.pack(pady=(0, 10))
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while showing related assets menu: {e}")
    
    def _select_related_asset(self, asset_value: str, popup_window):
        """Select and view a related asset from the popup menu."""
        popup_window.destroy()
        self._view_related_asset(asset_value)

    def _get_related_asset_info(self, related_value: str) -> Optional[Dict[str, Any]]:
        """Get related asset information from the database.
        
        Args:
            related_value: Asset number or serial number to search for
            
        Returns:
            Dictionary with asset info or None if not found
        """
        try:
            # First, search by Asset Number
            assets = self.db.search_assets({"asset_no": related_value})
            if assets:
                return assets[0]
            
            # If not found by Asset Number, search by Serial Number
            assets = self.db.search_assets({"serial_number": related_value})
            if assets:
                return assets[0]
        except Exception as e:
            print(f"Error fetching related asset info for '{related_value}': {e}")
        
        return None
    
    def _format_related_asset_display(self, related_value: str) -> str:
        """Format related asset for display with Make and Model.
        
        Args:
            related_value: Asset number or serial number
            
        Returns:
            Formatted string with asset number and make/model
        """
        asset_info = self._get_related_asset_info(related_value)
        
        if asset_info:
            # Get manufacturer and model
            manufacturer = asset_info.get('manufacturer', '').strip()
            model = asset_info.get('model', '').strip()
            
            # Build display text
            if manufacturer and model:
                return f"{related_value} ({manufacturer} {model})"
            elif manufacturer:
                return f"{related_value} ({manufacturer})"
            elif model:
                return f"{related_value} ({model})"
        
        # Fallback to just the asset number/serial
        return related_value

    def _view_related_asset(self, related_value: str):
        """Open Asset Details window for the related asset."""
        try:
            # Search for the related asset in the database
            related_asset = self._get_related_asset_info(related_value)
            
            if related_asset:
                # Open Asset Details window for the related asset
                AssetDetailWindow(
                    parent=self.parent_frame.winfo_toplevel(),
                    asset=related_asset,
                    on_edit_callback=self.on_edit_callback
                )
            else:
                messagebox.showinfo(
                    "Asset Not Found", 
                    f"No asset found with Asset Number or Serial Number '{related_value}'"
                )
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while searching for related asset: {e}")

    def _edit_asset(self):
        """Open the edit asset window."""
        try:
            # Import here to avoid circular imports
            from edit_asset import EditAssetWindow
            
            asset_id = self.asset.get('id')
            if not asset_id:
                messagebox.showerror("Error", "Cannot edit asset: Asset ID not found.")
                return
            
            # Define callback to refresh asset data after edit
            def on_asset_updated():
                # Call the external callback if provided
                if self.on_edit_callback:
                    self.on_edit_callback()
            
            # Open edit window
            EditAssetWindow(
                parent=self.parent_frame.winfo_toplevel(),
                asset_id=asset_id,
                config=self.config,
                on_update_callback=on_asset_updated
            )
            
        except ImportError as e:
            messagebox.showerror("Error", f"Could not open edit window: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def _request_label(self):
        """Request a label for the asset."""
        try:
            from asset_database import AssetDatabase
            
            asset_id = self.asset.get('id')
            if not asset_id:
                messagebox.showerror("Error", "Cannot request label: Asset ID not found.")
                return
            
            # Initialize database connection
            db = AssetDatabase()
            
            # Update the label_requested_date field
            success = db.request_label(asset_id)
            
            if success:
                messagebox.showinfo("Success", "Label request submitted successfully!")
                # Call the external callback if provided to refresh the display
                if self.on_edit_callback:
                    self.on_edit_callback()
            else:
                messagebox.showerror("Error", "Failed to submit label request.")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while requesting label: {e}")


class AssetDetailWindow:
    """Window for displaying detailed asset information with tabs for Details and History.
    """
    
    def __init__(self, parent, asset: Dict[str, Any], on_edit_callback=None):
        """Initialize the asset detail window.
        
        Args:
            parent: Parent window
            asset: Dictionary containing asset data
            on_edit_callback: Optional callback function to call when asset is edited
        """
        self.parent = parent
        self.asset = asset
        self.on_edit_callback = on_edit_callback
        
        # Create database instance for history
        self.config_manager = ConfigManager()
        self.config = self.config_manager.get_config()
        self.db = AssetDatabase(self.config.database_path)
        
        # Create window
        self.window = ctk.CTkToplevel(parent)
        self.window.title(f"Asset Details - {asset.get('asset_no', 'Unknown')}")
        self.window.geometry("1000x700")  # Match Edit Asset window size
        self.window.minsize(900, 600)
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create tabbed interface
        self._create_tabbed_interface()
        
        # Create bottom action frame
        action_frame = ctk.CTkFrame(self.window)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        # Buttons - right-aligned
        ctk.CTkButton(action_frame, text="Close", fg_color="gray", 
                      command=self.window.destroy).pack(side="right", padx=10, pady=10)
        
        self.window.focus_force()
        self._center_window()
    
    def _create_tabbed_interface(self):
        """Create the tabbed interface with Details and History tabs."""
        # Create tab view
        self.tabview = ctk.CTkTabview(self.window)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        details_tab = self.tabview.add("Details")
        history_tab = self.tabview.add("History")
        
        # Set default tab
        self.tabview.set("Details")
        
        # Create Details tab content using existing EmbeddedAssetDetail
        self.embedded_detail = EmbeddedAssetDetail(
            details_tab, 
            self.asset, 
            on_edit_callback=self._on_asset_edited,
            show_edit_button=True
        )
        
        # Create History tab content
        self._create_history_tab(history_tab)
    
    def _create_history_tab(self, parent_tab):
        """Create the History tab content showing audit log."""
        # Title frame
        title_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        title_frame.pack(fill="x", pady=(10, 5), padx=15)
        
        title_label = ctk.CTkLabel(title_frame, text="Change History", 
                                  font=ctk.CTkFont(size=18, weight="bold"))
        title_label.pack(side="left")
        
        # Refresh button
        refresh_btn = ctk.CTkButton(title_frame, text="🔄 Refresh", command=self._refresh_history,
                                   width=100, height=28, fg_color="#1f538d", hover_color="#14375e")
        refresh_btn.pack(side="right")
        
        # Create scrollable frame for history
        self.history_frame = ctk.CTkScrollableFrame(parent_tab)
        self.history_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        
        # Load and display history
        self._load_history()
    
    def _load_history(self):
        """Load and display the audit history for this asset."""
        # Clear existing history
        for widget in self.history_frame.winfo_children():
            widget.destroy()
        
        try:
            asset_id = self.asset.get('id')
            if not asset_id:
                no_history_label = ctk.CTkLabel(self.history_frame, 
                                              text="No asset ID found - cannot retrieve history.",
                                              font=ctk.CTkFont(size=14))
                no_history_label.pack(pady=20)
                return
            
            # Get audit history from database
            history = self.db.get_audit_history(asset_id)
            
            if not history:
                no_history_label = ctk.CTkLabel(self.history_frame, 
                                              text="No change history found for this asset.",
                                              font=ctk.CTkFont(size=14))
                no_history_label.pack(pady=20)
                return
            
            # Display each history entry
            for i, entry in enumerate(history):
                self._create_history_entry(entry)
                
                # Add divider between entries (except after the last one)
                if i < len(history) - 1:
                    divider = ctk.CTkFrame(self.history_frame, height=2, fg_color=("gray70", "gray30"))
                    divider.pack(fill="x", padx=20, pady=(10, 5))
                
        except Exception as e:
            error_label = ctk.CTkLabel(self.history_frame, 
                                     text=f"Error loading history: {e}",
                                     font=ctk.CTkFont(size=14),
                                     text_color="red")
            error_label.pack(pady=20)
    
    def _create_history_entry(self, entry: Dict[str, Any]):
        """Create a single history entry widget."""
        # Main entry frame
        entry_frame = ctk.CTkFrame(self.history_frame)
        entry_frame.pack(fill="x", padx=5, pady=5)
        
        # Header with action and date
        header_frame = ctk.CTkFrame(entry_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        # Format date for display
        try:
            change_date = entry['change_date']
            if isinstance(change_date, str):
                # Parse ISO format date
                dt = datetime.fromisoformat(change_date.replace('Z', '+00:00'))
                formatted_date = dt.strftime("%m/%d/%Y %I:%M %p")
            else:
                formatted_date = str(change_date)
        except:
            formatted_date = str(entry['change_date'])
        
        # Action and date
        action_text = entry['action'].upper()
        if action_text == 'INSERT':
            action_text = 'CREATED'
            action_color = "green"
        elif action_text == 'UPDATE':
            action_text = 'MODIFIED'
            action_color = "orange"
        elif action_text == 'DELETE':
            action_text = 'DELETED'
            action_color = "red"
        else:
            action_color = "blue"
        
        action_label = ctk.CTkLabel(header_frame, text=action_text,
                                   font=ctk.CTkFont(size=16, weight="bold"),
                                   text_color=action_color,
                                   anchor="w")
        action_label.pack(side="left", anchor="w")
        
        # Changed by (pack immediately after action)
        if entry['changed_by']:
            changed_by_label = ctk.CTkLabel(header_frame, text=f"by {entry['changed_by']}",
                                          font=ctk.CTkFont(size=13),
                                          text_color="gray60",
                                          anchor="w")
            changed_by_label.pack(side="left", padx=(10, 0), anchor="w")
        
        # Date (pack after changed_by)
        date_label = ctk.CTkLabel(header_frame, text=formatted_date,
                                 font=ctk.CTkFont(size=14),
                                 anchor="w")
        date_label.pack(side="left", padx=(20, 0), anchor="w")
        
        # Details frame (only for field changes)
        if entry['field_name'] and entry['action'].upper() == 'UPDATE':
            details_frame = ctk.CTkFrame(entry_frame, fg_color="transparent")
            details_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            # Field name
            field_label = ctk.CTkLabel(details_frame, text=f"Field: {entry['field_name']}",
                                     font=ctk.CTkFont(size=14, weight="bold"),
                                     anchor="w", justify="left")
            field_label.pack(anchor="w", pady=(0, 5), fill="x")
            
            # Create table frame with outline
            table_frame = ctk.CTkFrame(details_frame, fg_color=("gray90", "gray20"), corner_radius=5)
            table_frame.pack(fill="x", pady=2)
            
            # Configure grid weights for equal column width
            table_frame.grid_columnconfigure(0, weight=1)
            table_frame.grid_columnconfigure(1, weight=1)
            
            # Create header row with colored background
            from_header = ctk.CTkLabel(table_frame, 
                                     text="From",
                                     font=ctk.CTkFont(size=13, weight="bold"),
                                     fg_color=("red", "darkred"),
                                     corner_radius=3,
                                     text_color="white")
            from_header.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
            
            to_header = ctk.CTkLabel(table_frame, 
                                   text="To",
                                   font=ctk.CTkFont(size=13, weight="bold"),
                                   fg_color=("green", "darkgreen"),
                                   corner_radius=3,
                                   text_color="white")
            to_header.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
            
            # Create value cells with selectable text
            from_value = entry['old_value'] if entry['old_value'] else '(empty)'
            to_value = entry['new_value'] if entry['new_value'] else '(empty)'
            
            # Calculate height based on content (count newlines and estimate wrapped lines)
            def calculate_cell_height(text):
                """Calculate appropriate height for textbox based on content."""
                if not text or text == '(empty)':
                    return 30  # Minimum height for empty/short content
                
                lines = str(text).split('\n')
                line_count = len(lines)
                
                # Estimate additional lines from wrapping (rough estimate: 80 chars per line)
                for line in lines:
                    if len(line) > 80:
                        line_count += len(line) // 80
                
                # Each line is roughly 20 pixels, with min 30 and max 200
                height = max(30, min(200, line_count * 20 + 10))
                return height
            
            from_height = calculate_cell_height(from_value)
            to_height = calculate_cell_height(to_value)
            # Use the larger height for both cells to keep them aligned
            cell_height = max(from_height, to_height)
            
            from_cell = ctk.CTkTextbox(table_frame,
                                     height=cell_height,
                                     wrap="word",
                                     activate_scrollbars=True,
                                     font=ctk.CTkFont(size=13),
                                     fg_color=("white", "gray25"),
                                     corner_radius=3)
            from_cell.insert("0.0", str(from_value))
            from_cell.configure(state="disabled")  # Read-only but selectable
            from_cell.grid(row=1, column=0, sticky="ew", padx=2, pady=(0, 2))
            
            to_cell = ctk.CTkTextbox(table_frame,
                                   height=cell_height,
                                   wrap="word",
                                   activate_scrollbars=True,
                                   font=ctk.CTkFont(size=13),
                                   fg_color=("white", "gray25"),
                                   corner_radius=3)
            to_cell.insert("0.0", str(to_value))
            to_cell.configure(state="disabled")  # Read-only but selectable
            to_cell.grid(row=1, column=1, sticky="ew", padx=2, pady=(0, 2))
        
        # For creation entries, show summary
        elif entry['action'].upper() == 'INSERT':
            details_frame = ctk.CTkFrame(entry_frame, fg_color="transparent")
            details_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            details_label = ctk.CTkLabel(details_frame, text="Asset created with initial data",
                                       font=ctk.CTkFont(size=13),
                                       text_color="gray60",
                                       anchor="w", justify="left")
            details_label.pack(anchor="w", fill="x")
    
    def _refresh_history(self):
        """Refresh the history tab content."""
        self._load_history()
    
    def _on_asset_edited(self):
        """Handle asset edit callback - refresh history and call external callback."""
        # Refresh the asset data
        try:
            asset_id = self.asset.get('id')
            if asset_id:
                updated_assets = self.db.search_assets({"id": asset_id})
                if updated_assets:
                    self.asset = updated_assets[0]
                    
                    # Refresh the embedded detail view with updated data
                    if hasattr(self, 'embedded_detail'):
                        self.embedded_detail.refresh(self.asset)
                    
                    # Update window title in case asset number changed
                    self.window.title(f"Asset Details - {self.asset.get('asset_no', 'Unknown')}")
        except Exception as e:
            print(f"Error refreshing asset data: {e}")
        
        # Refresh history tab
        self._refresh_history()
        
        # Call external callback if provided
        if self.on_edit_callback:
            self.on_edit_callback()
    
    def _center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        
        # Get screen dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calculate position
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Set window position with fixed size like the original
        self.window.geometry(f"1000x700+{x}+{y}")


class MultiAssetViewer:
    """Multi-asset viewer window with navigation controls for cycling through assets."""
    
    def __init__(self, parent, assets):
        """Initialize the multi-asset viewer.
        
        Args:
            parent: Parent window
            assets: List of asset dictionaries to display
        """
        self.parent = parent
        self.assets = assets
        self.current_index = 0
        
        # Create window
        self.window = ctk.CTkToplevel(parent)
        self.window.title(f"Asset Viewer - {len(assets)} Assets")
        self.window.geometry("1000x750")
        self.window.minsize(900, 650)
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create interface
        self._create_interface()
        
        # Set up keyboard bindings
        self.window.bind("<Left>", lambda e: self._navigate(-1))
        self.window.bind("<Right>", lambda e: self._navigate(1))
        self.window.bind("<Escape>", lambda e: self.window.destroy())
        
        self.window.focus_force()
        self._center_window()
    
    def _create_interface(self):
        """Create the multi-asset viewer interface."""
        # Top navigation frame
        nav_frame = ctk.CTkFrame(self.window)
        nav_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        # Previous button
        self.prev_btn = ctk.CTkButton(nav_frame, text="◀ Previous", 
                                     command=lambda: self._navigate(-1),
                                     width=100, height=35)
        self.prev_btn.pack(side="left", padx=10, pady=10)
        
        # Asset counter/info
        self.asset_info_label = ctk.CTkLabel(nav_frame, 
                                           font=ctk.CTkFont(size=16, weight="bold"))
        self.asset_info_label.pack(side="left", expand=True, padx=20)
        
        # Next button
        self.next_btn = ctk.CTkButton(nav_frame, text="Next ▶", 
                                     command=lambda: self._navigate(1),
                                     width=100, height=35)
        self.next_btn.pack(side="right", padx=10, pady=10)
        
        # Content frame for embedded asset detail
        self.content_frame = ctk.CTkFrame(self.window)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Bottom action frame
        action_frame = ctk.CTkFrame(self.window)
        action_frame.pack(fill="x", padx=10, pady=(5, 10))
        
        # Instructions
        instructions_label = ctk.CTkLabel(action_frame, 
                                        text="💡 Use arrow keys or buttons to navigate • Press Escape to close",
                                        font=ctk.CTkFont(size=12),
                                        text_color="gray")
        instructions_label.pack(side="left", padx=15, pady=10)
        
        # Close button
        close_btn = ctk.CTkButton(action_frame, text="Close", 
                                 command=self.window.destroy,
                                 fg_color="gray", hover_color="darkgray")
        close_btn.pack(side="right", padx=15, pady=10)
        
        # Load first asset
        self._load_current_asset()
    
    def _navigate(self, direction):
        """Navigate to the next or previous asset.
        
        Args:
            direction: -1 for previous, 1 for next
        """
        if direction == -1 and self.current_index > 0:
            self.current_index -= 1
            self._load_current_asset()
        elif direction == 1 and self.current_index < len(self.assets) - 1:
            self.current_index += 1
            self._load_current_asset()
        
        # Update button states
        self._update_navigation_buttons()
    
    def _load_current_asset(self):
        """Load and display the current asset."""
        if not self.assets or self.current_index >= len(self.assets):
            return
        
        # Clear existing content
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Get current asset
        current_asset = self.assets[self.current_index]
        
        # Update asset info label
        asset_no = current_asset.get('asset_no', 'Unknown')
        manufacturer = current_asset.get('manufacturer', '')
        model = current_asset.get('model', '')
        
        info_text = f"Asset {self.current_index + 1} of {len(self.assets)}: {asset_no}"
        if manufacturer or model:
            info_text += f" ({manufacturer} {model})".strip()
        
        self.asset_info_label.configure(text=info_text)
        
        # Create embedded asset detail
        self.embedded_detail = EmbeddedAssetDetail(
            self.content_frame,
            current_asset,
            on_edit_callback=self._on_asset_edited,
            show_edit_button=True
        )
        
        # Update navigation buttons
        self._update_navigation_buttons()
    
    def _update_navigation_buttons(self):
        """Update the state of navigation buttons."""
        # Previous button
        if self.current_index <= 0:
            self.prev_btn.configure(state="disabled", fg_color="gray")
        else:
            self.prev_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
        
        # Next button
        if self.current_index >= len(self.assets) - 1:
            self.next_btn.configure(state="disabled", fg_color="gray")
        else:
            self.next_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
    
    def _on_asset_edited(self):
        """Handle asset edit callback - refresh current asset data."""
        try:
            # Refresh the current asset data from database
            current_asset = self.assets[self.current_index]
            asset_id = current_asset.get('id')
            
            if asset_id:
                # Re-query the database for updated asset data
                config_manager = ConfigManager()
                config = config_manager.get_config()
                db = AssetDatabase(config.database_path)
                
                with db.get_connection() as conn:
                    cursor = conn.execute("SELECT * FROM assets WHERE id = ? AND is_deleted = 0", (asset_id,))
                    result = cursor.fetchone()
                    
                    if result:
                        # Update the asset in our list
                        self.assets[self.current_index] = dict(result)
                        # Reload the current asset display
                        self._load_current_asset()
                
        except Exception as e:
            print(f"Error refreshing asset data: {e}")
    
    def _center_window(self):
        """Center the window on the screen."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        
        # Get screen dimensions
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        # Calculate position
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Set window position
        self.window.geometry(f"1000x750+{x}+{y}")
