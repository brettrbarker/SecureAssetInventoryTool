"""
Centralized error handling and logging system.
Provides consistent error reporting and debugging capabilities.
"""

import logging
import os
from datetime import datetime
from typing import Optional, Any
from tkinter import messagebox
import traceback

class AppLogger:
    """Centralized logging system for the application."""
    
    def __init__(self, log_file: str = "assets/app.log", level: int = logging.INFO):
        self.log_file = log_file
        self._setup_logger(level)
    
    def _setup_logger(self, level: int):
        """Setup the logging configuration."""
        os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # File handler
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        
        # Setup logger
        self.logger = logging.getLogger('AssetInventoryTool')
        self.logger.setLevel(level)
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def info(self, message: str, **kwargs):
        """Log info message."""
        self.logger.info(message, **kwargs)
    
    def warning(self, message: str, **kwargs):
        """Log warning message."""
        self.logger.warning(message, **kwargs)
    
    def error(self, message: str, exception: Exception = None, **kwargs):
        """Log error message with optional exception details."""
        if exception:
            self.logger.error(f"{message}: {str(exception)}", exc_info=True, **kwargs)
        else:
            self.logger.error(message, **kwargs)
    
    def debug(self, message: str, **kwargs):
        """Log debug message."""
        self.logger.debug(message, **kwargs)

class ErrorHandler:
    """Centralized error handling with user-friendly messages."""
    
    def __init__(self, logger: AppLogger = None):
        self.logger = logger or AppLogger()
    
    def handle_exception(self, exception: Exception, context: str = "", 
                        show_to_user: bool = True, parent_window = None) -> bool:
        """Handle an exception with logging and optional user notification."""
        error_msg = f"Error in {context}: {str(exception)}" if context else str(exception)
        self.logger.error(error_msg, exception=exception)
        
        if show_to_user:
            user_msg = self._get_user_friendly_message(exception, context)
            messagebox.showerror("Error", user_msg, parent=parent_window)
        
        return False  # Indicate operation failed
    
    def _get_user_friendly_message(self, exception: Exception, context: str) -> str:
        """Convert technical exception to user-friendly message."""
        if "database" in str(exception).lower():
            return f"Database error: {context}\n\nPlease check your database connection and try again."
        elif "file" in str(exception).lower() or "csv" in str(exception).lower():
            return f"File error: {context}\n\nPlease check the file path and permissions."
        elif "permission" in str(exception).lower():
            return f"Permission error: {context}\n\nPlease check file/folder permissions."
        else:
            return f"An error occurred: {context}\n\n{str(exception)}"
    
    def log_operation(self, operation: str, success: bool, details: str = ""):
        """Log the result of an operation."""
        if success:
            self.logger.info(f"Operation successful: {operation}. {details}")
        else:
            self.logger.error(f"Operation failed: {operation}. {details}")

def safe_execute(func, *args, error_handler: ErrorHandler = None, 
                context: str = "", default_return: Any = None, **kwargs):
    """Safely execute a function with error handling."""
    if error_handler is None:
        error_handler = ErrorHandler()
    
    try:
        return func(*args, **kwargs)
    except Exception as e:
        error_handler.handle_exception(e, context)
        return default_return

# Global instances
app_logger = AppLogger()
error_handler = ErrorHandler(app_logger)
