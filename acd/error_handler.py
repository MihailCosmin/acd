"""
Centralized error handling and email notification system for CMM Automation.
This module provides decorators and context managers to automatically handle
errors and send debug emails without duplicating code across multiple files.
"""

import functools
import inspect
from datetime import datetime
from traceback import format_exc
from typing import Any, Callable, Dict, List, Optional, Union
from contextlib import contextmanager

from .send_email import send_email_with_attachments


class ErrorHandler:
    """Centralized error handler that automatically captures context and sends debug emails."""
    
    def __init__(self, 
                 default_recipients: List[str] = None,
                 debug: bool = False):
        self.default_recipients = default_recipients or ["munteanu@althom.de"]
        self.debug = debug
        
    def with_error_handling(self, 
                          function_name: str = None,
                          attachments: List[str] = None,
                          recipients: List[str] = None,
                          capture_args: bool = True,
                          reraise: bool = False):
        """
        Decorator that adds automatic error handling and email notifications.
        
        Args:
            function_name: Override the function name in email subject
            attachments: List of file paths to attach (can include function arguments)
            recipients: Override default recipients
            capture_args: Whether to capture function arguments for debugging
            reraise: Whether to reraise the exception after handling
        """
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                try:
                    return func(*args, **kwargs)
                except Exception as err:
                    self._handle_error(
                        error=err,
                        function_name=function_name or func.__name__,
                        function_args=args if capture_args else None,
                        function_kwargs=kwargs if capture_args else None,
                        attachments=attachments,
                        recipients=recipients
                    )
                    if reraise:
                        raise
                    return 1  # Return error code
            return wrapper
        return decorator
    
    @contextmanager
    def error_context(self, 
                     operation_name: str,
                     attachments: List[str] = None,
                     recipients: List[str] = None,
                     context_data: Dict[str, Any] = None,
                     reraise: bool = True,
                     console: Any = None):
        """
        Context manager for error handling in specific code blocks.
        
        Args:
            operation_name: Name of the operation for email subject
            attachments: List of file paths to attach
            recipients: Override default recipients
            context_data: Additional context data to include in error report
            reraise: Whether to re-raise the exception (True) or continue execution (False)
            console: Console/UI object to emit error messages to
        """
        try:
            yield
        except Exception as err:
            # Handle the error (send email, etc.)
            self._handle_error(
                error=err,
                function_name=operation_name,
                attachments=attachments,
                recipients=recipients,
                context_data=context_data
            )
            
            # Emit to console if provided
            if console is not None:
                console.emit(f"Error in {operation_name}: {err}\n{format_exc()}")
            
            if reraise:
                raise
    
    def _handle_error(self, 
                     error: Exception,
                     function_name: str,
                     function_args: tuple = None,
                     function_kwargs: dict = None,
                     attachments: List[str] = None,
                     recipients: List[str] = None,
                     context_data: Dict[str, Any] = None):
        """Internal method to handle error processing and email sending."""
        
        now = datetime.now()
        formatted_now = now.strftime("%Y-%m-%d %H:%M:%S")
        
        # Build error message
        error_message = f"{function_name} failed\n{error}\n{format_exc()}"
        
        # Add function arguments to error message if available
        if function_args or function_kwargs:
            error_message += "\n\nFunction Call Details:\n"
            if function_args:
                # Filter out sensitive data and keep only relevant file paths
                safe_args = self._extract_safe_args(function_args)
                error_message += f"Arguments: {safe_args}\n"
            if function_kwargs:
                safe_kwargs = self._extract_safe_kwargs(function_kwargs)
                error_message += f"Keyword Arguments: {safe_kwargs}\n"
        
        # Add additional context data
        if context_data:
            error_message += f"\nContext Data: {context_data}\n"
        
        # Process attachments - extract file paths from function arguments if needed
        final_attachments = self._process_attachments(
            attachments, function_args, function_kwargs
        )
        
        # Send email
        try:
            send_email_with_attachments(
                subject=f"{function_name} - {formatted_now}",
                body=error_message,
                attachments=final_attachments,
                recipients=recipients or self.default_recipients
            )
            if self.debug:
                print(f"Error email sent for {function_name}")
        except Exception as email_err:
            # Fallback: log email sending failure
            print(f"Failed to send error email: {email_err}")
    
    def _extract_safe_args(self, args: tuple) -> List[str]:
        """Extract safe, relevant arguments (mainly file paths) from function args."""
        safe_args = []
        for arg in args:
            if isinstance(arg, str) and ('/' in arg or '\\' in arg):
                # Likely a file path
                safe_args.append(arg)
            elif isinstance(arg, (list, tuple)) and arg:
                # Check if it's a list of file paths
                if isinstance(arg[0], str) and ('/' in arg[0] or '\\' in arg[0]):
                    safe_args.extend(arg)
        return safe_args
    
    def _extract_safe_kwargs(self, kwargs: dict) -> Dict[str, Any]:
        """Extract safe keyword arguments, filtering out sensitive data."""
        safe_kwargs = {}
        safe_keys = ['debug', 'checks', 'export_path', 'baseline_report']
        for key, value in kwargs.items():
            if key in safe_keys or (isinstance(value, str) and ('/' in value or '\\' in value)):
                safe_kwargs[key] = value
        return safe_kwargs
    
    def _process_attachments(self, 
                           explicit_attachments: List[str],
                           function_args: tuple,
                           function_kwargs: dict) -> List[str]:
        """Process and combine explicit attachments with file paths from function arguments."""
        attachments = []
        
        # Add explicit attachments
        if explicit_attachments:
            attachments.extend([att for att in explicit_attachments if att is not None])
        
        # Auto-detect file paths from function arguments
        if function_args:
            file_args = self._extract_safe_args(function_args)
            attachments.extend(file_args)
        
        if function_kwargs:
            for key, value in function_kwargs.items():
                if isinstance(value, str) and ('/' in value or '\\' in value):
                    attachments.append(value)
                elif isinstance(value, (list, tuple)) and value:
                    if isinstance(value[0], str) and ('/' in value[0] or '\\' in value[0]):
                        attachments.extend(value)
        
        # Remove duplicates and None values
        return list(set([att for att in attachments if att is not None]))


# Create a default instance
default_error_handler = ErrorHandler()

# Convenience functions
def with_email_on_error(function_name: str = None, 
                       attachments: List[str] = None,
                       recipients: List[str] = None):
    """Convenience decorator using the default error handler."""
    return default_error_handler.with_error_handling(
        function_name=function_name,
        attachments=attachments,
        recipients=recipients
    )

def error_context(operation_name: str,
                 attachments: List[str] = None,
                 recipients: List[str] = None,
                 context_data: Dict[str, Any] = None,
                 reraise: bool = True,
                 console: Any = None):
    """Convenience context manager using the default error handler."""
    return default_error_handler.error_context(
        operation_name=operation_name,
        attachments=attachments,
        recipients=recipients,
        context_data=context_data,
        reraise=reraise,
        console=console
    )
