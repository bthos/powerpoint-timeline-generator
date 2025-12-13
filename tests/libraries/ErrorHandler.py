"""
Error Handler Library for Robot Framework
Provides error handling, retry logic, and audit logging capabilities.
"""
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional


class ErrorHandler:
    """Error handling and retry logic for Robot Framework keywords."""

    def __init__(self):
        self.audit_log = []
        self.log_dir = Path("logs")
        self.log_dir.mkdir(exist_ok=True)

    def execute_with_retry(
        self,
        keyword: Callable,
        max_attempts: int = 3,
        retry_delay: float = 1.0,
        *args,
        **kwargs
    ) -> Any:
        """
        Execute a keyword with retry logic.

        Args:
            keyword: The keyword function to execute
            max_attempts: Maximum number of retry attempts
            retry_delay: Delay between retries in seconds
            *args: Positional arguments for the keyword
            **kwargs: Keyword arguments for the keyword

        Returns:
            Result from the keyword execution

        Raises:
            Exception: If all retry attempts fail
        """
        last_exception = None
        for attempt in range(1, max_attempts + 1):
            try:
                result = keyword(*args, **kwargs)
                if attempt > 1:
                    self.log_audit_entry(
                        action=f"Retry successful on attempt {attempt}",
                        status="success",
                        details={"keyword": keyword.__name__, "attempt": attempt}
                    )
                return result
            except Exception as e:
                last_exception = e
                if attempt < max_attempts:
                    self.log_audit_entry(
                        action=f"Retry attempt {attempt} failed",
                        status="warning",
                        details={
                            "keyword": keyword.__name__,
                            "attempt": attempt,
                            "error": str(e)
                        }
                    )
                    import time
                    time.sleep(retry_delay)
                else:
                    self.log_audit_entry(
                        action=f"All retry attempts exhausted",
                        status="error",
                        details={
                            "keyword": keyword.__name__,
                            "max_attempts": max_attempts,
                            "error": str(e)
                        }
                    )

        raise last_exception

    def log_audit_entry(
        self,
        action: str,
        status: str = "info",
        details: Optional[dict] = None
    ):
        """
        Log an audit entry for compliance and tracking.

        Args:
            action: Description of the action performed
            status: Status of the action (info, success, warning, error)
            details: Additional details as a dictionary
        """
        entry = {
            "timestamp": datetime.now().isoformat(),
            "action": action,
            "status": status,
            "details": details or {}
        }
        self.audit_log.append(entry)

    def save_audit_log(self, filename: Optional[str] = None):
        """
        Save audit log to a JSON file.

        Args:
            filename: Optional custom filename. If not provided, uses timestamp.
        """
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"audit_log_{timestamp}.json"

        log_path = self.log_dir / filename
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(self.audit_log, f, indent=2, ensure_ascii=False)

        return str(log_path)

    def handle_error(self, error: Exception, context: str = ""):
        """
        Handle errors with consistent logging.

        Args:
            error: The exception that occurred
            context: Additional context about where the error occurred
        """
        error_details = {
            "error_type": type(error).__name__,
            "error_message": str(error),
            "context": context
        }

        self.log_audit_entry(
            action=f"Error occurred: {context}",
            status="error",
            details=error_details
        )

        return error_details

    def clear_audit_log(self):
        """Clear the current audit log."""
        self.audit_log = []

