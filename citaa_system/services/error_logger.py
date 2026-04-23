# -*- coding: utf-8 -*-
"""Error Logger Service"""
import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional
import config


class ErrorLogger:
    def __init__(self):
        config.ensure_directories()
        self._log_file = config.LOG_DIR / "error_log.json"
        self._errors: List[Dict] = self._load()

    def _load(self) -> List[Dict]:
        if self._log_file.exists():
            try:
                with open(self._log_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return []

    def _save(self):
        with open(self._log_file, "w", encoding="utf-8") as f:
            json.dump(self._errors, f, ensure_ascii=False, indent=2)

    def log_error(self, module: str, action: str, error: Exception, context: Optional[str] = None):
        """Log an error with details"""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "module": module,
            "action": action,
            "error_type": type(error).__name__,
            "error_message": str(error),
            "traceback": traceback.format_exc(),
            "context": context,
        }
        self._errors.append(entry)
        # Keep only last 100 errors
        if len(self._errors) > 100:
            self._errors = self._errors[-100:]
        self._save()

    def get_errors(self, limit: int = 50) -> List[Dict]:
        """Get recent errors"""
        return self._errors[-limit:][::-1]

    def clear_errors(self):
        """Clear all errors"""
        self._errors = []
        self._save()

    def has_errors(self) -> bool:
        """Check if there are any errors"""
        return len(self._errors) > 0


_error_logger: Optional[ErrorLogger] = None


def get_error_logger() -> ErrorLogger:
    global _error_logger
    if _error_logger is None:
        _error_logger = ErrorLogger()
    return _error_logger
