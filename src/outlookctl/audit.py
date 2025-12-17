"""
Audit logging for outlookctl operations.

Provides secure logging of send operations with metadata-only by default.
Logs are stored in %LOCALAPPDATA%/outlookctl/audit.log on Windows.
"""

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


def _warn_audit_failure(operation: str, error: Exception) -> None:
    """Emit a warning to stderr when audit logging fails."""
    print(
        f"Warning: Failed to write audit log for {operation}: {error}",
        file=sys.stderr,
    )


def get_audit_log_path() -> Path:
    """Get the path to the audit log file."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base_dir = Path(local_app_data) / "outlookctl"
    else:
        # Fallback for non-Windows or missing env var
        base_dir = Path.home() / ".outlookctl"

    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir / "audit.log"


def log_send_operation(
    to: list[str],
    cc: list[str],
    bcc: list[str],
    subject: str,
    success: bool,
    error: Optional[str] = None,
    entry_id: Optional[str] = None,
    log_body: bool = False,
    body: Optional[str] = None,
) -> None:
    """
    Log a send operation to the audit log.

    Args:
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        subject: Email subject
        success: Whether the send was successful
        error: Error message if failed
        entry_id: Message entry ID if available
        log_body: Whether to include body in log (default False)
        body: Email body (only logged if log_body is True)
    """
    log_entry = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "operation": "send",
        "success": success,
        "recipients": {
            "to_count": len(to),
            "cc_count": len(cc),
            "bcc_count": len(bcc),
        },
        "subject_length": len(subject) if subject else 0,
    }

    # Add entry_id if available
    if entry_id:
        log_entry["entry_id"] = entry_id

    # Add error if present
    if error:
        log_entry["error"] = error

    # Only log body if explicitly requested
    if log_body and body:
        log_entry["body"] = body

    # Write to log file
    log_path = get_audit_log_path()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(log_entry) + "\n")
    except OSError as e:
        # Warn but don't fail - the operation should still succeed
        _warn_audit_failure("send", e)


def log_draft_operation(
    to: list[str],
    cc: list[str],
    bcc: list[str],
    subject: str,
    success: bool,
    entry_id: Optional[str] = None,
    error: Optional[str] = None,
) -> None:
    """
    Log a draft creation operation to the audit log.

    Args:
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        subject: Email subject
        success: Whether the draft was created successfully
        entry_id: Draft entry ID if available
        error: Error message if failed
    """
    log_entry = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "operation": "draft",
        "success": success,
        "recipients": {
            "to_count": len(to),
            "cc_count": len(cc),
            "bcc_count": len(bcc),
        },
        "subject_length": len(subject) if subject else 0,
    }

    if entry_id:
        log_entry["entry_id"] = entry_id

    if error:
        log_entry["error"] = error

    log_path = get_audit_log_path()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(log_entry) + "\n")
    except OSError as e:
        # Warn but don't fail - the operation should still succeed
        _warn_audit_failure("draft", e)
