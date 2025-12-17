"""
Data models for outlookctl JSON output.

All models use dataclasses and provide to_dict() methods for JSON serialization.
"""

from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Optional


@dataclass
class MessageId:
    """Stable identifier for an Outlook message."""
    entry_id: str
    store_id: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class EmailAddress:
    """Email address with optional display name."""
    name: str
    email: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class FolderInfo:
    """Basic folder information."""
    name: str
    path: Optional[str] = None
    store_id: Optional[str] = None

    def to_dict(self) -> dict:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class MessageSummary:
    """Summary of an email message for list/search results."""
    id: MessageId
    received_at: str
    subject: str
    sender: EmailAddress
    to: list[str]
    cc: list[str]
    unread: bool
    has_attachments: bool
    body_snippet: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "id": self.id.to_dict(),
            "received_at": self.received_at,
            "subject": self.subject,
            "from": self.sender.to_dict(),
            "to": self.to,
            "cc": self.cc,
            "unread": self.unread,
            "has_attachments": self.has_attachments,
        }
        if self.body_snippet is not None:
            result["body_snippet"] = self.body_snippet
        return result


@dataclass
class MessageDetail:
    """Full message details for get command."""
    id: MessageId
    received_at: str
    subject: str
    sender: EmailAddress
    to: list[str]
    cc: list[str]
    bcc: list[str]
    unread: bool
    has_attachments: bool
    attachments: list[str]
    body: Optional[str] = None
    body_html: Optional[str] = None
    headers: Optional[dict[str, str]] = None

    def to_dict(self) -> dict:
        result = {
            "id": self.id.to_dict(),
            "received_at": self.received_at,
            "subject": self.subject,
            "from": self.sender.to_dict(),
            "to": self.to,
            "cc": self.cc,
            "bcc": self.bcc,
            "unread": self.unread,
            "has_attachments": self.has_attachments,
            "attachments": self.attachments,
        }
        if self.body is not None:
            result["body"] = self.body
        if self.body_html is not None:
            result["body_html"] = self.body_html
        if self.headers is not None:
            result["headers"] = self.headers
        return result


@dataclass
class ListResult:
    """Result of a list operation."""
    version: str = "1.0"
    folder: FolderInfo = field(default_factory=lambda: FolderInfo(name="Inbox"))
    items: list[MessageSummary] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "folder": self.folder.to_dict(),
            "items": [item.to_dict() for item in self.items],
        }


@dataclass
class SearchResult:
    """Result of a search operation."""
    version: str = "1.0"
    query: dict = field(default_factory=dict)
    items: list[MessageSummary] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "query": self.query,
            "items": [item.to_dict() for item in self.items],
        }


@dataclass
class DraftResult:
    """Result of a draft operation."""
    version: str = "1.0"
    success: bool = True
    id: Optional[MessageId] = None
    saved_to: str = "Drafts"
    subject: Optional[str] = None
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "saved_to": self.saved_to,
        }
        if self.id:
            result["id"] = self.id.to_dict()
        if self.subject:
            result["subject"] = self.subject
        if self.to:
            result["to"] = self.to
        if self.cc:
            result["cc"] = self.cc
        if self.attachments:
            result["attachments"] = self.attachments
        return result


@dataclass
class SendResult:
    """Result of a send operation."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    sent_at: Optional[str] = None
    to: list[str] = field(default_factory=list)
    subject: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
        }
        if self.sent_at:
            result["sent_at"] = self.sent_at
        if self.to:
            result["to"] = self.to
        if self.subject:
            result["subject"] = self.subject
        return result


@dataclass
class AttachmentSaveResult:
    """Result of saving attachments."""
    version: str = "1.0"
    success: bool = True
    saved_files: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "success": self.success,
            "saved_files": self.saved_files,
            "errors": self.errors,
        }


@dataclass
class MoveResult:
    """Result of moving a message."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    id: Optional[MessageId] = None
    moved_to: Optional[str] = None
    subject: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
        }
        if self.id:
            result["id"] = self.id.to_dict()
        if self.moved_to:
            result["moved_to"] = self.moved_to
        if self.subject:
            result["subject"] = self.subject
        return result


@dataclass
class DeleteResult:
    """Result of deleting a message."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    subject: Optional[str] = None
    permanent: bool = False

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
            "permanent": self.permanent,
        }
        if self.subject:
            result["subject"] = self.subject
        return result


@dataclass
class MarkReadResult:
    """Result of marking messages as read/unread."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    count: int = 0
    marked_as: str = "read"  # "read" or "unread"

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "success": self.success,
            "message": self.message,
            "count": self.count,
            "marked_as": self.marked_as,
        }


@dataclass
class ForwardResult:
    """Result of creating a forward draft."""
    version: str = "1.0"
    success: bool = True
    id: Optional[MessageId] = None
    saved_to: str = "Drafts"
    original_subject: Optional[str] = None
    to: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "saved_to": self.saved_to,
        }
        if self.id:
            result["id"] = self.id.to_dict()
        if self.original_subject:
            result["original_subject"] = self.original_subject
        if self.to:
            result["to"] = self.to
        return result


@dataclass
class DoctorCheck:
    """Single check result for doctor command."""
    name: str
    passed: bool
    message: str
    remediation: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "name": self.name,
            "passed": self.passed,
            "message": self.message,
        }
        if self.remediation:
            result["remediation"] = self.remediation
        return result


@dataclass
class DoctorResult:
    """Result of doctor command."""
    version: str = "1.0"
    all_passed: bool = True
    checks: list[DoctorCheck] = field(default_factory=list)
    outlook_path: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "all_passed": self.all_passed,
            "checks": [check.to_dict() for check in self.checks],
            "outlook_path": self.outlook_path,
        }


# =============================================================================
# Calendar Models
# =============================================================================


@dataclass
class EventId:
    """Stable identifier for an Outlook calendar event."""
    entry_id: str
    store_id: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class Attendee:
    """Meeting attendee with response status."""
    name: str
    email: str
    type: str  # "required", "optional", "resource"
    response: str  # "none", "accepted", "declined", "tentative", "organizer"

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class RecurrenceInfo:
    """Recurrence pattern information."""
    type: str  # "daily", "weekly", "monthly", "monthly_nth", "yearly"
    interval: int = 1
    days_of_week: list[str] = field(default_factory=list)  # ["monday", "wednesday"]
    day_of_month: Optional[int] = None
    month_of_year: Optional[int] = None
    instance: Optional[int] = None  # For "2nd Tuesday" patterns
    end_date: Optional[str] = None
    occurrences: Optional[int] = None

    def to_dict(self) -> dict:
        result = {
            "type": self.type,
            "interval": self.interval,
        }
        if self.days_of_week:
            result["days_of_week"] = self.days_of_week
        if self.day_of_month is not None:
            result["day_of_month"] = self.day_of_month
        if self.month_of_year is not None:
            result["month_of_year"] = self.month_of_year
        if self.instance is not None:
            result["instance"] = self.instance
        if self.end_date:
            result["end_date"] = self.end_date
        if self.occurrences is not None:
            result["occurrences"] = self.occurrences
        return result


@dataclass
class EventSummary:
    """Summary of a calendar event for list results."""
    id: EventId
    subject: str
    start: str
    end: str
    location: str
    organizer: str
    is_recurring: bool
    is_all_day: bool
    is_meeting: bool
    response_status: str  # "none", "organizer", "accepted", "declined", "tentative"
    busy_status: str  # "free", "tentative", "busy", "out_of_office", "working_elsewhere"

    def to_dict(self) -> dict:
        return {
            "id": self.id.to_dict(),
            "subject": self.subject,
            "start": self.start,
            "end": self.end,
            "location": self.location,
            "organizer": self.organizer,
            "is_recurring": self.is_recurring,
            "is_all_day": self.is_all_day,
            "is_meeting": self.is_meeting,
            "response_status": self.response_status,
            "busy_status": self.busy_status,
        }


@dataclass
class EventDetail:
    """Full event details for get command."""
    id: EventId
    subject: str
    start: str
    end: str
    location: str
    organizer: str
    is_recurring: bool
    is_all_day: bool
    is_meeting: bool
    response_status: str
    busy_status: str
    body: Optional[str] = None
    attendees: list[Attendee] = field(default_factory=list)
    recurrence: Optional[RecurrenceInfo] = None
    categories: list[str] = field(default_factory=list)
    reminder_minutes: Optional[int] = None
    sensitivity: str = "normal"  # "normal", "personal", "private", "confidential"

    def to_dict(self) -> dict:
        result = {
            "id": self.id.to_dict(),
            "subject": self.subject,
            "start": self.start,
            "end": self.end,
            "location": self.location,
            "organizer": self.organizer,
            "is_recurring": self.is_recurring,
            "is_all_day": self.is_all_day,
            "is_meeting": self.is_meeting,
            "response_status": self.response_status,
            "busy_status": self.busy_status,
            "sensitivity": self.sensitivity,
        }
        if self.body is not None:
            result["body"] = self.body
        if self.attendees:
            result["attendees"] = [a.to_dict() for a in self.attendees]
        if self.recurrence:
            result["recurrence"] = self.recurrence.to_dict()
        if self.categories:
            result["categories"] = self.categories
        if self.reminder_minutes is not None:
            result["reminder_minutes"] = self.reminder_minutes
        return result


@dataclass
class CalendarListResult:
    """Result of a calendar list operation."""
    version: str = "1.0"
    calendar: str = "Calendar"
    start_date: str = ""
    end_date: str = ""
    items: list[EventSummary] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "calendar": self.calendar,
            "start_date": self.start_date,
            "end_date": self.end_date,
            "items": [item.to_dict() for item in self.items],
        }


@dataclass
class EventCreateResult:
    """Result of creating a calendar event."""
    version: str = "1.0"
    success: bool = True
    id: Optional[EventId] = None
    saved_to: str = "Calendar"
    subject: Optional[str] = None
    start: Optional[str] = None
    attendees: list[str] = field(default_factory=list)
    is_draft: bool = True  # True if meeting not sent yet

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "saved_to": self.saved_to,
            "is_draft": self.is_draft,
        }
        if self.id:
            result["id"] = self.id.to_dict()
        if self.subject:
            result["subject"] = self.subject
        if self.start:
            result["start"] = self.start
        if self.attendees:
            result["attendees"] = self.attendees
        return result


@dataclass
class EventSendResult:
    """Result of sending a meeting invitation."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    sent_at: Optional[str] = None
    attendees: list[str] = field(default_factory=list)
    subject: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
        }
        if self.sent_at:
            result["sent_at"] = self.sent_at
        if self.attendees:
            result["attendees"] = self.attendees
        if self.subject:
            result["subject"] = self.subject
        return result


@dataclass
class EventRespondResult:
    """Result of responding to a meeting invitation."""
    version: str = "1.0"
    success: bool = True
    response: str = ""  # "accepted", "declined", "tentative"
    subject: Optional[str] = None
    organizer: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "response": self.response,
        }
        if self.subject:
            result["subject"] = self.subject
        if self.organizer:
            result["organizer"] = self.organizer
        return result


@dataclass
class EventUpdateResult:
    """Result of updating a calendar event."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    id: Optional[EventId] = None
    subject: Optional[str] = None
    start: Optional[str] = None
    updated_fields: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
            "updated_fields": self.updated_fields,
        }
        if self.id:
            result["id"] = self.id.to_dict()
        if self.subject:
            result["subject"] = self.subject
        if self.start:
            result["start"] = self.start
        return result


@dataclass
class EventDeleteResult:
    """Result of deleting a calendar event."""
    version: str = "1.0"
    success: bool = True
    message: str = ""
    subject: Optional[str] = None
    cancelled: bool = False  # True if meeting cancellation sent

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "message": self.message,
            "cancelled": self.cancelled,
        }
        if self.subject:
            result["subject"] = self.subject
        return result


@dataclass
class ErrorResult:
    """Error result for any command."""
    version: str = "1.0"
    success: bool = False
    error: str = ""
    error_code: Optional[str] = None
    remediation: Optional[str] = None

    def to_dict(self) -> dict:
        result = {
            "version": self.version,
            "success": self.success,
            "error": self.error,
        }
        if self.error_code:
            result["error_code"] = self.error_code
        if self.remediation:
            result["remediation"] = self.remediation
        return result
