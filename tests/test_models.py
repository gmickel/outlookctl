"""Tests for outlookctl data models."""

import pytest
from outlookctl.models import (
    MessageId,
    EmailAddress,
    FolderInfo,
    MessageSummary,
    MessageDetail,
    ListResult,
    SearchResult,
    DraftResult,
    SendResult,
    AttachmentSaveResult,
    MoveResult,
    DeleteResult,
    MarkReadResult,
    ForwardResult,
    DoctorCheck,
    DoctorResult,
    ErrorResult,
    CalendarInfo,
    CalendarsResult,
    EventId,
    EventUpdateResult,
    EventDeleteResult,
)


class TestMessageId:
    def test_to_dict(self):
        msg_id = MessageId(entry_id="abc123", store_id="store456")
        result = msg_id.to_dict()
        assert result == {"entry_id": "abc123", "store_id": "store456"}


class TestEmailAddress:
    def test_to_dict(self):
        addr = EmailAddress(name="John Doe", email="john@example.com")
        result = addr.to_dict()
        assert result == {"name": "John Doe", "email": "john@example.com"}


class TestFolderInfo:
    def test_to_dict_minimal(self):
        folder = FolderInfo(name="Inbox")
        result = folder.to_dict()
        assert result == {"name": "Inbox"}

    def test_to_dict_with_path(self):
        folder = FolderInfo(name="Subfolder", path="Inbox/Subfolder")
        result = folder.to_dict()
        assert result == {"name": "Subfolder", "path": "Inbox/Subfolder"}


class TestMessageSummary:
    def test_to_dict_without_snippet(self):
        summary = MessageSummary(
            id=MessageId(entry_id="e1", store_id="s1"),
            received_at="2025-01-15T10:00:00",
            subject="Test Subject",
            sender=EmailAddress(name="Sender", email="sender@test.com"),
            to=["recipient@test.com"],
            cc=[],
            unread=True,
            has_attachments=False,
        )
        result = summary.to_dict()
        assert result["subject"] == "Test Subject"
        assert result["unread"] is True
        assert "body_snippet" not in result

    def test_to_dict_with_snippet(self):
        summary = MessageSummary(
            id=MessageId(entry_id="e1", store_id="s1"),
            received_at="2025-01-15T10:00:00",
            subject="Test",
            sender=EmailAddress(name="S", email="s@t.com"),
            to=[],
            cc=[],
            unread=False,
            has_attachments=False,
            body_snippet="Hello world...",
        )
        result = summary.to_dict()
        assert result["body_snippet"] == "Hello world..."


class TestListResult:
    def test_to_dict(self):
        result = ListResult(
            folder=FolderInfo(name="Inbox"),
            items=[],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["folder"]["name"] == "Inbox"
        assert output["items"] == []


class TestSearchResult:
    def test_to_dict(self):
        result = SearchResult(
            query={"from": "test@example.com"},
            items=[],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["query"]["from"] == "test@example.com"


class TestDraftResult:
    def test_to_dict_success(self):
        result = DraftResult(
            success=True,
            id=MessageId(entry_id="draft1", store_id="store1"),
            subject="Draft Subject",
            to=["recipient@test.com"],
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["id"]["entry_id"] == "draft1"
        assert output["subject"] == "Draft Subject"


class TestSendResult:
    def test_to_dict_success(self):
        result = SendResult(
            success=True,
            message="Sent successfully",
            sent_at="2025-01-15T10:00:00",
            to=["recipient@test.com"],
            subject="Test",
        )
        output = result.to_dict()
        assert output["success"] is True
        assert output["sent_at"] == "2025-01-15T10:00:00"


class TestDoctorResult:
    def test_to_dict(self):
        result = DoctorResult(
            all_passed=True,
            checks=[
                DoctorCheck(
                    name="test_check",
                    passed=True,
                    message="Check passed",
                )
            ],
            outlook_path="C:\\OUTLOOK.EXE",
        )
        output = result.to_dict()
        assert output["all_passed"] is True
        assert len(output["checks"]) == 1
        assert output["outlook_path"] == "C:\\OUTLOOK.EXE"


class TestErrorResult:
    def test_to_dict(self):
        result = ErrorResult(
            error="Something went wrong",
            error_code="TEST_ERROR",
            remediation="Try again",
        )
        output = result.to_dict()
        assert output["success"] is False
        assert output["error"] == "Something went wrong"
        assert output["error_code"] == "TEST_ERROR"
        assert output["remediation"] == "Try again"


class TestMoveResult:
    def test_to_dict(self):
        result = MoveResult(
            success=True,
            message="Message moved",
            id=MessageId(entry_id="new_e1", store_id="new_s1"),
            moved_to="Archive",
            subject="Test Subject",
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["success"] is True
        assert output["moved_to"] == "Archive"
        assert output["id"]["entry_id"] == "new_e1"

    def test_to_dict_minimal(self):
        result = MoveResult(
            success=True,
            message="Moved",
            moved_to="Inbox",
        )
        output = result.to_dict()
        assert output["success"] is True
        assert "id" not in output


class TestDeleteResult:
    def test_to_dict(self):
        result = DeleteResult(
            success=True,
            message="Message deleted",
            subject="Deleted Subject",
            permanent=False,
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["success"] is True
        assert output["permanent"] is False

    def test_to_dict_permanent(self):
        result = DeleteResult(
            success=True,
            message="Permanently deleted",
            subject="Gone",
            permanent=True,
        )
        output = result.to_dict()
        assert output["permanent"] is True


class TestMarkReadResult:
    def test_to_dict_read(self):
        result = MarkReadResult(
            success=True,
            message="Marked as read",
            count=1,
            marked_as="read",
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["marked_as"] == "read"

    def test_to_dict_unread(self):
        result = MarkReadResult(
            success=True,
            message="Marked as unread",
            count=1,
            marked_as="unread",
        )
        output = result.to_dict()
        assert output["marked_as"] == "unread"


class TestForwardResult:
    def test_to_dict(self):
        result = ForwardResult(
            success=True,
            id=MessageId(entry_id="fwd1", store_id="s1"),
            saved_to="Drafts",
            original_subject="Original",
            to=["recipient@test.com"],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["success"] is True
        assert output["original_subject"] == "Original"
        assert output["to"] == ["recipient@test.com"]


class TestCalendarInfo:
    def test_to_dict(self):
        info = CalendarInfo(
            name="Family",
            path="Family",
            store="gordon@outlook.com",
        )
        output = info.to_dict()
        assert output["name"] == "Family"
        assert output["path"] == "Family"
        assert output["store"] == "gordon@outlook.com"


class TestCalendarsResult:
    def test_to_dict(self):
        result = CalendarsResult(
            calendars=[
                CalendarInfo(name="Calendar", path="Calendar", store="main"),
                CalendarInfo(name="Family", path="Family", store="shared"),
            ]
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert len(output["calendars"]) == 2
        assert output["calendars"][0]["name"] == "Calendar"

    def test_to_dict_empty(self):
        result = CalendarsResult(calendars=[])
        output = result.to_dict()
        assert output["calendars"] == []


class TestEventUpdateResult:
    def test_to_dict(self):
        result = EventUpdateResult(
            success=True,
            message="Event updated",
            id=EventId(entry_id="evt1", store_id="s1"),
            updated_fields=["subject", "location"],
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["success"] is True
        assert output["updated_fields"] == ["subject", "location"]


class TestEventDeleteResult:
    def test_to_dict(self):
        result = EventDeleteResult(
            success=True,
            message="Event deleted",
            subject="Team Meeting",
            cancelled=False,
        )
        output = result.to_dict()
        assert output["version"] == "1.0"
        assert output["success"] is True
        assert output["cancelled"] is False

    def test_to_dict_with_cancellations(self):
        result = EventDeleteResult(
            success=True,
            message="Event cancelled",
            subject="Team Meeting",
            cancelled=True,
        )
        output = result.to_dict()
        assert output["cancelled"] is True
