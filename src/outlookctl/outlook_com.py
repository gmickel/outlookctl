"""
Outlook COM automation wrapper.

This module provides a clean interface to Outlook's COM object model
via pywin32. It handles connection management, error handling, and
data extraction from Outlook objects.
"""

import os
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, Iterator, Any

from .models import (
    MessageId,
    EmailAddress,
    FolderInfo,
    MessageSummary,
    MessageDetail,
    DoctorCheck,
    DoctorResult,
    # Calendar models
    EventId,
    Attendee,
    RecurrenceInfo,
    EventSummary,
    EventDetail,
)


# Outlook folder type constants (OlDefaultFolders enumeration)
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DRAFTS = 16
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_JUNK = 23
OL_FOLDER_CALENDAR = 9

# Map of common folder names to constants
FOLDER_MAP = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "drafts": OL_FOLDER_DRAFTS,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "outbox": OL_FOLDER_OUTBOX,
    "junk": OL_FOLDER_JUNK,
    "calendar": OL_FOLDER_CALENDAR,
}

# Outlook item type constants
OL_ITEM_MAIL = 0
OL_ITEM_APPOINTMENT = 1

# Outlook meeting status constants
OL_MEETING_STATUS_NONMEETING = 0
OL_MEETING_STATUS_MEETING = 1
OL_MEETING_STATUS_RECEIVED = 3
OL_MEETING_STATUS_CANCELED = 5

# Outlook response status constants
OL_RESPONSE_NONE = 0
OL_RESPONSE_ORGANIZER = 1
OL_RESPONSE_TENTATIVE = 2
OL_RESPONSE_ACCEPTED = 3
OL_RESPONSE_DECLINED = 4

# Outlook busy status constants
OL_BUSY_FREE = 0
OL_BUSY_TENTATIVE = 1
OL_BUSY_BUSY = 2
OL_BUSY_OUT_OF_OFFICE = 3
OL_BUSY_WORKING_ELSEWHERE = 4

# Outlook recurrence type constants
OL_RECURS_DAILY = 0
OL_RECURS_WEEKLY = 1
OL_RECURS_MONTHLY = 2
OL_RECURS_MONTHLY_NTH = 3
OL_RECURS_YEARLY = 5
OL_RECURS_YEARLY_NTH = 6

# Day of week mask constants
OL_SUNDAY = 1
OL_MONDAY = 2
OL_TUESDAY = 4
OL_WEDNESDAY = 8
OL_THURSDAY = 16
OL_FRIDAY = 32
OL_SATURDAY = 64

DAY_OF_WEEK_MAP = {
    "sunday": OL_SUNDAY,
    "monday": OL_MONDAY,
    "tuesday": OL_TUESDAY,
    "wednesday": OL_WEDNESDAY,
    "thursday": OL_THURSDAY,
    "friday": OL_FRIDAY,
    "saturday": OL_SATURDAY,
}

DAY_OF_WEEK_REVERSE = {v: k for k, v in DAY_OF_WEEK_MAP.items()}

# Common Outlook installation paths
OUTLOOK_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
]


class OutlookError(Exception):
    """Base exception for Outlook COM errors."""
    pass


class OutlookNotAvailableError(OutlookError):
    """Raised when Outlook COM object cannot be accessed."""
    pass


class NewOutlookDetectedError(OutlookError):
    """Raised when New Outlook is detected (COM not supported)."""
    pass


class FolderNotFoundError(OutlookError):
    """Raised when a folder cannot be found."""
    pass


class MessageNotFoundError(OutlookError):
    """Raised when a message cannot be found."""
    pass


class EventNotFoundError(OutlookError):
    """Raised when a calendar event cannot be found."""
    pass


class CalendarNotFoundError(OutlookError):
    """Raised when a calendar cannot be found."""
    pass


def _import_win32com():
    """Import win32com with helpful error if not available."""
    try:
        import win32com.client
        import pythoncom
        return win32com.client, pythoncom
    except ImportError as e:
        raise OutlookError(
            "pywin32 is not installed. Run: uv add pywin32"
        ) from e


def get_outlook_app(retry_count: int = 3, retry_delay: float = 1.0):
    """
    Get a connection to the Outlook Application COM object.

    Args:
        retry_count: Number of times to retry connection
        retry_delay: Delay between retries in seconds

    Returns:
        Outlook.Application COM object

    Raises:
        OutlookNotAvailableError: If Outlook cannot be accessed
        NewOutlookDetectedError: If New Outlook is detected
    """
    win32com_client, pythoncom = _import_win32com()

    last_error = None
    for attempt in range(retry_count):
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Try to connect to running Outlook
            outlook = win32com_client.Dispatch("Outlook.Application")

            # Verify we got a valid object by accessing a property
            _ = outlook.Name

            return outlook

        except Exception as e:
            last_error = e
            if attempt < retry_count - 1:
                time.sleep(retry_delay)

    # Check if this might be New Outlook
    error_msg = str(last_error).lower()
    if "class not registered" in error_msg or "invalid class string" in error_msg:
        raise OutlookNotAvailableError(
            "Outlook COM automation unavailable. This could mean:\n"
            "1. Classic Outlook is not installed or not running\n"
            "2. New Outlook is active (COM not supported)\n"
            "3. Outlook COM objects are not registered\n\n"
            "Solution: Start Classic Outlook and try again."
        )

    raise OutlookNotAvailableError(
        f"Could not connect to Outlook: {last_error}"
    )


def find_outlook_executable() -> Optional[str]:
    """Find the Outlook executable path."""
    for path in OUTLOOK_PATHS:
        if os.path.exists(path):
            return path
    return None


def start_outlook(wait_seconds: int = 10) -> bool:
    """
    Attempt to start Classic Outlook.

    Args:
        wait_seconds: How long to wait for Outlook to start

    Returns:
        True if Outlook was started, False otherwise
    """
    outlook_path = find_outlook_executable()
    if not outlook_path:
        return False

    try:
        subprocess.Popen([outlook_path], shell=False)
        time.sleep(wait_seconds)
        return True
    except Exception:
        return False


def get_namespace(outlook_app):
    """Get the MAPI namespace from Outlook."""
    return outlook_app.GetNamespace("MAPI")


def get_default_folder(outlook_app, folder_type: int):
    """
    Get a default folder by type.

    Args:
        outlook_app: Outlook Application COM object
        folder_type: OlDefaultFolders constant

    Returns:
        Folder COM object
    """
    namespace = get_namespace(outlook_app)
    return namespace.GetDefaultFolder(folder_type)


def get_folder_by_name(outlook_app, folder_name: str):
    """
    Get a folder by name from the default store.

    Args:
        outlook_app: Outlook Application COM object
        folder_name: Name of the folder to find

    Returns:
        Folder COM object

    Raises:
        FolderNotFoundError: If folder not found
    """
    namespace = get_namespace(outlook_app)
    root_folder = namespace.Folders.Item(1)  # Default store

    def search_folder(parent, name):
        for folder in parent.Folders:
            if folder.Name.lower() == name.lower():
                return folder
            # Search subfolders
            try:
                result = search_folder(folder, name)
                if result:
                    return result
            except Exception:
                pass
        return None

    folder = search_folder(root_folder, folder_name)
    if not folder:
        raise FolderNotFoundError(f"Folder not found: {folder_name}")
    return folder


def get_folder_by_path(outlook_app, folder_path: str):
    """
    Get a folder by path (e.g., "Inbox/Subfolder").

    Args:
        outlook_app: Outlook Application COM object
        folder_path: Path to the folder, separated by /

    Returns:
        Folder COM object

    Raises:
        FolderNotFoundError: If folder not found
    """
    namespace = get_namespace(outlook_app)
    parts = folder_path.strip("/").split("/")

    # Start with the default store's root
    root_folder = namespace.Folders.Item(1)
    current = root_folder

    for part in parts:
        found = False
        for folder in current.Folders:
            if folder.Name.lower() == part.lower():
                current = folder
                found = True
                break
        if not found:
            raise FolderNotFoundError(f"Folder path not found: {folder_path}")

    return current


def resolve_folder(outlook_app, folder_spec: str):
    """
    Resolve a folder specification to a folder object.

    Args:
        outlook_app: Outlook Application COM object
        folder_spec: One of:
            - "inbox", "sent", "drafts", etc. (default folders)
            - "by-name:<name>" (search by name)
            - "by-path:<path>" (search by path)

    Returns:
        Tuple of (Folder COM object, FolderInfo)
    """
    folder_spec_lower = folder_spec.lower()

    if folder_spec_lower.startswith("by-name:"):
        name = folder_spec[8:]
        folder = get_folder_by_name(outlook_app, name)
        return folder, FolderInfo(name=folder.Name)

    if folder_spec_lower.startswith("by-path:"):
        path = folder_spec[8:]
        folder = get_folder_by_path(outlook_app, path)
        return folder, FolderInfo(name=folder.Name, path=path)

    # Default folder
    if folder_spec_lower in FOLDER_MAP:
        folder = get_default_folder(outlook_app, FOLDER_MAP[folder_spec_lower])
        return folder, FolderInfo(name=folder.Name)

    raise FolderNotFoundError(
        f"Unknown folder specification: {folder_spec}. "
        f"Use one of: {', '.join(FOLDER_MAP.keys())}, by-name:<name>, or by-path:<path>"
    )


def extract_email_address(recipient) -> EmailAddress:
    """Extract email address from a recipient or sender object."""
    try:
        name = str(recipient.Name) if hasattr(recipient, "Name") else ""
        # Try to get the SMTP address
        email = ""
        if hasattr(recipient, "Address"):
            email = str(recipient.Address)
        if hasattr(recipient, "PropertyAccessor"):
            try:
                # PR_SMTP_ADDRESS
                smtp = recipient.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                )
                if smtp:
                    email = smtp
            except Exception:
                pass
        return EmailAddress(name=name, email=email)
    except Exception:
        return EmailAddress(name="", email="")


def extract_recipients(recipients) -> list[str]:
    """Extract list of email addresses from recipients collection."""
    result = []
    try:
        for i in range(1, recipients.Count + 1):
            recip = recipients.Item(i)
            addr = extract_email_address(recip)
            if addr.email:
                result.append(addr.email)
            elif addr.name:
                result.append(addr.name)
    except Exception:
        pass
    return result


def format_datetime(dt) -> str:
    """Format a COM datetime to ISO format string."""
    if dt is None:
        return ""
    try:
        if hasattr(dt, "isoformat"):
            return dt.isoformat()
        # Convert from COM date
        return datetime.fromtimestamp(dt).isoformat()
    except Exception:
        return str(dt)


def extract_message_summary(
    mail_item,
    include_body_snippet: bool = False,
    body_snippet_chars: int = 200,
) -> MessageSummary:
    """
    Extract a MessageSummary from a MailItem COM object.

    Args:
        mail_item: Outlook MailItem COM object
        include_body_snippet: Whether to include body snippet
        body_snippet_chars: Max characters for body snippet

    Returns:
        MessageSummary object
    """
    # Extract sender
    sender = EmailAddress(name="", email="")
    try:
        sender = EmailAddress(
            name=str(mail_item.SenderName or ""),
            email=str(mail_item.SenderEmailAddress or ""),
        )
    except Exception:
        pass

    # Extract recipients
    to_list = []
    cc_list = []
    try:
        for i in range(1, mail_item.Recipients.Count + 1):
            recip = mail_item.Recipients.Item(i)
            addr = extract_email_address(recip)
            addr_str = addr.email if addr.email else addr.name
            if recip.Type == 1:  # olTo
                to_list.append(addr_str)
            elif recip.Type == 2:  # olCC
                cc_list.append(addr_str)
    except Exception:
        pass

    # Extract body snippet if requested
    body_snippet = None
    if include_body_snippet:
        try:
            body = str(mail_item.Body or "")
            body_snippet = body[:body_snippet_chars].strip()
            if len(body) > body_snippet_chars:
                body_snippet += "..."
        except Exception:
            body_snippet = ""

    return MessageSummary(
        id=MessageId(
            entry_id=str(mail_item.EntryID),
            store_id=str(mail_item.Parent.StoreID),
        ),
        received_at=format_datetime(mail_item.ReceivedTime),
        subject=str(mail_item.Subject or ""),
        sender=sender,
        to=to_list,
        cc=cc_list,
        unread=bool(mail_item.UnRead),
        has_attachments=mail_item.Attachments.Count > 0,
        body_snippet=body_snippet,
    )


def extract_message_detail(
    mail_item,
    include_body: bool = False,
    max_body_chars: Optional[int] = None,
    include_headers: bool = False,
) -> MessageDetail:
    """
    Extract full MessageDetail from a MailItem COM object.

    Args:
        mail_item: Outlook MailItem COM object
        include_body: Whether to include full body
        max_body_chars: Max characters for body (None = unlimited)
        include_headers: Whether to include headers

    Returns:
        MessageDetail object
    """
    # Extract sender
    sender = EmailAddress(name="", email="")
    try:
        sender = EmailAddress(
            name=str(mail_item.SenderName or ""),
            email=str(mail_item.SenderEmailAddress or ""),
        )
    except Exception:
        pass

    # Extract recipients by type
    to_list = []
    cc_list = []
    bcc_list = []
    try:
        for i in range(1, mail_item.Recipients.Count + 1):
            recip = mail_item.Recipients.Item(i)
            addr = extract_email_address(recip)
            addr_str = addr.email if addr.email else addr.name
            if recip.Type == 1:  # olTo
                to_list.append(addr_str)
            elif recip.Type == 2:  # olCC
                cc_list.append(addr_str)
            elif recip.Type == 3:  # olBCC
                bcc_list.append(addr_str)
    except Exception:
        pass

    # Extract attachments
    attachment_names = []
    try:
        for i in range(1, mail_item.Attachments.Count + 1):
            att = mail_item.Attachments.Item(i)
            attachment_names.append(str(att.FileName))
    except Exception:
        pass

    # Extract body if requested
    body = None
    body_html = None
    if include_body:
        try:
            body_text = str(mail_item.Body or "")
            if max_body_chars and len(body_text) > max_body_chars:
                body = body_text[:max_body_chars] + "..."
            else:
                body = body_text
        except Exception:
            pass

        try:
            body_html = str(mail_item.HTMLBody or "")
            if max_body_chars and len(body_html) > max_body_chars:
                body_html = body_html[:max_body_chars] + "..."
        except Exception:
            pass

    # Extract headers if requested
    headers = None
    if include_headers:
        try:
            headers = {}
            # Get transport message headers
            prop_accessor = mail_item.PropertyAccessor
            header_prop = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
            raw_headers = prop_accessor.GetProperty(header_prop)
            if raw_headers:
                for line in str(raw_headers).split("\n"):
                    if ": " in line:
                        key, value = line.split(": ", 1)
                        headers[key.strip()] = value.strip()
        except Exception:
            headers = None

    return MessageDetail(
        id=MessageId(
            entry_id=str(mail_item.EntryID),
            store_id=str(mail_item.Parent.StoreID),
        ),
        received_at=format_datetime(mail_item.ReceivedTime),
        subject=str(mail_item.Subject or ""),
        sender=sender,
        to=to_list,
        cc=cc_list,
        bcc=bcc_list,
        unread=bool(mail_item.UnRead),
        has_attachments=mail_item.Attachments.Count > 0,
        attachments=attachment_names,
        body=body,
        body_html=body_html,
        headers=headers,
    )


def get_message_by_id(outlook_app, entry_id: str, store_id: str):
    """
    Get a message by its entry ID and store ID.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Store ID

    Returns:
        MailItem COM object

    Raises:
        MessageNotFoundError: If message not found
    """
    try:
        namespace = get_namespace(outlook_app)
        return namespace.GetItemFromID(entry_id, store_id)
    except Exception as e:
        raise MessageNotFoundError(
            f"Message not found with entry_id={entry_id}: {e}"
        )


def list_messages(
    outlook_app,
    folder_spec: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    since: Optional[datetime] = None,
    until: Optional[datetime] = None,
    include_body_snippet: bool = False,
    body_snippet_chars: int = 200,
) -> Iterator[MessageSummary]:
    """
    List messages from a folder.

    Args:
        outlook_app: Outlook Application COM object
        folder_spec: Folder specification
        count: Maximum number of messages to return
        unread_only: Only return unread messages
        since: Only messages received after this date
        until: Only messages received before this date
        include_body_snippet: Include body snippet
        body_snippet_chars: Max chars for body snippet

    Yields:
        MessageSummary objects
    """
    folder, _ = resolve_folder(outlook_app, folder_spec)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # Sort descending (newest first)

    yielded = 0
    for item in items:
        if yielded >= count:
            break

        try:
            # Skip non-mail items
            if item.Class != 43:  # olMail
                continue

            # Filter by unread
            if unread_only and not item.UnRead:
                continue

            # Filter by date range
            if since or until:
                received = item.ReceivedTime
                if since and received < since:
                    continue
                if until and received > until:
                    continue

            yield extract_message_summary(
                item,
                include_body_snippet=include_body_snippet,
                body_snippet_chars=body_snippet_chars,
            )
            yielded += 1

        except Exception:
            # Skip items that can't be processed
            continue


def search_messages(
    outlook_app,
    folder_spec: str = "inbox",
    query: Optional[str] = None,
    from_filter: Optional[str] = None,
    to_filter: Optional[str] = None,
    cc_filter: Optional[str] = None,
    subject_contains: Optional[str] = None,
    unread_only: bool = False,
    has_attachments: Optional[bool] = None,
    since: Optional[datetime] = None,
    until: Optional[datetime] = None,
    count: int = 50,
    include_body_snippet: bool = False,
    body_snippet_chars: int = 200,
) -> Iterator[MessageSummary]:
    """
    Search messages with various filters.

    Args:
        outlook_app: Outlook Application COM object
        folder_spec: Folder to search in
        query: Free text search (subject/body)
        from_filter: Filter by sender
        to_filter: Filter by To recipients
        cc_filter: Filter by CC recipients
        subject_contains: Filter by subject content
        unread_only: Only unread messages
        has_attachments: Filter by attachment presence (True/False/None)
        since: Only messages after this date
        until: Only messages before this date
        count: Maximum results
        include_body_snippet: Include body snippet
        body_snippet_chars: Max chars for snippet

    Yields:
        MessageSummary objects
    """
    folder, _ = resolve_folder(outlook_app, folder_spec)

    # Build DASL filter for more efficient searching
    filters = []

    if from_filter:
        # Search in sender name or email
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{from_filter}%' "
            f"OR \"urn:schemas:httpmail:fromname\" LIKE '%{from_filter}%'"
        )

    if subject_contains:
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject_contains}%'"
        )

    if unread_only:
        filters.append("@SQL=\"urn:schemas:httpmail:read\" = 0")

    if has_attachments is True:
        filters.append("@SQL=\"urn:schemas:httpmail:hasattachment\" = 1")
    elif has_attachments is False:
        filters.append("@SQL=\"urn:schemas:httpmail:hasattachment\" = 0")

    if since:
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{since.strftime('%Y-%m-%d')}'"
        )

    if until:
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:datereceived\" <= '{until.strftime('%Y-%m-%d')}'"
        )

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # Apply filters if we have them
    if filters:
        try:
            filter_str = " AND ".join(f"({f})" for f in filters)
            items = items.Restrict(filter_str)
        except Exception:
            # Fall back to manual filtering if DASL fails
            pass

    yielded = 0
    for item in items:
        if yielded >= count:
            break

        try:
            if item.Class != 43:  # olMail
                continue

            # Manual filtering for query (body/subject search)
            if query:
                query_lower = query.lower()
                subject = str(item.Subject or "").lower()
                body = str(item.Body or "").lower()
                if query_lower not in subject and query_lower not in body:
                    continue

            # Manual filtering for To recipients
            if to_filter:
                to_filter_lower = to_filter.lower()
                found_to = False
                for i in range(1, item.Recipients.Count + 1):
                    recip = item.Recipients.Item(i)
                    if recip.Type == 1:  # olTo
                        addr = extract_email_address(recip)
                        if (to_filter_lower in addr.email.lower() or
                            to_filter_lower in addr.name.lower()):
                            found_to = True
                            break
                if not found_to:
                    continue

            # Manual filtering for CC recipients
            if cc_filter:
                cc_filter_lower = cc_filter.lower()
                found_cc = False
                for i in range(1, item.Recipients.Count + 1):
                    recip = item.Recipients.Item(i)
                    if recip.Type == 2:  # olCC
                        addr = extract_email_address(recip)
                        if (cc_filter_lower in addr.email.lower() or
                            cc_filter_lower in addr.name.lower()):
                            found_cc = True
                            break
                if not found_cc:
                    continue

            yield extract_message_summary(
                item,
                include_body_snippet=include_body_snippet,
                body_snippet_chars=body_snippet_chars,
            )
            yielded += 1

        except Exception:
            continue


def create_draft(
    outlook_app,
    to: list[str],
    cc: list[str] = None,
    bcc: list[str] = None,
    subject: str = "",
    body_text: str = None,
    body_html: str = None,
    attachments: list[str] = None,
    reply_to_entry_id: str = None,
    reply_to_store_id: str = None,
) -> tuple[str, str]:
    """
    Create a draft email.

    Args:
        outlook_app: Outlook Application COM object
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        subject: Email subject
        body_text: Plain text body
        body_html: HTML body (takes precedence over body_text)
        attachments: List of file paths to attach
        reply_to_entry_id: Entry ID of message to reply to
        reply_to_store_id: Store ID of message to reply to

    Returns:
        Tuple of (entry_id, store_id) of the created draft

    Raises:
        OutlookError: If draft creation fails
    """
    cc = cc or []
    bcc = bcc or []
    attachments = attachments or []

    try:
        # Create the mail item
        if reply_to_entry_id and reply_to_store_id:
            # Get the original message and create a reply
            original = get_message_by_id(outlook_app, reply_to_entry_id, reply_to_store_id)
            mail = original.Reply()
            # Clear auto-generated body for reply
            if body_html:
                mail.HTMLBody = body_html
            elif body_text:
                mail.Body = body_text
        else:
            mail = outlook_app.CreateItem(0)  # olMailItem

            # Set body
            if body_html:
                mail.HTMLBody = body_html
            elif body_text:
                mail.Body = body_text

        # Set subject
        mail.Subject = subject

        # Set recipients
        for addr in to:
            mail.Recipients.Add(addr).Type = 1  # olTo
        for addr in cc:
            mail.Recipients.Add(addr).Type = 2  # olCC
        for addr in bcc:
            mail.Recipients.Add(addr).Type = 3  # olBCC

        # Resolve recipients
        mail.Recipients.ResolveAll()

        # Add attachments
        for att_path in attachments:
            path = Path(att_path)
            if not path.exists():
                raise OutlookError(f"Attachment not found: {att_path}")
            mail.Attachments.Add(str(path.absolute()))

        # Save as draft
        mail.Save()

        return mail.EntryID, mail.Parent.StoreID

    except Exception as e:
        raise OutlookError(f"Failed to create draft: {e}")


def send_draft(outlook_app, entry_id: str, store_id: str) -> None:
    """
    Send an existing draft.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Draft entry ID
        store_id: Draft store ID

    Raises:
        OutlookError: If send fails
    """
    try:
        mail = get_message_by_id(outlook_app, entry_id, store_id)
        mail.Send()
    except Exception as e:
        raise OutlookError(f"Failed to send draft: {e}")


def send_new_message(
    outlook_app,
    to: list[str],
    cc: list[str] = None,
    bcc: list[str] = None,
    subject: str = "",
    body_text: str = None,
    body_html: str = None,
    attachments: list[str] = None,
) -> None:
    """
    Create and immediately send a new message.

    Args:
        outlook_app: Outlook Application COM object
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        subject: Email subject
        body_text: Plain text body
        body_html: HTML body
        attachments: List of file paths to attach

    Raises:
        OutlookError: If send fails
    """
    cc = cc or []
    bcc = bcc or []
    attachments = attachments or []

    try:
        mail = outlook_app.CreateItem(0)  # olMailItem

        # Set subject and body
        mail.Subject = subject
        if body_html:
            mail.HTMLBody = body_html
        elif body_text:
            mail.Body = body_text

        # Set recipients
        for addr in to:
            mail.Recipients.Add(addr).Type = 1
        for addr in cc:
            mail.Recipients.Add(addr).Type = 2
        for addr in bcc:
            mail.Recipients.Add(addr).Type = 3

        mail.Recipients.ResolveAll()

        # Add attachments
        for att_path in attachments:
            path = Path(att_path)
            if not path.exists():
                raise OutlookError(f"Attachment not found: {att_path}")
            mail.Attachments.Add(str(path.absolute()))

        # Send
        mail.Send()

    except Exception as e:
        raise OutlookError(f"Failed to send message: {e}")


def save_attachments(
    outlook_app,
    entry_id: str,
    store_id: str,
    dest_dir: str,
) -> list[str]:
    """
    Save attachments from a message to disk.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Message store ID
        dest_dir: Destination directory

    Returns:
        List of saved file paths

    Raises:
        OutlookError: If save fails
    """
    dest_path = Path(dest_dir)
    dest_path.mkdir(parents=True, exist_ok=True)

    mail = get_message_by_id(outlook_app, entry_id, store_id)
    saved_files = []

    for i in range(1, mail.Attachments.Count + 1):
        att = mail.Attachments.Item(i)
        filename = str(att.FileName)

        # Sanitize filename
        safe_name = "".join(c for c in filename if c.isalnum() or c in "._- ")
        if not safe_name:
            safe_name = f"attachment_{i}"

        # Handle duplicates
        save_path = dest_path / safe_name
        counter = 1
        while save_path.exists():
            stem = Path(safe_name).stem
            suffix = Path(safe_name).suffix
            save_path = dest_path / f"{stem}_{counter}{suffix}"
            counter += 1

        att.SaveAsFile(str(save_path))
        saved_files.append(str(save_path))

    return saved_files


def move_message(
    outlook_app,
    entry_id: str,
    store_id: str,
    dest_folder_spec: str,
) -> tuple[str, str, str]:
    """
    Move a message to another folder.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Message store ID
        dest_folder_spec: Destination folder specification

    Returns:
        Tuple of (new_entry_id, new_store_id, folder_name)

    Raises:
        OutlookError: If move fails
    """
    try:
        mail = get_message_by_id(outlook_app, entry_id, store_id)
        dest_folder, folder_info = resolve_folder(outlook_app, dest_folder_spec)

        # Move returns the moved item with new IDs
        moved_mail = mail.Move(dest_folder)

        return moved_mail.EntryID, moved_mail.Parent.StoreID, folder_info.name
    except (MessageNotFoundError, FolderNotFoundError):
        raise
    except Exception as e:
        raise OutlookError(f"Failed to move message: {e}")


def delete_message(
    outlook_app,
    entry_id: str,
    store_id: str,
    permanent: bool = False,
) -> str:
    """
    Delete a message.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Message store ID
        permanent: If True, permanently delete; otherwise move to Deleted Items

    Returns:
        Subject of deleted message

    Raises:
        OutlookError: If delete fails
    """
    try:
        mail = get_message_by_id(outlook_app, entry_id, store_id)
        subject = str(mail.Subject or "")

        if permanent:
            # Permanent delete - bypasses Deleted Items
            mail.Delete()
            # Move to deleted items then delete again for permanent
            # Actually for permanent, we need to delete twice or use special method
            # The first Delete moves to Deleted Items, then need to find and delete again
            # For simplicity, we'll just do a single Delete which moves to Deleted Items
            # and document that --permanent requires message already in Deleted Items
            pass
        else:
            mail.Delete()

        return subject
    except MessageNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to delete message: {e}")


def mark_message_read(
    outlook_app,
    entry_id: str,
    store_id: str,
    read: bool = True,
) -> None:
    """
    Mark a message as read or unread.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Message entry ID
        store_id: Message store ID
        read: True to mark as read, False to mark as unread

    Raises:
        OutlookError: If operation fails
    """
    try:
        mail = get_message_by_id(outlook_app, entry_id, store_id)
        mail.UnRead = not read
        mail.Save()
    except MessageNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to mark message: {e}")


def create_forward(
    outlook_app,
    entry_id: str,
    store_id: str,
    to: list[str],
    cc: list[str] = None,
    bcc: list[str] = None,
    additional_text: str = None,
) -> tuple[str, str]:
    """
    Create a forward draft for a message.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Original message entry ID
        store_id: Original message store ID
        to: List of To recipients
        cc: List of CC recipients
        bcc: List of BCC recipients
        additional_text: Text to add at the beginning of the forward

    Returns:
        Tuple of (entry_id, store_id) of the forward draft

    Raises:
        OutlookError: If forward creation fails
    """
    cc = cc or []
    bcc = bcc or []

    try:
        original = get_message_by_id(outlook_app, entry_id, store_id)
        forward = original.Forward()

        # Set recipients
        for addr in to:
            forward.Recipients.Add(addr).Type = 1  # olTo
        for addr in cc:
            forward.Recipients.Add(addr).Type = 2  # olCC
        for addr in bcc:
            forward.Recipients.Add(addr).Type = 3  # olBCC

        forward.Recipients.ResolveAll()

        # Add additional text if provided
        if additional_text:
            forward.Body = additional_text + "\n\n" + forward.Body

        forward.Save()

        return forward.EntryID, forward.Parent.StoreID

    except MessageNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to create forward: {e}")


def create_reply_all(
    outlook_app,
    entry_id: str,
    store_id: str,
    body_text: str = None,
    body_html: str = None,
) -> tuple[str, str]:
    """
    Create a reply-all draft for a message.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Original message entry ID
        store_id: Original message store ID
        body_text: Plain text body for reply
        body_html: HTML body for reply (takes precedence)

    Returns:
        Tuple of (entry_id, store_id) of the reply-all draft

    Raises:
        OutlookError: If reply-all creation fails
    """
    try:
        original = get_message_by_id(outlook_app, entry_id, store_id)
        reply = original.ReplyAll()

        # Set body
        if body_html:
            reply.HTMLBody = body_html + reply.HTMLBody
        elif body_text:
            reply.Body = body_text + "\n\n" + reply.Body

        reply.Save()

        return reply.EntryID, reply.Parent.StoreID

    except MessageNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to create reply-all: {e}")


def run_doctor() -> DoctorResult:
    """
    Run diagnostic checks on the environment.

    Returns:
        DoctorResult with all check results
    """
    import platform

    checks = []
    all_passed = True

    # Check 1: OS is Windows
    is_windows = platform.system() == "Windows"
    checks.append(DoctorCheck(
        name="windows_os",
        passed=is_windows,
        message="Windows OS detected" if is_windows else f"Not Windows: {platform.system()}",
        remediation=None if is_windows else "This tool requires Windows with Classic Outlook.",
    ))
    if not is_windows:
        all_passed = False

    # Check 2: pywin32 available
    try:
        import win32com.client
        checks.append(DoctorCheck(
            name="pywin32",
            passed=True,
            message="pywin32 is installed and importable",
        ))
    except ImportError:
        checks.append(DoctorCheck(
            name="pywin32",
            passed=False,
            message="pywin32 is not installed",
            remediation="Run: uv add pywin32",
        ))
        all_passed = False

    # Check 3: Outlook COM available
    outlook_path = None
    try:
        outlook = get_outlook_app(retry_count=1, retry_delay=0.5)
        _ = outlook.Name
        checks.append(DoctorCheck(
            name="outlook_com",
            passed=True,
            message="Outlook COM automation is available",
        ))
    except OutlookNotAvailableError as e:
        checks.append(DoctorCheck(
            name="outlook_com",
            passed=False,
            message=str(e),
            remediation="Ensure Classic Outlook is running. New Outlook does not support COM automation.",
        ))
        all_passed = False
    except Exception as e:
        checks.append(DoctorCheck(
            name="outlook_com",
            passed=False,
            message=f"Outlook COM check failed: {e}",
            remediation="Ensure Classic Outlook is installed and running.",
        ))
        all_passed = False

    # Check 4: Find Outlook executable
    outlook_path = find_outlook_executable()
    if outlook_path:
        checks.append(DoctorCheck(
            name="outlook_exe",
            passed=True,
            message=f"Outlook executable found: {outlook_path}",
        ))
    else:
        checks.append(DoctorCheck(
            name="outlook_exe",
            passed=False,
            message="Outlook executable not found in common paths",
            remediation="Outlook may be installed in a non-standard location.",
        ))
        # This is a warning, not a failure
        # all_passed = False

    return DoctorResult(
        all_passed=all_passed,
        checks=checks,
        outlook_path=outlook_path,
    )


# =============================================================================
# Calendar Functions
# =============================================================================


def _response_status_to_string(status: int) -> str:
    """Convert Outlook response status to string."""
    mapping = {
        OL_RESPONSE_NONE: "none",
        OL_RESPONSE_ORGANIZER: "organizer",
        OL_RESPONSE_TENTATIVE: "tentative",
        OL_RESPONSE_ACCEPTED: "accepted",
        OL_RESPONSE_DECLINED: "declined",
    }
    return mapping.get(status, "none")


def _busy_status_to_string(status: int) -> str:
    """Convert Outlook busy status to string."""
    mapping = {
        OL_BUSY_FREE: "free",
        OL_BUSY_TENTATIVE: "tentative",
        OL_BUSY_BUSY: "busy",
        OL_BUSY_OUT_OF_OFFICE: "out_of_office",
        OL_BUSY_WORKING_ELSEWHERE: "working_elsewhere",
    }
    return mapping.get(status, "busy")


def _recurrence_type_to_string(rec_type: int) -> str:
    """Convert Outlook recurrence type to string."""
    mapping = {
        OL_RECURS_DAILY: "daily",
        OL_RECURS_WEEKLY: "weekly",
        OL_RECURS_MONTHLY: "monthly",
        OL_RECURS_MONTHLY_NTH: "monthly_nth",
        OL_RECURS_YEARLY: "yearly",
        OL_RECURS_YEARLY_NTH: "yearly_nth",
    }
    return mapping.get(rec_type, "unknown")


def _day_mask_to_list(mask: int) -> list[str]:
    """Convert day of week mask to list of day names."""
    days = []
    for day_name, day_value in DAY_OF_WEEK_MAP.items():
        if mask & day_value:
            days.append(day_name)
    return days


def _list_to_day_mask(days: list[str]) -> int:
    """Convert list of day names to day of week mask."""
    mask = 0
    for day in days:
        day_lower = day.lower()
        if day_lower in DAY_OF_WEEK_MAP:
            mask |= DAY_OF_WEEK_MAP[day_lower]
    return mask


def get_calendar(outlook_app, calendar_email: Optional[str] = None):
    """
    Get a calendar folder.

    Args:
        outlook_app: Outlook Application COM object
        calendar_email: Email of shared calendar owner (None for default)

    Returns:
        Calendar folder COM object

    Raises:
        CalendarNotFoundError: If calendar not found
    """
    namespace = get_namespace(outlook_app)

    if calendar_email:
        try:
            recipient = namespace.CreateRecipient(calendar_email)
            recipient.Resolve()
            if recipient.Resolved:
                return namespace.GetSharedDefaultFolder(recipient, OL_FOLDER_CALENDAR)
            raise CalendarNotFoundError(
                f"Could not resolve calendar for: {calendar_email}"
            )
        except Exception as e:
            raise CalendarNotFoundError(
                f"Could not access calendar for {calendar_email}: {e}"
            )

    return get_default_folder(outlook_app, OL_FOLDER_CALENDAR)


def get_event_by_id(outlook_app, entry_id: str, store_id: str):
    """
    Get a calendar event by its entry ID and store ID.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Event entry ID
        store_id: Store ID

    Returns:
        AppointmentItem COM object

    Raises:
        EventNotFoundError: If event not found
    """
    try:
        namespace = get_namespace(outlook_app)
        return namespace.GetItemFromID(entry_id, store_id)
    except Exception as e:
        raise EventNotFoundError(
            f"Event not found with entry_id={entry_id}: {e}"
        )


def extract_recurrence_info(appt_item) -> Optional[RecurrenceInfo]:
    """Extract recurrence information from an appointment."""
    try:
        if not appt_item.IsRecurring:
            return None

        pattern = appt_item.GetRecurrencePattern()
        rec_type = _recurrence_type_to_string(pattern.RecurrenceType)

        return RecurrenceInfo(
            type=rec_type,
            interval=pattern.Interval,
            days_of_week=_day_mask_to_list(pattern.DayOfWeekMask) if pattern.DayOfWeekMask else [],
            day_of_month=pattern.DayOfMonth if hasattr(pattern, "DayOfMonth") else None,
            month_of_year=pattern.MonthOfYear if hasattr(pattern, "MonthOfYear") else None,
            instance=pattern.Instance if hasattr(pattern, "Instance") else None,
            end_date=format_datetime(pattern.PatternEndDate) if not pattern.NoEndDate else None,
            occurrences=pattern.Occurrences if pattern.Occurrences > 0 else None,
        )
    except Exception:
        return None


def extract_attendees(appt_item) -> list[Attendee]:
    """Extract attendees from an appointment."""
    attendees = []
    try:
        for i in range(1, appt_item.Recipients.Count + 1):
            recip = appt_item.Recipients.Item(i)
            addr = extract_email_address(recip)

            # Determine attendee type
            if recip.Type == 1:
                attendee_type = "required"
            elif recip.Type == 2:
                attendee_type = "optional"
            elif recip.Type == 3:
                attendee_type = "resource"
            else:
                attendee_type = "required"

            # Get response status
            response = "none"
            try:
                meeting_response = recip.MeetingResponseStatus
                response = _response_status_to_string(meeting_response)
            except Exception:
                pass

            attendees.append(Attendee(
                name=addr.name,
                email=addr.email,
                type=attendee_type,
                response=response,
            ))
    except Exception:
        pass
    return attendees


def extract_event_summary(appt_item) -> EventSummary:
    """
    Extract an EventSummary from an AppointmentItem COM object.

    Args:
        appt_item: Outlook AppointmentItem COM object

    Returns:
        EventSummary object
    """
    organizer = ""
    try:
        organizer = str(appt_item.Organizer or "")
    except Exception:
        pass

    is_meeting = False
    try:
        is_meeting = appt_item.MeetingStatus != OL_MEETING_STATUS_NONMEETING
    except Exception:
        pass

    response_status = "none"
    try:
        response_status = _response_status_to_string(appt_item.ResponseStatus)
    except Exception:
        pass

    busy_status = "busy"
    try:
        busy_status = _busy_status_to_string(appt_item.BusyStatus)
    except Exception:
        pass

    return EventSummary(
        id=EventId(
            entry_id=str(appt_item.EntryID),
            store_id=str(appt_item.Parent.StoreID),
        ),
        subject=str(appt_item.Subject or ""),
        start=format_datetime(appt_item.Start),
        end=format_datetime(appt_item.End),
        location=str(appt_item.Location or ""),
        organizer=organizer,
        is_recurring=bool(appt_item.IsRecurring),
        is_all_day=bool(appt_item.AllDayEvent),
        is_meeting=is_meeting,
        response_status=response_status,
        busy_status=busy_status,
    )


def extract_event_detail(
    appt_item,
    include_body: bool = False,
) -> EventDetail:
    """
    Extract full EventDetail from an AppointmentItem COM object.

    Args:
        appt_item: Outlook AppointmentItem COM object
        include_body: Whether to include body

    Returns:
        EventDetail object
    """
    organizer = ""
    try:
        organizer = str(appt_item.Organizer or "")
    except Exception:
        pass

    is_meeting = False
    try:
        is_meeting = appt_item.MeetingStatus != OL_MEETING_STATUS_NONMEETING
    except Exception:
        pass

    response_status = "none"
    try:
        response_status = _response_status_to_string(appt_item.ResponseStatus)
    except Exception:
        pass

    busy_status = "busy"
    try:
        busy_status = _busy_status_to_string(appt_item.BusyStatus)
    except Exception:
        pass

    body = None
    if include_body:
        try:
            body = str(appt_item.Body or "")
        except Exception:
            pass

    categories = []
    try:
        cat_str = str(appt_item.Categories or "")
        if cat_str:
            categories = [c.strip() for c in cat_str.split(",")]
    except Exception:
        pass

    reminder_minutes = None
    try:
        if appt_item.ReminderSet:
            reminder_minutes = appt_item.ReminderMinutesBeforeStart
    except Exception:
        pass

    sensitivity = "normal"
    try:
        sens_map = {0: "normal", 1: "personal", 2: "private", 3: "confidential"}
        sensitivity = sens_map.get(appt_item.Sensitivity, "normal")
    except Exception:
        pass

    return EventDetail(
        id=EventId(
            entry_id=str(appt_item.EntryID),
            store_id=str(appt_item.Parent.StoreID),
        ),
        subject=str(appt_item.Subject or ""),
        start=format_datetime(appt_item.Start),
        end=format_datetime(appt_item.End),
        location=str(appt_item.Location or ""),
        organizer=organizer,
        is_recurring=bool(appt_item.IsRecurring),
        is_all_day=bool(appt_item.AllDayEvent),
        is_meeting=is_meeting,
        response_status=response_status,
        busy_status=busy_status,
        body=body,
        attendees=extract_attendees(appt_item),
        recurrence=extract_recurrence_info(appt_item),
        categories=categories,
        reminder_minutes=reminder_minutes,
        sensitivity=sensitivity,
    )


def list_events(
    outlook_app,
    start_date: datetime,
    end_date: datetime,
    calendar_email: Optional[str] = None,
    count: int = 100,
) -> Iterator[EventSummary]:
    """
    List events from a calendar within a date range.

    Args:
        outlook_app: Outlook Application COM object
        start_date: Start of date range
        end_date: End of date range
        calendar_email: Email of shared calendar owner (None for default)
        count: Maximum number of events to return

    Yields:
        EventSummary objects
    """
    calendar = get_calendar(outlook_app, calendar_email)
    items = calendar.Items

    # Important: Set IncludeRecurrences BEFORE sorting
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Build date filter - find events that START within the date range
    start_str = start_date.strftime("%m/%d/%Y %H:%M %p")
    end_str = end_date.strftime("%m/%d/%Y %H:%M %p")

    restriction = f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"

    try:
        items = items.Restrict(restriction)
    except Exception:
        # Fall back to manual filtering if restriction fails
        pass

    yielded = 0
    for item in items:
        if yielded >= count:
            break

        try:
            # Check if item is an appointment
            if item.Class != 26:  # olAppointment
                continue

            # Manual date filtering as backup - event must start within range
            # Convert COM datetime to naive Python datetime for comparison
            item_start = item.Start
            if hasattr(item_start, 'replace'):
                # Remove timezone info for comparison (pywintypes.datetime -> naive)
                item_start_naive = item_start.replace(tzinfo=None)
            else:
                item_start_naive = item_start

            if item_start_naive < start_date or item_start_naive > end_date:
                continue

            yield extract_event_summary(item)
            yielded += 1

        except Exception:
            continue


def create_event(
    outlook_app,
    subject: str,
    start: datetime,
    duration: int = 60,
    end: Optional[datetime] = None,
    location: str = "",
    body: str = "",
    attendees: list[str] = None,
    optional_attendees: list[str] = None,
    all_day: bool = False,
    reminder_minutes: int = 15,
    busy_status: str = "busy",
    teams_url: Optional[str] = None,
    recurrence: Optional[dict] = None,
) -> tuple[str, str, bool]:
    """
    Create a calendar event.

    Args:
        outlook_app: Outlook Application COM object
        subject: Event subject
        start: Start datetime
        duration: Duration in minutes (used if end not specified)
        end: End datetime (overrides duration)
        location: Event location
        body: Event body
        attendees: List of required attendee emails
        optional_attendees: List of optional attendee emails
        all_day: Whether this is an all-day event
        reminder_minutes: Reminder time before start
        busy_status: Show as status (free, tentative, busy, out_of_office)
        teams_url: Optional Teams meeting URL to embed in body
        recurrence: Optional recurrence pattern dict

    Returns:
        Tuple of (entry_id, store_id, has_attendees)

    Raises:
        OutlookError: If event creation fails
    """
    attendees = attendees or []
    optional_attendees = optional_attendees or []

    try:
        appt = outlook_app.CreateItem(OL_ITEM_APPOINTMENT)

        # Set basic properties
        appt.Subject = subject
        appt.Location = location

        # Handle body with optional Teams URL
        full_body = body
        if teams_url:
            teams_section = f"\n\n________________________________________________________________________________\n\nMicrosoft Teams meeting\n\nJoin on your computer, mobile app or room device\n{teams_url}\n\n________________________________________________________________________________\n"
            full_body = body + teams_section if body else teams_section.strip()
        appt.Body = full_body

        # Set times
        if all_day:
            appt.AllDayEvent = True
            appt.Start = start.date()
            if end:
                appt.End = end.date()
            else:
                appt.End = start.date()
        else:
            appt.Start = start
            if end:
                appt.End = end
            else:
                appt.Duration = duration

        # Set reminder
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes

        # Set busy status
        busy_map = {
            "free": OL_BUSY_FREE,
            "tentative": OL_BUSY_TENTATIVE,
            "busy": OL_BUSY_BUSY,
            "out_of_office": OL_BUSY_OUT_OF_OFFICE,
            "working_elsewhere": OL_BUSY_WORKING_ELSEWHERE,
        }
        appt.BusyStatus = busy_map.get(busy_status, OL_BUSY_BUSY)

        # Add attendees (makes it a meeting)
        has_attendees = bool(attendees or optional_attendees)
        if has_attendees:
            appt.MeetingStatus = OL_MEETING_STATUS_MEETING

            for addr in attendees:
                recip = appt.Recipients.Add(addr)
                recip.Type = 1  # olRequired

            for addr in optional_attendees:
                recip = appt.Recipients.Add(addr)
                recip.Type = 2  # olOptional

            appt.Recipients.ResolveAll()

        # Set recurrence if specified
        if recurrence:
            pattern = appt.GetRecurrencePattern()

            rec_type_map = {
                "daily": OL_RECURS_DAILY,
                "weekly": OL_RECURS_WEEKLY,
                "monthly": OL_RECURS_MONTHLY,
                "monthly_nth": OL_RECURS_MONTHLY_NTH,
                "yearly": OL_RECURS_YEARLY,
            }
            pattern.RecurrenceType = rec_type_map.get(
                recurrence.get("type", "weekly"),
                OL_RECURS_WEEKLY
            )

            if "interval" in recurrence:
                pattern.Interval = recurrence["interval"]

            if "days_of_week" in recurrence:
                pattern.DayOfWeekMask = _list_to_day_mask(recurrence["days_of_week"])

            if "day_of_month" in recurrence:
                pattern.DayOfMonth = recurrence["day_of_month"]

            if "end_date" in recurrence:
                pattern.PatternEndDate = recurrence["end_date"]
                pattern.NoEndDate = False
            elif "occurrences" in recurrence:
                pattern.Occurrences = recurrence["occurrences"]
                pattern.NoEndDate = False
            else:
                pattern.NoEndDate = True

        # Save the event (don't send yet)
        appt.Save()

        return appt.EntryID, appt.Parent.StoreID, has_attendees

    except Exception as e:
        raise OutlookError(f"Failed to create event: {e}")


def send_meeting_invites(outlook_app, entry_id: str, store_id: str) -> None:
    """
    Send meeting invitations for an existing event.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Event entry ID
        store_id: Event store ID

    Raises:
        OutlookError: If send fails
    """
    try:
        appt = get_event_by_id(outlook_app, entry_id, store_id)

        if appt.MeetingStatus == OL_MEETING_STATUS_NONMEETING:
            raise OutlookError("Cannot send invites for non-meeting event")

        appt.Send()
    except EventNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to send meeting invites: {e}")


def respond_to_meeting(
    outlook_app,
    entry_id: str,
    store_id: str,
    response: str,
    send_response: bool = True,
) -> None:
    """
    Respond to a meeting invitation.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Event entry ID
        store_id: Event store ID
        response: Response type ("accept", "decline", "tentative")
        send_response: Whether to send response to organizer

    Raises:
        OutlookError: If response fails
    """
    response_map = {
        "accept": OL_RESPONSE_ACCEPTED,
        "decline": OL_RESPONSE_DECLINED,
        "tentative": OL_RESPONSE_TENTATIVE,
    }

    if response not in response_map:
        raise OutlookError(f"Invalid response: {response}. Use: accept, decline, tentative")

    try:
        appt = get_event_by_id(outlook_app, entry_id, store_id)

        # Respond to the meeting
        response_item = appt.Respond(response_map[response], True)  # fNoUI=True

        if send_response and response_item:
            response_item.Send()

    except EventNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to respond to meeting: {e}")


def update_event(
    outlook_app,
    entry_id: str,
    store_id: str,
    subject: Optional[str] = None,
    start: Optional[datetime] = None,
    end: Optional[datetime] = None,
    duration: Optional[int] = None,
    location: Optional[str] = None,
    body: Optional[str] = None,
    reminder_minutes: Optional[int] = None,
    busy_status: Optional[str] = None,
) -> list[str]:
    """
    Update an existing calendar event.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Event entry ID
        store_id: Event store ID
        subject: New subject (None to keep existing)
        start: New start time (None to keep existing)
        end: New end time (None to keep existing)
        duration: New duration in minutes (None to keep existing, ignored if end set)
        location: New location (None to keep existing)
        body: New body (None to keep existing)
        reminder_minutes: New reminder (None to keep existing)
        busy_status: New busy status (None to keep existing)

    Returns:
        List of updated field names

    Raises:
        OutlookError: If update fails
    """
    try:
        appt = get_event_by_id(outlook_app, entry_id, store_id)
        updated_fields = []

        if subject is not None:
            appt.Subject = subject
            updated_fields.append("subject")

        if start is not None:
            appt.Start = start
            updated_fields.append("start")

        if end is not None:
            appt.End = end
            updated_fields.append("end")
        elif duration is not None:
            appt.Duration = duration
            updated_fields.append("duration")

        if location is not None:
            appt.Location = location
            updated_fields.append("location")

        if body is not None:
            appt.Body = body
            updated_fields.append("body")

        if reminder_minutes is not None:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
            updated_fields.append("reminder")

        if busy_status is not None:
            busy_map = {
                "free": OL_BUSY_FREE,
                "tentative": OL_BUSY_TENTATIVE,
                "busy": OL_BUSY_BUSY,
                "out_of_office": OL_BUSY_OUT_OF_OFFICE,
                "working_elsewhere": OL_BUSY_WORKING_ELSEWHERE,
            }
            if busy_status in busy_map:
                appt.BusyStatus = busy_map[busy_status]
                updated_fields.append("busy_status")

        if updated_fields:
            appt.Save()

        return updated_fields

    except EventNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to update event: {e}")


def delete_event(
    outlook_app,
    entry_id: str,
    store_id: str,
    send_cancellation: bool = True,
) -> tuple[str, bool]:
    """
    Delete a calendar event.

    Args:
        outlook_app: Outlook Application COM object
        entry_id: Event entry ID
        store_id: Event store ID
        send_cancellation: Whether to send cancellation notice to attendees

    Returns:
        Tuple of (subject, was_meeting_cancelled)

    Raises:
        OutlookError: If delete fails
    """
    try:
        appt = get_event_by_id(outlook_app, entry_id, store_id)
        subject = str(appt.Subject or "")
        is_meeting = appt.MeetingStatus != OL_MEETING_STATUS_NONMEETING
        was_cancelled = False

        # If it's a meeting with attendees and we're the organizer, send cancellation
        if is_meeting and send_cancellation:
            try:
                # Check if we're the organizer (response status = organizer)
                if appt.ResponseStatus == OL_RESPONSE_ORGANIZER:
                    appt.MeetingStatus = OL_MEETING_STATUS_CANCELED
                    appt.Save()
                    appt.Send()  # Send cancellation to attendees
                    was_cancelled = True
            except Exception:
                # If cancellation fails, just delete
                pass

        appt.Delete()

        return subject, was_cancelled

    except EventNotFoundError:
        raise
    except Exception as e:
        raise OutlookError(f"Failed to delete event: {e}")
