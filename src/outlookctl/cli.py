"""
Command-line interface for outlookctl.

This module provides the main entry point and argument parsing for the CLI.
"""

import argparse
import functools
import json
import sys
from datetime import datetime
from typing import Optional, Callable

from . import __version__
from .models import (
    ListResult,
    SearchResult,
    DraftResult,
    SendResult,
    AttachmentSaveResult,
    ErrorResult,
    FolderInfo,
    MessageId,
    # New email operation models
    MoveResult,
    DeleteResult,
    MarkReadResult,
    ForwardResult,
    # Calendar models
    CalendarInfo,
    CalendarsResult,
    CalendarListResult,
    EventCreateResult,
    EventSendResult,
    EventRespondResult,
    EventUpdateResult,
    EventDeleteResult,
    EventId,
)
from .outlook_com import (
    get_outlook_app,
    resolve_folder,
    list_messages,
    search_messages,
    get_message_by_id,
    extract_message_detail,
    create_draft,
    send_draft,
    send_new_message,
    save_attachments,
    run_doctor,
    OutlookError,
    OutlookNotAvailableError,
    FolderNotFoundError,
    MessageNotFoundError,
    # New email operations
    move_message,
    delete_message,
    mark_message_read,
    create_forward,
    create_reply_all,
    # Calendar functions
    list_events,
    list_all_calendars,
    get_event_by_id,
    extract_event_detail,
    create_event,
    send_meeting_invites,
    respond_to_meeting,
    update_event,
    delete_event,
    EventNotFoundError,
    CalendarNotFoundError,
)
from .safety import (
    validate_send_confirmation,
    validate_unsafe_send_new,
    check_recipients,
    SendConfirmationError,
)
from .audit import log_send_operation, log_draft_operation


def output_json(data: dict, output_format: str = "json") -> None:
    """Output data in the specified format."""
    if output_format == "json":
        print(json.dumps(data, indent=2, ensure_ascii=False))
    else:
        # Simple text format
        print(json.dumps(data, indent=2, ensure_ascii=False))


def output_error(error: str, error_code: str = None, remediation: str = None) -> None:
    """Output an error in JSON format."""
    result = ErrorResult(
        error=error,
        error_code=error_code,
        remediation=remediation,
    )
    print(json.dumps(result.to_dict(), indent=2, ensure_ascii=False))
    sys.exit(1)


def handle_outlook_errors(error_code: str) -> Callable:
    """
    Decorator to handle common Outlook errors consistently.

    Args:
        error_code: The error code to use for unhandled exceptions

    Returns:
        Decorated function
    """
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except OutlookNotAvailableError as e:
                output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
            except FolderNotFoundError as e:
                output_error(str(e), "FOLDER_NOT_FOUND")
            except MessageNotFoundError as e:
                output_error(str(e), "MESSAGE_NOT_FOUND")
            except EventNotFoundError as e:
                output_error(str(e), "EVENT_NOT_FOUND")
            except CalendarNotFoundError as e:
                output_error(str(e), "CALENDAR_NOT_FOUND")
            except SendConfirmationError as e:
                output_error(str(e), "CONFIRMATION_REQUIRED", "Use --confirm-send YES")
            except ValueError as e:
                output_error(str(e), "VALIDATION_ERROR")
            except OutlookError as e:
                output_error(str(e), error_code)
            except Exception as e:
                output_error(str(e), error_code)
        return wrapper
    return decorator


def parse_recipient_args(
    to: Optional[str] = None,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
) -> tuple[list[str], list[str], list[str]]:
    """
    Parse recipient arguments from comma-separated strings.

    Args:
        to: Comma-separated To addresses
        cc: Comma-separated CC addresses
        bcc: Comma-separated BCC addresses

    Returns:
        Tuple of (to_list, cc_list, bcc_list)
    """
    to_list = [addr.strip() for addr in to.split(",") if addr.strip()] if to else []
    cc_list = [addr.strip() for addr in cc.split(",") if addr.strip()] if cc else []
    bcc_list = [addr.strip() for addr in bcc.split(",") if addr.strip()] if bcc else []
    return to_list, cc_list, bcc_list


def parse_date(date_str: str) -> Optional[datetime]:
    """Parse an ISO date string."""
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str)
    except ValueError:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"Invalid date format: {date_str}. Use ISO format (YYYY-MM-DD).")


def cmd_doctor(args: argparse.Namespace) -> None:
    """Run diagnostic checks."""
    result = run_doctor()
    output_json(result.to_dict(), args.output)
    if not result.all_passed:
        sys.exit(1)


@handle_outlook_errors("LIST_ERROR")
def cmd_list(args: argparse.Namespace) -> None:
    """List messages from a folder."""
    outlook = get_outlook_app()
    _folder, folder_info = resolve_folder(outlook, args.folder)

    since = parse_date(args.since) if args.since else None
    until = parse_date(args.until) if args.until else None

    messages = list(list_messages(
        outlook,
        folder_spec=args.folder,
        count=args.count,
        unread_only=args.unread_only,
        since=since,
        until=until,
        include_body_snippet=args.include_body_snippet,
        body_snippet_chars=args.body_snippet_chars,
    ))

    result = ListResult(
        folder=folder_info,
        items=messages,
    )
    output_json(result.to_dict(), args.output)


@handle_outlook_errors("GET_ERROR")
def cmd_get(args: argparse.Namespace) -> None:
    """Get a single message by ID."""
    outlook = get_outlook_app()
    mail = get_message_by_id(outlook, args.id, args.store)

    detail = extract_message_detail(
        mail,
        include_body=args.include_body,
        max_body_chars=args.max_body_chars,
        include_headers=args.include_headers,
    )

    output_json({"version": "1.0", **detail.to_dict()}, args.output)


@handle_outlook_errors("SEARCH_ERROR")
def cmd_search(args: argparse.Namespace) -> None:
    """Search messages."""
    outlook = get_outlook_app()

    since = parse_date(args.since) if args.since else None
    until = parse_date(args.until) if args.until else None

    # Parse has_attachments flag
    has_attachments = None
    if args.has_attachments:
        has_attachments = True
    elif args.no_attachments:
        has_attachments = False

    messages = list(search_messages(
        outlook,
        folder_spec=args.folder,
        query=args.query,
        from_filter=getattr(args, "from", None),
        to_filter=args.to,
        cc_filter=args.cc,
        subject_contains=args.subject_contains,
        unread_only=args.unread_only,
        has_attachments=has_attachments,
        since=since,
        until=until,
        count=args.count,
        include_body_snippet=args.include_body_snippet,
        body_snippet_chars=args.body_snippet_chars,
    ))

    query_info = {}
    if args.query:
        query_info["text"] = args.query
    if getattr(args, "from", None):
        query_info["from"] = getattr(args, "from")
    if args.to:
        query_info["to"] = args.to
    if args.cc:
        query_info["cc"] = args.cc
    if args.subject_contains:
        query_info["subject_contains"] = args.subject_contains
    if args.unread_only:
        query_info["unread_only"] = True
    if has_attachments is not None:
        query_info["has_attachments"] = has_attachments
    if since:
        query_info["since"] = since.isoformat()
    if until:
        query_info["until"] = until.isoformat()

    result = SearchResult(
        query=query_info,
        items=messages,
    )
    output_json(result.to_dict(), args.output)


def cmd_draft(args: argparse.Namespace) -> None:
    """Create a draft message."""
    # Initialize recipient lists before try block for error logging
    to_list, cc_list, bcc_list = [], [], []
    try:
        to_list, cc_list, bcc_list = parse_recipient_args(args.to, args.cc, args.bcc)

        # For reply-all, we don't require recipients (they come from original)
        if not args.reply_all:
            check_recipients(to_list, cc_list, bcc_list)

        outlook = get_outlook_app()

        # Handle reply-all differently
        if args.reply_all and args.reply_to_id and args.reply_to_store:
            entry_id, store_id = create_reply_all(
                outlook,
                entry_id=args.reply_to_id,
                store_id=args.reply_to_store,
                body_text=args.body_text,
                body_html=args.body_html,
            )
        else:
            entry_id, store_id = create_draft(
                outlook,
                to=to_list,
                cc=cc_list,
                bcc=bcc_list,
                subject=args.subject or "",
                body_text=args.body_text,
                body_html=args.body_html,
                attachments=args.attach or [],
                reply_to_entry_id=args.reply_to_id,
                reply_to_store_id=args.reply_to_store,
            )

        log_draft_operation(
            to=to_list,
            cc=cc_list,
            bcc=bcc_list,
            subject=args.subject or "",
            success=True,
            entry_id=entry_id,
        )

        result = DraftResult(
            success=True,
            id=MessageId(entry_id=entry_id, store_id=store_id),
            saved_to="Drafts",
            subject=args.subject,
            to=to_list,
            cc=cc_list,
            attachments=args.attach or [],
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        log_draft_operation(
            to=to_list,
            cc=cc_list,
            bcc=bcc_list,
            subject=args.subject or "",
            success=False,
            error=str(e),
        )
        output_error(str(e), "DRAFT_ERROR")
    except Exception as e:
        output_error(str(e), "DRAFT_ERROR")


def cmd_send(args: argparse.Namespace) -> None:
    """Send a draft or new message."""
    try:
        outlook = get_outlook_app()

        # Case 1: Sending an existing draft
        if args.draft_id and args.draft_store:
            validate_send_confirmation(args.confirm_send, args.confirm_send_file)

            # Get draft info for logging
            mail = get_message_by_id(outlook, args.draft_id, args.draft_store)
            to_list = [str(r.Address) for r in mail.Recipients if r.Type == 1]
            subject = str(mail.Subject)

            send_draft(outlook, args.draft_id, args.draft_store)

            log_send_operation(
                to=to_list,
                cc=[],
                bcc=[],
                subject=subject,
                success=True,
                entry_id=args.draft_id,
                log_body=args.log_body,
            )

            result = SendResult(
                success=True,
                message="Draft sent successfully",
                sent_at=datetime.now().isoformat(),
                to=to_list,
                subject=subject,
            )
            output_json(result.to_dict(), args.output)

        # Case 2: Sending a new message directly (requires --unsafe-send-new)
        elif args.to:
            validate_unsafe_send_new(
                args.unsafe_send_new,
                args.confirm_send,
                args.confirm_send_file,
            )

            to_list, cc_list, bcc_list = parse_recipient_args(args.to, args.cc, args.bcc)
            check_recipients(to_list, cc_list, bcc_list)

            send_new_message(
                outlook,
                to=to_list,
                cc=cc_list,
                bcc=bcc_list,
                subject=args.subject or "",
                body_text=args.body_text,
                body_html=args.body_html,
                attachments=args.attach or [],
            )

            log_send_operation(
                to=to_list,
                cc=cc_list,
                bcc=bcc_list,
                subject=args.subject or "",
                success=True,
                log_body=args.log_body,
                body=args.body_text or args.body_html,
            )

            result = SendResult(
                success=True,
                message="Message sent successfully",
                sent_at=datetime.now().isoformat(),
                to=to_list,
                subject=args.subject,
            )
            output_json(result.to_dict(), args.output)

        else:
            output_error(
                "Either --draft-id/--draft-store or --to is required",
                "MISSING_ARGUMENTS",
                "Use --draft-id and --draft-store to send an existing draft, "
                "or use --to with --unsafe-send-new to send a new message directly.",
            )

    except SendConfirmationError as e:
        output_error(str(e), "CONFIRMATION_REQUIRED")
    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except MessageNotFoundError as e:
        output_error(str(e), "DRAFT_NOT_FOUND")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        log_send_operation(
            to=[],
            cc=[],
            bcc=[],
            subject="",
            success=False,
            error=str(e),
        )
        output_error(str(e), "SEND_ERROR")
    except Exception as e:
        output_error(str(e), "SEND_ERROR")


def cmd_attachments_save(args: argparse.Namespace) -> None:
    """Save attachments from a message."""
    try:
        outlook = get_outlook_app()

        saved_files = save_attachments(
            outlook,
            entry_id=args.id,
            store_id=args.store,
            dest_dir=args.dest,
        )

        result = AttachmentSaveResult(
            success=True,
            saved_files=saved_files,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except MessageNotFoundError as e:
        output_error(str(e), "MESSAGE_NOT_FOUND")
    except OutlookError as e:
        output_error(str(e), "ATTACHMENT_ERROR")
    except Exception as e:
        output_error(str(e), "ATTACHMENT_ERROR")


@handle_outlook_errors("MOVE_ERROR")
def cmd_move(args: argparse.Namespace) -> None:
    """Move a message to another folder."""
    outlook = get_outlook_app()

    # Get message info before moving
    mail = get_message_by_id(outlook, args.id, args.store)
    subject = str(mail.Subject or "")

    new_entry_id, new_store_id, folder_name = move_message(
        outlook,
        entry_id=args.id,
        store_id=args.store,
        dest_folder_spec=args.dest,
    )

    result = MoveResult(
        success=True,
        message=f"Message moved to {folder_name}",
        id=MessageId(entry_id=new_entry_id, store_id=new_store_id),
        moved_to=folder_name,
        subject=subject,
    )
    output_json(result.to_dict(), args.output)


@handle_outlook_errors("DELETE_ERROR")
def cmd_delete(args: argparse.Namespace) -> None:
    """Delete a message."""
    outlook = get_outlook_app()

    subject = delete_message(
        outlook,
        entry_id=args.id,
        store_id=args.store,
        permanent=args.permanent,
    )

    result = DeleteResult(
        success=True,
        message="Message deleted" if not args.permanent else "Message permanently deleted",
        subject=subject,
        permanent=args.permanent,
    )
    output_json(result.to_dict(), args.output)


@handle_outlook_errors("MARK_ERROR")
def cmd_mark_read(args: argparse.Namespace) -> None:
    """Mark a message as read or unread."""
    outlook = get_outlook_app()

    # Determine if marking as read or unread
    mark_as_read = not args.unread

    mark_message_read(
        outlook,
        entry_id=args.id,
        store_id=args.store,
        read=mark_as_read,
    )

    status = "read" if mark_as_read else "unread"
    result = MarkReadResult(
        success=True,
        message=f"Message marked as {status}",
        count=1,
        marked_as=status,
    )
    output_json(result.to_dict(), args.output)


@handle_outlook_errors("FORWARD_ERROR")
def cmd_forward(args: argparse.Namespace) -> None:
    """Create a forward draft for a message."""
    to_list, cc_list, bcc_list = parse_recipient_args(args.to, args.cc, args.bcc)
    check_recipients(to_list, cc_list, bcc_list)

    outlook = get_outlook_app()

    # Get original message info
    original = get_message_by_id(outlook, args.id, args.store)
    original_subject = str(original.Subject or "")

    entry_id, store_id = create_forward(
        outlook,
        entry_id=args.id,
        store_id=args.store,
        to=to_list,
        cc=cc_list,
        bcc=bcc_list,
        additional_text=args.message,
    )

    result = ForwardResult(
        success=True,
        id=MessageId(entry_id=entry_id, store_id=store_id),
        saved_to="Drafts",
        original_subject=original_subject,
        to=to_list,
    )
    output_json(result.to_dict(), args.output)


# =============================================================================
# Calendar Commands
# =============================================================================


def parse_datetime(dt_str: str) -> datetime:
    """Parse a datetime string in various formats."""
    formats = [
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            continue
    raise ValueError(
        f"Invalid datetime format: {dt_str}. "
        "Use 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD'."
    )


@handle_outlook_errors("CALENDARS_LIST_ERROR")
def cmd_calendar_calendars(args: argparse.Namespace) -> None:
    """List all available calendars."""
    outlook = get_outlook_app()
    calendars = list_all_calendars(outlook)

    result = CalendarsResult(
        calendars=[
            CalendarInfo(
                name=cal["name"],
                path=cal["path"],
                store=cal["store"],
            )
            for cal in calendars
        ]
    )
    output_json(result.to_dict(), args.output)


def cmd_calendar_list(args: argparse.Namespace) -> None:
    """List calendar events."""
    try:
        outlook = get_outlook_app()

        # Determine date range
        from datetime import timedelta
        if args.start:
            start_date = parse_datetime(args.start)
        else:
            start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        if args.end:
            end_date = parse_datetime(args.end)
        elif args.days:
            end_date = start_date + timedelta(days=args.days)
        else:
            end_date = start_date + timedelta(days=7)

        # Handle --all flag to query all calendars
        if getattr(args, 'all', False):
            from outlookctl.outlook_com import list_all_calendars, list_events_from_folder

            calendars = list_all_calendars(outlook)
            all_events = []

            for cal in calendars:
                try:
                    namespace = outlook.GetNamespace("MAPI")
                    folder = namespace.GetFolderFromID(cal["entry_id"], cal["store_id"])
                    events = list(list_events_from_folder(
                        folder,
                        start_date=start_date,
                        end_date=end_date,
                        count=args.count,
                    ))
                    # Add calendar name to each event
                    for event in events:
                        event.calendar_name = cal["name"]
                    all_events.extend(events)
                except Exception:
                    continue  # Skip calendars that fail

            # Sort all events by start time
            all_events.sort(key=lambda e: e.start)

            # Limit total count
            if len(all_events) > args.count:
                all_events = all_events[:args.count]

            result = {
                "version": "1.0",
                "calendars": "all",
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "items": [
                    {**event.to_dict(), "calendar": event.calendar_name}
                    for event in all_events
                ],
            }
            output_json(result, args.output)
        else:
            events = list(list_events(
                outlook,
                start_date=start_date,
                end_date=end_date,
                calendar_email=args.calendar,
                count=args.count,
            ))

            result = CalendarListResult(
                calendar=args.calendar or "Calendar",
                start_date=start_date.isoformat(),
                end_date=end_date.isoformat(),
                items=events,
            )
            output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except CalendarNotFoundError as e:
        output_error(str(e), "CALENDAR_NOT_FOUND")
    except Exception as e:
        output_error(str(e), "CALENDAR_LIST_ERROR")


def cmd_calendar_get(args: argparse.Namespace) -> None:
    """Get a single calendar event by ID."""
    try:
        outlook = get_outlook_app()
        appt = get_event_by_id(outlook, args.id, args.store)

        detail = extract_event_detail(
            appt,
            include_body=args.include_body,
        )

        output_json({"version": "1.0", **detail.to_dict()}, args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except EventNotFoundError as e:
        output_error(str(e), "EVENT_NOT_FOUND")
    except Exception as e:
        output_error(str(e), "CALENDAR_GET_ERROR")


def cmd_calendar_create(args: argparse.Namespace) -> None:
    """Create a calendar event."""
    try:
        outlook = get_outlook_app()

        # Parse start time
        start = parse_datetime(args.start)

        # Parse end time if provided
        end = parse_datetime(args.end) if args.end else None

        # Parse attendees
        attendees = [a.strip() for a in args.attendees.split(",")] if args.attendees else None
        optional_attendees = [a.strip() for a in args.optional_attendees.split(",")] if args.optional_attendees else None

        entry_id, store_id, has_attendees = create_event(
            outlook,
            subject=args.subject,
            start=start,
            duration=args.duration,
            end=end,
            location=args.location or "",
            body=args.body or "",
            attendees=attendees,
            optional_attendees=optional_attendees,
            all_day=args.all_day,
            reminder_minutes=args.reminder,
            busy_status=args.busy_status or "busy",
            teams_url=args.teams_url,
            recurrence=args.recurrence,
        )

        # If --send-now is specified with proper confirmation, send immediately
        send_now = False
        if args.send_now and has_attendees:
            validate_send_confirmation(args.confirm_send, None)
            send_meeting_invites(outlook, entry_id, store_id)
            send_now = True

        result = EventCreateResult(
            success=True,
            id=EventId(entry_id=entry_id, store_id=store_id),
            saved_to="Calendar",
            subject=args.subject,
            start=start.isoformat(),
            attendees=(attendees or []) + (optional_attendees or []),
            is_draft=has_attendees and not send_now,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        output_error(str(e), "CALENDAR_CREATE_ERROR")
    except Exception as e:
        output_error(str(e), "CALENDAR_CREATE_ERROR")


def cmd_calendar_send(args: argparse.Namespace) -> None:
    """Send meeting invitations for an existing calendar event."""
    try:
        # Validate confirmation
        validate_send_confirmation(args.confirm_send, args.confirm_send_file)

        outlook = get_outlook_app()

        # Get event info for response
        appt = get_event_by_id(outlook, args.id, args.store)
        subject = str(appt.Subject)
        attendees = []
        for r in appt.Recipients:
            try:
                attendees.append(str(r.Address))
            except Exception:
                attendees.append(str(r.Name))

        send_meeting_invites(outlook, args.id, args.store)

        result = EventSendResult(
            success=True,
            message="Meeting invitations sent",
            sent_at=datetime.now().isoformat(),
            attendees=attendees,
            subject=subject,
        )
        output_json(result.to_dict(), args.output)

    except SendConfirmationError as e:
        output_error(str(e), "CONFIRMATION_REQUIRED")
    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except EventNotFoundError as e:
        output_error(str(e), "EVENT_NOT_FOUND")
    except OutlookError as e:
        output_error(str(e), "CALENDAR_SEND_ERROR")
    except Exception as e:
        output_error(str(e), "CALENDAR_SEND_ERROR")


def cmd_calendar_respond(args: argparse.Namespace) -> None:
    """Respond to a meeting invitation."""
    try:
        outlook = get_outlook_app()

        # Validate response value
        response = args.response.lower()
        if response not in ("accept", "decline", "tentative"):
            raise ValueError(
                f"Invalid response '{response}'. Must be 'accept', 'decline', or 'tentative'."
            )

        # Get event info for response
        appt = get_event_by_id(outlook, args.id, args.store)
        subject = str(appt.Subject)
        try:
            organizer = str(appt.Organizer)
        except Exception:
            organizer = None

        respond_to_meeting(
            outlook,
            args.id,
            args.store,
            response=response,
            send_response=not args.no_response,
        )

        result = EventRespondResult(
            success=True,
            response=response,
            subject=subject,
            organizer=organizer,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except EventNotFoundError as e:
        output_error(str(e), "EVENT_NOT_FOUND")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        output_error(str(e), "CALENDAR_RESPOND_ERROR")
    except Exception as e:
        output_error(str(e), "CALENDAR_RESPOND_ERROR")


def cmd_calendar_update(args: argparse.Namespace) -> None:
    """Update an existing calendar event."""
    try:
        outlook = get_outlook_app()

        # Parse times if provided
        start = parse_datetime(args.start) if args.start else None
        end = parse_datetime(args.end) if args.end else None

        # Get event info before update
        appt = get_event_by_id(outlook, args.id, args.store)
        subject = str(appt.Subject or "")

        updated_fields = update_event(
            outlook,
            entry_id=args.id,
            store_id=args.store,
            subject=args.subject,
            start=start,
            end=end,
            duration=args.duration,
            location=args.location,
            body=args.body,
            reminder_minutes=args.reminder,
            busy_status=args.busy_status,
        )

        # Get updated subject if it changed
        if "subject" in updated_fields:
            subject = args.subject

        result = EventUpdateResult(
            success=True,
            message=f"Event updated: {', '.join(updated_fields)}" if updated_fields else "No changes made",
            id=EventId(entry_id=args.id, store_id=args.store),
            subject=subject,
            start=start.isoformat() if start else None,
            updated_fields=updated_fields,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except EventNotFoundError as e:
        output_error(str(e), "EVENT_NOT_FOUND")
    except ValueError as e:
        output_error(str(e), "VALIDATION_ERROR")
    except OutlookError as e:
        output_error(str(e), "CALENDAR_UPDATE_ERROR")
    except Exception as e:
        output_error(str(e), "CALENDAR_UPDATE_ERROR")


def cmd_calendar_delete(args: argparse.Namespace) -> None:
    """Delete a calendar event."""
    try:
        outlook = get_outlook_app()

        subject, was_cancelled = delete_event(
            outlook,
            entry_id=args.id,
            store_id=args.store,
            send_cancellation=not args.no_cancel,
        )

        if was_cancelled:
            message = "Meeting deleted and cancellation sent to attendees"
        else:
            message = "Event deleted"

        result = EventDeleteResult(
            success=True,
            message=message,
            subject=subject,
            cancelled=was_cancelled,
        )
        output_json(result.to_dict(), args.output)

    except OutlookNotAvailableError as e:
        output_error(str(e), "OUTLOOK_UNAVAILABLE", "Start Classic Outlook and try again.")
    except EventNotFoundError as e:
        output_error(str(e), "EVENT_NOT_FOUND")
    except OutlookError as e:
        output_error(str(e), "CALENDAR_DELETE_ERROR")
    except Exception as e:
        output_error(str(e), "CALENDAR_DELETE_ERROR")


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser."""
    parser = argparse.ArgumentParser(
        prog="outlookctl",
        description="Local CLI bridge for Outlook Classic automation via COM",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Doctor command
    doctor_parser = subparsers.add_parser(
        "doctor", help="Validate environment and prerequisites"
    )
    doctor_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    doctor_parser.set_defaults(func=cmd_doctor)

    # List command
    list_parser = subparsers.add_parser(
        "list", help="List messages from a folder"
    )
    list_parser.add_argument(
        "--folder", default="inbox",
        help="Folder: inbox|sent|drafts|by-name:<name>|by-path:<path> (default: inbox)"
    )
    list_parser.add_argument(
        "--count", type=int, default=10,
        help="Number of messages to return (default: 10)"
    )
    list_parser.add_argument(
        "--unread-only", action="store_true",
        help="Only return unread messages"
    )
    list_parser.add_argument(
        "--since", help="Only messages received after this date (ISO format)"
    )
    list_parser.add_argument(
        "--until", help="Only messages received before this date (ISO format)"
    )
    list_parser.add_argument(
        "--include-body-snippet", action="store_true",
        help="Include a snippet of the message body"
    )
    list_parser.add_argument(
        "--body-snippet-chars", type=int, default=200,
        help="Maximum characters for body snippet (default: 200)"
    )
    list_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    list_parser.set_defaults(func=cmd_list)

    # Get command
    get_parser = subparsers.add_parser(
        "get", help="Get a single message by ID"
    )
    get_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    get_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    get_parser.add_argument(
        "--include-body", action="store_true",
        help="Include message body"
    )
    get_parser.add_argument(
        "--include-headers", action="store_true",
        help="Include message headers"
    )
    get_parser.add_argument(
        "--max-body-chars", type=int,
        help="Maximum characters for body"
    )
    get_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    get_parser.set_defaults(func=cmd_get)

    # Search command
    search_parser = subparsers.add_parser(
        "search", help="Search messages"
    )
    search_parser.add_argument(
        "--folder", default="inbox",
        help="Folder to search in (default: inbox)"
    )
    search_parser.add_argument(
        "--query", help="Free text search (subject/body)"
    )
    search_parser.add_argument(
        "--from", dest="from", help="Filter by sender email or name"
    )
    search_parser.add_argument(
        "--to", help="Filter by To recipient email or name"
    )
    search_parser.add_argument(
        "--cc", help="Filter by CC recipient email or name"
    )
    search_parser.add_argument(
        "--subject-contains", help="Filter by subject content"
    )
    search_parser.add_argument(
        "--unread-only", action="store_true",
        help="Only return unread messages"
    )
    search_parser.add_argument(
        "--has-attachments", action="store_true",
        help="Only return messages with attachments"
    )
    search_parser.add_argument(
        "--no-attachments", action="store_true",
        help="Only return messages without attachments"
    )
    search_parser.add_argument(
        "--since", help="Only messages after this date (ISO format)"
    )
    search_parser.add_argument(
        "--until", help="Only messages before this date (ISO format)"
    )
    search_parser.add_argument(
        "--count", type=int, default=50,
        help="Maximum results (default: 50)"
    )
    search_parser.add_argument(
        "--include-body-snippet", action="store_true",
        help="Include a snippet of the message body"
    )
    search_parser.add_argument(
        "--body-snippet-chars", type=int, default=200,
        help="Maximum characters for body snippet (default: 200)"
    )
    search_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    search_parser.set_defaults(func=cmd_search)

    # Draft command
    draft_parser = subparsers.add_parser(
        "draft", help="Create a draft message"
    )
    draft_parser.add_argument(
        "--to", help="To recipients (comma-separated)"
    )
    draft_parser.add_argument(
        "--cc", help="CC recipients (comma-separated)"
    )
    draft_parser.add_argument(
        "--bcc", help="BCC recipients (comma-separated)"
    )
    draft_parser.add_argument(
        "--subject", help="Email subject"
    )
    draft_parser.add_argument(
        "--body-text", help="Plain text body"
    )
    draft_parser.add_argument(
        "--body-html", help="HTML body"
    )
    draft_parser.add_argument(
        "--attach", action="append",
        help="File path to attach (can be used multiple times)"
    )
    draft_parser.add_argument(
        "--reply-to-id", help="Entry ID of message to reply to"
    )
    draft_parser.add_argument(
        "--reply-to-store", help="Store ID of message to reply to"
    )
    draft_parser.add_argument(
        "--reply-all", action="store_true",
        help="Create reply-all instead of reply (use with --reply-to-id/--reply-to-store)"
    )
    draft_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    draft_parser.set_defaults(func=cmd_draft)

    # Send command
    send_parser = subparsers.add_parser(
        "send", help="Send a draft or new message"
    )
    # For sending existing draft
    send_parser.add_argument(
        "--draft-id", help="Entry ID of draft to send"
    )
    send_parser.add_argument(
        "--draft-store", help="Store ID of draft to send"
    )
    # For sending new message directly
    send_parser.add_argument(
        "--to", help="To recipients (comma-separated)"
    )
    send_parser.add_argument(
        "--cc", help="CC recipients (comma-separated)"
    )
    send_parser.add_argument(
        "--bcc", help="BCC recipients (comma-separated)"
    )
    send_parser.add_argument(
        "--subject", help="Email subject"
    )
    send_parser.add_argument(
        "--body-text", help="Plain text body"
    )
    send_parser.add_argument(
        "--body-html", help="HTML body"
    )
    send_parser.add_argument(
        "--attach", action="append",
        help="File path to attach (can be used multiple times)"
    )
    # Safety flags
    send_parser.add_argument(
        "--confirm-send",
        help="Confirmation string (must be exactly 'YES')"
    )
    send_parser.add_argument(
        "--confirm-send-file",
        help="Path to file containing confirmation string"
    )
    send_parser.add_argument(
        "--unsafe-send-new", action="store_true",
        help="Allow sending new message directly (not recommended)"
    )
    send_parser.add_argument(
        "--log-body", action="store_true",
        help="Include body in audit log (default: metadata only)"
    )
    send_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    send_parser.set_defaults(func=cmd_send)

    # Attachments subcommand
    attachments_parser = subparsers.add_parser(
        "attachments", help="Attachment operations"
    )
    attachments_subparsers = attachments_parser.add_subparsers(
        dest="attachments_command", help="Attachment commands"
    )

    # Attachments save
    save_parser = attachments_subparsers.add_parser(
        "save", help="Save attachments from a message"
    )
    save_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    save_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    save_parser.add_argument(
        "--dest", required=True, help="Destination directory"
    )
    save_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    save_parser.set_defaults(func=cmd_attachments_save)

    # Move command
    move_parser = subparsers.add_parser(
        "move", help="Move a message to another folder"
    )
    move_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    move_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    move_parser.add_argument(
        "--dest", required=True,
        help="Destination folder: inbox|sent|drafts|deleted|by-name:<name>|by-path:<path>"
    )
    move_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    move_parser.set_defaults(func=cmd_move)

    # Delete command
    delete_parser = subparsers.add_parser(
        "delete", help="Delete a message"
    )
    delete_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    delete_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    delete_parser.add_argument(
        "--permanent", action="store_true",
        help="Permanently delete (skip Deleted Items)"
    )
    delete_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    delete_parser.set_defaults(func=cmd_delete)

    # Mark-read command
    mark_parser = subparsers.add_parser(
        "mark-read", help="Mark a message as read or unread"
    )
    mark_parser.add_argument(
        "--id", required=True, help="Message entry ID"
    )
    mark_parser.add_argument(
        "--store", required=True, help="Message store ID"
    )
    mark_parser.add_argument(
        "--unread", action="store_true",
        help="Mark as unread instead of read"
    )
    mark_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    mark_parser.set_defaults(func=cmd_mark_read)

    # Forward command
    forward_parser = subparsers.add_parser(
        "forward", help="Create a forward draft for a message"
    )
    forward_parser.add_argument(
        "--id", required=True, help="Original message entry ID"
    )
    forward_parser.add_argument(
        "--store", required=True, help="Original message store ID"
    )
    forward_parser.add_argument(
        "--to", required=True, help="To recipients (comma-separated)"
    )
    forward_parser.add_argument(
        "--cc", help="CC recipients (comma-separated)"
    )
    forward_parser.add_argument(
        "--bcc", help="BCC recipients (comma-separated)"
    )
    forward_parser.add_argument(
        "--message", help="Additional message to prepend to the forward"
    )
    forward_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    forward_parser.set_defaults(func=cmd_forward)

    # =========================================================================
    # Calendar subcommand
    # =========================================================================
    calendar_parser = subparsers.add_parser(
        "calendar", help="Calendar operations"
    )
    calendar_subparsers = calendar_parser.add_subparsers(
        dest="calendar_command", help="Calendar commands"
    )

    # Calendar calendars (list all calendars)
    cal_calendars_parser = calendar_subparsers.add_parser(
        "calendars", help="List all available calendars"
    )
    cal_calendars_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_calendars_parser.set_defaults(func=cmd_calendar_calendars)

    # Calendar list
    cal_list_parser = calendar_subparsers.add_parser(
        "list", help="List calendar events"
    )
    cal_list_parser.add_argument(
        "--start", help="Start date (YYYY-MM-DD or YYYY-MM-DD HH:MM)"
    )
    cal_list_parser.add_argument(
        "--end", help="End date (YYYY-MM-DD or YYYY-MM-DD HH:MM)"
    )
    cal_list_parser.add_argument(
        "--days", type=int,
        help="Number of days from start (default: 7, ignored if --end is set)"
    )
    cal_list_parser.add_argument(
        "--calendar",
        help="Calendar: name, 'by-name:Name', or email for shared calendar"
    )
    cal_list_parser.add_argument(
        "--all", action="store_true",
        help="Query ALL calendars and merge results (overrides --calendar)"
    )
    cal_list_parser.add_argument(
        "--count", type=int, default=100,
        help="Maximum events to return (default: 100)"
    )
    cal_list_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_list_parser.set_defaults(func=cmd_calendar_list)

    # Calendar get
    cal_get_parser = calendar_subparsers.add_parser(
        "get", help="Get a single calendar event by ID"
    )
    cal_get_parser.add_argument(
        "--id", required=True, help="Event entry ID"
    )
    cal_get_parser.add_argument(
        "--store", required=True, help="Event store ID"
    )
    cal_get_parser.add_argument(
        "--include-body", action="store_true",
        help="Include event body/description"
    )
    cal_get_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_get_parser.set_defaults(func=cmd_calendar_get)

    # Calendar create
    cal_create_parser = calendar_subparsers.add_parser(
        "create", help="Create a calendar event"
    )
    cal_create_parser.add_argument(
        "--subject", required=True, help="Event subject/title"
    )
    cal_create_parser.add_argument(
        "--start", required=True,
        help="Start time (YYYY-MM-DD HH:MM or YYYY-MM-DD for all-day)"
    )
    cal_create_parser.add_argument(
        "--duration", type=int, default=60,
        help="Duration in minutes (default: 60, ignored if --end is set)"
    )
    cal_create_parser.add_argument(
        "--end", help="End time (YYYY-MM-DD HH:MM)"
    )
    cal_create_parser.add_argument(
        "--location", help="Event location"
    )
    cal_create_parser.add_argument(
        "--body", help="Event description/body"
    )
    cal_create_parser.add_argument(
        "--attendees", help="Required attendees (comma-separated emails)"
    )
    cal_create_parser.add_argument(
        "--optional-attendees", help="Optional attendees (comma-separated emails)"
    )
    cal_create_parser.add_argument(
        "--all-day", action="store_true",
        help="Create as all-day event"
    )
    cal_create_parser.add_argument(
        "--reminder", type=int, default=15,
        help="Reminder minutes before event (default: 15)"
    )
    cal_create_parser.add_argument(
        "--busy-status", choices=["free", "tentative", "busy", "out_of_office", "working_elsewhere"],
        default="busy", help="Show as status (default: busy)"
    )
    cal_create_parser.add_argument(
        "--teams-url", help="Teams meeting URL to include in body"
    )
    cal_create_parser.add_argument(
        "--recurrence",
        help="Recurrence pattern (e.g., 'weekly:monday,wednesday:until:2025-12-31')"
    )
    cal_create_parser.add_argument(
        "--send-now", action="store_true",
        help="Send meeting invites immediately (requires --confirm-send YES)"
    )
    cal_create_parser.add_argument(
        "--confirm-send", help="Confirmation string (must be 'YES') for --send-now"
    )
    cal_create_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_create_parser.set_defaults(func=cmd_calendar_create)

    # Calendar send (send meeting invites)
    cal_send_parser = calendar_subparsers.add_parser(
        "send", help="Send meeting invitations for an existing event"
    )
    cal_send_parser.add_argument(
        "--id", required=True, help="Event entry ID"
    )
    cal_send_parser.add_argument(
        "--store", required=True, help="Event store ID"
    )
    cal_send_parser.add_argument(
        "--confirm-send", help="Confirmation string (must be exactly 'YES')"
    )
    cal_send_parser.add_argument(
        "--confirm-send-file", help="Path to file containing confirmation string"
    )
    cal_send_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_send_parser.set_defaults(func=cmd_calendar_send)

    # Calendar respond
    cal_respond_parser = calendar_subparsers.add_parser(
        "respond", help="Respond to a meeting invitation"
    )
    cal_respond_parser.add_argument(
        "--id", required=True, help="Event entry ID"
    )
    cal_respond_parser.add_argument(
        "--store", required=True, help="Event store ID"
    )
    cal_respond_parser.add_argument(
        "--response", required=True, choices=["accept", "decline", "tentative"],
        help="Response type"
    )
    cal_respond_parser.add_argument(
        "--no-response", action="store_true",
        help="Don't send response to organizer"
    )
    cal_respond_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_respond_parser.set_defaults(func=cmd_calendar_respond)

    # Calendar update
    cal_update_parser = calendar_subparsers.add_parser(
        "update", help="Update an existing calendar event"
    )
    cal_update_parser.add_argument(
        "--id", required=True, help="Event entry ID"
    )
    cal_update_parser.add_argument(
        "--store", required=True, help="Event store ID"
    )
    cal_update_parser.add_argument(
        "--subject", help="New event subject"
    )
    cal_update_parser.add_argument(
        "--start", help="New start time (YYYY-MM-DD HH:MM)"
    )
    cal_update_parser.add_argument(
        "--end", help="New end time (YYYY-MM-DD HH:MM)"
    )
    cal_update_parser.add_argument(
        "--duration", type=int, help="New duration in minutes"
    )
    cal_update_parser.add_argument(
        "--location", help="New event location"
    )
    cal_update_parser.add_argument(
        "--body", help="New event description"
    )
    cal_update_parser.add_argument(
        "--reminder", type=int, help="New reminder minutes before event"
    )
    cal_update_parser.add_argument(
        "--busy-status", choices=["free", "tentative", "busy", "out_of_office", "working_elsewhere"],
        help="New show-as status"
    )
    cal_update_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_update_parser.set_defaults(func=cmd_calendar_update)

    # Calendar delete
    cal_delete_parser = calendar_subparsers.add_parser(
        "delete", help="Delete a calendar event"
    )
    cal_delete_parser.add_argument(
        "--id", required=True, help="Event entry ID"
    )
    cal_delete_parser.add_argument(
        "--store", required=True, help="Event store ID"
    )
    cal_delete_parser.add_argument(
        "--no-cancel", action="store_true",
        help="Don't send cancellation to attendees (for meetings)"
    )
    cal_delete_parser.add_argument(
        "--output", choices=["json", "text"], default="json",
        help="Output format (default: json)"
    )
    cal_delete_parser.set_defaults(func=cmd_calendar_delete)

    return parser


def main() -> None:
    """Main entry point."""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Handle attachments subcommand
    if args.command == "attachments" and not args.attachments_command:
        parser.parse_args(["attachments", "-h"])
        sys.exit(1)

    # Handle calendar subcommand
    if args.command == "calendar" and not args.calendar_command:
        parser.parse_args(["calendar", "-h"])
        sys.exit(1)

    if hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
