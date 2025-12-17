# JSON Schema Reference

All `outlookctl` commands output JSON with consistent schemas. This document describes the output formats.

> **Schema Version:** `"1.0"` - All outputs include a `version` field for forward compatibility.

---

## Table of Contents

- [Common Types](#common-types)
  - [MessageId](#messageid)
  - [EmailAddress](#emailaddress)
  - [EventId](#eventid)
- [Email Command Outputs](#email-command-outputs)
  - [ListResult](#listresult)
  - [SearchResult](#searchresult)
  - [GetResult](#getresult-message-detail)
  - [DraftResult](#draftresult)
  - [SendResult](#sendresult)
  - [AttachmentSaveResult](#attachmentsaveresult)
- [Calendar Command Outputs](#calendar-command-outputs)
  - [CalendarListResult](#calendarlistresult)
  - [EventDetail](#eventdetail)
  - [EventCreateResult](#eventcreateresult)
- [System Outputs](#system-outputs)
  - [DoctorResult](#doctorresult)
  - [ErrorResult](#errorresult)
- [Error Codes](#error-codes)

---

## Common Types

### MessageId

Stable identifier for referencing messages across commands. **Both fields are required** when referencing a message (e.g., in `get`, `send`, `attachments save`).

```json
{
  "entry_id": "00000000...",
  "store_id": "00000000..."
}
```

### EmailAddress

Represents a sender or recipient with display name and email address.

```json
{
  "name": "John Doe",
  "email": "john.doe@example.com"
}
```

### EventId

Stable identifier for referencing calendar events. Same structure as MessageId.

```json
{
  "entry_id": "00000000...",
  "store_id": "00000000..."
}
```

---

## Email Command Outputs

### ListResult

Output of `outlookctl list`:

```json
{
  "version": "1.0",
  "folder": {
    "name": "Inbox",
    "path": null,
    "store_id": null
  },
  "items": [
    {
      "id": {
        "entry_id": "00000000...",
        "store_id": "00000000..."
      },
      "received_at": "2025-01-15T09:30:00",
      "subject": "Meeting Tomorrow",
      "from": {
        "name": "Jane Smith",
        "email": "jane@example.com"
      },
      "to": ["you@example.com"],
      "cc": [],
      "unread": true,
      "has_attachments": false,
      "body_snippet": "Hi, just wanted to confirm..."
    }
  ]
}
```

> **Notes:**
> - `body_snippet` only included when `--include-body-snippet` is used
> - Items are sorted by received time (newest first)

### SearchResult

Output of `outlookctl search`:

```json
{
  "version": "1.0",
  "query": {
    "from": "boss@company.com",
    "since": "2025-01-01T00:00:00",
    "unread_only": true
  },
  "items": [
    // Same structure as ListResult items
  ]
}
```

### GetResult (Message Detail)

Output of `outlookctl get`:

```json
{
  "version": "1.0",
  "id": {
    "entry_id": "00000000...",
    "store_id": "00000000..."
  },
  "received_at": "2025-01-15T09:30:00",
  "subject": "Meeting Tomorrow",
  "from": {
    "name": "Jane Smith",
    "email": "jane@example.com"
  },
  "to": ["you@example.com"],
  "cc": ["team@example.com"],
  "bcc": [],
  "unread": true,
  "has_attachments": true,
  "attachments": ["document.pdf", "image.png"],
  "body": "Full message body text...",
  "body_html": "<html>...</html>",
  "headers": {
    "Message-ID": "<abc123@mail.example.com>",
    "Date": "Wed, 15 Jan 2025 09:30:00 +0000"
  }
}
```

> **Notes:**
> - `body` and `body_html` only included with `--include-body`
> - `headers` only included with `--include-headers`

### DraftResult

Output of `outlookctl draft`:

```json
{
  "version": "1.0",
  "success": true,
  "id": {
    "entry_id": "00000000...",
    "store_id": "00000000..."
  },
  "saved_to": "Drafts",
  "subject": "Re: Meeting",
  "to": ["recipient@example.com"],
  "cc": [],
  "attachments": ["./report.pdf"]
}
```

### SendResult

Output of `outlookctl send`:

```json
{
  "version": "1.0",
  "success": true,
  "message": "Draft sent successfully",
  "sent_at": "2025-01-15T10:00:00",
  "to": ["recipient@example.com"],
  "subject": "Re: Meeting"
}
```

### AttachmentSaveResult

Output of `outlookctl attachments save`:

```json
{
  "version": "1.0",
  "success": true,
  "saved_files": [
    "C:\\Users\\user\\downloads\\document.pdf",
    "C:\\Users\\user\\downloads\\image.png"
  ],
  "errors": []
}
```

---

## Calendar Command Outputs

### CalendarListResult

Output of `outlookctl calendar list`:

```json
{
  "version": "1.0",
  "calendar": "Calendar",
  "start_date": "2025-01-15",
  "end_date": "2025-01-22",
  "items": [
    {
      "id": {
        "entry_id": "00000000...",
        "store_id": "00000000..."
      },
      "subject": "Team Sync",
      "start": "2025-01-16T10:00:00",
      "end": "2025-01-16T11:00:00",
      "location": "Room A",
      "organizer": "organizer@example.com",
      "is_recurring": false,
      "is_all_day": false,
      "response_status": "accepted"
    }
  ]
}
```

### EventDetail

Output of `outlookctl calendar get`:

```json
{
  "version": "1.0",
  "id": {
    "entry_id": "00000000...",
    "store_id": "00000000..."
  },
  "subject": "Team Sync",
  "start": "2025-01-16T10:00:00",
  "end": "2025-01-16T11:00:00",
  "location": "Room A",
  "organizer": "organizer@example.com",
  "is_recurring": false,
  "is_all_day": false,
  "response_status": "accepted",
  "body": "Weekly team sync meeting...",
  "attendees": [
    {
      "name": "Alice",
      "email": "alice@example.com",
      "type": "required",
      "response": "accepted"
    }
  ],
  "categories": ["Meetings"],
  "reminder_minutes": 15
}
```

### EventCreateResult

Output of `outlookctl calendar create`:

```json
{
  "version": "1.0",
  "success": true,
  "id": {
    "entry_id": "00000000...",
    "store_id": "00000000..."
  },
  "saved_to": "Calendar",
  "subject": "Team Sync",
  "start": "2025-01-16T10:00:00",
  "attendees": ["alice@example.com", "bob@example.com"],
  "is_draft": true
}
```

> **Note:** `is_draft` is `true` when attendees are present but invitations haven't been sent yet.

---

## System Outputs

### DoctorResult

Output of `outlookctl doctor`:

```json
{
  "version": "1.0",
  "all_passed": true,
  "checks": [
    {
      "name": "windows_os",
      "passed": true,
      "message": "Windows OS detected",
      "remediation": null
    },
    {
      "name": "pywin32",
      "passed": true,
      "message": "pywin32 is installed and importable",
      "remediation": null
    },
    {
      "name": "outlook_com",
      "passed": true,
      "message": "Outlook COM automation is available",
      "remediation": null
    },
    {
      "name": "outlook_exe",
      "passed": true,
      "message": "Outlook executable found: C:\\...\\OUTLOOK.EXE",
      "remediation": null
    }
  ],
  "outlook_path": "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
}
```

### ErrorResult

All commands return this format on error:

```json
{
  "version": "1.0",
  "success": false,
  "error": "Description of what went wrong",
  "error_code": "OUTLOOK_UNAVAILABLE",
  "remediation": "Start Classic Outlook and try again."
}
```

---

## Error Codes

| Code | Description |
|------|-------------|
| `OUTLOOK_UNAVAILABLE` | Cannot connect to Outlook COM interface |
| `FOLDER_NOT_FOUND` | Specified folder doesn't exist |
| `MESSAGE_NOT_FOUND` | Message ID is invalid or expired |
| `EVENT_NOT_FOUND` | Calendar event ID is invalid or expired |
| `CONFIRMATION_REQUIRED` | Send attempted without proper confirmation |
| `VALIDATION_ERROR` | Invalid arguments or missing required fields |
| `DRAFT_ERROR` | Failed to create draft |
| `SEND_ERROR` | Failed to send message or meeting invitation |
| `ATTACHMENT_ERROR` | Failed to save attachments |
