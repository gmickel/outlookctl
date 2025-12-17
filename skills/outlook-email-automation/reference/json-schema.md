# JSON Schema Reference

All `outlookctl` commands output JSON with consistent schemas. This document describes the output formats.

## Common Types

### MessageId

Stable identifier for referencing messages across commands.

```json
{
  "entry_id": "00000000...",
  "store_id": "00000000..."
}
```

Both fields are required when referencing a message (e.g., in `get`, `send`, `attachments save`).

### EmailAddress

```json
{
  "name": "John Doe",
  "email": "john.doe@example.com"
}
```

## Command Outputs

### List Result

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

**Notes:**
- `body_snippet` is only included when `--include-body-snippet` is used
- Items are sorted by received time (newest first)

### Search Result

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
    // Same structure as list items
  ]
}
```

### Get Result (Message Detail)

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

**Notes:**
- `body` and `body_html` only included with `--include-body`
- `headers` only included with `--include-headers`

### Draft Result

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

### Send Result

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

### Attachment Save Result

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

### Doctor Result

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

### Error Result

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

**Common error codes:**
- `OUTLOOK_UNAVAILABLE` - Cannot connect to Outlook COM
- `FOLDER_NOT_FOUND` - Specified folder doesn't exist
- `MESSAGE_NOT_FOUND` - Message ID is invalid or expired
- `CONFIRMATION_REQUIRED` - Send attempted without proper confirmation
- `VALIDATION_ERROR` - Invalid arguments or missing required fields
- `DRAFT_ERROR` - Failed to create draft
- `SEND_ERROR` - Failed to send message
- `ATTACHMENT_ERROR` - Failed to save attachments

## Schema Versioning

All outputs include a `version` field. Current version is `"1.0"`. Future versions will maintain backward compatibility where possible.
