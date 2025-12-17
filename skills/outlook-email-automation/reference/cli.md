# CLI Reference

Complete reference for all `outlookctl` commands and options.

---

## Table of Contents

- [Global Options](#global-options)
- [Email Commands](#email-commands)
  - [doctor](#outlookctl-doctor)
  - [list](#outlookctl-list)
  - [get](#outlookctl-get)
  - [search](#outlookctl-search)
  - [draft](#outlookctl-draft)
  - [send](#outlookctl-send)
  - [attachments save](#outlookctl-attachments-save)
- [Calendar Commands](#calendar-commands)
  - [calendar list](#outlookctl-calendar-list)
  - [calendar get](#outlookctl-calendar-get)
  - [calendar create](#outlookctl-calendar-create)
  - [calendar send](#outlookctl-calendar-send)
  - [calendar respond](#outlookctl-calendar-respond)

---

## Global Options

All commands support:

| Option | Default | Description |
|--------|---------|-------------|
| `--output` | json | Output format: `json` or `text` |
| `--version` | - | Show version number and exit |
| `--help` | - | Show help message and exit |

---

## Email Commands

### `outlookctl doctor`

Validates environment and prerequisites. **Run this first** to ensure everything is set up correctly.

```bash
uv run python -m outlookctl.cli doctor
```

**Checks performed:**

| Check | Description |
|-------|-------------|
| `windows_os` | Windows OS detection |
| `pywin32` | pywin32 installation |
| `outlook_com` | Outlook COM availability |
| `outlook_exe` | Outlook executable location |

**Example output:**

```json
{
  "version": "1.0",
  "all_passed": true,
  "checks": [
    {"name": "windows_os", "passed": true, "message": "Windows OS detected"},
    {"name": "pywin32", "passed": true, "message": "pywin32 is installed"},
    {"name": "outlook_com", "passed": true, "message": "Outlook COM available"},
    {"name": "outlook_exe", "passed": true, "message": "Found: C:\\...\\OUTLOOK.EXE"}
  ],
  "outlook_path": "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
}
```

---

### `outlookctl list`

List messages from a folder.

```bash
uv run python -m outlookctl.cli list [OPTIONS]
```

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--folder` | inbox | Folder specification (see below) |
| `--count` | 10 | Number of messages to return |
| `--unread-only` | false | Only return unread messages |
| `--since` | - | ISO date filter (messages after) |
| `--until` | - | ISO date filter (messages before) |
| `--include-body-snippet` | false | Include body preview |
| `--body-snippet-chars` | 200 | Max chars for snippet |

**Folder specifications:**

| Folder | Description |
|--------|-------------|
| `inbox` | Default inbox |
| `sent` | Sent items |
| `drafts` | Drafts folder |
| `deleted` | Deleted items |
| `outbox` | Outbox |
| `junk` | Junk/spam |
| `by-name:<name>` | Find folder by name |
| `by-path:<path>` | Find folder by path (e.g., `Inbox/Subfolder`) |

**Example:**

```bash
uv run python -m outlookctl.cli list --folder inbox --count 5 --unread-only --include-body-snippet
```

---

### `outlookctl get`

Get a single message by ID.

```bash
uv run python -m outlookctl.cli get --id <entry_id> --store <store_id> [OPTIONS]
```

**Required:**

| Option | Description |
|--------|-------------|
| `--id` | Message entry ID |
| `--store` | Message store ID |

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--include-body` | false | Include full message body |
| `--include-headers` | false | Include message headers |
| `--max-body-chars` | - | Limit body length |

**Example:**

```bash
uv run python -m outlookctl.cli get --id "00000..." --store "00000..." --include-body --max-body-chars 5000
```

---

### `outlookctl search`

Search messages with various filters.

```bash
uv run python -m outlookctl.cli search [OPTIONS]
```

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--folder` | inbox | Folder to search |
| `--query` | - | Free text search (subject/body) |
| `--from` | - | Filter by sender |
| `--subject-contains` | - | Filter by subject text |
| `--unread-only` | false | Only unread messages |
| `--since` | - | ISO date filter |
| `--until` | - | ISO date filter |
| `--count` | 50 | Maximum results |
| `--include-body-snippet` | false | Include body preview |
| `--body-snippet-chars` | 200 | Max chars for snippet |

**Example:**

```bash
uv run python -m outlookctl.cli search --from "boss@company.com" --since 2025-01-01 --unread-only
```

---

### `outlookctl draft`

Create a draft message.

```bash
uv run python -m outlookctl.cli draft [OPTIONS]
```

**Options:**

| Option | Description |
|--------|-------------|
| `--to` | To recipients (comma-separated) |
| `--cc` | CC recipients (comma-separated) |
| `--bcc` | BCC recipients (comma-separated) |
| `--subject` | Email subject |
| `--body-text` | Plain text body |
| `--body-html` | HTML body (mutually exclusive with --body-text) |
| `--attach` | File path to attach (repeatable) |
| `--reply-to-id` | Entry ID of message to reply to |
| `--reply-to-store` | Store ID of message to reply to |

**Example:**

```bash
uv run python -m outlookctl.cli draft \
  --to "recipient@example.com" \
  --cc "cc@example.com" \
  --subject "Meeting Follow-up" \
  --body-text "Thank you for the meeting today." \
  --attach "./report.pdf"
```

**Reply example:**

```bash
uv run python -m outlookctl.cli draft \
  --to "sender@example.com" \
  --subject "Re: Original Subject" \
  --body-text "Reply content" \
  --reply-to-id "00000..." \
  --reply-to-store "00000..."
```

---

### `outlookctl send`

Send a draft or new message. **Requires explicit confirmation.**

#### Sending an existing draft (recommended):

```bash
uv run python -m outlookctl.cli send \
  --draft-id <entry_id> \
  --draft-store <store_id> \
  --confirm-send YES
```

#### Sending a new message directly (not recommended):

```bash
uv run python -m outlookctl.cli send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body-text "Body" \
  --unsafe-send-new \
  --confirm-send YES
```

**Safety options:**

| Option | Description |
|--------|-------------|
| `--confirm-send` | Must be exactly "YES" to proceed |
| `--confirm-send-file` | Path to file containing "YES" |
| `--unsafe-send-new` | Required flag for sending new message directly |
| `--log-body` | Include body in audit log |

> **Warning:** The `--unsafe-send-new` flag is intentionally cumbersome to discourage bypassing the draft-first workflow.

---

### `outlookctl attachments save`

Save attachments from a message to disk.

```bash
uv run python -m outlookctl.cli attachments save \
  --id <entry_id> \
  --store <store_id> \
  --dest <directory>
```

**Required:**

| Option | Description |
|--------|-------------|
| `--id` | Message entry ID |
| `--store` | Message store ID |
| `--dest` | Destination directory (created if needed) |

**Example:**

```bash
uv run python -m outlookctl.cli attachments save --id "00000..." --store "00000..." --dest "./downloads"
```

**Output:**

```json
{
  "version": "1.0",
  "success": true,
  "saved_files": [
    "./downloads/document.pdf",
    "./downloads/image.png"
  ],
  "errors": []
}
```

---

## Calendar Commands

### `outlookctl calendar list`

List calendar events within a date range.

```bash
uv run python -m outlookctl.cli calendar list [OPTIONS]
```

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--start` | today | Start date (YYYY-MM-DD or YYYY-MM-DD HH:MM) |
| `--end` | - | End date (overrides --days) |
| `--days` | 7 | Number of days from start |
| `--calendar` | - | Email address for shared calendar |
| `--count` | 100 | Maximum events to return |

**Example:**

```bash
uv run python -m outlookctl.cli calendar list --start "2025-01-20" --days 14
```

**Output:**

```json
{
  "version": "1.0",
  "calendar": "Calendar",
  "start_date": "2025-01-20T00:00:00",
  "end_date": "2025-02-03T00:00:00",
  "items": [
    {
      "id": {"entry_id": "...", "store_id": "..."},
      "subject": "Team Meeting",
      "start": "2025-01-20T10:00:00",
      "end": "2025-01-20T11:00:00",
      "location": "Conference Room A",
      "organizer": "boss@example.com",
      "is_recurring": false,
      "is_all_day": false,
      "is_meeting": true,
      "response_status": "accepted",
      "busy_status": "busy"
    }
  ]
}
```

---

### `outlookctl calendar get`

Get detailed information about a calendar event.

```bash
uv run python -m outlookctl.cli calendar get --id <entry_id> --store <store_id> [OPTIONS]
```

**Required:**

| Option | Description |
|--------|-------------|
| `--id` | Event entry ID |
| `--store` | Event store ID |

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--include-body` | false | Include event description |

**Example:**

```bash
uv run python -m outlookctl.cli calendar get --id "00000..." --store "00000..." --include-body
```

---

### `outlookctl calendar create`

Create a calendar event or meeting.

```bash
uv run python -m outlookctl.cli calendar create --subject <subject> --start <datetime> [OPTIONS]
```

**Required:**

| Option | Description |
|--------|-------------|
| `--subject` | Event title |
| `--start` | Start time (YYYY-MM-DD HH:MM or YYYY-MM-DD for all-day) |

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--duration` | 60 | Duration in minutes (ignored if --end set) |
| `--end` | - | End time (YYYY-MM-DD HH:MM) |
| `--location` | - | Event location |
| `--body` | - | Event description |
| `--attendees` | - | Required attendees (comma-separated emails) |
| `--optional-attendees` | - | Optional attendees (comma-separated emails) |
| `--all-day` | false | Create as all-day event |
| `--reminder` | 15 | Reminder minutes before event |
| `--busy-status` | busy | Show as: free, tentative, busy, out_of_office, working_elsewhere |
| `--teams-url` | - | Teams meeting URL to embed in body |
| `--recurrence` | - | Recurrence pattern (see below) |
| `--send-now` | false | Send invites immediately (requires --confirm-send YES) |
| `--confirm-send` | - | Confirmation string for --send-now |

**Recurrence patterns:**

| Pattern | Example |
|---------|---------|
| Weekly on specific days | `--recurrence "weekly:monday,wednesday:until:2025-12-31"` |
| Daily with count | `--recurrence "daily:count:10"` |
| Monthly on day of month | `--recurrence "monthly:day:15:until:2025-06-01"` |

**Example (personal event):**

```bash
uv run python -m outlookctl.cli calendar create \
  --subject "Focus Time" \
  --start "2025-01-20 14:00" \
  --duration 120
```

**Example (meeting - saved as draft):**

```bash
uv run python -m outlookctl.cli calendar create \
  --subject "Team Sync" \
  --start "2025-01-20 10:00" \
  --duration 60 \
  --location "Room A" \
  --attendees "alice@example.com,bob@example.com"
```

> **Note:** Meetings with attendees are saved as drafts. Use `calendar send` to send invitations.

**Output:**

```json
{
  "version": "1.0",
  "success": true,
  "id": {"entry_id": "...", "store_id": "..."},
  "saved_to": "Calendar",
  "subject": "Team Sync",
  "start": "2025-01-20T10:00:00",
  "attendees": ["alice@example.com", "bob@example.com"],
  "is_draft": true
}
```

---

### `outlookctl calendar send`

Send meeting invitations for an existing event. **Requires explicit confirmation.**

```bash
uv run python -m outlookctl.cli calendar send --id <entry_id> --store <store_id> --confirm-send YES
```

**Required:**

| Option | Description |
|--------|-------------|
| `--id` | Event entry ID |
| `--store` | Event store ID |
| `--confirm-send` | Must be exactly "YES" |

**Options:**

| Option | Description |
|--------|-------------|
| `--confirm-send-file` | Path to file containing "YES" |

**Example:**

```bash
uv run python -m outlookctl.cli calendar send --id "00000..." --store "00000..." --confirm-send YES
```

**Output:**

```json
{
  "version": "1.0",
  "success": true,
  "message": "Meeting invitations sent",
  "sent_at": "2025-01-20T09:00:00",
  "attendees": ["alice@example.com", "bob@example.com"],
  "subject": "Team Sync"
}
```

---

### `outlookctl calendar respond`

Respond to a meeting invitation.

```bash
uv run python -m outlookctl.cli calendar respond --id <entry_id> --store <store_id> --response <response>
```

**Required:**

| Option | Description |
|--------|-------------|
| `--id` | Event entry ID |
| `--store` | Event store ID |
| `--response` | One of: `accept`, `decline`, `tentative` |

**Options:**

| Option | Default | Description |
|--------|---------|-------------|
| `--no-response` | false | Don't send response to organizer |

**Example:**

```bash
uv run python -m outlookctl.cli calendar respond --id "00000..." --store "00000..." --response accept
```

**Output:**

```json
{
  "version": "1.0",
  "success": true,
  "response": "accepted",
  "subject": "Team Meeting",
  "organizer": "boss@example.com"
}
```
