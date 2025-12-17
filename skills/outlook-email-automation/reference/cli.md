# CLI Reference

Complete reference for all `outlookctl` commands and options.

## Global Options

All commands support:
- `--output json|text` - Output format (default: json)

## Email Commands

### `outlookctl doctor`

Validates environment and prerequisites.

```bash
uv run outlookctl doctor
```

---

### `outlookctl list`

List messages from a folder.

```bash
uv run outlookctl list [OPTIONS]
```

**Options:**
- `--folder` - Folder: inbox|sent|drafts|by-name:<name>|by-path:<path> (default: inbox)
- `--count` - Number of messages (default: 10)
- `--unread-only` - Only unread messages
- `--since` - ISO date filter (messages after)
- `--until` - ISO date filter (messages before)
- `--include-body-snippet` - Include body preview
- `--body-snippet-chars` - Max chars for snippet (default: 200)

---

### `outlookctl get`

Get a single message by ID.

```bash
uv run outlookctl get --id <entry_id> --store <store_id> [OPTIONS]
```

**Options:**
- `--include-body` - Include full message body
- `--include-headers` - Include message headers
- `--max-body-chars` - Limit body length

---

### `outlookctl search`

Search messages with various filters.

```bash
uv run outlookctl search [OPTIONS]
```

**Options:**
- `--folder` - Folder to search (default: inbox)
- `--query` - Free text search (subject/body)
- `--from` - Filter by sender
- `--to` - Filter by To recipient
- `--cc` - Filter by CC recipient
- `--subject-contains` - Filter by subject text
- `--unread-only` - Only unread messages
- `--has-attachments` - Only messages with attachments
- `--no-attachments` - Only messages without attachments
- `--since` - ISO date filter
- `--until` - ISO date filter
- `--count` - Maximum results (default: 50)
- `--include-body-snippet` - Include body preview

---

### `outlookctl draft`

Create a draft message.

```bash
uv run outlookctl draft [OPTIONS]
```

**Options:**
- `--to` - To recipients (comma-separated)
- `--cc` - CC recipients (comma-separated)
- `--bcc` - BCC recipients (comma-separated)
- `--subject` - Email subject
- `--body-text` - Plain text body
- `--body-html` - HTML body
- `--attach` - File path to attach (repeatable)
- `--reply-to-id` - Entry ID of message to reply to
- `--reply-to-store` - Store ID of message to reply to
- `--reply-all` - Create reply-all instead of reply

---

### `outlookctl send`

Send a draft or new message. **Requires explicit confirmation.**

```bash
# Send existing draft (recommended)
uv run outlookctl send --draft-id <id> --draft-store <store> --confirm-send YES

# Send new directly (not recommended)
uv run outlookctl send --to "email" --subject "..." --body-text "..." --unsafe-send-new --confirm-send YES
```

**Safety options:**
- `--confirm-send` - Must be exactly "YES"
- `--confirm-send-file` - Path to file containing "YES"
- `--unsafe-send-new` - Required for direct send
- `--log-body` - Include body in audit log

---

### `outlookctl move`

Move a message to another folder.

```bash
uv run outlookctl move --id <entry_id> --store <store_id> --dest <folder>
```

**Options:**
- `--id` - Message entry ID (required)
- `--store` - Message store ID (required)
- `--dest` - Destination folder (required)

---

### `outlookctl delete`

Delete a message.

```bash
uv run outlookctl delete --id <entry_id> --store <store_id> [--permanent]
```

**Options:**
- `--id` - Message entry ID (required)
- `--store` - Message store ID (required)
- `--permanent` - Permanently delete (skip Deleted Items)

---

### `outlookctl mark-read`

Mark a message as read or unread.

```bash
uv run outlookctl mark-read --id <entry_id> --store <store_id> [--unread]
```

**Options:**
- `--id` - Message entry ID (required)
- `--store` - Message store ID (required)
- `--unread` - Mark as unread instead of read

---

### `outlookctl forward`

Create a forward draft for a message.

```bash
uv run outlookctl forward --id <entry_id> --store <store_id> --to "recipient@example.com"
```

**Options:**
- `--id` - Original message entry ID (required)
- `--store` - Original message store ID (required)
- `--to` - To recipients (required)
- `--cc` - CC recipients
- `--bcc` - BCC recipients
- `--message` - Additional text to prepend

---

### `outlookctl attachments save`

Save attachments from a message to disk.

```bash
uv run outlookctl attachments save --id <entry_id> --store <store_id> --dest <directory>
```

---

## Calendar Commands

### `outlookctl calendar list`

List calendar events within a date range.

```bash
uv run outlookctl calendar list [OPTIONS]
```

**Options:**
- `--start` - Start date (default: today)
- `--end` - End date (overrides --days)
- `--days` - Number of days (default: 7)
- `--calendar` - Email for shared calendar
- `--count` - Maximum events (default: 100)

---

### `outlookctl calendar get`

Get detailed information about a calendar event.

```bash
uv run outlookctl calendar get --id <entry_id> --store <store_id> [--include-body]
```

---

### `outlookctl calendar create`

Create a calendar event or meeting.

```bash
uv run outlookctl calendar create --subject <subject> --start <datetime> [OPTIONS]
```

**Options:**
- `--subject` - Event title (required)
- `--start` - Start time (required)
- `--duration` - Duration in minutes (default: 60)
- `--end` - End time (overrides duration)
- `--location` - Event location
- `--body` - Event description
- `--attendees` - Required attendees (comma-separated)
- `--optional-attendees` - Optional attendees
- `--all-day` - Create as all-day event
- `--reminder` - Reminder minutes (default: 15)
- `--busy-status` - Show as: free, tentative, busy, out_of_office, working_elsewhere
- `--teams-url` - Teams URL to embed
- `--recurrence` - Recurrence pattern
- `--send-now` - Send invites immediately
- `--confirm-send` - Confirmation for --send-now

---

### `outlookctl calendar send`

Send meeting invitations. **Requires confirmation.**

```bash
uv run outlookctl calendar send --id <entry_id> --store <store_id> --confirm-send YES
```

---

### `outlookctl calendar respond`

Respond to a meeting invitation.

```bash
uv run outlookctl calendar respond --id <entry_id> --store <store_id> --response accept|decline|tentative
```

**Options:**
- `--no-response` - Don't send response to organizer

---

### `outlookctl calendar update`

Update an existing calendar event.

```bash
uv run outlookctl calendar update --id <entry_id> --store <store_id> [OPTIONS]
```

**Options:**
- `--id` - Event entry ID (required)
- `--store` - Event store ID (required)
- `--subject` - New event subject
- `--start` - New start time
- `--end` - New end time
- `--duration` - New duration in minutes
- `--location` - New location
- `--body` - New description
- `--reminder` - New reminder minutes
- `--busy-status` - New show-as status

---

### `outlookctl calendar delete`

Delete a calendar event.

```bash
uv run outlookctl calendar delete --id <entry_id> --store <store_id> [--no-cancel]
```

**Options:**
- `--id` - Event entry ID (required)
- `--store` - Event store ID (required)
- `--no-cancel` - Don't send cancellation to attendees
