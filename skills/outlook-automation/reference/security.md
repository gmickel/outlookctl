# Security and Data Handling

This document describes the security model and data handling practices for `outlookctl`.

## Overview

`outlookctl` operates entirely locally using Outlook's COM automation interface. There are no network calls to external services, no cloud storage, and no OAuth/API tokens required.

## Data Minimization

### Default Behavior

By default, commands return **metadata only**:
- Subject line
- Sender/recipient email addresses
- Timestamps
- Attachment filenames (not content)
- Read/unread status

**Body content is never retrieved unless explicitly requested** using:
- `--include-body` (for `get` command)
- `--include-body-snippet` (for `list` and `search` commands)

### Opt-In Body Access

To retrieve message body content, you must explicitly request it:

```bash
# Get full body
uv run outlookctl get --id "..." --store "..." --include-body

# Get truncated snippet
uv run outlookctl list --include-body-snippet --body-snippet-chars 200
```

### Redaction Options

For sensitive environments, consider:
- Using `--max-body-chars` to limit body size
- Processing only metadata (default behavior)
- Implementing additional filtering in your workflow

## Send Safety Gates

### Two-Step Workflow

The recommended workflow for sending email:

1. **Create a draft** - `outlookctl draft ...`
2. **Review** - Show user the subject/recipients/body preview
3. **Explicit send** - `outlookctl send --draft-id ... --confirm-send YES`

### Required Confirmation

The `send` command will **refuse to execute** unless one of these conditions is met:

1. `--confirm-send YES` flag with exact string "YES"
2. `--confirm-send-file <path>` pointing to a file containing "YES"

```bash
# This will fail
uv run outlookctl send --draft-id "..." --draft-store "..."

# This will succeed
uv run outlookctl send --draft-id "..." --draft-store "..." --confirm-send YES
```

### Unsafe Direct Send

Sending a new message without first creating a draft requires **additional confirmation**:

```bash
# Requires BOTH flags
uv run outlookctl send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body-text "Body" \
  --unsafe-send-new \
  --confirm-send YES
```

This is intentionally cumbersome to discourage bypassing the draft workflow.

## Audit Logging

### Location

Audit logs are stored at:
- Windows: `%LOCALAPPDATA%\outlookctl\audit.log`
- Fallback: `~/.outlookctl/audit.log`

### What's Logged

For send operations, the audit log records:

```json
{
  "timestamp": "2025-01-15T10:00:00+00:00",
  "operation": "send",
  "success": true,
  "recipients": {
    "to_count": 1,
    "cc_count": 0,
    "bcc_count": 0
  },
  "subject_length": 25,
  "entry_id": "00000..."
}
```

**Note:** By default, the log contains **counts and lengths only**, not actual content.

### Logging Body Content

To include body content in the audit log (not recommended for sensitive data):

```bash
uv run outlookctl send --draft-id "..." --draft-store "..." --confirm-send YES --log-body
```

## Outlook Security Prompts

### Programmatic Access Warning

Outlook may display a security prompt when accessing certain properties programmatically:

> "A program is trying to access email addresses stored in Outlook..."

This is a Windows/Outlook security feature. Options:
1. Click "Allow" when prompted
2. Configure Outlook Trust Center settings
3. Use Group Policy (enterprise environments)

### COM Security

The tool uses standard COM automation (`Outlook.Application`). This:
- Runs in the user's security context
- Respects Outlook's security settings
- May be blocked by some antivirus software

## Best Practices

### For Claude/AI Usage

1. **Always use draft-first workflow** - Create drafts, show preview, send after confirmation
2. **Minimize data retrieval** - Only fetch body when explicitly needed
3. **Don't store credentials** - No passwords or tokens are needed
4. **Respect user intent** - Never auto-send without explicit user instruction

### For Sensitive Environments

1. Review audit logs periodically
2. Use `--max-body-chars` to limit data exposure
3. Consider disabling `--log-body` entirely
4. Monitor for unexpected send operations

## No Network Access Required

This tool:
- Does NOT require Microsoft Graph API
- Does NOT require Azure app registration
- Does NOT require OAuth tokens
- Does NOT make any network calls

All operations happen locally through the Outlook COM interface.

## Limitations

- Only works with **Classic Outlook** (not New Outlook)
- Requires Outlook to be running and logged in
- Subject to corporate Outlook policies
- COM automation may trigger security prompts
