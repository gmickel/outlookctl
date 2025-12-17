# Security & Data Handling

This document describes the security model and data handling practices for `outlookctl`.

---

## Table of Contents

- [Overview](#overview)
- [Data Minimization](#data-minimization)
- [Send Safety Gates](#send-safety-gates)
- [Audit Logging](#audit-logging)
- [Outlook Security Prompts](#outlook-security-prompts)
- [Best Practices](#best-practices)
- [Limitations](#limitations)

---

## Overview

`outlookctl` operates **entirely locally** using Outlook's COM automation interface.

| Feature | Status |
|---------|--------|
| Network calls to external services | **None** |
| Cloud storage | **None** |
| OAuth/API tokens required | **None** |
| Microsoft Graph API | **Not required** |
| Azure app registration | **Not required** |

All operations happen locally through the Outlook COM interface, using your existing Windows authentication.

---

## Data Minimization

### Default Behavior

By default, commands return **metadata only**:

- Subject line
- Sender/recipient email addresses
- Timestamps
- Attachment filenames (not content)
- Read/unread status

> **Important:** Body content is **never** retrieved unless explicitly requested.

### Opt-In Body Access

To retrieve message body content, you must explicitly request it:

```bash
# Get full body
uv run python -m outlookctl.cli get --id "..." --store "..." --include-body

# Get truncated snippet (list/search)
uv run python -m outlookctl.cli list --include-body-snippet --body-snippet-chars 200
```

### Redaction Options

For sensitive environments:

| Option | Purpose |
|--------|---------|
| `--max-body-chars` | Limit body size returned |
| Default (no flags) | Metadata only |
| Custom filtering | Implement in your workflow |

---

## Send Safety Gates

### Two-Step Workflow (Recommended)

```
1. Create draft    →    outlookctl draft ...
2. Review          →    Show user subject/recipients/body preview
3. Explicit send   →    outlookctl send --draft-id ... --confirm-send YES
```

### Required Confirmation

The `send` command will **refuse to execute** unless:

| Method | Example |
|--------|---------|
| Flag confirmation | `--confirm-send YES` (exact string) |
| File confirmation | `--confirm-send-file <path>` containing "YES" |

```bash
# This will FAIL
uv run python -m outlookctl.cli send --draft-id "..." --draft-store "..."

# This will SUCCEED
uv run python -m outlookctl.cli send --draft-id "..." --draft-store "..." --confirm-send YES
```

### Calendar Safety

Meeting invitations follow the same pattern:

```bash
# Create meeting (saved as draft with attendees)
uv run python -m outlookctl.cli calendar create --subject "..." --start "..." --attendees "..."

# Send invitations (requires confirmation)
uv run python -m outlookctl.cli calendar send --id "..." --store "..." --confirm-send YES
```

### Unsafe Direct Send

Sending a new message without first creating a draft requires **additional confirmation**:

```bash
# Requires BOTH flags - intentionally cumbersome
uv run python -m outlookctl.cli send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body-text "Body" \
  --unsafe-send-new \
  --confirm-send YES
```

> This is intentionally cumbersome to discourage bypassing the draft workflow.

---

## Audit Logging

### Location

| Platform | Path |
|----------|------|
| Windows | `%LOCALAPPDATA%\outlookctl\audit.log` |
| Fallback | `~/.outlookctl/audit.log` |

### What's Logged

For send operations (email and calendar):

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

> **Privacy:** By default, logs contain **counts and lengths only**, not actual content.

### Logging Body Content

To include body content in the audit log (not recommended for sensitive data):

```bash
uv run python -m outlookctl.cli send --draft-id "..." --draft-store "..." --confirm-send YES --log-body
```

---

## Outlook Security Prompts

### Programmatic Access Warning

Outlook may display a security prompt:

> "A program is trying to access email addresses stored in Outlook..."

**Solutions:**

| Option | Description |
|--------|-------------|
| Click "Allow" | One-time approval |
| Trust Center | File > Options > Trust Center > Programmatic Access |
| Group Policy | Enterprise configuration via IT |

### COM Security

The tool uses standard COM automation (`Outlook.Application`):

- Runs in the user's security context
- Respects Outlook's Trust Center settings
- May be blocked by some antivirus software

---

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

---

## Limitations

| Limitation | Details |
|------------|---------|
| Classic Outlook only | New Outlook does not support COM |
| Windows only | COM is a Windows technology |
| Outlook must be running | And logged into your account |
| Corporate policies | Subject to Outlook/Exchange policies |
| Security prompts | COM automation may trigger dialogs |
