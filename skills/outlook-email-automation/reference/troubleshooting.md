# Troubleshooting

Common issues and solutions for `outlookctl`.

---

## Table of Contents

- [Quick Diagnostics](#quick-diagnostics)
- [Common Issues](#common-issues)
  - [Outlook COM Unavailable](#outlook-com-unavailable)
  - [pywin32 Not Installed](#pywin32-not-installed)
  - [Folder Not Found](#folder-not-found)
  - [Message/Event Not Found](#messageevent-not-found)
  - [Send Confirmation Required](#send-confirmation-required)
  - [Attachment Errors](#attachment-errors)
  - [Unicode/Encoding Errors](#unicodeencoding-errors)
  - [Outlook Security Prompts](#outlook-security-prompts)
  - [Slow Performance](#slow-performance)
- [Environment-Specific Issues](#environment-specific-issues)
- [Getting Help](#getting-help)

---

## Quick Diagnostics

Always start with the doctor command:

```bash
uv run python -m outlookctl.cli doctor
```

This checks all prerequisites and provides specific remediation advice.

---

## Common Issues

### Outlook COM Unavailable

**Error:** `OUTLOOK_UNAVAILABLE` - "Could not connect to Outlook"

| Cause | Solution |
|-------|----------|
| Classic Outlook not running | Start Outlook manually, wait for full load |
| New Outlook is active | Switch to Classic: Settings > General > Use new Outlook > **Off** |
| Outlook not installed | Install Microsoft 365 / Office suite with Outlook |
| COM registration issue | Repair Office: Settings > Apps > Microsoft 365 > Modify > Repair |

---

### pywin32 Not Installed

**Error:** "pywin32 is not installed"

```bash
# Using uv (recommended)
uv add pywin32
uv sync

# Or with pip
pip install pywin32
```

---

### Folder Not Found

**Error:** `FOLDER_NOT_FOUND`

**Standard folders** (case-insensitive):

```bash
uv run python -m outlookctl.cli list --folder inbox
uv run python -m outlookctl.cli list --folder sent
uv run python -m outlookctl.cli list --folder drafts
```

**Custom folders:**

```bash
# By name
uv run python -m outlookctl.cli list --folder "by-name:Projects"

# By path
uv run python -m outlookctl.cli list --folder "by-path:Inbox/Projects/2025"
```

---

### Message/Event Not Found

**Error:** `MESSAGE_NOT_FOUND` or `EVENT_NOT_FOUND`

| Cause | Solution |
|-------|----------|
| Item was deleted or moved | Re-run `list` or `search` for fresh IDs |
| Entry ID expired | IDs can become stale - get new ones |
| Wrong store ID | Ensure `--id` and `--store` are from same item |

---

### Send Confirmation Required

**Error:** `CONFIRMATION_REQUIRED`

This is **intentional**. You must provide explicit confirmation:

```bash
# Sending a draft (email)
uv run python -m outlookctl.cli send \
  --draft-id "..." \
  --draft-store "..." \
  --confirm-send YES

# Sending meeting invitations
uv run python -m outlookctl.cli calendar send \
  --id "..." \
  --store "..." \
  --confirm-send YES

# Direct send (not recommended)
uv run python -m outlookctl.cli send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body-text "Body" \
  --unsafe-send-new \
  --confirm-send YES
```

---

### Attachment Errors

**When creating draft:**

- Verify the file path exists
- Use absolute paths for reliability
- Check file permissions

**When saving attachments:**

- Ensure message has attachments (`has_attachments: true`)
- Verify destination directory is writable

---

### Unicode/Encoding Errors

**Symptoms:** Garbled characters, `UnicodeDecodeError`

| Solution | Command |
|----------|---------|
| Use PowerShell Core | Instead of cmd.exe |
| Set encoding | `$env:PYTHONIOENCODING = "utf-8"` |
| Check terminal | Ensure UTF-8 support |

---

### Outlook Security Prompts

**Symptom:** Dialog: "A program is trying to access email..."

| Solution | Steps |
|----------|-------|
| Click "Allow" | One-time approval |
| Trust Center | File > Options > Trust Center > Programmatic Access > "Never warn" |
| Group Policy | Contact IT for enterprise configuration |

---

### Slow Performance

**Large mailboxes:**

```bash
# Limit results
uv run python -m outlookctl.cli list --count 10

# Use date filters
uv run python -m outlookctl.cli search --since 2025-01-01 --until 2025-01-31

# Search specific folders
uv run python -m outlookctl.cli list --folder inbox
```

**COM initialization:**

- First command may be slow (COM initialization)
- Subsequent commands are faster

---

## Environment-Specific Issues

### Windows Server / RDP

| Issue | Solution |
|-------|----------|
| COM not connecting | Ensure Outlook runs in same session |
| Cross-session issues | Run both Outlook and outlookctl in same RDP session |

### Virtual Machines

| Consideration | Notes |
|---------------|-------|
| COM support | Works in VMs with Outlook installed |
| Memory | Ensure adequate RAM for Outlook |

### WSL (Windows Subsystem for Linux)

COM automation requires Windows-native Python:

```bash
# Option 1: Call via PowerShell
powershell.exe -c "uv run python -m outlookctl.cli ..."

# Option 2: Install Python in Windows
# Use Windows Python directly
```

---

## Getting Help

If issues persist:

1. **Run diagnostics:**
   ```bash
   uv run python -m outlookctl.cli doctor
   ```

2. **Verify Outlook version:**
   - Must be Classic Outlook (not New Outlook)
   - Check: Settings > General > "Use new Outlook" should be **Off**

3. **Try restarting Outlook**

4. **Verify pywin32:**
   ```bash
   python -c "import win32com.client; print('OK')"
   ```

5. **Get version info:**
   ```bash
   uv run python -m outlookctl.cli --version
   ```

Include version info and `doctor` output when reporting issues.
