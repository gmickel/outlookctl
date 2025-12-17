# Troubleshooting

Common issues and solutions for `outlookctl`.

## Quick Diagnostics

Always start with the doctor command:

```bash
uv run outlookctl doctor
```

This checks all prerequisites and provides specific remediation advice.

## Common Issues

### "Outlook COM automation unavailable"

**Symptoms:**
- Error code: `OUTLOOK_UNAVAILABLE`
- Message: "Could not connect to Outlook"

**Causes and Solutions:**

1. **Classic Outlook not running**
   - Start Outlook manually
   - Wait for it to fully load before retrying

2. **New Outlook is active**
   - New Outlook (the modern app) does NOT support COM automation
   - Switch to Classic Outlook: Settings > General > Use new Outlook > Off
   - Or start Classic Outlook directly from the Office folder

3. **Outlook not installed**
   - Install Microsoft Office with Outlook
   - Classic Outlook is part of Microsoft 365 / Office suite

4. **COM registration issue**
   - Try repairing Office installation
   - Run Office repair: Settings > Apps > Microsoft 365 > Modify > Repair

### "pywin32 is not installed"

**Solution:**
```bash
uv add pywin32
uv sync
```

Or with pip:
```bash
pip install pywin32
```

### "Folder not found"

**Symptoms:**
- Error code: `FOLDER_NOT_FOUND`

**Solutions:**

1. **Check folder name spelling**
   ```bash
   # Use the default folders
   uv run outlookctl list --folder inbox
   uv run outlookctl list --folder sent
   uv run outlookctl list --folder drafts
   ```

2. **For custom folders, use by-name or by-path**
   ```bash
   # Search by name
   uv run outlookctl list --folder "by-name:Projects"

   # Search by path
   uv run outlookctl list --folder "by-path:Inbox/Projects/2025"
   ```

3. **Folder names are case-insensitive**

### "Message not found"

**Symptoms:**
- Error code: `MESSAGE_NOT_FOUND`

**Causes:**
1. Message was deleted or moved
2. Entry ID has expired (IDs can become stale)
3. Wrong store ID

**Solutions:**
1. Re-run `list` or `search` to get fresh IDs
2. Ensure both `--id` and `--store` are from the same message

### "Send confirmation required"

**Symptoms:**
- Error code: `CONFIRMATION_REQUIRED`

**Solution:**
This is intentional. You must provide explicit confirmation:

```bash
# For sending a draft
uv run outlookctl send \
  --draft-id "..." \
  --draft-store "..." \
  --confirm-send YES

# For sending new message directly (not recommended)
uv run outlookctl send \
  --to "recipient@example.com" \
  --subject "Subject" \
  --body-text "Body" \
  --unsafe-send-new \
  --confirm-send YES
```

### "Attachment not found"

**When creating draft:**
- Verify the file path exists
- Use absolute paths for reliability
- Check file permissions

**When saving attachments:**
- Ensure the message actually has attachments (`has_attachments: true`)
- Verify destination directory is writable

### Unicode/Encoding Errors

**Symptoms:**
- Garbled characters in output
- UnicodeDecodeError

**Solutions:**
1. Ensure terminal supports UTF-8
2. Try PowerShell Core instead of cmd.exe
3. Set environment: `$env:PYTHONIOENCODING = "utf-8"`

### Outlook Security Prompts

**Symptom:**
Dialog appears: "A program is trying to access email..."

**Solutions:**

1. **Click Allow** - One-time approval
2. **Trust Center settings** (Outlook):
   - File > Options > Trust Center > Trust Center Settings
   - Programmatic Access > Set to "Never warn me"
3. **Group Policy** (enterprise):
   - Contact IT for policy configuration

### Slow Performance

**Large mailboxes:**
- Use `--count` to limit results
- Use date filters (`--since`, `--until`)
- Search specific folders instead of all

**COM initialization:**
- First command may be slow as COM initializes
- Subsequent commands are faster

## Environment-Specific Issues

### Windows Server / RDP

- Ensure Outlook is running in the same session
- COM may not work across sessions
- Consider running Outlook and outlookctl in the same RDP session

### Virtual Machines

- COM works in VMs with Outlook installed
- Ensure adequate memory for Outlook

### WSL (Windows Subsystem for Linux)

- COM automation requires Windows-native Python
- Call outlookctl via PowerShell: `powershell.exe -c "uv run outlookctl ..."`
- Or install Python + dependencies in Windows

## Getting Help

If issues persist:

1. Run `uv run outlookctl doctor` and note all failures
2. Check Outlook is Classic (not New Outlook)
3. Try restarting Outlook
4. Verify pywin32: `python -c "import win32com.client; print('OK')"`

## Version Information

```bash
uv run outlookctl --version
```

Include version info when reporting issues.
