# CLAUDE.md - Development Guide

This document provides context for AI assistants (Claude) working on the outlookctl project.

## Project Overview

**outlookctl** is a local CLI tool for automating Classic Outlook on Windows via COM automation. It includes a Claude Code Skill that enables AI-assisted email management.

### Key Points

- **Local automation only** - No external APIs, no OAuth, no cloud services
- **COM-based** - Uses Windows COM (`Outlook.Application`) via pywin32
- **Classic Outlook required** - New Outlook does not support COM
- **Safety-first design** - Draft-first workflow, explicit send confirmation

## Architecture

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   Claude Code   │────▶│    outlookctl    │────▶│ Classic Outlook │
│     (Skill)     │     │   (Python CLI)   │     │     (COM)       │
└─────────────────┘     └──────────────────┘     └─────────────────┘
```

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `cli.py` | Argparse CLI entry point, command routing |
| `models.py` | Dataclasses for JSON serialization |
| `outlook_com.py` | COM automation wrapper, all Outlook interactions |
| `safety.py` | Send confirmation gates, validation |
| `audit.py` | Audit logging for send operations |

## Running Commands

### During Development

```bash
# From project directory
uv run python -m outlookctl.cli <command> [options]

# Examples
uv run python -m outlookctl.cli doctor
uv run python -m outlookctl.cli list --count 5
uv run python -m outlookctl.cli search --from "someone@example.com"
```

### From Any Directory

```bash
uv run --project "C:/Users/GordonMickel/work/outlookctl" python -m outlookctl.cli <command>
```

### Running Tests

```bash
uv run python -m pytest tests/ -v
```

## Code Patterns

### JSON Output

All commands output JSON via `models.py` dataclasses:

```python
from outlookctl.models import ListResult, MessageSummary, FolderInfo

result = ListResult(
    folder=FolderInfo(name="Inbox"),
    items=[...],
)
print(json.dumps(result.to_dict(), indent=2))
```

### Error Handling

Use `ErrorResult` for consistent error output:

```python
from outlookctl.models import ErrorResult

def output_error(error: str, error_code: str = None, remediation: str = None):
    result = ErrorResult(error=error, error_code=error_code, remediation=remediation)
    print(json.dumps(result.to_dict(), indent=2))
    sys.exit(1)
```

### COM Access Pattern

Always use `get_outlook_app()` which handles retries and error messages:

```python
from outlookctl.outlook_com import get_outlook_app, OutlookNotAvailableError

try:
    outlook = get_outlook_app()
    # Use outlook object...
except OutlookNotAvailableError as e:
    output_error(str(e), "OUTLOOK_UNAVAILABLE")
```

### Safety Gates

All send operations must go through safety validation:

```python
from outlookctl.safety import validate_send_confirmation, SendConfirmationError

try:
    validate_send_confirmation(confirm_send="YES")
except SendConfirmationError as e:
    output_error(str(e), "CONFIRMATION_REQUIRED")
```

## Skill System

### Skill Location

- **Personal**: `~/.claude/skills/outlook-automation/`
- **Project**: `.claude/skills/outlook-automation/`

### SKILL.md Structure

```yaml
---
name: outlook-automation
description: >
  Brief description for Claude to understand when to use this skill...
---

# Skill content with usage instructions...
```

### Installing/Updating Skill

```bash
uv run python tools/install_skill.py --personal
```

## Important Constraints

### Never Auto-Send

The skill instructs Claude to **always use draft-first workflow**:
1. Create draft with `draft` command
2. Show preview to user
3. Only send after explicit user confirmation

### Metadata by Default

Body content should only be retrieved when explicitly requested:
- `list` and `search` return metadata only by default
- Use `--include-body` or `--include-body-snippet` when needed

### Confirmation Required

The `send` command requires `--confirm-send YES` (exact string match).

## Common Development Tasks

### Adding a New Command

1. Add command handler in `cli.py`:
   ```python
   def cmd_newcommand(args: argparse.Namespace) -> None:
       # Implementation
   ```

2. Add subparser in `create_parser()`:
   ```python
   new_parser = subparsers.add_parser("newcommand", help="...")
   new_parser.add_argument(...)
   new_parser.set_defaults(func=cmd_newcommand)
   ```

3. Add result model in `models.py` if needed

4. Update SKILL.md and reference docs

### Adding COM Functionality

Add to `outlook_com.py`, following existing patterns:
- Use `get_outlook_app()` for COM connection
- Handle exceptions appropriately
- Return typed data (dataclasses from `models.py`)

### Updating the Skill

After any changes to CLI or behavior:

```bash
# Update skill files in skills/outlook-automation/
# Then reinstall:
uv run python tools/install_skill.py --personal
```

## Testing

### Unit Tests

```bash
uv run python -m pytest tests/ -v
```

### Manual Testing

```bash
# Test doctor
uv run python -m outlookctl.cli doctor

# Test list
uv run python -m outlookctl.cli list --count 3

# Test search
uv run python -m outlookctl.cli search --unread-only --count 5

# Test draft (creates actual draft in Outlook)
uv run python -m outlookctl.cli draft --to "test@example.com" --subject "Test" --body-text "Test body"
```

### Skill Testing

Start a new Claude Code session and test prompts:
- "Summarize my latest 5 unread emails"
- "Find emails from GitHub"
- "Draft a reply to the most recent email"

## Troubleshooting

### "Access is denied" on outlookctl.exe

The globally installed executable has permission issues in Git Bash. Use:
```bash
uv run python -m outlookctl.cli <command>
```

### COM Connection Fails

1. Ensure Classic Outlook is running
2. Check it's not New Outlook (toggle at top-right)
3. Run `doctor` to diagnose

### Tests Fail to Import

```bash
uv sync  # Ensure dependencies are installed
```

## File Locations

| File | Purpose |
|------|---------|
| `pyproject.toml` | Project config, dependencies |
| `src/outlookctl/` | Main Python package |
| `skills/outlook-automation/SKILL.md` | Skill definition |
| `skills/outlook-automation/reference/` | Skill documentation |
| `tools/install_skill.py` | Skill installer |
| `tests/` | pytest tests |
| `evals/` | Skill evaluation scenarios |

## Dependencies

- **pywin32** - Windows COM automation
- **pytest** (dev) - Testing

## Version History

- **0.1.0** - Initial implementation
  - CLI with doctor, list, get, search, draft, send, attachments save
  - Claude Code Skill with safety guardrails
  - Draft-first workflow, audit logging
