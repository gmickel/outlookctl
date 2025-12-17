## PRD: Outlook Classic Local Automation Skill for Claude Code

### 0. Key fixes vs the previous draft

* **Dependency management**: Use `uv` project workflows (`uv add`, `uv sync`, `uv run`) rather than Poetry or `uv pip`. `uv add` updates `pyproject.toml` and the lockfile, and `uv run` keeps the environment locked + synced automatically.
* **Skill install**: Install/copy the **entire skill folder** (not just a single markdown file). Claude Code Skills are discovered from `~/.claude/skills/` and `.claude/skills/`.
* **Correct Skill file naming**: The required entrypoint is `SKILL.md` with YAML frontmatter (`name`, `description`, optional `allowed-tools`).
* **Outlook compatibility**: The automation path relies on Outlook’s **COM** object model → requires **Classic Outlook**. New Outlook for Windows does not support COM add-ins (strong signal it won’t expose the same COM surface for automation).
* **Safer sending**: Add explicit guardrails: default to **draft**, require explicit confirmation for **send**, and provide a two-step workflow.

---

## 1. Product summary

Build a **Claude Code Skill** plus a small **local CLI bridge** that can **read**, **search**, **draft**, and **send** Outlook email from a Windows Devbox **using the currently authenticated Classic Outlook session** (COM automation via `pywin32`). No Graph API, no OAuth, no external app registration.

The system is designed to:

* Work under strict corporate access policies by staying within the local user session.
* Be safe-by-default (draft-first, explicit send confirmation).
* Be operationally simple (single repo; `uv` for Python deps; optional global installer).

---

## 2. Background and problem statement

### Problem

Corporate restrictions prevent using Microsoft Graph / OAuth app registrations. The user can access Outlook only via approved endpoints (corp laptop/iPhone/cloud Devbox). Automation is desired from the Devbox while Outlook is running.

### Opportunity

Classic Outlook exposes a COM automation model accessible locally via `pywin32` in the same Windows session. This enables local-only automation while respecting “authorized boundary” constraints.

---

## 3. Goals

### Primary goals

1. **Read emails**: List recent messages from selected folders (Inbox by default) with configurable metadata and optional body snippet.
2. **Search emails**: Filter by sender/subject/date range/unread/has attachments, returning stable identifiers.
3. **Draft emails**: Create drafts (To/CC/BCC/Subject/Body/Attachments) without sending.
4. **Send emails**: Send only with explicit confirmation/flags; log what was sent (metadata only by default).
5. **Claude Code Skill**: Provide a robust Skill that guides Claude to use the CLI bridge correctly with deterministic outputs, minimal ambiguity, and safe workflows.
6. **Installation tooling**: Provide a simple installer that can:

   * Install the skill **personally** (`~/.claude/skills/...`)
   * Install into the **current repo** (`.claude/skills/...`)
   * Optionally install the CLI globally as a `uv` tool.

### Secondary goals

* Support saving attachments to disk.
* Provide a “doctor” command to validate prerequisites and troubleshoot quickly.
* Provide an evaluation harness (prompt scenarios + expected behavior) to validate Skill usefulness.

---

## 4. Non-goals

* Microsoft Graph API integration.
* Cross-platform support (macOS/Linux native). (WSL may call Windows executables; treat as optional.)
* Full fidelity mailbox operations (rules, categories, archiving policies) unless explicitly added later.
* Bypassing Outlook “programmatic access” security prompts via unsupported hacks.

---

## 5. Users and use cases

### Personas

* **Primary**: Technical operator on a Windows Devbox (PowerShell), with Claude Code, `uv`, and Classic Outlook available.
* **Secondary**: Team members using the same repo as a project Skill checked into git.

### Core use cases

1. “Summarize my latest 10 unread emails from Inbox.”
2. “Find all emails from `<sender>` in the last 7 days and extract action items.”
3. “Draft an email to `<recipient>` responding to `<message-id>` and attach `<file>`.”
4. “Send the draft only after I approve.”

---

## 6. Constraints and assumptions

### Hard constraints

* Must run on **Windows** with **Classic Outlook** available.
* Must not require Azure app registration/admin consent.
* Must use **`uv`** for dependency management (no Poetry; avoid `uv pip` flows where possible).

### Assumptions

* User has permission to run local scripts on the Devbox.
* Outlook mailbox is already configured and logged in.
* Corporate policy permits use of Claude Code for the intended data classification (solution includes data minimization controls regardless).

---

## 7. Proposed solution architecture

### Components

1. **Python CLI bridge (`outlookctl`)**

   * A local command-line tool that executes Outlook COM operations.
   * Produces deterministic **JSON** for Claude to parse.
   * Contains safety gates for sending.

2. **Claude Code Skill (`outlook-automation`)**

   * `SKILL.md` + minimal supporting docs + scripts
   * Provides workflows and “guardrail patterns” for Claude (draft-first; send-confirm).

3. **Installer (`outlook-skill-installer`)**

   * A small Python entrypoint (or script) that copies skill folders into:

     * `~/.claude/skills/<skill>/...` (personal)
     * `.claude/skills/<skill>/...` (project)
   * Optional: install the CLI as a persistent uv tool.

### Why this design

* **Thin waist**: Claude interacts with a single stable interface (CLI + JSON).
* **Progressive disclosure**: Skill stays concise; deeper docs are in separate files, consistent with Skill best practices.
* **Safety**: Sending is a separate explicit action with confirmations and logging.

---

## 8. Functional requirements

### 8.1 CLI: `outlookctl` commands

The CLI must be runnable via `uv run outlookctl ...` (project mode) and optionally via a globally installed tool.

#### 8.1.1 `outlookctl doctor`

Validates environment:

* OS is Windows
* Python can import `win32com`
* Outlook COM object can be created (`Dispatch("Outlook.Application")`)
* Detect whether Outlook is Classic-capable (best-effort)
* Optionally locate `OUTLOOK.EXE` common paths and report findings

**Output**: JSON report with pass/fail checks and recommended remediations.

#### 8.1.2 `outlookctl list`

List messages from a folder.

**Options**

* `--folder inbox|sent|drafts|by-name:<name>|by-path:<path>`
* `--count N` default 10
* `--unread-only`
* `--since ISO_DATE` / `--until ISO_DATE`
* `--include-body-snippet` (off by default)
* `--body-snippet-chars N` default 200
* `--output json` (default) or `text`

**Output JSON schema (v1)**

```json
{
  "version": "1.0",
  "folder": {"name": "Inbox"},
  "items": [
    {
      "id": {"entry_id": "...", "store_id": "..."},
      "received_at": "2025-12-16T09:12:33",
      "subject": "...",
      "from": {"name": "...", "email": "..."},
      "to": ["..."],
      "cc": ["..."],
      "unread": true,
      "has_attachments": false,
      "body_snippet": "..." 
    }
  ]
}
```

#### 8.1.3 `outlookctl get --id <entry_id> --store <store_id>`

Fetch a single message with controlled fields.

**Options**

* `--include-body` (off by default)
* `--include-headers` (optional)
* `--max-body-chars N`

#### 8.1.4 `outlookctl search`

Search by:

* `--query "<free text>"` (subject/body best-effort)
* `--from "<email or name>"`
* `--subject-contains "<text>"`
* `--unread-only`
* date range

Return list results with stable ids.

#### 8.1.5 `outlookctl draft`

Create a draft message.

**Inputs**

* `--to`, `--cc`, `--bcc` (comma-separated)
* `--subject`
* `--body-text` or `--body-html` (mutually exclusive)
* `--attach <path>` (repeatable)
* `--reply-to-id <entry_id> --reply-to-store <store_id>` (optional, for reply/forward flows)
* `--save-to drafts` (default)

**Output**

* Draft identifiers (entry_id/store_id) + where it was saved.

#### 8.1.6 `outlookctl send`

Send an existing draft (preferred) or send a new message with explicit confirmation.

**Safety requirements**

* Default behavior: **refuse** unless one of:

  * `--confirm-send "YES"` (exact string)
  * or `--confirm-send-file <path>` containing `YES`
* If sending “new” directly, CLI should recommend “draft then send” and require a stronger flag:

  * `--unsafe-send-new --confirm-send "YES"`

**Logging requirements**

* Write a local audit log entry (metadata only) to:

  * `%LOCALAPPDATA%/outlookctl/audit.log` (Windows)
* Must never log full bodies unless `--log-body` is explicitly set.

#### 8.1.7 `outlookctl attachments save`

Given a message id:

* `--dest <dir>`
* Save attachments with safe filenames
* Output list of saved paths

---

### 8.2 Outlook instance management

The CLI must:

* Attempt `Dispatch("Outlook.Application")` first.
* If that fails, optionally attempt to start Classic Outlook (best effort):

  * Try common paths and report which was used.
  * Wait/retry COM attach for a limited time.
* If “New Outlook” is active and Classic is unavailable, return a clear diagnostic:

  * “Classic Outlook COM automation unavailable. Switch to Classic Outlook.”

Rationale: New Outlook does not support COM add-ins; treat it as non-automatable via COM.

---

### 8.3 Data minimization and redaction

Because email content may be sensitive:

* Default outputs should be **headers + minimal metadata**.
* Body access must be opt-in (`--include-body*` flags).
* Provide optional `--redact` mode:

  * Remove email addresses (hash or replace with tokens)
  * Truncate subjects
  * Remove bodies entirely

---

## 9. Claude Code Skill requirements

### 9.1 Skill placement and discovery

* Support as:

  * **Personal Skill** under `~/.claude/skills/<skill-name>/SKILL.md`
  * **Project Skill** under `.claude/skills/<skill-name>/SKILL.md`

### 9.2 `SKILL.md` frontmatter rules

* Must include `name` and `description`.
* Description must be **third person** and include both “what it does” and “when to use it”.
* Optional: `allowed-tools` to restrict what Claude can do while the Skill is active.

### 9.3 Skill scope and triggers

**Skill name (example)**: `outlook-automation`

**Description (example)**

* Must include triggers such as: “Outlook”, “Inbox”, “email”, “draft”, “send”, “reply”, “attachments”, “Devbox”, “COM”, “pywin32”.
* Example phrasing (third person):

  * “Automates reading, searching, drafting, and sending emails in Classic Outlook on Windows using local COM automation. Use when the user asks to process Outlook emails, create drafts, or send messages from the authenticated Outlook session.”

### 9.4 Progressive disclosure file structure

Keep `SKILL.md` concise and link to one-level-deep references.

Recommended skill folder (both personal and project installs):

```
outlook-automation/
  SKILL.md
  reference/
    cli.md
    security.md
    troubleshooting.md
    json-schema.md
  scripts/
    outlookctl.py
```

### 9.5 Skill workflows (must be included)

The Skill must instruct Claude to follow these patterns:

**Workflow A: Read/search**

1. Run `outlookctl list` or `outlookctl search` with minimal fields.
2. Summarize based on JSON results.
3. Only fetch bodies (`outlookctl get --include-body`) if the user explicitly asks.

**Workflow B: Draft-first response**

1. Identify target message (search/list → pick id).
2. Generate draft via `outlookctl draft ... --reply-to-id ...`.
3. Show the user a preview (subject/body/recipients).
4. Only send after explicit user instruction; then use `outlookctl send ... --confirm-send YES`.

**Workflow C: Attachments**

1. Save attachments via `outlookctl attachments save`.
2. Reference saved paths when drafting/sending.

### 9.6 Skill safety rules (must be explicit)

* Never auto-send emails.
* Always require explicit user instruction + CLI confirmation flag.
* Avoid leaking full content; default to metadata-only.

---

## 10. Installer requirements

### 10.1 Installer behavior

Provide a Python executable entrypoint (stdlib-only preferred) that can:

**Install personal skill**

* Copy `outlook-automation/` → `~/.claude/skills/outlook-automation/`

**Install project skill**

* Copy `outlook-automation/` → `<cwd>/.claude/skills/outlook-automation/`

**Idempotency**

* If destination exists:

  * Default: update in place (overwrite files)
  * Option: `--backup` to zip existing destination first

**Uninstall**

* Remove the installed skill folder from selected target

**Verify**

* Confirm that `SKILL.md` exists at destination and passes basic YAML fence checks.

Claude Code documents the personal/project skill directories and discovery model; the installer must adhere to these exact locations.

### 10.2 Optional: install CLI globally via uv tools

Support an optional command:

* `uv tool install .` to install the package’s console script(s) into a persistent tool environment (document this as an option, not required).

---

## 11. Python project packaging and dependency management (uv-first)

### 11.1 `pyproject.toml` format

Use PEP 621 (`[project]`) rather than Poetry tables.

Minimum requirements:

* Python `>=3.12`
* Dependency: `pywin32`

Add a build system so `uv sync` installs the project in editable mode when appropriate; uv notes that if no build system is defined, the project won’t be installed.

### 11.2 uv workflows

Document these flows:

**Bootstrap**

* `uv sync` (create/sync project environment)
* `uv run outlookctl doctor`

**Add dependencies**

* `uv add pywin32` (writes to `pyproject.toml` + updates lock/environment)

**Run**

* `uv run outlookctl list --count 10`

---

## 12. Non-functional requirements

### 12.1 Security and compliance

* No network calls required for mailbox access.
* No caching of email bodies by default.
* Audit logging for send operations (metadata-only).
* Provide explicit “data handling” documentation (`reference/security.md`).

### 12.2 Reliability

* Deterministic JSON output with schema versioning.
* Robust error messages with actionable remediation.

### 12.3 Usability

* One command to verify environment (`doctor`).
* Clear examples in Skill and README.
* Minimal required flags for common operations.

---

## 13. Error handling and troubleshooting requirements

Must detect and clearly report:

* Outlook COM attach failures
* Outlook not running
* Classic vs New Outlook mismatch (“COM automation unavailable” guidance)
* “Programmatic access” prompt risk when accessing body/email addresses
* Missing attachments / invalid paths
* Unicode/encoding issues when extracting body

Provide a dedicated troubleshooting doc referenced from `SKILL.md`.

---

## 14. Testing and evaluation

### 14.1 Unit tests (CI-safe)

* JSON formatting and schema validation
* Argument parsing and safety gates
* File copying installer logic (temp directories)

### 14.2 Manual integration tests (requires Outlook)

* `doctor` passes on the Devbox
* list/search works on Inbox
* draft created in Drafts folder
* send works only with confirm flags
* attachments saved correctly

### 14.3 Skill evaluations (prompt-based)

Create `evals/` with 3+ scenarios:

* Summarize unread emails
* Draft reply to a message
* Attempt to send without confirmation (expected: refuse)

---

## 15. Acceptance criteria

1. `uv run outlookctl doctor` returns all green checks on a properly configured Devbox.
2. `outlookctl list` returns valid JSON and includes stable identifiers.
3. `outlookctl draft` creates a visible draft in Outlook Drafts.
4. `outlookctl send` refuses without explicit confirmation; sends successfully with confirmation.
5. Installer correctly installs as both personal and project skill locations.
6. `SKILL.md` follows required structure, and description is third person + includes usage triggers.

---

## 16. Recommended repo layout

```
outlook-claude-skill/
  pyproject.toml
  uv.lock
  README.md

  src/outlookctl/               # Python package
    __init__.py
    cli.py                      # argparse entrypoint
    outlook_com.py              # COM wrapper + helpers
    models.py                   # dataclasses for JSON
    audit.py                    # audit log writer
    safety.py                   # send confirmation gates

  skills/
    outlook-automation/
      SKILL.md
      reference/
        cli.md
        security.md
        troubleshooting.md
        json-schema.md
      scripts/
        # (optional) thin wrappers, or omit if using uv run + console scripts

  tools/
    install_skill.py            # installer entrypoint
    uninstall_skill.py          # optional
```

---

## 17. Implementation notes Claude Code should follow (explicit instructions for the builder)

* Prioritize **deterministic CLI JSON** over Claude parsing console prose.
* Implement **draft-first** as the default high-level behavior.
* Keep `SKILL.md` short; push details into `reference/*.md`.
* Add `allowed-tools` only after confirming the exact tool names available in Claude Code; restrict as tightly as practical.
* Treat New Outlook as non-automatable; ensure docs and errors point users to Classic Outlook.

---

If you want, I can also draft the **actual `SKILL.md` + reference docs outlines** (not the full code) so Claude Code can immediately start implementing against a concrete Skill contract.
