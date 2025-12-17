#!/usr/bin/env python3
"""
Skill installer for outlook-email-automation.

This script installs the Skill to:
- Claude Code personal: ~/.claude/skills/outlook-email-automation/
- Claude Code project: .claude/skills/outlook-email-automation/
- OpenAI Codex: ~/.codex/skills/outlook-email-automation/

Usage:
    python tools/install_skill.py --personal    # Install for Claude Code (personal)
    python tools/install_skill.py --project     # Install for Claude Code (project)
    python tools/install_skill.py --codex       # Install for OpenAI Codex
    python tools/install_skill.py --uninstall --personal  # Remove personal install
    python tools/install_skill.py --verify --personal     # Verify installation
"""

import argparse
import os
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path


SKILL_NAME = "outlook-email-automation"


def get_skill_source_dir() -> Path:
    """Get the source skill directory."""
    # When run from project root
    script_dir = Path(__file__).parent.parent
    skill_dir = script_dir / "skills" / SKILL_NAME

    if skill_dir.exists():
        return skill_dir

    # Try relative to current working directory
    cwd_skill = Path.cwd() / "skills" / SKILL_NAME
    if cwd_skill.exists():
        return cwd_skill

    raise FileNotFoundError(
        f"Skill source directory not found. Expected at: {skill_dir}"
    )


def get_personal_skill_dir() -> Path:
    """Get the personal skill installation directory."""
    home = Path.home()
    return home / ".claude" / "skills" / SKILL_NAME


def get_project_skill_dir() -> Path:
    """Get the project skill installation directory."""
    return Path.cwd() / ".claude" / "skills" / SKILL_NAME


def get_codex_skill_dir() -> Path:
    """Get the OpenAI Codex skill installation directory."""
    home = Path.home()
    return home / ".codex" / "skills" / SKILL_NAME


def backup_existing(dest_dir: Path) -> Path | None:
    """Create a backup of existing installation."""
    if not dest_dir.exists():
        return None

    backup_name = f"{SKILL_NAME}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    backup_path = dest_dir.parent / backup_name

    print(f"Creating backup: {backup_path}")

    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in dest_dir.rglob("*"):
            if file_path.is_file():
                arc_name = file_path.relative_to(dest_dir)
                zf.write(file_path, arc_name)

    return backup_path


def copy_skill(source_dir: Path, dest_dir: Path) -> None:
    """Copy skill directory to destination."""
    if dest_dir.exists():
        # Remove existing files but keep directory
        shutil.rmtree(dest_dir)

    # Create parent directories
    dest_dir.parent.mkdir(parents=True, exist_ok=True)

    # Copy entire skill directory
    shutil.copytree(source_dir, dest_dir)


def verify_installation(dest_dir: Path) -> bool:
    """Verify that the skill is properly installed."""
    skill_md = dest_dir / "SKILL.md"

    if not skill_md.exists():
        print(f"ERROR: SKILL.md not found at {skill_md}")
        return False

    # Check for YAML frontmatter
    content = skill_md.read_text(encoding="utf-8")
    if not content.startswith("---"):
        print("ERROR: SKILL.md missing YAML frontmatter")
        return False

    # Check for required fields
    frontmatter_end = content.find("---", 3)
    if frontmatter_end == -1:
        print("ERROR: SKILL.md has malformed frontmatter")
        return False

    frontmatter = content[3:frontmatter_end]
    if "name:" not in frontmatter:
        print("ERROR: SKILL.md missing 'name' field")
        return False
    if "description:" not in frontmatter:
        print("ERROR: SKILL.md missing 'description' field")
        return False

    # Check reference docs exist
    reference_dir = dest_dir / "reference"
    required_refs = ["cli.md", "security.md", "troubleshooting.md", "json-schema.md"]

    for ref in required_refs:
        if not (reference_dir / ref).exists():
            print(f"WARNING: Missing reference doc: {ref}")

    print(f"SUCCESS: Skill verified at {dest_dir}")
    return True


def uninstall_skill(dest_dir: Path) -> bool:
    """Remove installed skill."""
    if not dest_dir.exists():
        print(f"Skill not installed at {dest_dir}")
        return True

    print(f"Removing: {dest_dir}")
    shutil.rmtree(dest_dir)

    # Clean up empty parent directories
    parent = dest_dir.parent
    try:
        if parent.exists() and not any(parent.iterdir()):
            parent.rmdir()
            grandparent = parent.parent
            if grandparent.exists() and not any(grandparent.iterdir()):
                grandparent.rmdir()
    except OSError:
        pass  # Directory not empty or permission issue

    print("Uninstall complete")
    return True


def install(
    target: str,
    backup: bool = False,
    verify: bool = True,
) -> bool:
    """
    Install the skill to the specified target.

    Args:
        target: "personal" or "project"
        backup: Create backup before overwriting
        verify: Verify installation after copying

    Returns:
        True if successful
    """
    try:
        source_dir = get_skill_source_dir()
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        return False

    if target == "personal":
        dest_dir = get_personal_skill_dir()
    elif target == "project":
        dest_dir = get_project_skill_dir()
    elif target == "codex":
        dest_dir = get_codex_skill_dir()
    else:
        print(f"ERROR: Invalid target: {target}")
        return False

    print(f"Source: {source_dir}")
    print(f"Destination: {dest_dir}")

    # Backup if requested and destination exists
    if backup and dest_dir.exists():
        backup_path = backup_existing(dest_dir)
        if backup_path:
            print(f"Backup created: {backup_path}")

    # Copy skill
    print("Copying skill...")
    copy_skill(source_dir, dest_dir)
    print(f"Skill installed to: {dest_dir}")

    # Verify
    if verify:
        return verify_installation(dest_dir)

    return True


def main():
    parser = argparse.ArgumentParser(
        description="Install outlook-email-automation skill for Claude Code or OpenAI Codex"
    )

    target_group = parser.add_mutually_exclusive_group(required=True)
    target_group.add_argument(
        "--personal",
        action="store_true",
        help="Install to ~/.claude/skills/ (Claude Code personal)",
    )
    target_group.add_argument(
        "--project",
        action="store_true",
        help="Install to .claude/skills/ (Claude Code project)",
    )
    target_group.add_argument(
        "--codex",
        action="store_true",
        help="Install to ~/.codex/skills/ (OpenAI Codex)",
    )

    action_group = parser.add_mutually_exclusive_group()
    action_group.add_argument(
        "--uninstall",
        action="store_true",
        help="Remove the installed skill",
    )
    action_group.add_argument(
        "--verify",
        action="store_true",
        help="Only verify existing installation",
    )

    parser.add_argument(
        "--backup",
        action="store_true",
        help="Create backup before overwriting",
    )

    parser.add_argument(
        "--no-verify",
        action="store_true",
        help="Skip verification after install",
    )

    args = parser.parse_args()

    if args.personal:
        target = "personal"
    elif args.codex:
        target = "codex"
    else:
        target = "project"

    def get_dest_dir():
        if args.personal:
            return get_personal_skill_dir()
        elif args.codex:
            return get_codex_skill_dir()
        else:
            return get_project_skill_dir()

    if args.uninstall:
        dest_dir = get_dest_dir()
        success = uninstall_skill(dest_dir)
    elif args.verify:
        dest_dir = get_dest_dir()
        success = verify_installation(dest_dir)
    else:
        success = install(
            target=target,
            backup=args.backup,
            verify=not args.no_verify,
        )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
