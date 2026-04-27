"""Apply a decisions-JSON (exported from the Streamlit Cloud dashboard) to local ADR files.

Usage:
    python scripts/apply_decisions.py path/to/parthenon-decisions-YYYYMMDD-HHMM.json
    python scripts/apply_decisions.py path/to/file.json --dry-run
    python scripts/apply_decisions.py path/to/file.json --force   # overwrite existing answers

Workflow:
    1. SME (Anchal) answers decisions in the Streamlit Cloud dashboard.
    2. They click "Download answers JSON" — receives a file like
       parthenon-decisions-20260427-1430.json.
    3. They share it with Ankit.
    4. Ankit runs this script locally → it patches the matching ADR .md files.
    5. Ankit reviews `git diff docs/decisions/` and commits.

Refuses to overwrite an ADR that already has a different answer (use --force to override).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).parent.parent
DECISIONS_DIR = ROOT / "docs" / "decisions"
ADR_PATTERN = re.compile(r"^---\n(.*?)\n---\n(.*)$", re.DOTALL)


def find_adr_by_id(adr_id: str) -> Path | None:
    """Locate the ADR file whose frontmatter id matches."""
    for p in sorted(DECISIONS_DIR.glob("*.md")):
        if p.stem.upper() in ("README", "TEMPLATE"):
            continue
        text = p.read_text(encoding="utf-8")
        m = ADR_PATTERN.match(text)
        if not m:
            continue
        try:
            fm = yaml.safe_load(m.group(1)) or {}
        except yaml.YAMLError:
            continue
        if str(fm.get("id", "")) == str(adr_id):
            return p
    return None


def apply_answer(adr_path: Path, answer: dict, force: bool, dry_run: bool) -> tuple[bool, str]:
    """Patch one ADR with one answer. Returns (changed, message)."""
    text = adr_path.read_text(encoding="utf-8")
    m = ADR_PATTERN.match(text)
    if not m:
        return False, f"  malformed ADR (no frontmatter): {adr_path.name}"
    fm = yaml.safe_load(m.group(1)) or {}
    body = m.group(2)

    existing_decision = fm.get("decision")
    incoming_decision = answer.get("decision")

    if fm.get("status") == "answered" and existing_decision and existing_decision != incoming_decision:
        if not force:
            return False, (
                f"  CONFLICT: {adr_path.name} already answered with "
                f"'{existing_decision}'; incoming is '{incoming_decision}'. "
                f"Use --force to overwrite."
            )

    if fm.get("status") == "answered" and existing_decision == incoming_decision:
        return False, f"  unchanged: {adr_path.name} already has same answer"

    fm["status"] = "answered"
    fm["decision"] = incoming_decision
    fm["decided_by"] = answer.get("decided_by")
    fm["decided_on"] = answer.get("decided_on")
    if answer.get("notes"):
        fm["notes"] = answer["notes"]

    if dry_run:
        return True, f"  WOULD UPDATE: {adr_path.name} → decision='{incoming_decision}', by='{fm['decided_by']}'"

    yaml_text = yaml.safe_dump(fm, sort_keys=False, allow_unicode=True, width=1000)
    adr_path.write_text(f"---\n{yaml_text}---\n{body}", encoding="utf-8")
    return True, f"  updated: {adr_path.name} → decision='{incoming_decision}'"


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply decisions JSON to ADR files.")
    parser.add_argument("json_file", type=Path, help="Path to the decisions JSON exported from the dashboard.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would change without writing.")
    parser.add_argument("--force", action="store_true", help="Overwrite existing answers if they differ.")
    args = parser.parse_args()

    if not args.json_file.exists():
        print(f"ERROR: {args.json_file} not found", file=sys.stderr)
        return 1

    payload = json.loads(args.json_file.read_text(encoding="utf-8"))
    answers = payload.get("answers", [])
    if not answers:
        print("No answers in JSON.")
        return 0

    print(f"Applying {len(answers)} answer(s) from {args.json_file.name}"
          f"{' (DRY RUN)' if args.dry_run else ''}")
    if payload.get("session_exported_at"):
        print(f"  session exported: {payload['session_exported_at']}")
    print()

    n_changed = 0
    n_skipped = 0
    n_missing = 0
    for ans in answers:
        adr_id = str(ans.get("id", ""))
        title = ans.get("title", "(no title)")
        print(f"ADR {adr_id}: {title[:70]}")

        adr_path = find_adr_by_id(adr_id)
        if adr_path is None:
            print(f"  ERROR: no ADR file found with id={adr_id!r}")
            n_missing += 1
            continue

        changed, msg = apply_answer(adr_path, ans, force=args.force, dry_run=args.dry_run)
        print(msg)
        if changed:
            n_changed += 1
        else:
            n_skipped += 1

    print()
    print(f"Summary: {n_changed} updated, {n_skipped} skipped, {n_missing} missing")
    if not args.dry_run and n_changed:
        print()
        print("Next: review with `git diff docs/decisions/` and commit.")
    return 0 if n_missing == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
