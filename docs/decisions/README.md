# Decisions (ADRs)

This directory holds **Architecture / Implementation Decision Records (ADRs)** for
Project Parthenon — every choice that needs SME input or that materially shapes
the deliverable.

## How they're used

1. When a question comes up that needs Anchal Gupta (or another SME) to answer,
   it becomes a numbered ADR file here.
2. The dev dashboard (`devtools/workflow_explorer.py`, **🤔 Decisions** tab) lists
   all pending ADRs and lets the SME answer them in-app via radio buttons + an
   optional notes box.
3. On submit, the dashboard does **two** things:
   - **Always:** appends the answer to a session-state list (offered as a
     downloadable JSON via the "Download answers JSON" button at the top).
   - **Best-effort:** writes back to the ADR `.md` file directly. Works when the
     filesystem is writable (running locally). Will gracefully fail on Streamlit
     Cloud — the JSON download is the canonical path there.
4. The git commit of the updated `.md` file (or of the apply-script output) is
   the audit trail.

### Workflow A — Local (Anchal at Ankit's desk)

1. Ankit runs `streamlit run devtools/workflow_explorer.py`.
2. Anchal answers decisions in the **🤔 Decisions** tab.
3. Each submit updates the corresponding `.md` file directly.
4. Ankit reviews `git diff docs/decisions/`, commits, pushes.

### Workflow B — Cloud (Anchal answers remotely on Streamlit Cloud)

1. Anchal opens the cloud-deployed dashboard.
2. Answers one or more decisions; each submit accumulates in the session.
3. Clicks **Download answers JSON** at the top of the Decisions tab.
4. Sends Ankit the file (Slack / email / shared drive — whatever is fast).
5. Ankit applies it locally:
   ```bash
   python scripts/apply_decisions.py path/to/parthenon-decisions-YYYYMMDD-HHMM.json
   # or, to preview without writing:
   python scripts/apply_decisions.py path/to/file.json --dry-run
   # or, to overwrite an existing answer that differs:
   python scripts/apply_decisions.py path/to/file.json --force
   ```
6. Ankit reviews `git diff docs/decisions/`, commits, pushes.

The apply script refuses to overwrite an ADR that has a different existing
answer unless `--force` is passed — so re-applying a JSON is safe.

## File naming

`NNNN-short-kebab-title.md` — 4-digit zero-padded counter, hyphenated title.
Examples: `0001-active-flag-canonical-test.md`, `0007-revenue-curve-source.md`.

## Frontmatter schema

```yaml
id: "0001"             # 4-digit string
title: ...             # one-line question
status: pending        # pending | answered | superseded
gates: [P2, ...]       # which work items this blocks
asked: 2026-04-27      # ISO date
decided_by: null       # auto-fills to "Anchal Gupta" on answer if null
decided_on: null       # ISO date, set on answer
decision: null         # the chosen option id ('a', 'b', ...) or free text
notes: null            # optional free-text reasoning
options:               # 2+ options the SME picks among
  - id: a
    label: ...
    consequence: ...
```

## Body sections (markdown, after the frontmatter)

- `## Where in Excel` — sheet, cell/named range, formula or VBA snippet
- `## What's ambiguous` — plain English, 1-2 sentences
- `## Quick check Anchal can do` — concrete look-at-Excel verification

## When NOT to write an ADR

- Trivial implementation choices (file format, variable names) — pick and go.
- Decisions whose answer is "obvious to anyone who knows the code" — just code it.
- Decisions where the SME has no context (purely engineering) — don't waste their time.

ADRs are for choices where **(a)** the SME's answer materially changes the
deliverable AND **(b)** the choice will be re-examined later (or by someone else).

## See also

- [TEMPLATE.md](TEMPLATE.md) — copy this when drafting a new ADR
- [CLAUDE.md](../../CLAUDE.md) §"Behavioral Guidelines · Think Before Coding"
