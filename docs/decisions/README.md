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
3. Submitting an answer writes back to the ADR file: sets `status: answered`,
   fills `decision`, `decided_on`, and (if non-default) `decided_by`.
4. The git commit of that file is the audit trail.

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
