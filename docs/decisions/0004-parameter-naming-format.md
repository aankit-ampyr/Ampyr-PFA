---
id: "0004"
title: How should we record parameter names — raw Excel label or derived snake_case?
status: pending
gates: [P1, schema_#15]
asked: 2026-04-27
decided_by: null
decided_on: null
decision: null
notes: null
options:
  - id: a
    label: Raw Excel labels as-is
    consequence: 'Preserves typos and inconsistencies exactly (e.g. "Land acqusition" stays misspelled). Easiest to trace back to Excel. Hard to reference in code (string with spaces, brackets, slashes).'
  - id: b
    label: Derived snake_case canonical form
    consequence: 'Code-friendly (e.g. contracted_volume_input). Diverges from Excel; we lose ability to grep Excel labels directly. Need a transformation rule documented + applied consistently.'
  - id: c
    label: Hybrid — store both as display_name + param_id
    consequence: 'param_id is canonical for code/joins; display_name preserves the Excel label verbatim. Slightly more storage. Best of both for traceability AND code ergonomics.'
---

## Where in Excel

- **Sheet:** Project Info
- **Parameter labels live in column C** (sometimes B for section headers). Spans
  rows 18-596 across the 21 named sections.
- **Sample labels** (real, from F1):
  - `Negotiation Start` (r22)
  - `Energy yield - P50` (r87)
  - `P50 - P90 conversion rate` (r90)
  - `Contracted Volume (input)` (r181)
  - `Contracted Volume (sensi)` (r182)
  - `Land acqusition amount` (r135)  ← typo present in Excel
  - `Land acquisition date [@RtB]` (r129)  ← bracketed qualifier
  - `Solar Power Prices - Sizing Case` (r278)
  - `PPA on / off` (r177)  ← contains slash + spaces

## What's ambiguous

Many parameter labels in Project Info are not code-friendly: spaces, slashes,
brackets, parentheses, the occasional typo. The same name often appears in
multiple sections (e.g. "PPA on / off" appears under each PPA contract block).
We need a stable identifier strategy before scanning across all 120 slots.

The choice affects:
- Whether engineers grep Excel labels in Python code (raw) or use mangled names
- Whether the apply_decisions / scenario engine can reference parameters cleanly
- Whether we can mechanically convert Excel labels → IDs without ambiguity

## Quick check Anchal can do

1. Look at `Project Info` rows 176-189 in F1 (the "PPA contract 1" block).
2. Notice that "PPA on / off" appears at row 177, AND a similar label appears
   in the "PPA contract 2" block (row 192) and contracts 3+. They're the *same
   parameter shape* but different *contract instances*.
3. Confirm with Anchal: when ASE/GTC analysts refer to these in conversation,
   do they say "PPA contract 1's on/off" (treating the contract index as a
   qualifier) or "PPA1 on/off" (treating it as a single name)? This affects
   whether `param_id` should encode the contract index or whether the schema
   needs a separate `contract_index` column.

## Why we're asking now

P1's deliverable is a **parameter inventory** — name + (eventually) value for
each (asset slot × parameter). The naming choice is the row-key for that
table, so it determines:
- Schema column type / length
- Whether the inventory is row-indexable by canonical name
- How brittle the ingestion is to Excel label changes between quarters

If we pick raw labels and ASE fixes the "acqusition" typo next quarter, our
keys break. If we pick derived names and the transformation rule changes,
historical data needs a migration. Hybrid (option C) sidesteps both at minor
storage cost.
