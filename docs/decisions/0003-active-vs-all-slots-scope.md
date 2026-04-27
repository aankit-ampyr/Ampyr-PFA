---
id: "0003"
title: Asset-list scope — all 120 slots, or only the 78 actives?
status: pending
gates: [P2, engine_#14]
asked: 2026-04-27
decided_by: null
decided_on: null
decision: null
notes: null
options:
  - id: a
    label: All 120 slots, with an `is_active` boolean column
    consequence: Captures the full template space. Placeholders (42 of them) become real assets in future quarters - keeping their slots avoids re-numbering. Larger working set.
  - id: b
    label: Only the 78 active assets, no placeholder slots
    consequence: Smaller and cleaner. But when ASE activates a placeholder next quarter, we'd need a separate flow to add it. Risks coupling slot_index to row position in Excel.
  - id: c
    label: All slots, but a `slot_status` enum (active / placeholder / retired)
    consequence: Most expressive. Captures the activation lifecycle. Slightly more code.
---

## Where in Excel

- **Sheet:** Project Info, **row 7** (`projectactiveflag` named range, cols G:DR).
- **Slot count:** 120 columns. Of these, **78 are active in F1** (per the deep
  re-analysis), 42 are placeholders.
- **Quarterly churn pattern:** F3 → F1 saw 6 slots flip False→True (placeholders
  becoming real assets) and 2 flip True→False (retired). So the placeholder pool
  is a real, used pipeline, not dead space.

## What's ambiguous

The model has 120 template slots. Some are active assets, some are reserved for
future projects. The question is whether our `assets` table mirrors the full
120 (preserving the slot structure) or only stores the 78 currently active.

## Quick check Anchal can do

1. Look at any "placeholder" slot in F1 — pick a column where row 7 is False
   (e.g. ask for one to inspect).
2. Check rows 18-596 (parameter rows) for that slot — are they:
   - **Empty / zeroed** → placeholder is purely a reservation. Option B is fine.
   - **Populated with provisional data** (capacity, country, COD) → placeholders
     carry meaning. Option A or C preserves it.
3. Confirm whether `slot_index` is referenced anywhere downstream (e.g. in
   `Time Inputs (Q)` cluster financing rows that may hard-code slot positions).
   If yes → keep all 120 to preserve the index.

## Why we're asking now

Decides the row count of the `assets` table (and downstream `asset_parameters`).
Wrong answer creates schema churn when ASE activates a placeholder next quarter.
Closely tied to ADR 0001 — once we know how "active" is tested, the count is
deterministic.
