---
id: '0005'
title: What metadata to capture per parameter, beyond name + value?
status: answered
gates:
- P1
- 'schema_#15'
asked: 2026-04-27
decided_by: Anchal Gupta
decided_on: '2026-04-27'
decision: c
notes: 'Full metadata: section, row, source_cell, sheet, units, scale, section_path, data_type, is_input_vs_derived. Highest-cost choice up front but pre-empts the need for backfill migrations later (asset_parameters is the largest table — getting columns right now is cheap). Enables: Excel cell traceability for SME validation calls, structured quarterly diffs (row moves vs value changes), unit checking, and input-vs-derived discrimination for the differ.'
options:
- id: a
  label: Minimal — (asset_id, param_id, value)
  consequence: Smallest schema. Fastest ingestion. Cannot trace a value back to its Excel cell when validating. No units capture (so the engine has to encode unit assumptions). Brittle when Excel structure shifts.
- id: b
  label: Traceable — adds (section, row, source_cell, sheet)
  consequence: 'Recommended baseline. Lets us answer "where did this value come from?" — critical for SME validation calls and quarterly diff reports. Roughly doubles the row width but each field is small (int row, string source_cell like "G87"). Schema #15 will need these for the differ.'
- id: c
  label: Full — adds (units, scale, section_path, data_type, is_input_vs_derived)
  consequence: Maximally expressive. Captures whether a cell is a raw input (analyst types it) or a derived formula (Excel calculates it from other inputs); units (MWh / GBP / %); scale factor (1 vs 1000 vs 1e6 — many financial models mix these). Heavy ingestion overhead. Useful for auto-generated reports and unit-checking, but most fields are not strictly needed for engine correctness.
---

## Where in Excel

- **Sheet:** Project Info, columns C-F across rows 18-596 (the 120-slot params).
- **Sample structure** (rows 87-91, "Energy yield" block):

```text
Row 87: C='Energy yield - P50',   E='P50',  F=(value per asset across G:DX)
Row 88: C='P50 - P75 conversion rate'                F=(value per asset)
Row 89: C='Energy yield - P75',   E='P75',  F=(value per asset)
Row 90: C='P50 - P90 conversion rate'                F=(value per asset)
Row 91: C='Energy yield - P90',   E='P90',  F=(value per asset)
```

Notice column E sometimes carries a **selector / unit** (P50, P90) and column
F sometimes carries a **default** (e.g. `Aurora` curve selector at r268-273).
Some rows (r274) have a `False` flag in column E — meta about the parameter,
not the asset value.

## What's ambiguous

The minimum to make the engine work is (asset_id, param_id, value). But several
real downstream needs argue for capturing more:

1. **Validation calls with Anchal/ASE finance:** "what's slot 7's contracted
   period?" — answerable only if we know `(source_cell, sheet, row)` to link
   back to the Excel cell.
2. **Quarterly differ:** detecting that ASE moved a row from r129 to r131 needs
   the row number stored.
3. **Unit ambiguity:** PPA prices appear in multiple currencies (£/MWh, €/MWh)
   and scales (raw, ×1000); without unit capture the engine has to hard-code
   per-section assumptions.
4. **Input vs derived:** some rows are raw analyst inputs, others are formulas
   that depend on other inputs. The differ should treat these differently
   (input change = real change; derived change = caused by an input change).

The trade-off is ingestion cost vs schema simplicity vs downstream flexibility.

## Quick check Anchal can do

1. In F1's Project Info, find a parameter whose value Anchal recently
   changed for an assessment (e.g. an Opex bump). Ask:
   - Was the change made directly in column F of the input row, or via a
     scenario lever in Sensis?
   - Does Anchal ever need to know the **history** of values at that cell
     (e.g. "what was it last quarter")?
   - Do Anchal and Daniel ever disagree about units / scale of a value?
2. If the answer to all three is "no" → option **A** is fine.
   If "yes" to (1)/(3) → option **B** at minimum.
   If "yes" to all three → option **C**.

## Why we're asking now

The schema column count for `asset_parameters` is the most expensive table
to migrate later (currently sized at 120 slots × 590+ params = ~70k rows
with versioning blowing it up further). Adding columns later means a
backfill migration; getting it right up front is cheap.

The choice also gates how the **quarterly differ** (Bucket 2 #17) is
written: option A means it can only diff `value` changes, options B/C
let it diff structure (row moves, unit changes).
