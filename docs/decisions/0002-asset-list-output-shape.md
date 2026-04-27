---
id: '0002'
title: Output shape for the 78-asset list (P2 deliverable)
status: answered
gates:
- P2
asked: 2026-04-27
decided_by: Anchal Gupta
decided_on: '2026-04-27'
decision: b
notes: JSON. Easy to extend with the richer schema implied by ADR 0003 (status enum, slot_index that can extend past 120). Alembic seed will consume the JSON later when schema work begins.
options:
- id: a
  label: Flat CSV (asset_name, country, technology, slot_index, is_active)
  consequence: Fastest to produce. Easy to eyeball. No use beyond P2 — schema work has to re-process.
- id: b
  label: JSON (list of objects, same fields)
  consequence: Slightly slower. Easier to extend with nested fields later (e.g. capacities, COD).
- id: c
  label: Alembic seed migration that creates an `assets` table and inserts rows
  consequence: 'Most useful downstream — schema work (Bucket 2 #15) consumes it directly. Highest commitment now (locks initial schema columns).'
---

## Where in Excel

- **Sheet:** Project Info, **rows 5-7** (asset name / technology / active flag), columns G:DR (120 slots horizontal).
- **Asset registry derived from:** F1 (`Converted 2026-01-08 ... v18 ... .xlsm`).

The data being extracted is just three rows × 120 columns. The question is what
**format** the extracted data should land in.

## What's ambiguous

Three reasonable formats exist. Each commits us to more or less downstream work
when schema design (Bucket 2 #15) starts.

## Quick check Anchal can do

This is more of a **process question** than an Excel question — Anchal's input
is helpful but not strictly required. The main consideration:

> Will the asset list be hand-edited between now and when Bucket 2 #15 starts?

- If **yes** (likely, as ASE updates assets every quarter) → option A or B is safer,
  because seed migrations are awkward to re-run.
- If **no** → option C saves a step.

## Why we're asking now

P2's deliverable shape is a 30-minute decision but determines whether Bucket 2 #15
(schema) starts from raw data or has already-typed rows. It's the cheap lever
that decides how P2's output composes with downstream work.
