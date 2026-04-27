---
id: '0001'
title: Which "active flag" test is canonical?
status: answered
gates:
- P2
- engine_#14
asked: 2026-04-27
decided_by: Anchal Gupta
decided_on: '2026-04-27'
decision: b
notes: 'Use boolean compare at both code paths in the Python port (the equivalent of option b uniformly). EMPIRICAL CHECK: F1 row 7 cells are 118 of 120 actual `bool` (probe in devtools/probe_active_flags.py); zero cells contain text ''True''. In VBA, loose typing means ''= "True"'' and ''= True'' both return TRUE for a boolean cell, so the Excel model produces identical output via either path — there''s no actual behavioural deviation. Python should use idiomatic `if flag:` (boolean truthiness). NO `DEVIATES_FROM_EXCEL` annotation needed.'
options:
- id: a
  label: String compare ('= "True"')
  consequence: Matches PlatformConsolidation (the production consolidation macro). Current production output behaviour. Risk if cell ever holds a real boolean.
- id: b
  label: Boolean compare ('= True')
  consequence: Matches Sens_platformconsol (sensitivity macro). Differs from main consolidation if cells hold text.
- id: c
  label: Replicate both, log mismatches
  consequence: Most defensive. Surfaces the bug rather than hiding it. More code in the Python port.
---

## Where in Excel

- **Sheet:** Project Info
- **Named range:** `projectactiveflag` — row 7, columns G:DR (one cell per asset slot 1-120)
- **VBA snippets that expose the choice:**

```vba
' Consolidation.bas (PlatformConsolidation - the production macro):
If projectactiveflag(i) = "True" Then         ' STRING compare against text "True"
    Range("Project_View") = projectList(i)
    Debt_sizing
    ' ... paste cashflows
End If

' Module1.bas (Sens_platformconsol - the sensitivity macro):
If projectactiveflag(i) = True Then            ' BOOLEAN compare against TRUE
    Range("Project_View") = projectList(i)
    ' ...
End If
```

## What's ambiguous

The two macros test the same flag differently. If the cell holds the *text*
`"True"`, the first test passes but the second fails. If the cell holds the
*boolean* `TRUE`, the first fails and the second passes. The two macros could
therefore see different "active" sets — and the Python port has to pick which
behaviour to replicate (or replicate both).

## Quick check Anchal can do

1. Open F1 (the latest quarter file) in Excel.
2. Go to the `Project Info` sheet.
3. Click any cell in row 7 of an active asset (try column G first — slot 1).
4. Look at the formula bar:
   - If it shows `True` (text, left-aligned) → the production macro works, sensitivity macro fails.
   - If it shows `TRUE` (boolean, centre-aligned, possibly all caps) → sensitivity works, production fails.
   - If it shows a formula like `=DropdownChoice` → look at what the dropdown returns.
5. Sanity-check by repeating on 2-3 other slots in different countries.

## Why we're asking now

- Blocks **P2** (the canonical 78-asset list — depends on which test counts an asset as active).
- Downstream gates **engine block #14** (validation): the Python port has to match
  Excel exactly for tests to pass, so we need to know which macro's count is the
  reference.
- Also touches the `DEVIATES_FROM_EXCEL` audit (CLAUDE.md, Coding Standards):
  if we pick "fix the bug", every other call site needs the deviation comment.
