---
id: "NNNN"
title: One-line question
status: pending
gates: [P?, engine_#?]
asked: 2026-MM-DD
decided_by: null
decided_on: null
decision: null
notes: null
options:
  - id: a
    label: Option A short description
    consequence: What happens if we pick this
  - id: b
    label: Option B short description
    consequence: What happens if we pick this
# YAML gotcha: if a label/consequence contains '#' followed by a space (e.g. "Bucket 2 #15"),
# wrap the whole value in single quotes — otherwise YAML treats the rest as a comment and truncates.
---

## Where in Excel

- **Sheet:** ...
- **Row / Cell / Named range:** ...
- **Formula or VBA snippet:**

```text
(paste the relevant formula or VBA here)
```

## What's ambiguous

One or two sentences in plain English. What's the choice and why does it matter?

## Quick check Anchal can do

A concrete instruction the SME can follow in 30 seconds to confirm an answer:
"Open F1, click cell X, look at the formula bar..."

## Why we're asking now

Which work item this gates (e.g. "Blocks P2 deliverable; downstream gates Bucket 2 #15 schema").
