# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Project Parthenon** — Digitize Ampyr's Excel-based renewable energy financial model into a Python web application. The Excel model analyzes a portfolio of 78 solar/BESS assets across DE/UK/NL with 170 metrics each over a 35-year monthly horizon (3.1M output cells). Current Excel workflow takes 2-3 hours; target is <5 seconds.

## Tech Stack

Python 3.12 · FastAPI · PostgreSQL 16 · SQLAlchemy 2.0 + Alembic · NumPy/SciPy (vectorized calc engine) · Streamlit (UI) · openpyxl (Excel read) · xlsxwriter (report export) · Docker

## Repository Structure

- `docs/` — All project documentation (specs, analysis, plans)
  - `DEVELOPMENT_SPEC.md` — **Master spec v1.1** (2026-04-27): architecture, schema, quarterly workflow, risk register, deferred-investigation gates, re-baselined timeline
  - `STRUCTURAL_MAP.md` — Row-by-row blueprint of all 45 Excel sheets, named ranges, section boundaries
  - `analysis_plw_formulas.txt` — Deep analysis of PLW formulas — superseded for counts by `Reference/2026-04-23_Excel_Deep_Analysis.md`
  - `analysis_gtc_inputs.txt` — GTC reporting sheet and input sheet dependency analysis
- `Reference/` — Authoritative deep re-analysis
  - `2026-04-23_Excel_Deep_Analysis.md` — Re-baselined formula counts, edge cases, churn metrics, GTC reshape findings
- `.claude/analysis_2026_04/` — Raw analysis artefacts (JSON per-sheet breakdowns, PLW row inventory, named ranges, GTC dependency matrix, freshly-extracted VBA)
- `data/` — Source Excel financial models (.xlsb originals)
  - `converted/` — .xlsm versions for openpyxl reading
- `macros/` — Exported VBA modules (8 modules, 751–792 LOC across files) documenting current Excel workflows
- `scripts/` — Python utility/analysis scripts (formula extraction, timeline generation)
- `devtools/` — Internal dashboards and one-off probes (workflow_explorer.py, oracle_probe.py). NOT part of the production app — see "Core / Dev separation" section below.
- `archive/` — Superseded files (JS scripts, node config)

## Architecture

**Pattern:** Hybrid Ingestion + Python Engine
1. Ingest Excel outputs into PostgreSQL for immediate reporting
2. Build Python calculation engine incrementally, validated against Excel
3. Excel serves as validation oracle throughout development

**Planned app structure:**
- `app/models/` — SQLAlchemy ORM (asset, time_series, scenario, version)
- `app/ingestion/` — Excel → DB pipeline (xlsm_reader, normalizer, differ, structural_compare)
- `app/engine/` — Vectorized calc engine translating 915 PLW formula rows (with 75 edge-case per-column overrides) to NumPy
- `app/api/` — FastAPI endpoints (assets, scenarios, reports, diff)
- `app/reports/` — GTC report generation + Excel export
- `app/ui/` — Streamlit pages (dashboard, scenario builder, sensitivity runner)

## Calculation Engine Execution Order

The engine must execute blocks in this order (mirrors Excel sheet structure):
1. Header & Checks → 2. Flags & Timing → 3. Production → 4. Generation Waterfall →
5. **Revenue** (largest: 200 formulas, country-specific DE/UK/NL branching) →
6. OpEx → 7. Depreciation → 8. VAT → 9. **Senior Debt & DSCR** (circular ref solver) →
10. Tax → 11. SHL & Distribution → 12. IRR & Consolidation

## Critical Implementation Details

- **DSCR Circular Reference:** Debt sizing requires fixed-point iteration (see `macros/DebtSizing.bas`). **TWO convergence criteria — both must pass:** `debt_delta < 0.2` AND `Use_delta < 0.2`. The VBA has no max-iter cap; the Python port must add one (`max_iter=50`) and emit a hard error on non-convergence. Typical: 3-8 iterations per asset.
- **Revenue Block Complexity:** 200 formulas with country-specific logic (PPA, CfD, FiT, EEG, merchant, GoO, tolling, capacity market). Highest risk area.
- **PLW edge cases:** **75 rows** have formulas that differ across the 421 time columns (not 13 as originally stated). Translator must support per-column override maps keyed by row × column.
- **PLW time axis:** Cols **AB (28) → QF (448)**, not M–PQ. 915 rows with time-axis formulas; 876 span the full 421 cols, 39 are partial-span.
- **Formula churn:** 184 PLW rows changed between F3→F1 (11.5%). The translator must be re-runnable each quarter with a formula-delta report — formulas are not frozen between quarterly releases.
- **Scenarios:** 24 scenario slots × **15+ levers across 4 sections** (Operations / Production / Uncontracted revenues / Contracted revenues). Mixed types (boolean / enum / percentage). Schema needs a `lever_type` discriminator. Sensis layout is stable between quarters.
- **VBA flag-comparison subtlety:** `PlatformConsolidation` uses `projectactiveflag(i) = "True"` (string compare); `Sens_platformconsol` uses `= True` (boolean compare). Replicate exactly to match Excel.
- **Validation:** Compare Python output against Excel for all 3.1M cells. Tolerance-based matching (`np.allclose(a, b, rtol=1e-4, atol=1e-4)`).
- **Validation oracle: confirmed available** (probed 2026-04-27). F1 PLW deep-region returns 64% populated cached values, Quarterly Output 62%. No manual `F9-and-save` step needed. See [devtools/oracle_probe.py](devtools/oracle_probe.py).
- **External workbook links: out of scope.** F1 references 4 SharePoint files (Project Canopy v14/v24/v30, CIP v7 Capacity); ingestion snapshots externally-resolved cached values inline and warns on stale cache. Do not attempt to ingest the external workbooks themselves.
- **Live `#REF!` errors:** F2 has 10 in `PnL projection - aggregate` and 30 in `Summary sheet`. Ingestion must log, not silently coerce to zero.
- **Hidden rows hold real data:** `Time Inputs (A)` 69% hidden, `HoldCo income` 80% hidden. Read regardless of `sheet_state`.
- **Named-range filtering at ingestion:** F2 carries ~490 garbage names from add-ins (Capital IQ `IQ_*`, Bloomberg, Smartview `AS2*`, Access `BNE_*`). Keep only those referencing ASE sheets.
- **Asset workings is a reshape, not logic:** F2's `Asset workings` (1.19M formula cells, 849k IFERROR + 843k SUMIFS) collapses to a single groupby/pivot on `quarterly_metrics`. Do not translate per-cell.

## Excel Functions to Translate (26 total)

Real per-cell call counts (PLW time-axis region) — vectorisation payoff is much larger than template counts suggest:

Top calls: SUM (79,957), IF (73,649), SUMIFS (41,257), MATCH (40,416), MAX (32,839), XLOOKUP (29,470), MIN (28,626), INDEX (25,260), IFERROR (24,417), AND (20,206), SINGLE (10,102), OR (8,819), MONTH (8,799), ROUND (5,894), YEAR (2,526), EOMONTH/NOT (1,684 each), MOD/SUMIF (1,263 each), SUMPRODUCT (841), ROUNDUP/YEARFRAC/LOOKUP/ABS/OFFSET (421 each).

## Core / Dev Separation (READ BEFORE EDITING)

The repo is deliberately split between the production app and the development
scaffolding that supports building it. Keep these separable.

**Core app (ships):**
- `app/` — FastAPI + SQLAlchemy + calc engine + Streamlit production UI
- `migrations/` — Alembic versioned schema
- `tests/` — engine + validation tests
- `requirements.txt` — **runtime dependencies** (everything anything imports at
  runtime: prod app + dashboards). Streamlit Cloud reads this file directly.

**Dev scaffolding (does NOT ship):**
- `devtools/` — internal Streamlit dashboards, one-off probes, exploration tools
- `scripts/` — analysis scripts, timeline generator, formula extractors
- `.claude/analysis_2026_04/` — raw JSON analysis artefacts
- `archive/` — superseded files
- `requirements-dev.txt` — **dev-only tools** (pytest, ruff, httpx) layered on
  top of `requirements.txt` via `-r`

**Source data / docs (neither ships nor "dev"):**
- `data/` — Excel files (input fixtures); `macros/` — extracted VBA reference
- `docs/`, `Reference/` — specs, structural map, formula analyses, deep checkpoints

### Boundary rules

1. **Nothing in `app/` may import from `devtools/`, `scripts/`, or `archive/`.**
   Production code does not depend on dev tooling.
2. **`tests/` may import from `app/` only** — not from `devtools/` or `scripts/`.
3. **`devtools/` and `scripts/` may import from `app/`** (to exercise it) and from
   any dev dependency. They may NOT be imported by `app/`.
4. **`app/` code must not import `pandas`, `plotly`, or `matplotlib`.** The engine
   is pure NumPy/SciPy. The boundary is enforced at **import statements** —
   `requirements.txt` carries pandas/plotly/pyyaml because the dev dashboards
   need them at runtime, but `app/` code cannot reach for them. Verify before
   merge with: `grep -rn "^\(import\|from\) \(pandas\|plotly\|matplotlib\)" app/`
   (must return zero matches).
5. **When adding a new file, ask "does this ship?"** — if no, it goes in `devtools/`
   or `scripts/`. If yes, it goes in `app/`. New runtime deps go in `requirements.txt`;
   new dev-only tools (test, lint, format) go in `requirements-dev.txt`.
6. **Streamlit is shared infrastructure.** `app/ui/` (production pages) and
   `devtools/` (internal dashboards) both use Streamlit, but they live in different
   directories so the boundary stays visible.

## Development Environment

- **Package manager:** pip
  - Production install: `pip install -r requirements.txt`
  - Development install: `pip install -r requirements-dev.txt` (includes prod)
- **Linting/formatting:** `ruff check --fix && ruff format`
- **Testing:** `pytest tests/`
- **Database:** `docker-compose up -d db` then `alembic upgrade head`
- **App server:** `fastapi dev app/main.py`
- **Streamlit (production):** `streamlit run app/ui/dashboard.py`
- **Streamlit (dev dashboard):** `streamlit run devtools/workflow_explorer.py`

## Coding Standards

- Use `ruff` for all linting and formatting (replaces black + flake8)
- Type hints required on all public functions
- `snake_case` for variables, functions, DB columns; `PascalCase` for SQLAlchemy models
- Engine block modules must include a docstring referencing the PLW row range they implement
- **`DEVIATES_FROM_EXCEL` convention.** During the Excel→Python port, replicate Excel
  behaviour exactly — including known quirks. If a deviation is genuinely required
  (e.g. adding the `max_iter=50` cap to the DSCR solver that the VBA lacks), mark
  it with a comment in this exact form so it's greppable:
  ```python
  # DEVIATES_FROM_EXCEL: <one-line reason>
  ```
  Examples that MUST be flagged: the DSCR `max_iter` cap; any "fix" of the
  `PlatformConsolidation` string-vs-boolean flag bug (default: replicate the bug,
  do not fix); any rounding/precision choice that differs from openpyxl's reading
  of cached values. Run `grep -rn DEVIATES_FROM_EXCEL app/` before each validation
  pass to enumerate every divergence and confirm it's intentional.

## Technical Constraints

- **Vectorization:** ALL time-series calculations MUST use NumPy vectorized ops. NO `for` loops along the time axis (421 periods) or asset axis (78 assets). This is the core performance requirement.
- **DSCR solver:** Explicit `while` loop with convergence check `delta < 0.2` — match VBA exactly (see `macros/DebtSizing.bas`). No recursion.
- **Calc engine:** Pure NumPy arrays. No pandas DataFrames in the engine — pandas is fine for ingestion and reports.
- **Excel reading:** Use `openpyxl` with `read_only=True` and `data_only=True` for large workbooks.

## Validation Rules

- Compare Python output vs Excel cached values: `np.allclose(a, b, rtol=1e-4, atol=1e-4)`
- Divergence target: < 0.01% for all metrics; investigate anything > 0.001%
- Test with 5 representative assets: mix of DE/UK/NL + Solar/BESS/Solar+BESS
- DSCR convergence must match VBA: **both** `debt_delta < 0.2` AND `Use_delta < 0.2`, typically 3-8 iterations per asset
- Every engine block module must have a corresponding validation test
- All 75 PLW edge-case rows need explicit per-row tests (not just broadcast verification)

## Behavioral Guidelines (Karpathy Skills)

> Source: https://raw.githubusercontent.com/forrestchang/andrej-karpathy-skills/main/CLAUDE.md
> Merged into this file 2026-04-27. **Tradeoff:** these guidelines bias toward
> caution over speed — for trivial tasks, use judgment.

### 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them — don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

### 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

### 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it — don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: every changed line should trace directly to the user's request.

### 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:

```text
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.

### How they're applied in this repo

The four principles above are general. Below are the concrete operational forms
they take in Project Parthenon.

**Goal-Driven Execution → todos must include a verify check.**
TodoWrite items must state what success looks like, not just the action.

- ❌ Weak: `"Implement SQLAlchemy models"`
- ✅ Strong: `"Implement SQLAlchemy models for Asset/TimeSeries/Scenario/Version → verify: alembic upgrade head succeeds AND each model docstring references its DEVELOPMENT_SPEC §10 section"`

When a multi-step plan is stated in chat, use the `→ verify: [check]` form for each step.

**Surgical Changes → flag every Excel deviation.**
See the `DEVIATES_FROM_EXCEL` convention under Coding Standards. Default is to
replicate Excel/VBA behaviour exactly, including known quirks. Deviations are
greppable and must be justified per call site.

**Think Before Coding → Phase 1 prerequisites get a "Questions to settle first" list.**
Before starting P1-P5 (see Current Phase below), surface 2-3 specific questions
whose answers determine the deliverable shape. Settle them with the user before
producing data, schema rows, or code.

**Scope of these conventions:** they're the default for `app/` and `tests/` work
(production code, where the bar is highest). They're advisory for `docs/` and
`devtools/` (where iteration and exploration are the point).

## Git & Version Control

- Git LFS is configured for `*.xlsb` and `*.xlsm` files
- Excel source files are in `data/` (original .xlsb) and `data/converted/` (.xlsm for openpyxl)
- Three quarterly snapshots exist for diff testing between quarters

## Current Phase

**Phase 0 COMPLETE — Project scaffolding done (2026-03-14)**
**Deep re-analysis COMPLETE (2026-04-23)** — see `Reference/2026-04-23_Excel_Deep_Analysis.md`
**Spec re-baselined to v1.1 (2026-04-27)** — DEVELOPMENT_SPEC.md

What's been set up:
- Python 3.12 venv created with all dependencies installed (`requirements.txt`)
- `docker-compose.yml` configured for PostgreSQL 16 (local Docker)
- Alembic initialized (`alembic.ini` + `migrations/`)
- App package structure created: `app/{models,ingestion,engine,api,reports,ui}/`
- `app/config.py` and `app/database.py` — settings and DB session management
- Test directories: `tests/{engine,validation}/`
- Ruff configured (`ruff.toml`), `.env.example` provided
- Project files reorganized from flat `Ref Docs/` into `docs/`, `data/`, `macros/`, `scripts/`, `archive/`

**Phase 1 prerequisites** (2 days, must complete before schema work).
Per the "Think Before Coding" convention, each item starts with **Questions to
settle first** captured as ADRs in [docs/decisions/](docs/decisions/) and answered
by the SME via the dev dashboard's 🤔 Decisions tab.

**P1 · Per-asset parameter inventory** (horizontal scan of Project Info × 120 slots)
ADRs to draft (not yet authored):
- ADR: parameter-name canonical format (snake_case derived vs raw Excel labels)
- ADR: do we capture units / scale / source cell, or just (name, value)
- (P1 scope re partial slots is covered by ADR 0003)

**P2 · Full 78-asset list by country/tech in F1** — ADRs drafted:
- [0001 · "active flag" canonical test](docs/decisions/0001-active-flag-canonical-test.md) (string vs boolean)
- [0002 · asset-list output shape](docs/decisions/0002-asset-list-output-shape.md) (CSV / JSON / Alembic seed)
- [0003 · scope — all 120 slots vs only the 78 actives](docs/decisions/0003-active-vs-all-slots-scope.md)

**P3 · Time Inputs (M) annual→monthly disagg mechanic** (partial — just the mechanic)
ADRs to draft on start (disagg pattern: uniform / weighted / season-specific; rule location).

**P4 · 75 PLW edge-case rows — per-row inspection** (gates Translator framework)
ADRs to draft on start (categorisation axis; output: row-by-row catalogue vs per-block summary).

**P5 · 184 changed PLW rows — F3 vs F1 formula comparison**
ADRs to draft on start (diff format).

**Then Phase 1 — Database schema & ingestion**

To resume development:
1. `python -m venv .venv && .venv/Scripts/activate && pip install -r requirements.txt`
2. `docker-compose up -d db`
3. Read `docs/DEVELOPMENT_SPEC.md` v1.1 (especially §17 deferred-investigation gates) for full architecture and schema design
4. Complete Phase 1 prerequisites above
5. Implement SQLAlchemy models in `app/models/` (Asset, TimeSeries, Scenario, Version) — note `scenario_overrides` needs `lever_type` discriminator (see DEVELOPMENT_SPEC §10)
6. Generate and run Alembic migration: `alembic revision --autogenerate -m "initial schema" && alembic upgrade head`
7. Do NOT begin engine work until the DB schema is complete and migrated

Active branch: `Prototype`
