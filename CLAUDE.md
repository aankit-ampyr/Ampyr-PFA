# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Project Parthenon** — Digitize Ampyr's Excel-based renewable energy financial model into a Python web application. The Excel model analyzes a portfolio of 78 solar/BESS assets across DE/UK/NL with 170 metrics each over a 35-year monthly horizon (3.1M output cells). Current Excel workflow takes 2-3 hours; target is <5 seconds.

## Tech Stack

Python 3.12 · FastAPI · PostgreSQL 16 · SQLAlchemy 2.0 + Alembic · NumPy/SciPy (vectorized calc engine) · Streamlit (UI) · openpyxl (Excel read) · xlsxwriter (report export) · Docker

## Repository Structure

- `docs/` — All project documentation (specs, analysis, plans)
  - `DEVELOPMENT_SPEC.md` — Master spec: architecture, database schema, quarterly workflow, risk register
  - `STRUCTURAL_MAP.md` — Row-by-row blueprint of all 45 Excel sheets, named ranges, section boundaries
  - `analysis_plw_formulas.txt` — Deep analysis of 877 calculation engine formulas (complexity, reuse, dependencies)
  - `analysis_gtc_inputs.txt` — GTC reporting sheet and input sheet dependency analysis
- `data/` — Source Excel financial models (.xlsb originals)
  - `converted/` — .xlsm versions for openpyxl reading
- `macros/` — Exported VBA modules (557 LOC) documenting current Excel workflows
- `scripts/` — Python utility/analysis scripts (formula extraction, timeline generation)
- `archive/` — Superseded files (JS scripts, node config)

## Architecture

**Pattern:** Hybrid Ingestion + Python Engine
1. Ingest Excel outputs into PostgreSQL for immediate reporting
2. Build Python calculation engine incrementally, validated against Excel
3. Excel serves as validation oracle throughout development

**Planned app structure:**
- `app/models/` — SQLAlchemy ORM (asset, time_series, scenario, version)
- `app/ingestion/` — Excel → DB pipeline (xlsm_reader, normalizer, differ)
- `app/engine/` — Vectorized calc engine translating 877 Excel formulas to NumPy
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

- **DSCR Circular Reference:** Debt sizing requires fixed-point iteration (see `macros/DebtSizing.bas`). Loop until delta < 0.2. This is the hardest piece to get right.
- **Revenue Block Complexity:** 200 formulas with country-specific logic (PPA, CfD, FiT, EEG, merchant, GoO, tolling, capacity market). Highest risk area.
- **Formula Reuse:** 98.5% of formulas repeat identically across 421 time columns — vectorize along the time axis, don't loop.
- **Scenarios:** 24 scenario slots × 6 levers each (production, capex, opex, revenue/pricing, devex, financing).
- **Validation:** Compare Python output against Excel for all 3.1M cells. Tolerance-based matching for float precision.

## Excel Functions to Translate (26 total)

Most frequent: SUM(160), IF(123), SUMIFS(81), MAX(75), MIN(62), INDEX/MATCH(60 each), IFERROR(57), XLOOKUP(50), AND(41). Also: custom INPUTS(58) and SINGLE(22) functions.

## Development Environment

- **Package manager:** pip (with requirements.txt)
- **Linting/formatting:** `ruff check --fix && ruff format`
- **Testing:** `pytest tests/`
- **Database:** `docker-compose up -d db` then `alembic upgrade head`
- **App server:** `fastapi dev app/main.py`
- **Streamlit:** `streamlit run app/ui/dashboard.py`

## Coding Standards

- Use `ruff` for all linting and formatting (replaces black + flake8)
- Type hints required on all public functions
- `snake_case` for variables, functions, DB columns; `PascalCase` for SQLAlchemy models
- Engine block modules must include a docstring referencing the PLW row range they implement

## Technical Constraints

- **Vectorization:** ALL time-series calculations MUST use NumPy vectorized ops. NO `for` loops along the time axis (421 periods) or asset axis (78 assets). This is the core performance requirement.
- **DSCR solver:** Explicit `while` loop with convergence check `delta < 0.2` — match VBA exactly (see `macros/DebtSizing.bas`). No recursion.
- **Calc engine:** Pure NumPy arrays. No pandas DataFrames in the engine — pandas is fine for ingestion and reports.
- **Excel reading:** Use `openpyxl` with `read_only=True` and `data_only=True` for large workbooks.

## Validation Rules

- Compare Python output vs Excel cached values: `np.allclose(a, b, rtol=1e-4, atol=1e-4)`
- Divergence target: < 0.01% for all metrics; investigate anything > 0.001%
- Test with 5 representative assets: mix of DE/UK/NL + Solar/BESS/Solar+BESS
- DSCR convergence must match VBA (< 0.2 delta, typically 3-8 iterations per asset)
- Every engine block module must have a corresponding validation test

## Git & Version Control

- Git LFS is configured for `*.xlsb` and `*.xlsm` files
- Excel source files are in `data/` (original .xlsb) and `data/converted/` (.xlsm for openpyxl)
- Three quarterly snapshots exist for diff testing between quarters

## Current Phase

**Phase 0 COMPLETE — Project scaffolding done (2026-03-14)**

What's been set up:
- Python 3.12 venv created with all dependencies installed (`requirements.txt`)
- `docker-compose.yml` configured for PostgreSQL 16 (local Docker)
- Alembic initialized (`alembic.ini` + `migrations/`)
- App package structure created: `app/{models,ingestion,engine,api,reports,ui}/`
- `app/config.py` and `app/database.py` — settings and DB session management
- Test directories: `tests/{engine,validation}/`
- Ruff configured (`ruff.toml`), `.env.example` provided
- Project files reorganized from flat `Ref Docs/` into `docs/`, `data/`, `macros/`, `scripts/`, `archive/`

**Next: Phase 1 — Database schema & ingestion**

To resume development:
1. `python -m venv .venv && .venv/Scripts/activate && pip install -r requirements.txt`
2. `docker-compose up -d db`
3. Read `docs/DEVELOPMENT_SPEC.md` for full architecture and schema design
4. Implement SQLAlchemy models in `app/models/` (Asset, TimeSeries, Scenario, Version)
5. Generate and run Alembic migration: `alembic revision --autogenerate -m "initial schema" && alembic upgrade head`
6. Do NOT begin engine work until the DB schema is complete and migrated

Active branch: `Prototype`
