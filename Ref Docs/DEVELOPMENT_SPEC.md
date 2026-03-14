# Project Parthenon -- Development Specification

> **Version:** 1.0 | **Date:** March 2026 | **Author:** GTC Product & Technology
> **Status:** Pre-Development | **Classification:** Internal / Confidential

---

## 1. Executive Summary

Project Parthenon replaces a 2-3 hour Excel-based scenario/sensitivity analysis workflow with a Python web application. The core financial model (.xlsb, ~34MB, owned by ASE) calculates 170 metrics across 120 asset slots (78 real assets) over a 35-year monthly horizon. The application will replicate the Excel calculation engine in Python, enable instant scenario analysis, and handle quarterly model updates automatically.

**Key numbers from deep analysis:**
- 877 formula templates in the calculation engine
- 26 unique Excel functions to translate
- 98.5% formula reuse across time periods (only 13 edge-case rows)
- 421 time period columns (monthly), single-asset-at-a-time model
- Python engine will vectorize across all 78 assets simultaneously

---

## 2. Business Context

### 2.1 Organization

- **Ampyr Energy Tech Solutions (GTC)** -- Owner of 200-asset pipeline (Solar, Solar+BESS) across Europe, USA, Australia
- **Ampyr Solar Europe (ASE)** -- Owns and maintains the core financial model, updates quarterly
- **Other entities** -- BD team, Investment team, Debt team -- provide inputs to ASE, consume outputs

### 2.2 Current Workflow

```
ASE collects changes from BD/Investment/Debt teams
  --> Updates core .xlsb quarterly
  --> Sends to GTC
  --> GTC adds 21 reporting sheets
  --> GTC runs scenario/sensitivity analysis (2-3 hours per run)
  --> Results used for investment decisions, debt covenants, board reporting
```

### 2.3 Pain Points

1. **Scenario analysis takes 2-3 hours** -- Excel recalculates all 120 assets sequentially even when only 2-3 change
2. **Manual quarterly update process** -- No automated way to detect what changed between file versions
3. **No portfolio-level view** -- Analysis is asset-by-asset in Excel
4. **GTC cannot modify the core model** -- Must add separate sheets and work around ASE's structure

---

## 3. Architecture Decision

### 3.1 Options Evaluated

| Option | Approach | Verdict |
|--------|----------|---------|
| **A: Formula Extraction from .xlsb** | Extract every formula, build Python computation graph | **Rejected** -- pyxlsb cannot extract formulas (only cached values). Solved by converting to .xlsm. |
| **B: Rebuild from First Principles** | Study structure, rebuild each block in Python | Best long-term but slow to start |
| **C: Hybrid (CHOSEN)** | Ingest Excel outputs first for reporting; build Python calc engine incrementally, validated against Excel | Delivers value early, Excel remains validation oracle |
| **D: Excel-as-a-Service** | Automate Excel via COM/xlwings | 100% fidelity but doesn't solve the speed problem |

### 3.2 Chosen Architecture: Option C (Hybrid)

```
+-----------------+     +------------------------------------+
|  .xlsm file     |---->|  Ingestion Pipeline (openpyxl)     |
|  (quarterly)    |     |  - Parse all sheets                |
+-----------------+     |  - Normalize into tables           |
                        |  - Assign version_id               |
                        +----------------+-------------------+
                                         |
                        +----------------v-------------------+
                        |        PostgreSQL                   |
                        |                                     |
                        |  assets (78 real + 42 placeholder)  |
                        |  asset_parameters (per-asset)       |
                        |  time_series_inputs (M/Q/A curves)  |
                        |  country_parameters                 |
                        |  financing_terms                    |
                        |  quarterly_metrics (3.1M cells)     |
                        |  holdco_metrics                     |
                        |  model_versions (quarterly track)   |
                        |  scenarios + scenario_overrides     |
                        +-------+------------------+---------+
                                |                  |
                   +------------v---+    +---------v-----------+
                   |  FastAPI        |   |  Python Calc Engine  |
                   |  - CRUD APIs    |   |  - NumPy vectorized  |
                   |  - Diff API     |   |  - Selective recalc  |
                   |  - Scenario API |   |  - Dependency graph  |
                   |  - Report Gen   |   |  - Asset-level       |
                   +--------+-------+   +----------------------+
                            |
                   +--------v--------+
                   |  Streamlit UI    |
                   |  - Asset browser |
                   |  - Diff viewer   |
                   |  - Scenario UI   |
                   |  - Report export |
                   +------------------+
```

### 3.3 Tech Stack

| Layer | Technology | Rationale |
|-------|-----------|-----------|
| Backend | Python 3.12 + FastAPI | Async, fast, great ecosystem for financial math |
| Database | PostgreSQL 16 | Partitioning for large datasets, JSONB for flexible params |
| ORM | SQLAlchemy 2.0 + Alembic | Migrations, type safety |
| Calc Engine | NumPy + SciPy | Vectorized operations across all assets simultaneously |
| Frontend | Streamlit | Fast prototyping, built-in auth, charts |
| Excel I/O | openpyxl (read .xlsm), xlsxwriter (write reports) | Formula extraction from .xlsm format |
| Deployment | Docker + docker-compose | PostgreSQL + app containers |

---

## 4. Source File Analysis

### 4.1 File Inventory

| File | Size | Sheets | Role |
|------|------|--------|------|
| F1: Latest ASE Original (.xlsm) | 46.2 MB | 25 | **Primary** -- calc engine source, formula extraction |
| F2: GTC Enhanced (.xlsm) | 70.5 MB | 45 | Reporting templates (21 GTC sheets added) |
| F3: Previous Quarter (.xlsm) | 47.8 MB | 23 | Diff testing only |

### 4.2 Asset Portfolio

- **120 total slots** in the model (78 real assets, 42 placeholders)
- **Countries:** Germany (29), UK (27), Netherlands (22)
- **Technologies:** Solar & BESS (39), Solar (28), BESS (11)
- **35-year horizon**, monthly granularity (421 time periods)
- **170 metrics** per asset in Quarterly Output

### 4.3 Sheet Structure (F2 -- 45 sheets)

| Category | Sheets | Purpose |
|----------|--------|---------|
| ASE Input (6) | Project Info, Country Inputs, Financing Inputs, Time Inputs (A/M/Q) | Raw parameters, pricing curves, dates |
| ASE Scenario (1) | Sensis | 24 scenario slots, 6 levers |
| ASE Calc Engine (1) | Project Level Workings | **The core** -- 1,609 rows x 421 time cols |
| ASE Output (2) | Quarterly Output, Project Level Macro Paste | Consolidated results (3.1M cells) |
| ASE HoldCo (3) | HoldCo CFs & Valuation, HoldCo income, HoldCo_Facility | Group-level aggregation |
| ASE Presentation (5) | Dashboard, ProjectSummary, HoldCo_Summary, Charts, Checks | Reporting views |
| ASE Helper (3) | FX, Lists, Scenarios | Lookup tables |
| GTC Reporting (21) | PnL/BS projections, CF Capital Strategy, QRep, Valuation, Asset workings, etc. | GTC-added aggregation/reporting |

---

## 5. Calculation Engine -- Deep Analysis

### 5.1 Critical Discovery: Single-Asset-at-a-Time Model

The PLW sheet does NOT calculate all assets simultaneously. It is a **template-based model**:

- **Columns M-PQ (421 cols)** = monthly time periods for ONE asset
- **Row 15** (`Project_View`) = INDEX/MATCH to select the active asset
- **VBA macro** (Consolidation.bas) loops through all assets: sets `Project_View`, triggers recalc, pastes results to Quarterly Output
- This sequential loop is WHY it takes 2-3 hours

**Implication for Python engine:** We translate 877 formula templates ONCE, then vectorize execution across all 78 assets as a `(78 x 421)` NumPy array. What takes Excel 2-3 hours becomes ~2-5 seconds.

### 5.2 Formula Statistics

| Metric | Value |
|--------|-------|
| Total formula rows (Period 1) | 877 |
| Simple (basic arithmetic) | 381 (43.4%) |
| Medium (1-2 functions) | 265 (30.2%) |
| Complex (nested IF/INDEX/MATCH) | 170 (19.4%) |
| Cross-sheet references | 61 (7.0%) |
| Formula reuse across time periods | 98.5% identical |
| Edge-case rows (unique logic) | 13 |
| Unique Excel functions used | 26 |

### 5.3 Functions to Implement in Python

| Function | Count | Python Equivalent |
|----------|:-----:|-------------------|
| SUM | 160 | `np.sum()` |
| IF | 123 | `np.where()` |
| SUMIFS | 81 | Filtered `np.sum()` with conditions |
| MAX | 75 | `np.maximum()` |
| MIN | 62 | `np.minimum()` |
| INDEX | 60 | Array indexing |
| MATCH | 60 | `np.searchsorted()` or `np.argwhere()` |
| INPUTS (custom) | 58 | Named range lookup helper |
| IFERROR | 57 | `try/except` or `np.nan_to_num()` |
| XLOOKUP | 50 | Array indexing with search |
| AND | 41 | `np.logical_and()` |
| SINGLE | 22 | Identity (Excel 365 implicit intersection) |
| OR | 12 | `np.logical_or()` |
| ROUND | 10 | `np.round()` |
| MONTH/YEAR | 10 | `pd.Timestamp` properties |
| EOMONTH | 3 | `pd.offsets.MonthEnd` |
| MOD | 3 | `np.mod()` |
| SUMIF | 3 | Filtered sum |
| YEARFRAC | 1 | Day-count calculation |
| SUMPRODUCT | 1 | `np.dot()` |
| OFFSET | 1 | Dynamic range indexing |
| LOOKUP | 1 | `np.searchsorted()` |
| ABS | 1 | `np.abs()` |
| ROUNDUP | 1 | `np.ceil()` with scaling |
| NOT | 4 | `np.logical_not()` |

### 5.4 Calculation Block Map

The 877 formulas are organized into these major sections within PLW:

| Block | Row Range | Formula Rows | Complexity | Key Logic |
|-------|-----------|:------------:|------------|-----------|
| Header & Checks | 1-30 | ~25 | Low | Project selection, validation flags, technology/country lookup |
| Financial Statements | 32-134 | ~80 | Low | Income statement, cashflow, balance sheet -- all are SUMIFS aggregation of lower blocks |
| Flags & Timing | 136-245 | ~100 | Low-Medium | COD flags, period flags, inflation indices (XLOOKUP into Time Inputs A) |
| Production | 247-271 | ~20 | Medium | Solar yield, degradation, curtailment, BESS dispatch/repower |
| Generation Waterfall | 273-309 | ~30 | Medium | UK CfD metering, net generation |
| **Revenue** | **310-619** | **~200** | **High** | **PPA, CfD, FiT/FiP, EEG, merchant, GoO, tolling, floor, capacity market. Country-specific branching (DE/UK/NL). LARGEST BLOCK.** |
| OpEx | 621-710 | ~60 | Medium | O&M Solar/BESS, insurance, land lease (fixed + revenue-dependent), community benefit |
| Depreciation | 712-807 | ~70 | Medium | 7 depreciation categories, fully-depreciated checks, BESS repower |
| VAT | 809-924 | ~50 | Medium | VAT on CapEx/OpEx/Revenue, receivable/payable, VAT facility |
| Senior Debt & DSCR | 926-1090 | ~100 | High | Drawdown, repayment, cash sweep, interest, DSCR sculpting (iterative solver), DSRA, MRA |
| Tax | 1092-1211 | ~80 | Medium-High | Corporate tax, loss carry-forward, local tax basis, interest deduction. Country-specific |
| Funding & Sources/Uses | 1157-1211 | ~30 | Medium | Funding calcs, sources/uses reconciliation |
| SHL & Distribution | 1214-1385 | ~110 | Medium-High | Shareholder loans, distribution waterfall, lock-up DSCR, dividends, WHT |
| IRR & Quarterly Consolidation | 1387-1596 | ~140 | Medium | XIRR, SUMIFS quarterly rollup, financial statements, BS check |

### 5.5 Calculation Dependency Chain

```
Flags/Timings --> Production --> Revenue --> OpEx --> EBITDA
                                                       |
                                Working Capital <-- Tax <-- Depreciation
                                                       |
                                Senior Debt <--> DSCR Sculpting (iterative solver)
                                                       |
                                Reserve Security --> Distribution Waterfall --> IRR
                                                       |
                                Quarterly Consolidation --> HoldCo Aggregation
```

**Selective recalculation:** Changing "Net Production" for 3 German assets only triggers `Production -> Revenue -> EBITDA -> Tax -> Debt -> IRR` for those 3 assets. Estimated ~2 seconds in NumPy vs 2-3 hours in Excel.

### 5.6 Cross-Sheet Dependencies (PLW)

PLW formulas reference these sheets:
- `Time Inputs (A)` -- 46 cross-sheet formulas (inflation indices, pricing curves via XLOOKUP)
- `Project Info` -- 10 cross-sheet formulas (asset parameters via INDEX/MATCH)
- `Country Inputs` -- 3 cross-sheet formulas (tax rates, regulatory params)
- `Sensis` -- 2 cross-sheet formulas (scenario lever values)
- `Dashboard` -- implicit references via named ranges

---

## 6. VBA Macro Analysis

### 6.1 Consolidation.bas -- The Asset Loop (CRITICAL)

This is the macro that takes 2-3 hours. Our Python engine replaces this entirely.

**What it does:**
1. Sets calculation to manual
2. Sets `Live_case = 1` (base case)
3. Loops through `ProjectList` (120 slots)
4. For each active project:
   - Sets `Project_View` = project name (triggers PLW to recalculate for that asset)
   - Calls `Debt_sizing` (the circular reference solver)
   - Pastes `SPV_ConsolidatedCashflows` range values into `Quarterly Output`
5. Advances paste row by `rangeRowsCF` for next asset
6. Restores original `Live_case`

**Python replacement:** No looping needed. All 78 assets calculated in one vectorized pass.

### 6.2 DebtSizing.bas -- The DSCR Solver (CRITICAL)

**What it does:**
1. Sets `Sizing_Active = 1`
2. Enters a `Do...Loop While` that:
   - Copies `Use_live` values to `Use_paste` (offset by project number)
   - Copies `d_service_live` to `d_service_paste`
   - Triggers `Calculate`
   - Repeats until `debt_delta < 0.2` AND `Use_delta < 0.2`
3. Saves `SeniorDebtOptimalValue`
4. Sets `Sizing_Active = 0`

**Key insight:** This is a **simple fixed-point iteration**, NOT Newton-Raphson. The convergence criterion is `delta < 0.2`, which is relatively loose. Python implementation is straightforward: iterate until convergence, typically 3-8 iterations per asset.

### 6.3 Module1.bas -- Sensitivity Macro

**What it does:**
1. Reads `Sens_list` to count active scenarios
2. For each scenario:
   - Sets `Live_case = j`
   - Calls `Sens_platformconsol` (a variant of the consolidation loop that optionally skips debt sizing based on `Sens_override`)
   - Pastes `Sens_live` range into Sensis sheet
3. Restores original `Live_case`

**Python replacement:** Scenario engine simply sets lever values and triggers the calc engine. With vectorized computation, all 24 scenarios can run in parallel.

### 6.4 Module4.bas -- Junior Debt Sizing

Similar fixed-point iteration for junior/mezzanine debt (currently commented out for Junior1, active for Junior2). Convergence criterion: `Junior2_delta < 1`.

### 6.5 Other Modules

- **Module2.bas** -- Timer utilities for benchmarking recalc performance
- **Module3.bas** -- Style cleanup utility (deletes non-built-in styles)
- **Module5.bas** -- Empty
- **Sheet1.cls** -- Empty event handler
- **Sensitivity.bas** -- Empty (logic moved to Module1)

---

## 7. Input Sheets Analysis

### 7.1 Summary

| Sheet | Total Cells | Formulas | Hardcoded | Input Levers | Key Functions |
|-------|:-----------:|:--------:|:---------:|:------------:|---------------|
| Project Info | 41,626 | 16,277 | 25,349 | 17,321 | INDEX/MATCH, IF, IFERROR |
| Country Inputs | 483 | 22 | 461 | 253 | Simple references |
| Financing Inputs | 583 | 105 | 478 | 333 | EOMONTH, IF, AVERAGE |
| Time Inputs (A) | 20,444 | 16,578 | 3,866 | 3,084 | INDEX/MATCH, XLOOKUP, SUMIFS |
| Time Inputs (M) | 152,605 | 151,838 | 767 | 253 | SUMIFS, IF, AND (monthly disagg) |
| Time Inputs (Q) | 12,054 | 8,147 | 3,907 | 3,793 | IF, EOMONTH |
| Sensis | 252 | 49 | 203 | 26 | LOOKUP |
| **Total** | **228,047** | **193,016** | **35,031** | **25,063** | |

### 7.2 Input Levers

There are **25,063 hardcoded numeric input values** across the input sheets. These are the cells that can change between quarterly updates. The quarterly comparison engine must detect changes in all of these.

The largest concentration is in **Project Info (17,321 levers)** -- this contains per-asset parameters (590+ per asset) including COD dates, capacities, contract terms, pricing assumptions, financing terms.

---

## 8. GTC Sheets Analysis

### 8.1 Summary

| Metric | Value |
|--------|-------|
| Total GTC sheets | 21 |
| Total formula cells | 142,589 |
| Complex formulas | 16 (out of 142K) |
| Primary functions | SUM, SUMIF, SUMIFS, ROUND, IF |
| Sheets referencing ASE data | 11 of 21 |

### 8.2 Complexity Assessment

The 21 GTC sheets are **overwhelmingly simple**:
- 59,011 simple formulas (basic arithmetic, SUM)
- 83,562 medium formulas (SUMIF/SUMIFS with one condition)
- Only **16 complex formulas** (XNPV in Valuation sheet)

### 8.3 Largest GTC Sheet

`Asset workings` (105,325 cells, 105,044 formulas) is the largest GTC sheet. It pulls data from `Quarterly Output` via SUMIFS and reorganizes it by asset. This is pure data aggregation -- no financial logic.

### 8.4 Structural Similarity

- `FY26B PL Mapping` and `FY26B BS Mapping` are structurally identical (same dimensions, same function set)
- PnL/BS "projection - aggregate" sheets use identical ROUND(SUMIF(...)) patterns
- All breakdown sheets share the same SUM-based column aggregation

### 8.5 GTC-ASE Dependency Flow

```
ASE Layer:
  Time Inputs (A)  <---- Most referenced ASE sheet (by 8 GTC sheets)
  HoldCo_Facility  <---- Referenced by 5 GTC sheets
  HoldCo CFs & Val <---- Referenced by 5 GTC sheets
  Quarterly Output <---- Referenced by Asset workings
  Project Info     <---- Referenced by Asset workings

GTC Layer:
  Breakdown sheets (PL/BS/CF) --> read ASE data + aggregate by period
  Projection sheets           --> SUMIF from Breakdown sheets
  Summary sheet               --> SUMs from Projection sheets
  Valuation sheets            --> XLOOKUP/XNPV from HoldCo data
  QRep (EUR/USD)              --> FX conversion of Valuation output
```

---

## 9. Scenario & Sensitivity Engine

### 9.1 Sensis Sheet Structure

- **24 scenario slots** (columns L-AJ, rows 11-12 for names/IDs)
- **6 scenario levers:**
  1. Devex adjustment (%)
  2. Capex adjustment (%)
  3. Opex adjustment (%)
  4. Production adjustment (%)
  5. Revenue/pricing adjustment (%)
  6. Financing terms adjustment (%)
- `Live_case` named range selects the active scenario (1-24)
- Lever values are read by PLW formulas via cross-sheet references to Sensis

### 9.2 How Scenarios Work (from VBA)

```
1. Set Live_case = scenario_number
2. Excel formulas in PLW read lever values from Sensis via named ranges
3. Formulas apply adjustments: e.g., Revenue = Base_Revenue * (1 + Sensis_Revenue_Adj)
4. Full consolidation loop runs (all assets recalculated)
5. Results pasted to output area
6. Repeat for next scenario
```

### 9.3 Python Implementation

In Python, scenarios become parameter overrides:
```python
# Pseudocode
scenario = {"production_adj": -0.05, "opex_adj": 0.10}  # -5% prod, +10% opex
results = calc_engine.run(assets=all_assets, overrides=scenario)
# Executes in ~2-5 seconds for all 78 assets
```

For sensitivity analysis, the VBA runs 1,872 recalculations (24 scenarios x 78 assets). The Python engine does this as matrix operations in a single pass.

---

## 10. Database Schema

```sql
-- Asset registry (from Project Info)
CREATE TABLE assets (
    id              SERIAL PRIMARY KEY,
    name            VARCHAR(255) NOT NULL,
    country         VARCHAR(50),        -- DE, UK, NL
    technology      VARCHAR(50),        -- Solar PV, BESS, Solar & BESS
    capacity_mw     DECIMAL(10,2),
    slot_index      INTEGER,            -- Position in ProjectList (1-120)
    include_in_consolidation BOOLEAN DEFAULT TRUE,
    version_id      INTEGER REFERENCES model_versions(id)
);

-- Per-asset parameters (590+ per asset, from Project Info)
CREATE TABLE asset_parameters (
    id              SERIAL PRIMARY KEY,
    asset_id        INTEGER REFERENCES assets(id),
    parameter_name  VARCHAR(255),       -- e.g., 'COD_Date', 'PPA_Price', 'Capacity_MW'
    value_numeric   DECIMAL(20,6),
    value_text      TEXT,
    value_date      DATE,
    source_sheet    VARCHAR(100),       -- 'Project Info', 'Financing Inputs', etc.
    source_cell     VARCHAR(20),        -- e.g., 'I42'
    version_id      INTEGER REFERENCES model_versions(id)
);

-- Time series inputs (monthly power prices, inflation, interest rates)
CREATE TABLE time_series_inputs (
    id              SERIAL PRIMARY KEY,
    series_name     VARCHAR(255),       -- e.g., 'DE_Solar_Price_P50', 'UK_CPI'
    frequency       VARCHAR(10),        -- 'M', 'Q', 'A'
    period_start    DATE,
    value           DECIMAL(20,8),
    source_sheet    VARCHAR(100),
    version_id      INTEGER REFERENCES model_versions(id)
);

-- Core output: calculated metrics per asset per period
CREATE TABLE quarterly_metrics (
    asset_id        INTEGER REFERENCES assets(id),
    metric_name     VARCHAR(255),       -- e.g., 'Revenue', 'EBITDA', 'DSCR'
    metric_category VARCHAR(100),       -- 'Income Statement', 'Cashflow', 'Balance Sheet'
    period_start    DATE,
    year            INTEGER,
    quarter         INTEGER,
    value           DECIMAL(20,6),
    unit            VARCHAR(50),        -- 'LC 000', '%', 'x'
    version_id      INTEGER REFERENCES model_versions(id),
    scenario_id     INTEGER REFERENCES scenarios(id)
) PARTITION BY RANGE (year);
-- Indexed on (asset_id, metric_name, version_id, scenario_id)

-- Version tracking
CREATE TABLE model_versions (
    id              SERIAL PRIMARY KEY,
    quarter         VARCHAR(10),        -- 'Q1 2026'
    upload_date     TIMESTAMP DEFAULT NOW(),
    file_name       VARCHAR(500),
    file_hash       VARCHAR(64),        -- SHA-256 for dedup
    is_baseline     BOOLEAN DEFAULT FALSE,
    structural_fingerprint JSONB        -- For quarterly comparison
);

-- Scenario management
CREATE TABLE scenarios (
    id              SERIAL PRIMARY KEY,
    name            VARCHAR(255),
    description     TEXT,
    base_version_id INTEGER REFERENCES model_versions(id),
    created_by      VARCHAR(100),
    created_at      TIMESTAMP DEFAULT NOW()
);

CREATE TABLE scenario_overrides (
    id              SERIAL PRIMARY KEY,
    scenario_id     INTEGER REFERENCES scenarios(id),
    lever_name      VARCHAR(100),       -- 'production_adj', 'opex_adj', etc.
    lever_value     DECIMAL(10,6),      -- e.g., -0.05 for -5%
    lever_type      VARCHAR(50),        -- 'percentage', 'absolute'
    asset_ids       INTEGER[],          -- NULL = all assets
    target_block    VARCHAR(100)        -- 'Revenue', 'Production', etc.
);
```

---

## 11. Quarterly Update Workflow

### 11.1 Process Flow

```
1. ASE delivers new .xlsb/.xlsm
2. Upload to app --> ingestion pipeline parses all sheets
3. Automated diff against previous baseline:
   - Structural fingerprint comparison (using STRUCTURAL_MAP.md baseline)
   - Asset parameter changes (new assets, changed values in 25,063 levers)
   - Time series curve updates (prices, rates)
   - Financing term changes
   - Formula changes (row insertions/deletions detected via label matching)
4. Change report generated:
   - Summary: X assets changed, Y parameters modified, Z new assets
   - Drill-down: exact cell-level diffs with old/new values
   - Heatmap: metric-level impact visualization
5. User reviews and approves
6. New version becomes baseline; previous version archived
7. Calc engine re-validated against new Excel output
```

### 11.2 Structural Fingerprint

The `STRUCTURAL_MAP.md` (1,998 lines) documents every row label, section boundary, and named range. The comparison engine uses this to:
- Detect inserted/deleted rows (row labels shift)
- Identify renamed sections
- Flag structural drift (new calc blocks, removed blocks)
- Map formula changes to affected calc engine modules

---

## 12. Project Structure

```
Ampyr-PFA/
+-- app/
|   +-- models/                    # SQLAlchemy models
|   |   +-- asset.py
|   |   +-- time_series.py
|   |   +-- quarterly_metric.py
|   |   +-- scenario.py
|   |   +-- version.py
|   +-- ingestion/                 # Excel --> DB pipeline
|   |   +-- xlsm_reader.py        # openpyxl-based sheet reader
|   |   +-- normalizer.py         # Raw cells --> normalized tables
|   |   +-- differ.py             # Quarter-over-quarter diff
|   |   +-- structural_compare.py # STRUCTURAL_MAP comparison
|   +-- engine/                    # Python calc engine
|   |   +-- translator.py         # Excel formula --> Python translator
|   |   +-- flags.py              # Flags & Timing block
|   |   +-- production.py         # Production engine
|   |   +-- revenue.py            # Revenue engine (largest)
|   |   +-- opex.py               # OpEx block
|   |   +-- depreciation.py       # Depreciation & Fixed Assets
|   |   +-- vat.py                # VAT block
|   |   +-- debt.py               # Senior Debt + DSCR solver
|   |   +-- tax.py                # Tax block
|   |   +-- shl.py                # Shareholder Loans
|   |   +-- distribution.py       # Distribution Waterfall
|   |   +-- irr.py                # IRR & Valuation (XIRR)
|   |   +-- statements.py         # Financial statements assembly
|   |   +-- dependency_graph.py   # Selective recalculation
|   |   +-- scenario_engine.py    # Scenario lever application
|   +-- api/                       # FastAPI endpoints
|   |   +-- assets.py
|   |   +-- scenarios.py
|   |   +-- reports.py
|   |   +-- versions.py
|   |   +-- diff.py
|   +-- reports/                   # GTC report generation
|   |   +-- gtc_generator.py      # 21 GTC sheet templates
|   |   +-- excel_exporter.py     # Formatted .xlsx output
|   +-- ui/                        # Streamlit pages
|       +-- dashboard.py
|       +-- asset_browser.py
|       +-- diff_viewer.py
|       +-- scenario_builder.py
|       +-- sensitivity_runner.py
|       +-- report_builder.py
+-- migrations/                    # Alembic migrations
+-- tests/
|   +-- validation/                # Excel vs Python comparison tests
|   +-- engine/                    # Calc engine unit tests
+-- Ref Docs/                      # Reference Excel files & analysis
|   +-- STRUCTURAL_MAP.md          # Complete structural blueprint
|   +-- analysis_plw_formulas.txt  # PLW formula analysis (6,454 lines)
|   +-- analysis_gtc_inputs.txt    # GTC & input sheet analysis
|   +-- Macros/                    # Exported VBA modules
+-- requirements.txt
+-- docker-compose.yml             # PostgreSQL + app
+-- .env                           # Environment configuration
```

---

## 13. Verification Plan

### 13.1 Per-Block Validation

For each calc engine block, after Python implementation:
1. Select 5 representative assets (mix of DE/UK/NL, Solar/BESS/Solar+BESS)
2. Run Python engine for those assets
3. Compare every calculated value against Excel's cached output
4. **Divergence target:** < 0.01% for all metrics
5. Investigate and fix any divergence > 0.001%

### 13.2 Full Model Validation

Once all blocks are complete:
1. Run Python engine for all 78 assets, base case
2. Compare all 170 metrics x 421 periods x 78 assets against Quarterly Output
3. Run 3 different scenarios, compare against Excel sensitivity output
4. Verify DSCR solver convergence matches VBA (< 0.2 delta threshold)

### 13.3 Performance Targets

| Operation | Target | Excel Baseline |
|-----------|--------|---------------|
| Single asset recalc | < 0.5 seconds | ~1.5 minutes |
| Full portfolio (78 assets) | < 5 seconds | 2-3 hours |
| Scenario comparison (24 scenarios) | < 30 seconds | 12+ hours |
| Quarterly file ingestion | < 2 minutes | Manual (hours) |
| Change report generation | < 30 seconds | Manual (hours) |

---

## 14. Risk Register

| Risk | Impact | Likelihood | Mitigation |
|------|--------|------------|------------|
| Revenue engine complexity (200 formula rows, country-specific) | High | Medium | Start Revenue block early; get Finance SME dedicated time for DE/UK/NL contract logic review |
| DSCR solver convergence edge cases | Medium | Low | VBA shows simple fixed-point iteration with loose tolerance (0.2); replicate exact same approach |
| Formula translation errors (rounding, date boundaries) | Medium | Medium | Automated validation suite comparing every cell against Excel; fix divergences iteratively |
| Finance team SME availability | High | Medium | Schedule weekly 30-min validation sessions; prepare specific questions in advance |
| Quarterly file structure changes | Medium | Low | Structural fingerprint comparison engine detects changes; alerts for manual review |
| NumPy vectorization edge cases | Low | Low | Fallback to per-asset iteration for any blocks that resist vectorization |

---

## 15. Key Reference Documents

| Document | Location | Contents |
|----------|----------|----------|
| Structural Map | `Ref Docs/STRUCTURAL_MAP.md` | Complete row-level blueprint of all 45 sheets (1,998 lines) |
| PLW Formula Analysis | `Ref Docs/analysis_plw_formulas.txt` | All 877 formula templates, complexity classification, reuse analysis (6,454 lines) |
| GTC & Input Analysis | `Ref Docs/analysis_gtc_inputs.txt` | Formula counts, cross-sheet dependencies, structural similarity |
| VBA Macros | `Ref Docs/Macros/` | 8 exported modules (Consolidation, DebtSizing, Sensitivity, etc.) |
| Development Timeline | `Ref Docs/Project_Parthenon_Timeline.xlsx` | 25 work items, 3 buckets, ~28.5 weeks at 25 hrs/week |
| Named Ranges | `Ref Docs/STRUCTURAL_MAP.md` Section 2 | 67 relevant named ranges with references and purposes |

---

## 16. Glossary

| Term | Definition |
|------|-----------|
| **ASE** | Ampyr Solar Europe -- entity that owns and maintains the core financial model |
| **GTC** | Ampyr Energy Tech Solutions (the user's organization) |
| **PLW** | Project Level Workings -- the core calculation sheet (1,609 rows) |
| **DSCR** | Debt Service Coverage Ratio -- key covenant metric, creates circular reference with debt sizing |
| **COD** | Commercial Operations Date -- when an asset starts generating revenue |
| **PPA** | Power Purchase Agreement -- contracted electricity price |
| **CfD** | Contract for Difference -- UK government subsidy mechanism |
| **FiT/FiP** | Feed-in Tariff / Feed-in Premium -- German subsidy mechanisms |
| **EEG** | Erneuerbare-Energien-Gesetz -- German renewable energy law |
| **GoO** | Guarantees of Origin -- tradeable renewable energy certificates |
| **SHL** | Shareholder Loan -- intercompany debt |
| **DSRA** | Debt Service Reserve Account |
| **MRA** | Maintenance Reserve Account |
| **WHT** | Withholding Tax |
| **IRR** | Internal Rate of Return |
| **XIRR** | Excel function for IRR with irregular dates |
| **HoldCo** | Holding Company level aggregation |
| **Sensis** | Sensitivity/Scenario control sheet |
| **Live_case** | Named range selecting the active scenario (1-24) |
| **Project_View** | Named range selecting the active asset for PLW calculation |
