# Project Parthenon -- Development Specification

> **Version:** 1.1 | **Date:** April 2026 | **Author:** GTC Product & Technology
> **Status:** Phase 0 Complete · Phase 1 Pending | **Classification:** Internal / Confidential
>
> **Change log v1.1 (2026-04-27):** Reconciled with deep re-analysis dated 2026-04-23
> ([Reference/2026-04-23_Excel_Deep_Analysis.md](../Reference/2026-04-23_Excel_Deep_Analysis.md)).
> Updated formula counts, edge-case count, time-axis column range, DSCR convergence
> criteria, scenario lever inventory, GTC reshape interpretation, validation-oracle
> confirmation. External workbook dependencies declared **out of scope** — values are
> snapshot-as-hardcoded.

---

## 1. Executive Summary

Project Parthenon replaces a 2-3 hour Excel-based scenario/sensitivity analysis workflow with a Python web application. The core financial model (.xlsb, 46.2 MB, owned by ASE) calculates 170 metrics across 120 asset slots (78 real assets) over a 35-year monthly horizon. The application will replicate the Excel calculation engine in Python, enable instant scenario analysis, and handle quarterly model updates automatically.

**Key numbers (re-confirmed against 23-Apr-2026 deep re-analysis):**
- **915** rows with time-axis formulas in PLW (was 877 in v1.0)
- **75 edge-case rows** with formulas that differ across time columns (was 13 in v1.0)
- **876 rows span the full 421 time cols**, 39 rows are partial-span
- **26 unique Excel functions** (confirmed)
- **421 time period columns** spanning **cols AB (28) → QF (448)** (was incorrectly stated as M–PQ in v1.0)
- Single-asset-at-a-time model in Excel; Python engine vectorizes across all 78 assets
- **Validation oracle confirmed:** F1 PLW deep-region cells are 64% cached, Quarterly Output 62% cached. `data_only=True` returns usable values without manual recalc.

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
| F1: Latest ASE Original (.xlsm) | 46.2 MB | **25** (was 23 in v1.0) | **Primary** -- calc engine source, formula extraction |
| F2: GTC Enhanced (.xlsm) | 70.5 MB | 45 | Reporting templates (21 GTC sheets added) |
| F3: Previous Quarter (.xlsm) | 47.8 MB | 23 | Diff testing only |

**External workbook links (F1 only):** F1 references 4 external SharePoint workbooks
(Project Canopy v14/v24/v30, CIP v7 Capacity Analysis). **Decision: out of scope.**
Ingestion pipeline must snapshot externally-resolved cached values as hardcoded inputs
and warn if the link target's cache is stale; the external files themselves are not ingested.

**Sheets new in F1 vs F3:** `Pltfrm Costs Devex Analysis`, `BESS DCF Multiple Analysis`.

**Hidden sheets in F2:** Summary sheet, Q Rep (USD), QRep(EUR) — must still be ingested
(content lives behind `sheet_state="hidden"`).

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

- **Columns AB-QF (421 cols, indices 28-448)** = monthly time periods for ONE asset
- **Row 15** (`Project_View`) = INDEX/MATCH to select the active asset
- **VBA macro** (Consolidation.bas) loops through all assets: sets `Project_View`, triggers recalc, pastes results to Quarterly Output
- This sequential loop is WHY it takes 2-3 hours

**Implication for Python engine:** We translate 915 formula rows ONCE, then vectorize execution across all 78 assets as a `(78 x 421)` NumPy array. What takes Excel 2-3 hours becomes ~2-5 seconds.

> ⚠ **75 edge-case rows** (not 13) have formulas that differ across time columns.
> A "translate once, broadcast 421×" approach mistranslates these. Translator framework
> must support per-column overrides keyed by row × column. See §5.7.

### 5.2 Formula Statistics

| Metric | Value |
|--------|-------|
| Rows with time-axis formulas | **915** |
| Rows spanning full 421 time cols | **876** |
| Rows with partial span | **39** |
| **Edge-case rows (formula differs across cols)** | **75** |
| Time-axis column range | **AB (28) → QF (448)** |
| Unique Excel functions used | 26 |
| Cross-sheet references from PLW | **64,413** to `Time Inputs (A)`, 13,472 to `Time Inputs (M)`, 12,630 self-ref, 842 to `Project Info` |

**True formula scale (cell counts, not template counts):**

| Sheet | F1 formula cells |
|---|---:|
| HoldCo CFs & Valuation | 663,939 |
| Time Inputs (M) | 602,728 |
| Project Level Workings | 401,759 |
| Charts | 304,729 |
| Quarterly Output | 180,779 (formula) + 2,924,999 (hardcoded paste target) |
| HoldCo income | 155,117 |
| HoldCo_Facility | 92,387 |

The "877 formulas" in v1.0 was *distinct rows in Period 1 of PLW*. Real cell counts are 1,000× higher (e.g. SUM has 79,957 calls, not 160). **Vectorisation payoff is therefore much larger than v1.0 communicated.**

### 5.3 Functions to Implement in Python

Counts below are **real cell-call counts** across PLW's time-axis region (cols AB–QF),
not template counts. Source: `.claude/analysis_2026_04/04_plw_fns_crosssheet.json`.

| Function | Calls | Python Equivalent |
|----------|------:|-------------------|
| SUM | 79,957 | `np.sum()` |
| IF | 73,649 | `np.where()` |
| SUMIFS | 41,257 | Filtered `np.sum()` with conditions |
| MATCH | 40,416 | `np.searchsorted()` or `np.argwhere()` |
| MAX | 32,839 | `np.maximum()` |
| XLOOKUP | 29,470 | Array indexing with search |
| MIN | 28,626 | `np.minimum()` |
| INDEX | 25,260 | Array indexing |
| IFERROR | 24,417 | `np.nan_to_num()` / masked ops |
| AND | 20,206 | `np.logical_and()` |
| SINGLE | 10,102 | Identity (Excel 365 implicit intersection) |
| OR | 8,819 | `np.logical_or()` |
| MONTH | 8,799 | `pd.Timestamp.month` |
| ROUND | 5,894 | `np.round()` |
| YEAR | 2,526 | `pd.Timestamp.year` |
| EOMONTH | 1,684 | `pd.offsets.MonthEnd` |
| NOT | 1,684 | `np.logical_not()` |
| MOD | 1,263 | `np.mod()` |
| SUMIF | 1,263 | Filtered sum |
| SUMPRODUCT | 841 | `np.dot()` |
| ROUNDUP | 421 | `np.ceil()` with scaling |
| YEARFRAC | 421 | Day-count calculation |
| LOOKUP | 421 | `np.searchsorted()` |
| ABS | 421 | `np.abs()` |
| OFFSET | 421 | Dynamic range indexing (1 template × 421 cols) |
| INPUTS (custom) | n/a | Named-range lookup helper (per-template, not per-cell) |

26 unique functions total — confirmed unchanged.

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

Real cross-sheet reference counts (from PLW formulas, full time-axis region):

| Source sheet | Refs from PLW |
|---|---:|
| `Time Inputs (A)` | **64,413** |
| `Time Inputs (M)` | 13,472 |
| PLW (self-ref) | 12,630 |
| `Project Info` | 842 |
| `Project Level Macro Paste` | 842 |
| `Country Inputs` | (low) — tax rates, regulatory params |
| `Sensis` | (low) — scenario lever values |

→ `Time Inputs (A)` is **~5× more load-bearing** than `Time Inputs (M)` for PLW.
Prioritise it in the ingestion pipeline.

### 5.7 Edge-case rows (75)

The 75 PLW rows whose formula text differs across the 421 time columns require
per-column handling. Categories include: COD-anchored period flags, partial-span
rows (39), terminal-period adjustments, and BESS-repower discontinuities. Full
list: `.claude/analysis_2026_04/04_plw_summary.json` → `edge_case_rows`.

Translator design: the per-block module declares a default (broadcast) translation
and a dict of `{column_index: override_translation}` for affected rows. Engine
applies the override map per-asset at execution time.

### 5.8 PLW formula churn between quarters

184 of 1,597 PLW formula rows (**11.5%**) changed between F3 and F1, concentrated
in rows 3, 17–22 (header/flags), 139–141, 190–210 (flags & timing), and 242–327
(production + early revenue). **Formulas are not frozen between releases** — the
translator must be re-runnable each quarter and produce a formula-delta report
against the previous baseline so affected blocks can be re-validated.

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
   - Repeats while `debt_delta > 0.2` **OR** `Use_delta > 0.2`
3. Saves `SeniorDebtOptimalValue`
4. Sets `Sizing_Active = 0`

**Key insight:** This is a **simple fixed-point iteration**, NOT Newton-Raphson.
**TWO convergence criteria — both must pass:** `debt_delta < 0.2` AND `Use_delta < 0.2`.
v1.0 of this spec mentioned only one criterion; the VBA actually checks both.
There is **no max-iteration cap** in the VBA — the Python port should add one
(e.g. `max_iter=50`) and emit a hard error on non-convergence to avoid edge-case
infinite loops. Typical convergence is 3-8 iterations per asset.

**⚠ Subtle bug risk:** `PlatformConsolidation` checks `projectactiveflag(i) = "True"`
(string comparison) while `Sens_platformconsol` uses `= True` (boolean comparison).
Different semantics — the Python port must replicate exactly to match Excel outputs.

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
- **15+ scenario levers across 4 sections** (v1.0 said 6 — incorrect):

| Section | Rows | Levers |
|---|---|---|
| Operations | 15-19 | Devex %, Capex %, Opex %, O&M % |
| Production | 21-25 | Net production %, Yield (P50/P90 selector), Curtailment %, Quarterly generation toggle (Annual/Quarterly) |
| Uncontracted revenues | 27-35 | Sensitivity price curve flag, Price curve selector (Low/Base/High), Indexed price curve toggle, Merchant power prices adj, Breakeven power price (UK Solar / DE Solar / NL Solar) |
| Contracted revenues | 37+ | PPA on, Contracted share %, Contracted offtake prices |

Lever **types are mixed** (boolean, enum, percentage). The scenario_overrides table
must carry a `lever_type` discriminator (see §10).

- `Live_case` named range selects the active scenario (1-24)
- Lever values are read by PLW formulas via cross-sheet references to Sensis
- **Stable between quarters:** F3→F1 only changed 2 trivial cells in Sensis. Scenario
  engine layout is effectively frozen.

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
    lever_value_num DECIMAL(20,8),      -- numeric value (when lever_type in 'percentage','absolute')
    lever_value_text VARCHAR(100),      -- enum value (e.g., 'P50','P90','Low','Base','High')
    lever_value_bool BOOLEAN,           -- boolean flag (e.g., PPA on, indexed price curve)
    lever_type      VARCHAR(50),        -- 'percentage','absolute','enum','boolean'
    asset_ids       INTEGER[],          -- NULL = all assets
    target_block    VARCHAR(100)        -- 'Revenue','Production', etc.
);
-- lever_type discriminator added in v1.1 — Sensis carries 15+ levers of mixed type
-- (boolean / enum / percentage), not just numeric percentages.
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

**Row-shift alignment, not equality.** F1 shifted Project Info rows by +2 vs F3 — 13
named ranges (`modelStartDate`, `months`, `mths_per`, `round`, `outputSheetName`,
`PPA_type`, `LL_revenue`, `P50_P90`, etc.) all moved by exactly 2 rows. The structural
fingerprint comparator must align by section label and named-range identity, then
detect shift offsets — equality on row numbers will produce false positives every
quarter.

### 11.3 Empirical churn observed (F3 → F1)

| Sheet | Cells scanned | Changed | % |
|---|---:|---:|---:|
| Project Info | 138,635 | 50,330 | **36.3%** |
| Country Inputs | 1,463 | 2 | 0.14% |
| Financing Inputs | 25,025 | 29 | 0.12% (all date shifts of exactly +2 years) |
| Time Inputs (A) | 211,575 | 1,129 | 0.53% |
| **Time Inputs (Q)** | 85,344 | **35,955** | **42.1%** (entire "Existing Financing" section rebuilt) |
| Sensis | 7,840 | 2 | 0.03% |
| **PLW formula rows** | 1,597 | **184** | **11.5%** |

Project Info churn = asset renames (DE: Arendsee/Dalum cluster replacing
Adamshoffnung/Parchim/Bückwitz/Dalum), UK typo fix (`Whitney` → `Witney`),
placeholders becoming real assets (Jerriestown, Harker 2, North Newton Phase 2),
3 tech-type flips Solar→Solar+BESS, 6 active-flag flips False→True, 2 True→False.

### 11.4 Robustness rules for ingestion

1. **Hidden rows still hold data.** `Time Inputs (A)` has 448 hidden rows (69%);
   `HoldCo income` has 609 (80%). Read regardless of `sheet_state`.
2. **Live `#REF!` errors exist** in F2 (`PnL projection - aggregate`: 10,
   `Summary sheet`: 30). Log them, do not silently coerce to zero.
3. **Named-range noise:** F2 carries ~490 garbage names from add-ins (Capital IQ
   `IQ_*`, Bloomberg, Smartview `AS2*`, Access `BNE_*`). Filter at ingestion to
   only those referencing ASE sheets.
4. **Data validations + conditional formatting** carry input-integrity rules
   (`Project Info` has 15 DV + 52 CF rules). Replicate as Pydantic validators.
5. **External workbook references in F1** (Project Canopy v14/v24/v30, CIP v7
   Capacity Analysis): out of scope. Snapshot externally-resolved cached values
   as inline hardcoded inputs; warn if the cached value is `#REF!` or stale.

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

**Validation oracle confirmed available** (probed 2026-04-27 via
[devtools/oracle_probe.py](../devtools/oracle_probe.py)). F1 PLW deep-region
(rows 250-260, cols AB-AZ) returns 64% populated cells via `data_only=True`;
Quarterly Output returns 62%. Both intermediate (PLW) and end-state
(Quarterly Output) values are usable as oracles — no manual `F9-and-save`
prerequisite. Block-level unit tests against PLW intermediates remain viable.

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
| Revenue engine complexity (200 formula rows, country-specific DE/UK/NL) | High | Medium | Start Revenue block early; Finance SME for DE/UK/NL contract logic review |
| **75 PLW edge-case rows mistranslated by broadcast** (revised up from 13) | High | Medium-High | Translator framework supports per-column override map; explicit test for every edge-case row |
| DSCR solver convergence edge cases (TWO criteria, no VBA cap) | Medium | Low-Medium | Add `max_iter` cap with hard error; verify both `debt_delta` AND `Use_delta < 0.2` |
| Formula translation errors (rounding, date boundaries) | Medium | Medium | Automated validation suite comparing every cell against Excel; fix divergences iteratively |
| Finance team SME availability | High | Medium | Schedule weekly 30-min validation sessions; prepare specific questions in advance |
| **Quarterly PLW formula churn** (184 rows changed F3→F1) | Medium | High | Translator re-runnable each quarter + formula-delta report; affected blocks re-validated |
| Quarterly structural drift (row shifts, renamed sections) | Medium | Medium | Fingerprint comparator aligns by section label / named-range identity, not row number |
| HoldCo CFs / Time Inputs (M) still black-boxed | Medium | Medium | Investigation gated before respective blocks (see §17) |
| External SharePoint workbooks unavailable | Low | Low | Declared out of scope (§4.1, §11.4) — values snapshot inline at ingestion |
| NumPy vectorization edge cases | Low | Low | Fallback to per-asset iteration for any blocks that resist vectorization |
| Live `#REF!` errors propagate as zeros | Medium | Low | Ingestion logs `#REF!` cells, refuses to silently coerce |

---

## 15. Key Reference Documents

| Document | Location | Contents |
|----------|----------|----------|
| **Deep Re-analysis Checkpoint (2026-04-23)** | [Reference/2026-04-23_Excel_Deep_Analysis.md](../Reference/2026-04-23_Excel_Deep_Analysis.md) | **Authoritative source for v1.1 numbers.** Re-baselined formula counts, edge cases, churn metrics |
| Raw analysis artefacts | `.claude/analysis_2026_04/*.json` | Per-sheet breakdowns, PLW row inventory, named ranges, GTC dependency matrix |
| Structural Map | `docs/STRUCTURAL_MAP.md` | Complete row-level blueprint of all 45 sheets (1,998 lines) |
| PLW Formula Analysis | `docs/analysis_plw_formulas.txt` | PLW formula templates analysis (6,454 lines) — superseded for counts by §5 above |
| GTC & Input Analysis | `docs/analysis_gtc_inputs.txt` | Formula counts, cross-sheet dependencies, structural similarity |
| VBA Macros | `macros/` | 8 exported modules (Consolidation, DebtSizing, Sensitivity, etc.) |
| Development Timeline | `docs/Ampyr Financial Model Digitisation timeline.xlsx` | 25 work items, 3 buckets — re-baselined v1.1 (see §17) |
| Validation oracle probe | [devtools/oracle_probe.py](../devtools/oracle_probe.py) | Empirical proof cached values exist in F1 PLW + Quarterly Output |

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

---

## 17. Deferred Investigation Gates (v1.1)

The 23-Apr re-analysis closed 8 of 15 investigation areas and left 7 open.
Each remaining item is gated to the latest engineering activity it would otherwise
block, so the program can proceed without finishing all investigation up-front.

| # | Investigation | Effort | Gates which work item | Status |
|---|---|---|---|---|
| 1 | Per-asset parameter inventory (590+ × 120 slots, Project Info horizontal scan) | 1 day | Bucket 2 #15 (Schema), #16 (Ingestion) | Pending — must complete before Phase 1 |
| 2 | Full 78-asset list by country/tech in F1 | 0.25 day | Bucket 2 #15 (Schema) | Pending — must complete before Phase 1 |
| 3 | Time Inputs (M) monthly disagg mechanics (603k formulas) | 2-3 days | Bucket 1 #4 (Revenue) | Pending — gate Revenue block |
| 4 | HoldCo CFs & Valuation (664k formulas, 3,379 rows) | 3-4 days | Bucket 1 #11 (IRR), #12-#13 implicitly | Pending — gate IRR / consolidation |
| 5 | HoldCo_Facility (92k formulas, facility-level debt) | 1-2 days | Bucket 1 #8 (Senior Debt + DSCR) | Pending — gate Senior Debt block |
| 6 | HoldCo income (155k formulas, 80% rows hidden) | 1 day | Bucket 1 #11 (IRR) | Pending — gate IRR |
| 7 | 184 PLW rows that changed F3→F1 (side-by-side formula comparison) | 0.5 day | Translator framework (Bucket 1 #1) | Pending — informs translator design |
| 8 | 75 PLW edge-case rows (per-row inspection) | 1 day | Translator framework (Bucket 1 #1) | Pending — informs translator design |
| — | Charts sheet (305k formulas) | — | None — staging only | **Deferred indefinitely** |
| — | Excel cached-value validation oracle | — | All engine work | **CLOSED** ✅ — confirmed 2026-04-27 |
| — | External SharePoint workbooks | — | Ingestion | **CLOSED** ✅ — declared out of scope |
| — | `ProjectActiveFlag` vs `ProjectconsolidateFlag` discrepancy | — | DSCR + Sensitivity blocks | Open — inspect at block-implementation time |
| — | Colour coding / fill rules | — | None unless GTC export must be pixel-perfect | **Deferred** |
| — | Print areas / page setup | — | None unless GTC export must be pixel-perfect | **Deferred** |

### Net effort impact

- **Before Phase 1** (~2 days): items 1, 2, plus partial item 3 (Time Inputs (M) annual→monthly disagg only)
- **Before Bucket 1 #1 Translator** (~1.5 days): items 7, 8 — design the per-column override map
- **Before Bucket 1 #4 Revenue** (~2 days): finish item 3
- **Before Bucket 1 #8 Senior Debt** (~1-2 days): item 5
- **Before Bucket 1 #11 IRR** (~4-5 days): items 4, 6

Total net new investigation: **~10 days (~2 weeks)**. The existing 2.2-week buffer
absorbs this with margin. **No timeline slip projected.**

### Re-baselined timeline (v1.1)

The 25-item plan in `docs/Ampyr Financial Model Digitisation timeline.xlsx` remains
structurally valid. Adjustments to apply when next regenerated:

| Item | v1.0 estimate | v1.1 adjustment | Reason |
|---|---|---|---|
| #1 Translator framework | 2.0 w | **+0.5 w → 2.5 w** | 75 edge-case rows + per-column override map |
| #4 Revenue engine | 2.8 w | unchanged | Investigation absorbs into existing buffer |
| #8 Senior Debt + DSCR | 1.4 w | **+0.2 w → 1.6 w** | TWO convergence criteria + max_iter cap |
| #12 Scenario engine | 1.0 w | **+0.2 w → 1.2 w** | 15+ levers across 4 sections, mixed types |
| #14 Validation | 1.6 w | unchanged | Oracle confirmed; plan stands |
| #16 Ingestion pipeline | 1.4 w | **+0.3 w → 1.7 w** | Hidden rows, #REF! handling, named-range filtering, row-shift alignment |
| #19 GTC reporting | 1.4 w | **−0.4 w → 1.0 w** | Asset workings collapses to one groupby/pivot, not 1.19M lines |
| Buffer | 2.2 w | **−0.8 w → 1.4 w** | Absorbs net adjustments |

Net Bucket totals remain ~28.5 weeks. The buffer shrinks but stays positive.
