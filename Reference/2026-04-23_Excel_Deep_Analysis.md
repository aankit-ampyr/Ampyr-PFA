# Deep Excel Re-analysis — Project Parthenon

> **Generated:** 2026-04-23
> **Scope:** All three source .xlsm workbooks in [data/converted/](../data/converted/)
> **Supersedes:** partial analysis in [docs/DEVELOPMENT_SPEC.md](../docs/DEVELOPMENT_SPEC.md) sections 4–6 and in `docs/analysis_plw_formulas.txt`
> **Purpose:** Checkpoint — revisit regularly to track what's been covered vs what's still missing.
> **Raw artefacts:** [.claude/analysis_2026_04/](../.claude/analysis_2026_04/) (full JSONs, per-row data, extracted VBA)

---

## 0 · Source files analysed

| Key | File | Size | Role |
|---|---|---:|---|
| F1 | `Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm` | 46.2 MB | **Primary** — ASE's latest quarterly model |
| F2 | `Converted FY27 Business Planv8 copy - file with GTC sheets added.xlsm` | 70.5 MB | GTC-enhanced (21 reporting sheets added) |
| F3 | `converted Project Parthenon - FM - raw file previous quarter.xlsm` | 47.8 MB | Previous quarter — used for diff |

---

## 1 · File inventory (corrected)

| | F1 | F2 | F3 |
|---|---:|---:|---:|
| Sheets | **25** (not 23) | 45 | 23 |
| PLW rows | 1,597 | **1,609** (+12 GTC) | 1,597 |
| Hidden sheets | 0 | 3 (Summary sheet, Q Rep (USD), QRep(EUR)) | 0 |
| External workbook links | **4** | 1 | **0** |
| `vbaProject.bin` bytes | 91,136 | 100,864 | 132,608 |
| VBA code identical across all 3 | ✅ | ✅ | ✅ |

**New sheets in F1 vs F3:**
- `Pltfrm Costs Devex Analysis` (42 rows × 325 cols, 5 CF rules)
- `BESS DCF Multiple Analysis` (196 rows × 383 cols, 382 formulas)

**External workbook dependencies (F1 only — new this quarter!):**
- SharePoint: `Project Canopy FM _v30 - Firsfield standalone.xlsm`
- SharePoint: `Project Canopy FM _v14 (wip).xlsm`
- SharePoint: `Project Canopy FM _v24.xlsm`
- SharePoint: `Project Parthenon - CIP v7 - Capacity Analysis.xlsb`

F2's one external link: `20250314 - Fin Rep.xlsb` (FY26 financial reporting workbook)

→ *F3 had zero external links. ASE started cross-referencing external acquisition/capacity models this quarter. These workbooks are not in our data/ folder.*

---

## 2 · True formula scale (previous analysis was template-count, not cell-count)

The old number "877 formulas" referred to **distinct formula rows in Period 1** of PLW. Real cell counts:

| Sheet | F1 formula cells |
|---|---:|
| HoldCo CFs & Valuation | 663,939 |
| Time Inputs (M) | 602,728 (+ 94k hardcoded) |
| **Project Level Workings** | **401,759** |
| Charts | 304,729 |
| Quarterly Output | 180,779 f + **2,924,999 hardcoded (paste target)** |
| HoldCo income | 155,117 |
| HoldCo_Facility | 92,387 |
| Time Inputs (A) | 19,542 (+ 12k hardcoded) |
| Project Info | 19,908 (+ 18,652 hardcoded) |
| Time Inputs (Q) | **8,147** (F3 had 25,804 — simplified in F1) |

**F2 adds (GTC layer):**
- `Asset workings` alone = **1,191,376 formula cells** — 849k `IFERROR`, 843k `SUMIFS`, 824k `IF`. It reshapes Quarterly Output via 2,529,973 cross-sheet lookups.

---

## 3 · PLW reality-check

Old spec assumed 877 templates, 98.5% reuse, 13 edge-case rows. **Actual:**

| Metric | Old spec | Real |
|---|---:|---:|
| Rows with time-axis formulas | 877 | **915** |
| Rows spanning full 421 time cols | — | 876 |
| Rows with partial span | — | 39 |
| **Rows with formula differing across time cols (edge cases)** | 13 | **75** |
| Time axis start column | col M (13) | col AB (28) |
| Time axis end column | col PQ (433) | col QF (448) |

The 75 edge-case rows are the ones where a "translate once, broadcast 421×" approach will mistranslate. Require per-column handling.

**Cross-sheet references from PLW (real counts):**

| Source | Refs |
|---|---:|
| Time Inputs (A) | **64,413** |
| Time Inputs (M) | 13,472 |
| PLW (self-ref) | 12,630 |
| Project Info | 842 |
| Project Level Macro Paste | 842 |

→ `Time Inputs (A)` is ~5× more load-bearing than (M) for PLW — prioritise it in ingestion.

---

## 4 · Function usage (real PLW time-axis counts)

| Fn | Calls | Fn | Calls |
|---|---:|---|---:|
| SUM | 79,957 | AND | 20,206 |
| IF | 73,649 | SINGLE | 10,102 |
| SUMIFS | 41,257 | OR | 8,819 |
| MATCH | 40,416 | MONTH | 8,799 |
| MAX | 32,839 | ROUND | 5,894 |
| XLOOKUP | 29,470 | YEAR | 2,526 |
| MIN | 28,626 | EOMONTH / NOT | 1,684 each |
| INDEX | 25,260 | MOD / SUMIF | 1,263 each |
| IFERROR | 24,417 | SUMPRODUCT | 841 |
| | | ROUNDUP / YEARFRAC / LOOKUP / ABS / OFFSET | 421 each |

Total: **26 unique Excel functions** (confirmed). OFFSET appears exactly once per time column = 1 template × 421 cols (single dynamic-range pattern).

---

## 5 · VBA — exactly what the live macros do

All 3 files share byte-identical VBA. 751–792 LOC across modules. Live production-path macros:

### `PlatformConsolidation` (Consolidation.bas) — the 2–3 hour monster
1. `Application.Calculation = xlCalculationManual`
2. Force `Live_case = 1`
3. Clear quarterly output paste range (50,000 × 500 cells)
4. Loop `projectList.Count + 1`; for each active project:
   - Set `Project_View = projectList(i)`
   - Call `Debt_sizing`
   - Copy `SPV_ConsolidatedCashflows.Value` into `quarterly output` at `startingrow`
   - `startingrow += rangeRowsCF`
5. Restore `Live_case`

### `Debt_sizing` (DebtSizing.bas) — 42 lines, fixed-point solver
```vba
If debtapplicable = True:
    Sizing_Active = 1; Calculate
    Do
        Use_paste[offset projectnumber-1] := Use_live
        d_service_paste[offset projectnumber-1] := d_service_live
        Calculate
    Loop While debt_delta > 0.2 OR Use_delta > 0.2
    SeniorDebtOptimalValue := SeniorDebtOptimalLive
    Sizing_Active = 0; Calculate
```
- **TWO convergence criteria**, both must pass: `debt_delta < 0.2` AND `Use_delta < 0.2`
- No max-iteration cap (can loop indefinitely in edge cases)

### `Sensitivity` / `Sens_platformconsol` (Module1.bas)
- Reads `Sens_list`, counts non-zero entries
- For each active scenario: `Live_case = j` → `Sens_platformconsol` → paste `Sens_live` into Sensis
- `Sens_platformconsol` is a variant that **skips `Debt_sizing`** when `Sens_override = True`
- Uses `ProjectconsolidateFlag` (typo — main uses `ProjectActiveFlag`)

### `junior_sizing` (Module4.bas)
- Only Junior2 is active (Junior1 commented out)
- Single convergence: `Junior2_delta < 1`

### `Style_killer` (Module3.bas)
- Deletes non-built-in styles — one-shot maintenance utility

### Module2.bas
- Benchmarking timers (`RangeTimer`, `SheetTimer`, `RecalcTimer`, `FullcalcTimer`) — not part of production path

**⚠ Subtle bug risk:** `PlatformConsolidation` checks `projectactiveflag(i) = "True"` (string!); `Sens_platformconsol` uses `= True` (boolean). Different semantics — replicate exactly to match outputs.

---

## 6 · Named ranges

| | Workbook-scoped | Sheet-scoped |
|---|---:|---:|
| F1 | 167 | 31 |
| F2 | 665 | **1,469** |
| F3 | 110 | 0 |

**F1 added 57 names vs F3** — key additions:
- Debt-sizing consolidation: `Debt_copy_consol`, `Debt_paste_consol`, `Debt_delta_consol`
- Gearing: `Gearing_senior_consol`, `Gearing_cap_consol`
- Cash sweep: `CashSweepOn`, `CashSweep_switch`, `CashSweep_consol`
- DSRF: `DSRF_Comm_copy`, `DSRF_Comm_paste`
- 25+ `IQ_*` Capital IQ integration names (addin noise)

**F2 adds ~490 garbage names** from add-ins (`AAAA`, `AJAJJAJ`, `BHUVAN`, `BNE_*`, `AS2*` Smartview) — strip these during ingestion.

**F1 shifted Project Info rows by +2 vs F3** — 13 named ranges all moved by exactly 2 rows (`modelStartDate`, `months`, `mths_per`, `round`, `outputSheetName`, `PPA_type`, `LL_revenue`, `P50_P90`, etc.). → Structural fingerprint must handle row-shift matching, not equality.

---

## 7 · Scenario engine — Sensis structure (more than "6 levers")

Old docs said 6 levers. Actual Sensis layout (cols A-N, rows 6-40+):

**Meta (rows 6-12):** `Sensitivity` (int), `Description`, `Sens_startrow=65`, `Sens_startcol=6`, `Sensitivity in sizing case` (flag), `Selected case`

**Operations (rows 15-19):** Devex %, Capex %, Opex %, O&M %

**Production (rows 21-25):** Net production %, Yield (P50/P90 selector), Curtailment %, Quarterly generation toggle (Annual/Quarterly)

**Uncontracted revenues (rows 27-35):**
- Sensitivity price curve active flag
- Price curve selector (Low/Base/High)
- Indexed price curve toggle
- Merchant power prices adj
- Breakeven power price (UK Solar, Germany Solar, The Netherlands Solar)

**Contracted revenues (rows 37+):** PPA on, Contracted share %, Contracted offtake prices

→ **15+ distinct levers across 4 sections**, mixed types (boolean, enum, percentage).

---

## 8 · Project Info section structure

Scanned col B labels on F1:

| Row | Section |
|---:|---|
| 18 | Timings |
| 20 | Solar+BESS |
| 49 | BESS |
| 61 | Asset sales |
| 79 | Production |
| 119 | Construction |
| 173 | Revenues — PPA / Subsidy contracts |
| 267 | Merchant |
| 280 | Solar- Embedded Benefits - 11kV |
| 291-294 | Q1/Q2/Q3/Q4 |
| 296 | BESS offtake |
| 298 | Merchant |
| 305 | Continuous Intraday Trading (CIDT) |
| 313 | Contracted |
| 523 | Devex - Solar |
| 527 | Capex - BESS |
| 532 | O&M |
| 548 | D&A rates |
| 550 | Solar |
| 559 | BESS |
| 566 | Success Factor |
| 570 | Other |
| 584 | Inputs (DO NOT DELETE) |
| 596 | End |

120 asset slots span horizontally; asset names in row 5, tech type in row 6, active flag in row 7.

### Country Inputs (108 non-empty rows, 2 countries UK + DE)
- Rows 5-29: Tax (corporate tax 1-4, local tax, VAT, EBIT tax, interest deductibility caps, loss carry-forward, WHT)
- Rows 30-36: Shareholder loan (15% interest, annuity, 35yr period)
- Rows 37+: VAT application toggles per capex category
- UK defaults: CT 0%, LT 0%, VAT 20%, WHT 0%
- DE defaults: CT ~0%, Gewerbesteuer 13.65%, VAT 19%

### Financing Inputs
- VAT Loan: margin 1.6%, commitment fee 35% of margin, upfront 1%
- Senior Debt: 20yr tenor, P90 sizing, Low curve, SONIA base, margin 1.9% both periods, swap 4.1%
- DSCR thresholds: **1.15** Solar contracted / **1.30** non-contracted / **1.20** BESS contracted / **1.80** BESS non-contracted

### Time Inputs (A) — 16 major sections
| Row | Section | Rows wide |
|---:|---|---:|
| 8 | Inflation Profiles | ~38 |
| 46 | Financing Curves | ~9 |
| 55 | FX curves | ~6 |
| 61 | Market prices | ~64 |
| 125 | Clawback | ~5 |
| 130 | Negative price adjustment | ~7 |
| 137 | Balancing cost | ~16 |
| 153 | Cash sweep profile for collateral | ~4 |
| 157 | Tax Rates | ~7 |
| 164 | Sources | ~37 |
| 201 | MRA | ~165 |
| 366 | SLA | ~166 |
| 532 | LC amounts | ~113 |

MRA + SLA + LC = ~450 rows — per-asset reserve account timing curves (largest cluster).

---

## 9 · GTC layer — clean 4-tier dependency graph

| Tier | Sheets | Pattern |
|---|---|---|
| **1 · Raw bridge** | Asset workings (1.19M f), Q Rep (USD)/(EUR), Valuation, Valuation (2) | `IFERROR(SUMIFS(QuarterlyOutput!..., criteria), 0)` for each (asset × metric × period) |
| **2 · Breakdown** | FY26A-29F PL Breakdown (14,985 f), FY26A-29F BS Breakdown (6,867 f), CF Capital Strategy Breakdown (1,459 f) | SUM aggregations from Asset workings |
| **3 · Projection-aggregate** | PnL projection (2,018 f), BS projection (1,127 f), CF Capital Strategy (1,903 f) | `ROUND(SUMIF(..., category), 0)` patterns |
| **4 · Summary/viz** | Summary sheet, Cash Breakeven, Graphs, Tables, Net income bridge | SUMs from Tier 3 |

**Orphan sheets (no outgoing formula refs):** FX, Log, Instruction tab, Net income bridge, Disclaimer

**⚠ Live #REF! errors in F2:**
- `PnL projection - aggregate`: 10 broken refs
- `Summary sheet`: 30 broken refs

These propagate. Flag during ingestion.

---

## 10 · Sheet structure — hidden content

Hidden rows/columns in F1 (must still be ingested):

| Sheet | Hidden rows | Hidden cols |
|---|---:|---:|
| HoldCo income | 609 (of 756, = 80%) | 1 |
| Time Inputs (M) | 574 (of 1,581) | 5 |
| Time Inputs (A) | 448 (of 651, = 69%) | 14 |
| HoldCo_Summary | 4 | **141** |
| Charts | 90 | 2 |

### Data validations + conditional formatting (input-integrity enforcement)
| Sheet | Data validations | CF rules |
|---|---:|---:|
| Project Info | 15 | **52** |
| HoldCo_Facility | 0 | 49 |
| PLW | 0 | 34 |
| ProjectSummary | 0 | 12 |
| HoldCo CFs & Valuation | 0 | 10 |
| Sensis | 6 | 1 |

Replicate DV rules via Pydantic validators during ingestion.

---

## 11 · Quarterly diff (F1 vs F3) — the actual churn

| Sheet | Cells scanned | Changed | % |
|---|---:|---:|---:|
| Project Info | 138,635 | 50,330 | **36.3%** |
| Country Inputs | 1,463 | 2 | 0.14% |
| Financing Inputs | 25,025 | 29 | 0.12% |
| Time Inputs (A) | 211,575 | 1,129 | 0.53% |
| **Time Inputs (Q)** | 85,344 | **35,955** | **42.1%** |
| Sensis | 7,840 | 2 | 0.03% |
| **PLW formula rows** | 1,597 | **184** | **11.5%** |

### Decoded changes

**Project Info (36.3%)**:
- Row 5 = asset names. DE renamed: Arendsee/Dalum Chicken Farm/Dalum East/Dalum West replace Adamshoffnung 1/Parchim/Bückwitz/Dalum
- UK typo fix: `Whitney` → `Witney`
- Placeholders → real assets: Jerriestown, Harker 2, North Newton Phase 2
- Row 6 = tech: 3 slots flipped Solar → Solar & BESS
- Row 7 = active flag: 6 flipped False→True (new actives), 2 flipped True→False (retired)

**Country Inputs (stable)**: only 2 header cells. **Tax/VAT/SHL unchanged this quarter.**

**Financing Inputs (stable)**: 29 cells, all date shifts of **exactly +2 years** on Senior Debt period-end + facility maturities (2037→2039, 2047→2049, 2048→2050). Margin/DSCR unchanged.

**Time Inputs (A)**: 1,129 cells — inflation/margin adjustments. Margins dropping 0.021 → 0.02; 2029 reference year → 2030.

**Time Inputs (Q) (42%)**: entire "Assets with Existing Financing" section (KfW Term Loan / Cluster Capacity layout) **removed** and rebuilt with a simpler NL/ECHT-SUSTEREN/Echt-Susteren cluster format. Drove the −17,657 formula drop seen in per-sheet counts.

**Sensis (stable)**: 2 trivial cells. **Scenario engine frozen between quarters.**

**PLW (11.5%)**: 184 rows with formula changes, concentrated in:
- Rows 3, 17-22 (header/flags)
- Rows 139-141, 190-210 (flags & timing)
- Rows 242-327 (production + early revenue)

→ ASE actively modifies PLW every quarter — formulas are **not** frozen between releases.

---

## 12 · Implications for the Python build

1. **Formula stability is within-quarter only.** 184 PLW rows changed between F3 and F1. Engine translator must be re-runnable each quarter with a formula-delta report identifying blocks needing re-validation.

2. **External workbook dependencies are new this quarter.** F1 references 4 external .xlsm/.xlsb files not in our data/ folder. Before Phase 1: decide to ingest (needs SharePoint access), snapshot values inline, or declare out of scope.

3. **Asset workings is a reshape, not financial logic.** Don't translate 1.19M SUMIFS into 1.19M Python lines. In the DB it's a single groupby/pivot on `quarterly_metrics`. Similarly all GTC projection/breakdown sheets are generic `ROUND(SUMIF(...))` — one aggregator, not 21 bespoke reports.

4. **Structural fingerprint must handle row-shift alignment, not equality.** Between F3 and F1, 13 named ranges shifted by exactly +2 rows. Align by section label, not row number.

5. **DSCR solver has TWO convergence criteria.** Both `debt_delta < 0.2` AND `Use_delta < 0.2` must be satisfied. Current `docs/DEVELOPMENT_SPEC.md` and `CLAUDE.md` mention only one — update.

6. **26 functions in the translator is correct, but usage is 1000× higher than template counts** (e.g. SUM 79,957 calls not 160). Vectorisation payoff is larger than previously communicated.

7. **Scenario engine has 15+ levers across 4 sections**, not 6. Mixed types (boolean, enum, percentage). Schema needs lever type discriminator.

8. **Named-range noise must be filtered at ingestion** — F2 has ~490 garbage names from Capital IQ / Bloomberg / Smartview / Access add-ins. Keep only those referencing ASE sheets.

9. **Hidden rows/cols still hold data.** Time Inputs (A) has 448 hidden rows; HoldCo income 609. Read hidden cells; ignore sheet_state for ingestion.

10. **Live `#REF!` errors exist** in F2's `PnL projection - aggregate` (10) and `Summary sheet` (30). Ingestion should log these, not treat them as zero.

---

## 13 · What this analysis did NOT cover (still missing)

Things to revisit later — treat this checklist as living:

- [ ] **HoldCo CFs & Valuation deep dive** — 663,939 formulas, ~3,379 rows. Second-largest formula sheet after Asset workings; not yet mapped.
- [ ] **HoldCo_Facility deep dive** — 92,387 formulas, 49 CF rules, facility-level debt model. Not yet mapped.
- [ ] **HoldCo income** — 155,117 formulas, 80% rows hidden. Needs inspection to confirm what's alive vs legacy.
- [ ] **Time Inputs (M)** — 602,728 formulas. Monthly disaggregation from annual — mechanics not yet decoded.
- [ ] **Charts sheet** — 304,729 formulas. Almost certainly chart-data staging, but worth confirming.
- [ ] **Per-asset parameter inventory** — 590+ parameters per asset claim is from docs, not empirically verified here. Need to scan Project Info horizontally × 120 slots.
- [ ] **Actual list of assets by country/tech in F1** — names were decoded for the DIFF rows only; full 78-asset list not enumerated.
- [ ] **External workbook contents** — are the 4 SharePoint files actually relevant or stale legacy links? Need to open `Project Canopy` + `CIP v7 Capacity Analysis` to see.
- [ ] **`ProjectActiveFlag` vs `ProjectconsolidateFlag` discrepancy** — confirm which VBA path is used in production runs.
- [ ] **The 75 PLW edge-case rows** — need row-by-row inspection to understand what makes each one different across time cols.
- [ ] **The 184 PLW rows that changed between F3→F1** — need side-by-side formula comparison to understand what ASE is actively changing.
- [ ] **Validation oracle** — Excel cached values loaded for all 3.1M Quarterly Output cells? Not verified; `data_only=True` returned mostly None for PLW — confirm caching is intact in the source files.
- [ ] **Colour coding / fill rules** — Excel models often encode input vs formula vs link via colour. Not scanned here.
- [ ] **Print areas, named print titles, page setup** — only relevant if GTC report export must be pixel-perfect.

---

## 14 · Checkpoint status

| Area | Status |
|---|---|
| File inventory | ✅ Complete |
| Named ranges (all 3 files) | ✅ Complete |
| Per-sheet cell/formula counts | ✅ Complete |
| PLW formula templates + cross-sheet refs | ✅ Complete |
| Input sheet structure | ✅ Complete (Project Info, Country Inputs, Financing Inputs, Time Inputs A, Sensis) |
| GTC dependency graph | ✅ Complete |
| VBA macros (all modules) | ✅ Complete + freshly re-extracted |
| Quarterly diff F1 vs F3 | ✅ Complete (cell-level input sheets; row-hash PLW) |
| Data validations + CF + hidden state | ✅ Complete |
| HoldCo sheets (CFs, income, Facility) | ❌ Pending |
| Time Inputs (M) mechanics | ❌ Pending |
| External workbook investigation | ❌ Pending |
| 75 PLW edge-case rows — individual inspection | ❌ Pending |
| 184 changed PLW rows — formula comparison | ❌ Pending |
| Excel cached-value validation oracle | ❌ Not verified |

---

## 15 · Raw artefacts

All under [.claude/analysis_2026_04/](../.claude/analysis_2026_04/):

- `01_inventory.json` — sheet dimensions per file
- `02_named_ranges.json` — all workbook + sheet-scoped names
- `03_sheet_breakdown_{F1,F2,F3}.json` — per-sheet cells/formulas/complexity
- `04_plw_summary.json` + `04_plw_rows.json` — every PLW row with template, fns, cross-sheet refs
- `04_plw_fns_crosssheet.json` — aggregate function + cross-sheet usage
- `05_inputs_pi_params.json` — Project Info parameter labels
- `06_gtc.json` — full GTC dependency matrix
- `07_quarterly_diff.json` — cell-level diffs F1 vs F3
- `08_structure.json` — DV/CF/hidden map per sheet
- `vba/{F1,F2,F3}/*.bas` — freshly extracted VBA from each file
- `*.py` — the analysis scripts themselves (reproducible)
- `*.log` — raw stdout from each run
