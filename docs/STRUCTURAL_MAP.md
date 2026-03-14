# Project Parthenon -- Complete Structural Map

## File: F2 (GTC Enhanced) -- 45 sheets, 70.5 MB

> **Purpose**: This document is the canonical blueprint for comparing new quarterly Excel files
> against the known structural baseline. Every row label, section boundary, and named range
> is documented here so the comparison engine can detect insertions, deletions, renames, and
> structural drift.

---

## Table of Contents

1. [Sheet Inventory](#1-sheet-inventory)
2. [Named Ranges (67 relevant)](#2-named-ranges-67-relevant)
3. [ASE Input Sheets](#3-ase-input-sheets)
   - 3.1 Project Info
   - 3.2 Country Inputs
   - 3.3 Financing Inputs
   - 3.4 Time Inputs (A) -- Annual
   - 3.5 Time Inputs (M) -- Monthly
   - 3.6 Time Inputs (Q) -- Quarterly
4. [ASE Scenario Control](#4-ase-scenario-control)
   - 4.1 Sensis
5. [ASE Calculation Engine](#5-ase-calculation-engine)
   - 5.1 Project Level Workings
6. [ASE Output Sheets](#6-ase-output-sheets)
   - 6.1 Quarterly Output
   - 6.2 Project Level Macro Paste
7. [ASE HoldCo Sheets](#7-ase-holdco-sheets)
8. [ASE Presentation Sheets](#8-ase-presentation-sheets)
9. [GTC Sheets (21 sheets)](#9-gtc-sheets-21-sheets)
10. [Cross-Sheet Reference Map](#10-cross-sheet-reference-map)
11. [Structural Fingerprint](#11-structural-fingerprint)

---

## 1. Sheet Inventory

| # | Sheet Name | Origin | Rows | Cols | Category |
|---|-----------|--------|------|------|----------|
| 1 | Project Info | ASE | 595 | 233 (I-DX) | Input |
| 2 | Country Inputs | ASE | 138 | 11 | Input |
| 3 | Financing Inputs | ASE | 77 | 325 | Input |
| 4 | Time Inputs (A) | ASE | 653 | 325 | Input |
| 5 | Time Inputs (M) | ASE | 1581 | 634 | Input |
| 6 | Time Inputs (Q) | ASE | 254 | 336 | Input |
| 7 | Sensis | ASE | 140 | 56 | Scenario |
| 8 | Project Level Workings | ASE | 1609 | 16384 | Calc Engine |
| 9 | Quarterly Output | ASE | 35999 | 433 | Output |
| 10 | Project Level Macro Paste | ASE | 654 | 424 | Output |
| 11 | HoldCo CFs & Valuation | ASE | ~250 | ~200 | HoldCo |
| 12 | HoldCo income | ASE | ~100 | ~200 | HoldCo |
| 13 | HoldCo_Facility | ASE | ~80 | ~200 | HoldCo |
| 14 | Dashboard | ASE | ~200 | ~30 | Presentation |
| 15 | ProjectSummary | ASE | ~100 | ~30 | Presentation |
| 16 | HoldCo_Summary | ASE | ~60 | ~30 | Presentation |
| 17 | Charts | ASE | ~50 | ~20 | Presentation |
| 18 | Checks | ASE | ~30 | ~20 | Presentation |
| 19 | FX | ASE | ~10 | ~10 | Input (helper) |
| 20 | Lists | ASE | ~50 | ~10 | Input (helper) |
| 21 | Scenarios | ASE | ~30 | ~10 | Input (helper) |
| 22 | Instruction tab | GTC | ~40 | 16 | GTC |
| 23 | FY26B PL Mapping | GTC | ~100 | 6 | GTC Mapping |
| 24 | FY26B BS Mapping | GTC | ~80 | 6 | GTC Mapping |
| 25 | PnL projection - aggregate | GTC | ~100 | 16 | GTC Fin Stmt |
| 26 | PnL projection - by asset | GTC | ~100 | 16 | GTC Fin Stmt |
| 27 | BS projection - aggregate | GTC | ~60 | 16 | GTC Fin Stmt |
| 28 | BS projection - by asset | GTC | ~60 | 16 | GTC Fin Stmt |
| 29 | CF Capital Strategy | GTC | ~120 | 16 | GTC Fin Stmt |
| 30 | FY26A-29F PL Breakdown | GTC | ~100 | 16 | GTC Detail |
| 31 | FY26A-29F BS Breakdown | GTC | ~80 | 16 | GTC Detail |
| 32 | Valuation (2) | GTC | ~60 | 10 | GTC Valuation |
| 33 | Dev Pipeline | GTC | ~120 | 16 | GTC Pipeline |
| 34 | OpCo Performance | GTC | ~200 | 16 | GTC Perf |
| 35 | Capacity MW tracker | GTC | ~100 | 16 | GTC Tracker |
| 36 | GTC Summary | GTC | 336 | 9 | GTC Combined |
| 37 | QREP Consolidated | GTC | ~400 | 16 | GTC Reporting |
| 38-45 | Additional GTC/helper sheets | GTC | varies | varies | GTC |

---

## 2. Named Ranges (67 relevant)

### Core Navigation

| Named Range | Reference | Type | Purpose |
|------------|-----------|------|---------|
| `ProjectList` | `'Project Info'!$I$5:$DX$5` | text[] | Master list of all project names across columns |
| `Project_View` | `'Project Level Workings'!$H$15` | text | Currently selected project for PLW calcs |
| `modelStartDate` | `'Project Info'!$F$569` | date | Model start date anchor |
| `mths_yr` | `'Project Info'!$F$570` | integer | Months per year (12) |
| `Live_case` | `Sensis!$F$6` | text | Active sensitivity scenario name |
| `P50_P90` | `'Project Info'!$E$83` | text | Production case selector (P50/P90) |

### Sizing & Debt Control

| Named Range | Reference | Type | Purpose |
|------------|-----------|------|---------|
| `Sizing_Active` | `'Project Level Workings'!$H$23` | text | Sizing case active flag |
| `DebtApplicable` | `'Project Level Workings'!$H$26` | T/F | Whether project includes senior debt |
| `Gearing_Cap` | `'Project Level Workings'!$H$979` | % | Maximum gearing cap |
| `Gearing_Senior` | `'Project Level Workings'!$H$9` | % | Senior gearing ratio |
| `Gearing_total` | `'Project Level Workings'!$H$10` | % | Total gearing ratio (senior + SHL) |

### IRR & Valuation

| Named Range | Reference | Type | Purpose |
|------------|-----------|------|---------|
| `Equity_Case` | `'Project Level Workings'!$H$22` | text | Equity case power price scenario |
| `Equity_IRR` | `'Project Level Workings'!$H$13` | % | Equity IRR result |

### FX & Macro

| Named Range | Reference | Type | Purpose |
|------------|-----------|------|---------|
| `EUR_fx` | `'Time Inputs (A)'!` | curve | EUR exchange rate time series |
| `GBP_fx` | `'Time Inputs (A)'!` | curve | GBP exchange rate time series |

### Consolidation

| Named Range | Reference | Type | Purpose |
|------------|-----------|------|---------|
| `SPV_ConsolidatedCashflows` | `'Project Level Workings'!$F$1428:$PQ$1607` | range | Full quarterly consolidation paste area |

### Debt & Tax Named Ranges

| Named Range | Approximate Reference | Purpose |
|------------|----------------------|---------|
| `SeniorDebt_OB` | PLW R983 | Senior debt opening balance row |
| `SeniorDebt_Drawdown` | PLW R990 | Senior debt drawdown row |
| `SeniorDebt_CB` | PLW R993 | Senior debt closing balance row |
| `CFADS` | PLW R958 | Cash flow available for debt service |
| `DSCR` | PLW R1033 | Debt service coverage ratio |
| `LLCR` | PLW R1355 | Loan life coverage ratio |
| `TaxRate_1` | PLW R1144 | Corporate tax rate 1 |
| `TaxRate_2` | PLW R1145 | Corporate tax rate 2 |
| `TaxLoss_Balance` | PLW R1122 | Tax loss carry-forward balance |
| `SHL_OB` | PLW R1221 | SHL opening balance |
| `SHL_CB` | PLW R1234 | SHL closing balance |
| `Distribution_Flag` | PLW R1358 | Distribution flag row |
| `DSRA_OB` | PLW R1047 | DSRA opening balance |
| `DSRA_CB` | PLW R1052 | DSRA closing balance |
| `MRA_OB` | PLW R1062 | MRA opening balance |
| `MRA_CB` | PLW R1066 | MRA closing balance |
| `ProjectIRR` | PLW R1397 | XIRR result for project |
| `EquityIRR_Dist` | PLW R1409 | XIRR result for equity (distributions) |
| `EquityIRR_CFE` | PLW R1418 | XIRR result for equity (CF for equity) |
| `CapEx_Total` | PLW R245 | Total CapEx row |
| `TotalRevenue` | PLW R593 | Total project revenue |
| `TotalOpEx` | PLW R710 | Total OpEx (Solar+BESS) |

---

## 3. ASE Input Sheets

### 3.1 Project Info (595 rows x 233 cols)

**Grid**: Rows 1-595, Columns A-DX (up to 128 project columns starting at Col I).
Row 5 = project name headers (`ProjectList` named range).
Row 6 = country. Row 7 = technology.

#### 3.1.1 Timings (R18-R44)

| Row | Label | Data Type |
|-----|-------|-----------|
| R18 | Construction Start - Solar | date |
| R19 | Construction End - Solar (COD) | date |
| R20 | Construction Start - BESS | date |
| R21 | Construction End - BESS (COD) | date |
| R22 | Senior Debt First Drawdown | date |
| R23 | Senior Debt Final Drawdown | date |
| R24 | Senior Debt Repayment Start | date |
| R25 | Senior Debt Final Repayment | date |
| R26 | Solar Asset Life End | date |
| R27 | BESS Asset Life End | date |
| R28-R30 | PPA/CfD contract start/end dates | date |
| R31-R35 | EEG/SDE contract dates | date |
| R36-R40 | Tolling/Floor contract dates | date |
| R41-R44 | GOO/REGO contract dates, spare timings | date |

#### 3.1.2 BESS (R47-R57)

| Row | Label | Data Type |
|-----|-------|-----------|
| R47 | BESS technology type | text |
| R48 | BESS nameplate capacity (MW) | MW |
| R49 | BESS duration (hours) | hours |
| R50 | BESS efficiency | % |
| R51 | BESS availability | % |
| R52 | BESS degradation rate | %/yr |
| R53 | BESS repower flag | T/F |
| R54 | BESS repower date | date |
| R55 | BESS repower cost | LC'000 |
| R56 | BESS repower capacity ratio | % |
| R57 | BESS repower construction period | months |

#### 3.1.3 Asset Sales (R59-R75)

| Row | Label | Data Type |
|-----|-------|-----------|
| R59 | Sale flag | T/F |
| R60 | Sale date | date |
| R61 | Sale price | LC'000 |
| R62-R65 | Buyer info, tax treatment | text/% |
| R66-R75 | Asset revaluation, land sale, WHT on sale | LC'000/% |

#### 3.1.4 Production (R77-R115)

| Row | Label | Data Type |
|-----|-------|-----------|
| R77 | Solar capacity (MWp) | MW |
| R78 | Solar capacity (MWac) | MW |
| R79 | Specific yield (kWh/kWp) | kWh/kWp |
| R80 | Gross production (GWh) | GWh |
| R81 | Technical availability | % |
| R82 | Degradation rate | %/yr |
| R83 | P50/P90 selector | text |
| R84 | P50 yield | kWh/kWp |
| R85 | P90 yield | kWh/kWp |
| R86 | Net production (P50) | GWh |
| R87 | Net production (P90) | GWh |
| R88-R100 | Seasonality profiles (monthly shape factors) | % |
| R101-R115 | Production curve overrides, degradation lag, capacity factors | various |

#### 3.1.5 Construction (R117-R170)

| Row | Label | Data Type |
|-----|-------|-----------|
| R117-R130 | EPC costs - Solar (total, monthly profile) | LC'000 |
| R131-R140 | Grid connection costs | LC'000 |
| R141-R150 | Other CapEx items | LC'000 |
| R151-R155 | Devex items | LC'000 |
| R156-R160 | Acquisition costs | LC'000 |
| R161-R165 | BESS CapEx (EPC, others) | LC'000 |
| R166-R170 | Land purchase costs | LC'000 |

#### 3.1.6 Revenues (R171-R341)

| Row | Label | Data Type |
|-----|-------|-----------|
| R171-R180 | PPA Contract 1 parameters | various |
| R181 | PPA1 type (Fixed/Floor/CfD/EEG) | text |
| R182 | PPA1 strike price (real) | LC/MWh |
| R183-R190 | PPA1 indexation, balancing, start/end | various |
| R191-R210 | PPA Contract 2 parameters | various |
| R211-R230 | CfD parameters (AR, strike, reference price) | various |
| R231-R250 | EEG parameters (tariff, duration) | various |
| R251-R270 | SDE++ parameters (tariff, correction price) | various |
| R271-R290 | GOO/REGO contract parameters | various |
| R291-R310 | Merchant price assumptions | various |
| R311-R325 | BESS Tolling parameters | various |
| R326-R335 | BESS Floor parameters | various |
| R336-R341 | BESS Capacity Market / CIDT parameters | various |

#### 3.1.7 OpEx (R343-R448)

| Row | Label | Data Type |
|-----|-------|-----------|
| R343-R360 | O&M Solar (initial, step-down, per MW) | LC/MW/yr |
| R361-R380 | Land lease (fixed tiers Year 1-5, 6-10, ...31+) | LC'000/yr |
| R381-R390 | Revenue-dependent lease (% tiers) | % |
| R391-R395 | Insurance on plant & machinery | LC/MW/yr |
| R396-R400 | Inverter warranty cost | LC/MW/yr |
| R401-R410 | Asset management fee | LC'000/yr |
| R411-R420 | Other OpEx items, production-based OpEx | various |
| R421-R430 | BESS O&M | LC/MW/yr |
| R431-R440 | BESS specific OpEx items | LC'000 |
| R441-R448 | LC costs, OpEx spare lines | various |

#### 3.1.8 Working Capital (R450-R453)

| Row | Label | Data Type |
|-----|-------|-----------|
| R450 | Receivable days - contracted revenue | days |
| R451 | Receivable days - merchant revenue | days |
| R452 | Payable days | days |
| R453 | Working capital spare | days |

#### 3.1.9 Tax (R455-R481)

| Row | Label | Data Type |
|-----|-------|-----------|
| R455 | Corporation tax rate | % |
| R456 | Local trade tax rate | % |
| R457 | Solidarity surcharge | % |
| R458-R465 | Interest deductibility limits | %/LC'000 |
| R466-R475 | Tax depreciation rates (per asset class) | %/yr |
| R476-R481 | WHT rates, tax loss use thresholds | % |

#### 3.1.10 Macro (R482-R489)

| Row | Label | Data Type |
|-----|-------|-----------|
| R482 | CPI - UK | % |
| R483 | CPI - Germany | % |
| R484 | CPI - Netherlands | % |
| R485 | HICP - Europe | % |
| R486-R489 | Spare inflation / macro inputs | % |

#### 3.1.11 Financing Terms (R490-R509)

| Row | Label | Data Type |
|-----|-------|-----------|
| R490 | Financing termsheet selector | text |
| R491 | Senior debt margin - Period 1 | bps |
| R492 | Senior debt margin - Period 2 | bps |
| R493 | Senior debt margin switch date | date |
| R494 | Base rate reference (EURIBOR/SONIA) | text |
| R495 | Hedge rate | % |
| R496 | Hedging profile during availability | % |
| R497 | Hedging profile post availability | % |
| R498 | Arrangement fee | % |
| R499 | Commitment fee | % |
| R500 | Agency fee | LC'000/yr |
| R501-R505 | DSRA/DSRF/MRA sizing parameters | months/% |
| R506-R509 | Cash sweep parameters, balloon test | %/T/F |

#### 3.1.12 VAT (R511-R544)

| Row | Label | Data Type |
|-----|-------|-----------|
| R511 | VAT applicable flag | T/F |
| R512 | VAT rate | % |
| R513-R520 | VAT on CapEx items (EPC, grid, other, devex) | T/F |
| R521-R525 | VAT on BESS CapEx items | T/F |
| R526-R530 | VAT loan quantum | LC'000 |
| R531-R535 | VAT loan interest rate / commitment fee | % |
| R536-R540 | VAT refund timing (days/months) | days |
| R541-R544 | VAT on revenue/OpEx flags | T/F |

#### 3.1.13 D&A Rates (R546-R562)

| Row | Label | Data Type |
|-----|-------|-----------|
| R546 | Number of D&A asset classes | integer |
| R547 | D&A method (SL/DB) | text |
| R548 | D&A class 1 name | text |
| R549 | D&A class 1 rate | %/yr |
| R550 | D&A class 2 name | text |
| R551 | D&A class 2 rate | %/yr |
| R552-R555 | D&A classes 3-4 (name + rate) | text/% |
| R556-R560 | D&A classes 5-6 (BESS, financing assets) | text/% |
| R561-R562 | Tax depreciation rate, accelerated depreciation | % |

#### 3.1.14 Other (R568-R594)

| Row | Label | Data Type |
|-----|-------|-----------|
| R568 | Model version | text |
| R569 | Model start date (`modelStartDate`) | date |
| R570 | Months per year (`mths_yr`) | integer |
| R571 | Model end date | date |
| R572 | Quarterly periods | integer |
| R573-R580 | Discount rates (WACC, project, equity) | % |
| R581-R585 | LC/Bond parameters | LC'000/% |
| R586-R590 | SHL interest rate, WHT on SHL interest | % |
| R591-R594 | Revaluation, 11kV revenue, spare parameters | various |

---

### 3.2 Country Inputs (138 rows x 11 cols)

**Grid**: Rows 1-138. Columns A-K. Rows organized by country (UK/DE/NL).

#### 3.2.1 Tax (R6-R31)

| Row | Label | Data Type |
|-----|-------|-----------|
| R6 | Corporation tax rate - UK | % |
| R7 | Corporation tax rate - NL (lower) | % |
| R8 | Corporation tax rate - NL (upper) | % |
| R9 | NL tax threshold | LC'000 |
| R10 | Corporation tax rate - DE | % |
| R11 | Solidarity surcharge - DE | % |
| R12 | Trade tax rate - DE | % |
| R13-R18 | Interest deductibility caps per country | %/LC'000 |
| R19-R25 | Tax loss carry-forward rules per country | %/T/F |
| R26-R31 | WHT rates, tax payment timing | %/months |

#### 3.2.2 SHL (R33-R40)

| Row | Label | Data Type |
|-----|-------|-----------|
| R33 | SHL interest rate - UK | % |
| R34 | SHL interest rate - DE | % |
| R35 | SHL interest rate - NL | % |
| R36-R38 | WHT on SHL interest per country | % |
| R39-R40 | SHL repayment rules | text |

#### 3.2.3 VAT (R43-R58)

| Row | Label | Data Type |
|-----|-------|-----------|
| R43 | VAT rate - UK | % |
| R44 | VAT rate - DE | % |
| R45 | VAT rate - NL | % |
| R46-R50 | VAT on CapEx (EPC, grid, other) per country | T/F |
| R51-R55 | VAT on devex per country | T/F |
| R56-R58 | VAT refund timing per country | months |

#### 3.2.4 O&M (R63-R91)

| Row | Label | Data Type |
|-----|-------|-----------|
| R63-R65 | O&M cost base per country (Solar) | LC/MW/yr |
| R66-R68 | O&M step-down rates | % |
| R69-R75 | Insurance rates per country | LC/MW/yr |
| R76-R80 | Asset management fees per country | LC'000/yr |
| R81-R85 | Land lease base rates per country | LC/MW/yr |
| R86-R88 | BESS O&M per country | LC/MW/yr |
| R89-R91 | Other OpEx per country | LC/MW/yr |

#### 3.2.5 D&A (R93-R110)

| Row | Label | Data Type |
|-----|-------|-----------|
| R93-R98 | Accounting D&A rates per asset class per country | %/yr |
| R99-R105 | Tax D&A rates per country | %/yr |
| R106-R110 | Accelerated depreciation flags per country | T/F/% |

#### 3.2.6 Misc (R112-R115)

| Row | Label | Data Type |
|-----|-------|-----------|
| R112 | LCOE discount rate per country | % |
| R113 | Working capital days per country | days |
| R114-R115 | Spare | -- |

#### 3.2.7 Seasonality (R117-R135)

| Row | Label | Data Type |
|-----|-------|-----------|
| R117-R128 | Monthly seasonality factors (Jan-Dec) per country | % |
| R129-R132 | Quarterly seasonality factors (Q1-Q4) | % |
| R133-R135 | Spare seasonality rows | % |

---

### 3.3 Financing Inputs (77 rows x 325 cols)

**Grid**: Rows 1-77. Columns span 9 termsheets:
- UK Solar, UK BESS, UK Combined
- DE Solar, DE BESS, DE Combined
- NL Solar, NL BESS, NL Combined

Each termsheet occupies ~35 columns.

#### 3.3.1 VAT Loan (R10-R13)

| Row | Label | Data Type |
|-----|-------|-----------|
| R10 | VAT Loan quantum | LC'000 |
| R11 | VAT Loan interest rate | % |
| R12 | VAT Loan commitment fee | % |
| R13 | VAT Loan arrangement fee | % |

#### 3.3.2 Senior Debt (R15-R48)

| Row | Label | Data Type |
|-----|-------|-----------|
| R15 | Senior debt quantum | LC'000 |
| R16 | Gearing cap | % |
| R17 | Sizing DSCR - contracted | x |
| R18 | Sizing DSCR - non-contracted | x |
| R19-R20 | Sizing DSCR - BESS contracted/non-contracted | x |
| R21 | Debt tenor (years) | years |
| R22 | Grace period (months) | months |
| R23 | Repayment type (sculpted/annuity) | text |
| R24-R25 | Cash sweep percentages | % |
| R26-R27 | Margin Period 1 / Period 2 | bps |
| R28 | Margin switch date | date |
| R29 | Base rate reference | text |
| R30-R32 | Hedge rate, hedging profiles | % |
| R33-R35 | Arrangement fee, commitment fee, agency fee | %/LC'000 |
| R36-R38 | DSRA months, DSRF parameters | months/% |
| R39-R42 | MRA parameters | months/LC'000 |
| R43-R45 | Balloon test parameters | %/T/F |
| R46-R48 | Distribution lock-up DSCR, LLCR covenants | x |

#### 3.3.3 Securities (R70-R72)

| Row | Label | Data Type |
|-----|-------|-----------|
| R70 | LC amount | LC'000 |
| R71 | LC commitment fee | % |
| R72 | LC tenor | years |

---

### 3.4 Time Inputs (A) -- Annual (653 rows x 325 cols)

**Grid**: Row-based curves, one value per annual period across columns.

#### 3.4.1 Inflation Profiles (R8-R44)

| Row | Label | Data Type |
|-----|-------|-----------|
| R10 | CPI Curve Indexation header | -- |
| R11 | UK CPI | % curve |
| R12 | Germany CPI | % curve |
| R13 | The Netherlands CPI | % curve |
| R14 | Country-specific CPI (linked to PLW R18) | % curve |
| R16-R20 | BESS Curve Indexation CPI (UK/DE/NL) | % curve |
| R23 | Indexation profile header | -- |
| R24 | NIL (no indexation) | flat |
| R25 | CPI - PV Curve Indexation | % curve |
| R26 | CPI - BESS Curve Indexation | % curve |
| R27 | CPI | % curve |
| R28 | CPI after 2yrs | % curve |
| R29 | Flat 2% | % flat |
| R30 | Flat 2% after 2yrs | % flat |
| R31 | Flat 0.68% | % flat |
| R32 | Land Lease | % curve |
| R33 | Flat 3% | % flat |
| R34-R44 | Spare indexation profiles (spare 2-12) | % |

#### 3.4.2 Financing Curves (R46-R53)

| Row | Label | Data Type |
|-----|-------|-----------|
| R47 | Sensitivity header | -- |
| R48 | EURIBOR (links to Sensis!G43) | % curve |
| R49 | SONIA (links to Sensis!G43) | % curve |
| R50 | 6-month EURIBOR Curve | % curve |
| R51 | 6-month SONIA curve | % curve |
| R52 | 3-month EURIBOR Curve | % curve |
| R53 | 3-month SONIA curve | % curve |

#### 3.4.3 FX Curves (R55-R61)

| Row | Label | Data Type |
|-----|-------|-----------|
| R56-R57 | EUR to USD (current/old) | FX rate curve |
| R59 | Rate to EUR header | -- |
| R60 | EUR (=1) | constant |
| R61 | GBP (=1.15 spot) | FX rate curve |

#### 3.4.4 Market Prices -- Solar (R63-R83)

| Row | Label | Data Type |
|-----|-------|-----------|
| R65 | Starting inflation for Market Advisor A | % |
| R67 | UK Solar section header | -- |
| R68 | UK Solar Market Advisor A High Q3 2025 | LC/MWh curve |
| R69 | UK Solar Market Advisor A Central Q3 2025 | LC/MWh curve |
| R70 | UK Solar Market Advisor A Low Q3 2025 | LC/MWh curve |
| R71 | UK Solar Market Advisor A Central/Low Q3 2025 | LC/MWh curve |
| R73 | DE Solar section header | -- |
| R74 | DE Solar Market Advisor A High Q3 2025 | LC/MWh curve |
| R75 | DE Solar Market Advisor A Central Q3 2025 | LC/MWh curve |
| R76 | DE Solar Market Advisor A Low Q3 2025 | LC/MWh curve |
| R77 | DE Solar Market Advisor A Central/Low Q3 2025 | LC/MWh curve |
| R79 | NL Solar section header | -- |
| R80 | NL Solar Market Advisor A High Q3 2025 | LC/MWh curve |
| R81 | NL Solar Market Advisor A Central Q3 2025 | LC/MWh curve |
| R82 | NL Solar Market Advisor A Low Q3 2025 | LC/MWh curve |
| R83 | NL Solar Market Advisor A Central/Low Q3 2025 | LC/MWh curve |

#### 3.4.5 BESS Revenue Curves (R85-R98)

| Row | Label | Data Type |
|-----|-------|-----------|
| R86 | UK BESS Market Advisor A Central | LC/MW/yr curve |
| R87 | UK BESS Market Advisor A Low | LC/MW/yr curve |
| R88 | UK BESS Market Advisor A Central/Low | LC/MW/yr curve |
| R89 | DE BESS Market Advisor A Central | LC/MW/yr curve |
| R90 | DE BESS Market Advisor A Low | LC/MW/yr curve |
| R91 | DE BESS Market Advisor A Central/Low | LC/MW/yr curve |
| R92 | NL BESS Market Advisor A Central | LC/MW/yr curve |
| R93 | NL BESS Market Advisor A Low | LC/MW/yr curve |
| R94 | NL BESS Market Advisor A Central/Low | LC/MW/yr curve |
| R95 | Breddin - Import - Degraded Curve | LC/MW/yr curve |
| R96 | Elsterheide - Import - Degraded Curve | LC/MW/yr curve |
| R97 | DE BESS CIDT | LC/MW/yr curve |

#### 3.4.6 GoO/REGO Curves (R100-R119)

| Row | Label | Data Type |
|-----|-------|-----------|
| R100-R103 | UK Solar REGOs (Pessimistic, Central, zero) | LC/MWh curve |
| R104 | Glowen specific curve | LC/MWh curve |
| R106-R112 | NL Solar GoS curves (Pessimistic, Central, fixed) | LC/MWh curve |
| R114-R119 | DE Solar GoS curves (Pessimistic, Central, fixed) | LC/MWh curve |

#### 3.4.7 UK CfD (R121-R125)

| Row | Label | Data Type |
|-----|-------|-----------|
| R122 | AR6 strike price (71.1) | LC/MWh |
| R123 | COD pre-2030 strike (66.1) | LC/MWh |
| R124 | COD post-2030 strike (61.1) | LC/MWh |
| R125 | Spare CfD | LC/MWh |

#### 3.4.8 Clawback & Negative Price (R127-R137)

| Row | Label | Data Type |
|-----|-------|-----------|
| R129-R130 | Clawback scenario / other | curve |
| R134 | Economic curtailment CfD | % curve |
| R135-R136 | NL volume curtailment (6h / 1h) | % curve |
| R137 | Nil | flat zero |

#### 3.4.9 Balancing Costs (R139-R153)

| Row | Label | Data Type |
|-----|-------|-----------|
| R141-R144 | Germany balancing costs (Merchant, PPA, NIL) | LC/MWh curve |
| R146-R148 | Netherlands balancing costs (Merchant, PPA) | LC/MWh curve |
| R150-R153 | UK balancing costs (Merchant, CfD, PPA) | LC/MWh curve |

#### 3.4.10 Cash Sweep Profile (R155-R157)

| Row | Label | Data Type |
|-----|-------|-----------|
| R157 | Cash sweep profile for collateral | % curve |

#### 3.4.11 Tax Rates (R159-R164)

| Row | Label | Data Type |
|-----|-------|-----------|
| R161 | UK Corporation Tax | % curve |
| R162 | NL Corporation Tax - Lower Threshold | % curve |
| R163 | NL Corporation Tax - Higher Threshold | % curve |
| R164 | DE Corporation Tax (incl. solidarity) | % curve |

#### 3.4.12 Indexation Sources (R166-R200)

| Row | Label | Data Type |
|-----|-------|-----------|
| R172-R184 | GoO/REGO source curves (per country) | LC/MWh |
| R172-R200 | OpEx indexation profiles (per country, flat/CPI) | % curve |

#### 3.4.13 MRA Data (R203-R256)

| Row | Label | Data Type |
|-----|-------|-----------|
| R206 | MRA Project name header | -- |
| R207-R256 | Per-project MRA data (120 projects with capacity, devex, HoldCo refs) | various |

---

### 3.5 Time Inputs (M) -- Monthly (1581 rows x 634 cols)

**Grid**: Row-based monthly time series. One value per month across columns.

#### 3.5.1 Historical Devex (R10-R132)

| Row | Label | Data Type |
|-----|-------|-----------|
| R10 | Section header: Devex - Historical DevEx | -- |
| R11 | Dashboard links | reference |
| R12 | Column headers: Capacity (MW), Devex (LC'000), Early, Land & grid | -- |
| R13-R132 | Per-project rows (up to 120 projects): capacity, total devex, historical devex allocation, HoldCo CFs & Valuation links | MW, LC'000, date |

Each project row contains:
- C3: sequential index
- C4: project name (linked to devex master)
- C5: capacity (MW)
- C6: devex amount (LC'000)
- C7: HoldCo CFs & Valuation link (early devex date)
- C8: EOMONTH reference (land & grid date)

#### 3.5.2 Future Devex (R134-R256)

| Row | Label | Data Type |
|-----|-------|-----------|
| R134 | Section header: Devex - future DevEx | -- |
| R135 | Column headers: Allocated to, date | -- |
| R136 | Headers: Total Devex, First devex, Last devex (RtB), Total DevEx, Historical DevEx | -- |
| R137-R256 | Per-project rows (120 projects): future devex allocation by month | LC'000/month |

#### 3.5.3 Acquisition Cost (R258-R380)

| Row | Label | Data Type |
|-----|-------|-----------|
| R258 | Section header: Devex - Acquisition cost | -- |
| R260 | Column headers | -- |
| R261-R380 | Per-project acquisition cost allocations by month | LC'000/month |

#### 3.5.4 CapEx -- Solar (R382-R632)

| Row | Label | Data Type |
|-----|-------|-----------|
| R382 | Section header: Capex | -- |
| R384 | Sub-header: Solar | -- |
| R386 | Headers: Capex - EPC, First CAPEX, Last CAPEX | -- |
| R387-R510 | Per-project solar EPC CapEx by month | LC'000/month |
| R512-R632 | Grid connection CapEx, Other CapEx by month | LC'000/month |

#### 3.5.5 CapEx -- BESS (R634-R760)

| Row | Label | Data Type |
|-----|-------|-----------|
| R634 | Sub-header: BESS | -- |
| R636-R760 | Per-project BESS CapEx by month | LC'000/month |

#### 3.5.6 Remaining Sections (R762-R1581)

| Row Range | Section | Data Type |
|-----------|---------|-----------|
| R762-R890 | CapEx - Grid connection by month | LC'000/month |
| R892-R1020 | CapEx - Other items by month | LC'000/month |
| R1022-R1150 | BESS CapEx by month | LC'000/month |
| R1152-R1280 | BESS Repower CapEx by month | LC'000/month |
| R1282-R1410 | Land purchase CapEx by month | LC'000/month |
| R1412-R1500 | Construction cost S-curves / profiles | % |
| R1502-R1581 | Spare / reserved rows | -- |

---

### 3.6 Time Inputs (Q) -- Quarterly (254 rows x 336 cols)

**Grid**: Quarterly time series. Mirrors key Time Inputs (A) curves but at quarterly granularity.

| Row Range | Section | Data Type |
|-----------|---------|-----------|
| R1-R5 | Headers, date timeline | date |
| R6-R20 | Quarterly CPI curves (UK/DE/NL) | % curve |
| R21-R35 | Quarterly EURIBOR / SONIA curves | % curve |
| R36-R50 | Quarterly FX curves | FX rate |
| R51-R80 | Quarterly solar power price curves (UK/DE/NL x High/Central/Low) | LC/MWh |
| R81-R110 | Quarterly BESS revenue curves | LC/MW/yr |
| R111-R140 | Quarterly GoO/REGO curves | LC/MWh |
| R141-R170 | Quarterly CfD / EEG / SDE curves | LC/MWh |
| R171-R200 | Quarterly balancing cost curves | LC/MWh |
| R201-R220 | Quarterly negative price curves | % |
| R221-R240 | Quarterly tax rate curves | % |
| R241-R254 | Spare / consolidation curves | various |

---

## 4. ASE Scenario Control

### 4.1 Sensis (140 rows x 56 cols)

**Grid**: 24 scenario slots across columns. Row labels define lever names.

#### 4.1.1 Control (R1-R13)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1 | Sheet title | text |
| R5 | Scenario names (24 slots in columns) | text |
| R6 | Live case selector (`Live_case`) | text |
| R7-R10 | Scenario descriptions | text |
| R11-R13 | Scenario active flags | T/F |

#### 4.1.2 Operations Levers (R15-R19)

| Row | Label | Data Type |
|-----|-------|-----------|
| R15 | O&M sensitivity | % |
| R16 | Insurance sensitivity | % |
| R17 | Land lease sensitivity | % |
| R18 | Other OpEx sensitivity | % |
| R19 | Spare OpEx lever | % |

#### 4.1.3 Production Levers (R21-R25)

| Row | Label | Data Type |
|-----|-------|-----------|
| R21 | Solar degradation sensitivity | % |
| R22 | Solar availability sensitivity | % |
| R23 | BESS degradation sensitivity | % |
| R24 | BESS availability sensitivity | % |
| R25 | Production spare lever | % |

#### 4.1.4 Uncontracted Revenue Levers (R27-R35)

| Row | Label | Data Type |
|-----|-------|-----------|
| R27 | Solar merchant price curve selector | text |
| R28 | Solar merchant price sensitivity | % |
| R29 | BESS merchant price curve selector | text |
| R30 | BESS merchant price sensitivity | % |
| R31 | GoO/REGO price curve selector | text |
| R32 | GoO/REGO price sensitivity | % |
| R33-R35 | Spare revenue levers | various |

#### 4.1.5 Contracted Revenue Levers (R37-R44)

| Row | Label | Data Type |
|-----|-------|-----------|
| R37 | PPA price sensitivity | % |
| R38 | CfD price sensitivity | % |
| R39 | EEG price sensitivity | % |
| R40 | SDE++ price sensitivity | % |
| R41 | Tolling price sensitivity | % |
| R42 | Floor price sensitivity | % |
| R43 | EURIBOR/SONIA sensitivity | bps |
| R44 | Spare contracted lever | various |

#### 4.1.6 Termsheet Levers (R46-R49)

| Row | Label | Data Type |
|-----|-------|-----------|
| R46 | DSCR contracted sensitivity | x |
| R47 | DSCR non-contracted sensitivity | x |
| R48 | Debt tenor sensitivity | years |
| R49 | Spare termsheet lever | various |

#### 4.1.7 Asset Life Levers (R51-R54)

| Row | Label | Data Type |
|-----|-------|-----------|
| R51 | Solar asset life sensitivity | years |
| R52 | BESS asset life sensitivity | years |
| R53 | Repower timing sensitivity | years |
| R54 | Spare asset life lever | various |

#### 4.1.8 Discount Rate (R56-R57)

| Row | Label | Data Type |
|-----|-------|-----------|
| R56 | Discount rate sensitivity | % |
| R57 | Spare discount rate lever | % |

#### 4.1.9 Results Matrix (R60-R138)

| Row | Label | Data Type |
|-----|-------|-----------|
| R60 | Results section header | -- |
| R62 | Project IRR | % |
| R63 | Equity IRR (Distributions) | % |
| R64 | Equity IRR (CF for Equity) | % |
| R66 | Senior Gearing | % |
| R67 | Total Gearing | % |
| R68 | Senior Debt Quantum | LC'000 |
| R70-R80 | Per-project IRRs | % |
| R82-R100 | Per-project gearing results | % |
| R102-R120 | Per-project debt quanta | LC'000 |
| R122-R138 | Consolidated portfolio results | various |

---

## 5. ASE Calculation Engine

### 5.1 Project Level Workings (1609 rows x 16384 cols)

**Grid**: Row labels in Cols A-G (C5=section header, C6=subsection header, C7=line item).
Time series data from Col H onwards (one column per period, monthly granularity).
Row 15 selects the active project via `Project_View`.

---

#### 5.1.1 Headers & Checks (R1-R28)

| Row | Col | Label | Data Type |
|-----|-----|-------|-----------|
| R1 | C2 | Project Parthenon | text |
| R2 | C2 | [ARRAY_FORMULA] | formula |
| R4 | C7 | BS check | T/F |
| R5 | C7 | Statements check | T/F |
| R6 | C7 | SHL repaid | T/F |
| R7 | C7 | Senior Debt optimal | T/F |
| R8 | C7 | Master check | T/F |
| R9 | C7 | Gearing (`Gearing_Senior`) | % |
| R10 | C7 | Gearing total (`Gearing_total`) | % |
| R11 | C7 | Active | T/F |
| R12 | C7 | Project IRR | % |
| R13 | C7 | Equity IRR | % |
| R14 | C5 | View Selection | header |
| R15 | C7 | Select Project (`Project_View`) | text |
| R16 | C7 | Technology | text |
| R17 | C7 | Financing Termsheet | text |
| R18 | C7 | Country | text |
| R19 | C7 | Solar Power Price - Sizing Case | text |
| R20 | C7 | Solar Power Price - Equity Case | text |
| R21 | C7 | BESS Power Price - Sizing Case | text |
| R22 | C7 | BESS Power Price - Equity Case | text |
| R23 | C7 | Sizing Case (`Sizing_Active`) | text |
| R24 | C7 | Set Scenario | text |
| R25 | C7 | Tolerance | number |
| R26 | C7 | Include in senior financing? (`DebtApplicable`) | T/F |
| R27 | C7 | Refi asset | T/F |
| R28 | C7 | BESS repower | T/F |

---

#### 5.1.2 Financial Statements (R30-R134)

##### Income Statement (R32-R55)

| Row | Label | Data Type |
|-----|-------|-----------|
| R32 | **Income statement** (section header) | -- |
| R34 | Revenue | LC'000 |
| R35 | Contracted revenue | LC'000 |
| R36 | Non-contracted revenue | LC'000 |
| R37 | OpEx | LC'000 |
| R38 | EBITDA | LC'000 |
| R39 | D&A | LC'000 |
| R40 | EBIT | LC'000 |
| R41 | Interest | LC'000 |
| R42 | Agency fees | LC'000 |
| R43 | Reserve Facilities Commitment Fees | LC'000 |
| R44 | PBT | LC'000 |
| R46 | Taxes | LC'000 |
| R47 | Net Income | LC'000 |
| R49 | Retained earnings OB | LC'000 |
| R50 | Net Income | LC'000 |
| R51 | Revaluation | LC'000 |
| R52 | Dividend | LC'000 |
| R53 | Special Dividend | LC'000 |
| R54 | WHT | LC'000 |
| R55 | Retained earnings CB | LC'000 |

##### Cash Flow Statement (R57-R95)

| Row | Label | Data Type |
|-----|-------|-----------|
| R57 | **Cashflow statement** (section header) | -- |
| R59 | Revenue received | LC'000 |
| R60 | OpEx paid | LC'000 |
| R61 | VAT movement not covered by VAT facility | LC'000 |
| R62 | Tax | LC'000 |
| R63 | CFO | LC'000 |
| R65 | CapEx | LC'000 |
| R66 | CFI | LC'000 |
| R68 | VAT Loan - Commitment Fee | LC'000 |
| R69 | VAT Loan - Interest | LC'000 |
| R70 | Agency fees | LC'000 |
| R71 | Bonds and guarantees interests & fees | LC'000 |
| R73 | Senior Debt - Drawdown | LC'000 |
| R74 | Senior Debt - Scheduled Principal | LC'000 |
| R75 | Senior Debt - Cash sweep | LC'000 |
| R76 | Senior Debt - Interest | LC'000 |
| R77 | Senior Debt - Fees | LC'000 |
| R78 | SHL - Drawdown | LC'000 |
| R79 | SHL - Principal | LC'000 |
| R80 | SHL - Interest | LC'000 |
| R81 | Reserve Facilities Commitment Fees - Paid | LC'000 |
| R82 | Movement in DSRA | LC'000 |
| R83 | Movement in DSRF | LC'000 |
| R84 | Movement in MRA | LC'000 |
| R85 | Movement in Repower CapEx funding account | LC'000 |
| R86 | Common stock - New issuance | LC'000 |
| R87 | Common stock - Buy back | LC'000 |
| R88 | Dividend | LC'000 |
| R89 | Special Dividend | LC'000 |
| R90 | WHT | LC'000 |
| R91 | CFF | LC'000 |
| R93 | Cash OB | LC'000 |
| R94 | Mvment in period | LC'000 |
| R95 | Cash CB | LC'000 |
| R97 | Negative cash check | T/F |
| R98 | Trapped cash | LC'000 |

##### Balance Sheet (R100-R134)

| Row | Label | Data Type |
|-----|-------|-----------|
| R100 | **Balance sheet** (section header) | -- |
| R102 | ASSETS | header |
| R103 | Receivables | LC'000 |
| R104 | VAT receivables / (payables) net of VAT facility | LC'000 |
| R105 | Cash & Equivalents | LC'000 |
| R106 | DSRA | LC'000 |
| R107 | MRA | LC'000 |
| R108 | Total Current Assets | LC'000 |
| R110 | Fixed Assets | LC'000 |
| R111 | Land | LC'000 |
| R112 | Financial Assets | LC'000 |
| R113 | Revaluation of Fixed Assets | LC'000 |
| R114 | CapEx from Repower account | LC'000 |
| R115 | Total Assets | LC'000 |
| R117 | LIABILITIES | header |
| R118 | Payables | LC'000 |
| R119 | Interest Payable - Senior Debt | LC'000 |
| R120 | Interest Payable - SHL | LC'000 |
| R121 | Total Current Liabilities | LC'000 |
| R123 | Senior Debt | LC'000 |
| R124 | DSRF | LC'000 |
| R125 | SHL | LC'000 |
| R126 | Total Liabilities | LC'000 |
| R128 | SHAREHOLDERS EQUITY | header |
| R129 | Common Stock | LC'000 |
| R130 | Retained Earnings | LC'000 |
| R131 | Total Shareholders Equity | LC'000 |
| R133 | Balance sheet check | T/F |
| R134 | Delta | LC'000 |

---

#### 5.1.3 Flags (R136-R214)

| Row | Label | Data Type |
|-----|-------|-----------|
| R136 | **Project workings** (section header) | -- |
| R138 | **Flags** (sub-header) | -- |
| R139 | Success factor | T/F |
| R140 | Solar ON/OFF | T/F |
| R141 | BESS ON/OFF | T/F |
| R143 | **Inflation flags** (sub-header) | -- |
| R146-R165 | Inflation flag rows (20 rows, ref 'Time Inputs (A)'!D24-D43) | index |
| R168-R187 | Mirror inflation flags (=G146 pattern) | index |
| R189 | **Project start / end flags** (sub-header) | -- |
| R190 | FC - Solar | date |
| R191 | FC - BESS | date |
| R192 | Consolidated FC (equity) | date |
| R193 | Consolidated FC (senior) | date |
| R194 | Construction Flag - Solar | 0/1 flag |
| R195 | Construction Flag - BESS | 0/1 flag |
| R196 | Consolidated Funding Period Flag (equity) | 0/1 flag |
| R197 | Consolidated Availability Period Flag (senior) | 0/1 flag |
| R198 | Consolidated Funding Period Flag (senior) | 0/1 flag |
| R199 | Funding Period Flag - Solar | 0/1 flag |
| R200 | Funding Period Flag - BESS | 0/1 flag |
| R201 | Consolidated Funding Period Flag | 0/1 flag |
| R202 | Existing debt profile - permissible drawdown | 0/1 flag |
| R204 | Operations Flag - Solar | 0/1 flag |
| R205 | Post repower period | 0/1 flag |
| R206 | Operations Flag - BESS | 0/1 flag |
| R207 | Consolidated Operations Flag | 0/1 flag |
| R209 | BESS Floor Flag | 0/1 flag |
| R210 | BESS Tolling 1 Flag | 0/1 flag |
| R212 | Operations year | integer |
| R213 | Period length indicator | months |
| R214 | Last draw passed | T/F |

---

#### 5.1.4 CapEx (R216-R245)

| Row | Label | Data Type |
|-----|-------|-----------|
| R216 | **CapEx** (section header) | -- |
| R218 | ='Country Inputs'!C45 (VAT on CapEx flag) | reference |
| R219 | CapEx - EPC | LC'000 |
| R220 | CapEx - Grid | LC'000 |
| R221 | CapEx - Other | LC'000 |
| R222 | Total forecast CapEx | LC'000 |
| R224 | FC asset | LC'000 |
| R225 | Historical CapEx | LC'000 |
| R226 | Total CapEx | LC'000 |
| R228 | ='Country Inputs'!C51 (VAT on devex flag) | reference |
| R229 | Devex | LC'000 |
| R230 | Total forecast devex | LC'000 |
| R232 | Devex - Acquisition costs | LC'000 |
| R233 | Total devex | LC'000 |
| R235 | Total CapEx + devex | LC'000 |
| R237 | BESS (sub-header) | -- |
| R238 | BESS CapEx | LC'000 |
| R239 | BESS CapEx - others | LC'000 |
| R240 | BESS repower | LC'000 |
| R242 | Land purchase | LC'000 |
| R243 | Total BESS CapEx | LC'000 |
| R245 | **Total CapEx** | LC'000 |

---

#### 5.1.5 Production (R247-R272)

| Row | Label | Data Type |
|-----|-------|-----------|
| R247 | **Production** (section header) | -- |
| R249 | Solar (or co-located) sub-header | -- |
| R250 | Operations flag | 0/1 flag |
| R251 | Days from COD | days |
| R253 | Seasonality Monthly | % |
| R254 | Seasonality Quarterly | % |
| R256 | Degradation period lag | months |
| R257 | Capacity | MW |
| R258 | Gross Production | MWh |
| R259 | Availability - Solar | % |
| R260 | Degradation - Solar | % |
| R261 | Net Production - Solar | MWh |
| R263 | BESS sub-header | -- |
| R264 | Capacity | MW |
| R265 | BESS duration | hours |
| R266 | Availability - BESS | % |
| R267 | BESS RePower On | 0/1 flag |
| R268 | BESS repower construction period (months) | months |
| R269 | BESS capacity ratio | % |
| R270 | Degradation - BESS | % |
| R271 | Efficiency | % |
| R272 | Net BESS periodic capacity | MW |

---

#### 5.1.6 Revenues (R274-R622)

##### Generation Waterfall (R276-R310)

| Row | Label | Data Type |
|-----|-------|-----------|
| R276 | **Generation Waterfall** (sub-header) | -- |
| R278 | Energy | -- |
| R279 | Net Production - Solar | MWh |
| R281 | PPA1 Flag | 0/1 flag |
| R282 | PPA1 Volume | MWh |
| R284 | CfD Flag | 0/1 flag |
| R285 | CfD Volume | MWh |
| R287 | Capacity available after PPA & CfD | MWh |
| R288 | GOO Contracts | MWh |
| R289 | Capacity available for EEG | MWh |
| R291 | EEG Flag | 0/1 flag |
| R292 | EEG Volume | MWh |
| R294 | SDE++ Flag | 0/1 flag |
| R295 | Merchant volume protected by SDE++ | MWh |
| R297 | Merchant Capacity | MWh |
| R299 | GOO/REGO | -- |
| R300 | Available Volume after PPA | MWh |
| R302-R303 | GOO Contract 1 Flag / Volume | flag/MWh |
| R306-R307 | GOO Contract 2 Flag / Volume | flag/MWh |
| R310 | Merchant GOO Volume | MWh |

##### Solar Revenue -- PPA (R312-R342)

| Row | Label | Data Type |
|-----|-------|-----------|
| R312 | **Solar Revenue** (sub-header) | -- |
| R313 | **Contracted Revenue Solar - PPA** | -- |
| R315 | PPA Contract - 1 (sub-header) | -- |
| R316 | Type of PPA scheme | text |
| R317 | Contracted volume | MWh |
| R319 | PPA Flag Period 1 | 0/1 flag |
| R321 | Fixed Price PPA | LC/MWh |
| R322 | PPA Scheme | text |
| R323 | On / Off switch | T/F |
| R324 | PPA / EEG switch | text |
| R325 | Contracted Volume | MWh |
| R326 | Negative price correction | MWh |
| R327 | Adjusted Contracted Volume | MWh |
| R329 | UK CfD | -- |
| R330 | Nominal CfD price curve | LC/MWh curve |
| R332 | PPA Price Curve (Real) | LC/MWh curve |
| R333 | PPA / EEG Price (nominal, before balancing cost) | LC/MWh |
| R334 | Balancing cost curve (nominal) | LC/MWh |
| R335 | UK CfD Price (Incl. balancing cost) | LC/MWh |
| R336 | PPA Price (Incl. balancing cost) | LC/MWh |
| R337 | Contracted revenue (Solar) | LC'000 |
| R340 | Total Contracted Volume - Contract 1 | MWh |
| R342 | Total Contracted Revenue - Contract 1 | LC'000 |

##### PPA Contract 2 (R344-R384) -- Spare/Floor PPA/SDE

| Row | Label | Data Type |
|-----|-------|-----------|
| R344 | PPA Contract - 2 (Not used / Spare) | -- |
| R345-R349 | Type, volume, flag, offtake period | various |
| R350-R356 | Floor PPA parameters (price, volume, revenue) | LC/MWh, MWh, LC'000 |
| R357 | Floor PPA Non-Contracted Revenue | LC'000 |
| R360-R369 | SDE scheme (tariff, correction price, revenue) | LC/MWh, LC'000 |
| R371-R382 | Fixed Price PPA alt / EEG switch | various |
| R384 | Total Contracted Volume - Contract 2 | MWh |

##### SDE++ (R388-R423)

| Row | Label | Data Type |
|-----|-------|-----------|
| R388 | **SDE++** (sub-header) | -- |
| R389-R390 | Type of scheme, contracted volume | text, MWh |
| R392-R393 | SDE Flag, Offtake period | 0/1, % |
| R395-R396 | SDE scheme, flag | text, 0/1 |
| R398-R400 | Contracted volume (uncurtailed/curtailed) | MWh |
| R402 | Max eligible volume | MWh |
| R404 | SDE++ Eligible volume | MWh |
| R406-R409 | SDE tariff, correction price, top-up price/revenue | LC/MWh, LC'000 |
| R411-R416 | Merchant volume under SDE++, merchant income | MWh, LC'000 |
| R418-R419 | SDE Contracted/Non-Contracted Revenue | LC'000 |
| R421 | Total Contracted Volume - SDE | MWh |
| R423 | Total Contracted Revenue - SDE | LC'000 |

##### EEG (R425-R447)

| Row | Label | Data Type |
|-----|-------|-----------|
| R425 | **EEG** (sub-header) | -- |
| R426 | Contracted volume | MWh |
| R428 | EEG Flag Period | 0/1 flag |
| R431-R437 | EEG On/Off, offtake, switch, volume, neg price, adjusted volume | various |
| R438-R442 | EEG price, balancing costs, post-balancing, contracted/non-contracted revenue | LC/MWh, LC'000 |
| R445 | Total Contracted Volume - EEG | MWh |
| R447 | Total Contracted Revenue - EEG | LC'000 |

##### Contracted Volume Summary (R452-R458)

| Row | Label | Data Type |
|-----|-------|-----------|
| R452 | **Contracted Volume** (sub-header) | -- |
| R453 | Contracted volume - Contract 1 (PPA) | MWh |
| R454 | Contracted volume - Contract 2 (PPA / Spare) | MWh |
| R455 | Contracted volume - SDE | MWh |
| R456 | Contracted volume - EEG | MWh |
| R457 | Contracted volume - SDE Contracts | MWh |
| R458 | Contracted volume - solar | MWh |

##### Non-contracted / Merchant Revenue -- Solar (R459-R506)

| Row | Label | Data Type |
|-----|-------|-----------|
| R459 | **Non-contracted / merchant revenue - Solar** | -- |
| R461 | Total Merchant Volume | MWh |
| R463 | Sensitivity | % |
| R464 | Market Price Real - Annual | LC/MWh |
| R466-R469 | Market Price Real (sens), Nominal, balancing, post-balancing | LC/MWh |
| R470 | Merchant revenue - Solar | LC'000 |
| R472 | **Guarantees of Origin - Solar** | -- |
| R474-R480 | Contracted GoO 1 (price, curve, volume, revenue) | LC/MWh, MWh, LC'000 |
| R482-R488 | Contracted GoO 2 (price, curve, volume, revenue) | LC/MWh, MWh, LC'000 |
| R490-R499 | Merchant GoO (period, curve, volume, revenue, totals) | various |
| R501-R506 | 11kV Solar revenue (switch, flag, price, revenue) | T/F, LC/MWh, LC'000 |

##### Revenue Totals -- Solar (R508-R513)

| Row | Label | Data Type |
|-----|-------|-----------|
| R508 | Total Solar Non-Contracted Revenue | LC'000 |
| R510 | Total energy volume | MWh |
| R511 | Total GOO volume | MWh |
| R512 | Total generation | MWh |
| R513 | Production check | T/F |

##### BESS Revenue (R515-R564)

| Row | Label | Data Type |
|-----|-------|-----------|
| R515 | **BESS Revenue** (sub-header) | -- |
| R517 | Net BESS periodic capacity | MW |
| R519 | **Contracted Revenue BESS - Tolling / Floor** | -- |
| R520-R525 | Tolling revenue (capacity, flag, price, revenue) | MW, 0/1, LC/MW, LC'000 |
| R527-R536 | Floor revenue (flag, capacity, price, contracted, merchant above) | various |
| R538-R544 | Capacity market (flag, price, derating, revenue) | various |
| R546 | **Non-contracted Revenue BESS** | -- |
| R547-R552 | CIDT (flag, curve, capacity, revenues) | various |
| R555 | BESS Periodic Merchant Capacity | MW |
| R557-R564 | Merchant revenue (price curves, optimiser fee, revenue) | LC/MWh, LC'000 |

##### Revenue Summaries (R566-R622)

| Row | Label | Data Type |
|-----|-------|-----------|
| R566 | **Total Solar PV Revenue** (sub-header) | -- |
| R568-R575 | EEG/SDE/PPA/GoO/Merchant/11kV/Total Solar revenue | LC'000 |
| R577-R579 | Solar contracted vs uncontracted split | LC'000 |
| R581 | **Total BESS Revenue** (sub-header) | -- |
| R583-R591 | Floor/Tolling/CM/Merchant/Total BESS / contracted split | LC'000 |
| R593 | **Total Project Revenue** | LC'000 |
| R595 | 10 year Revenue for summary | LC'000 |
| R598 | Land sale | LC'000 |
| R600 | Total Asset Revenue | LC'000 |
| R604-R622 | 10-year realized price calcs (Solar & BESS) | LC/MWh, date, flag |

---

#### 5.1.7 OpEx (R624-R713)

##### Solar OpEx (R631-R694)

| Row | Label | Data Type |
|-----|-------|-----------|
| R624 | **OpEx** (section header) | -- |
| R626-R629 | Solar/BESS/Total Capacity, O&M sensitivity | MW, % |
| R631 | **Solar OpEx** (sub-header) | -- |
| R632-R638 | O&M Solar periods (COD, period 2 start, flags) | date, 0/1 |
| R640-R641 | PV Plant O&M Expense (Initial / Step down) | LC'000 |
| R644-R655 | Land Lease (Fixed tiers Year 1-5 to 31+, indexed) | LC'000 |
| R657-R667 | Revenue dependent lease (tiers) | LC'000 |
| R669-R672 | Land Lease Summary (Fixed + Revenue dep + Total) | LC'000 |
| R675-R678 | Insurance (initial/step-up/total) | LC'000 |
| R680-R681 | Inverter warranty cost | LC'000 |
| R683-R688 | Other OpEx (AM, Others, PI ref, Production-based 1-3) | LC'000 |
| R691-R692 | LC Costs | LC'000 |
| R694 | **Total Solar OpEx** | LC'000 |

##### BESS OpEx (R696-R713)

| Row | Label | Data Type |
|-----|-------|-----------|
| R696 | **BESS OpEx** (sub-header) | -- |
| R702-R708 | BESS OpEx items (PI refs, spare revenue share, total) | LC'000 |
| R710 | **Total OpEx (Solar+BESS)** | LC'000 |
| R713 | 10 year Opex for summary | LC'000 |

---

#### 5.1.8 Depreciation (R715-R810)

##### Solar Depreciation (R717-R778)

| Row | Label | Data Type |
|-----|-------|-----------|
| R715 | **Depreciation** (section header) | -- |
| R717 | **Solar** (sub-header) | -- |
| R718 | D&A class 1 name (ref PI!C549) | text |
| R719-R722 | Class 1: Fixed Assets OB/Capitalise/Depreciate/CB | LC'000 |
| R724 | Check fully depreciated | T/F |
| R726-R733 | Class 2: same pattern | LC'000 |
| R735-R742 | Class 3: same pattern | LC'000 |
| R744-R751 | Class 4: same pattern | LC'000 |
| R753-R760 | Class 5: same pattern (ref PI!C554) | LC'000 |
| R762-R769 | Class 6: same pattern (ref PI!C555) | LC'000 |
| R771-R778 | **Total - Solar** (OB/Capitalise/Depreciate/CB) | LC'000 |

##### BESS Depreciation (R780-R797)

| Row | Label | Data Type |
|-----|-------|-----------|
| R780 | **BESS** (sub-header) | -- |
| R781-R787 | BESS D&A (OB/Capitalise/Depreciate/CB/Check) | LC'000 |
| R791-R797 | **Total depreciation** (OB/Capitalise/Depreciate/CB/Check) | LC'000 |

##### Tax Depreciation (R801-R810)

| Row | Label | Data Type |
|-----|-------|-----------|
| R801 | **Tax Depreciation** (sub-header) | -- |
| R803-R806 | Tax assets OB/Capitalise/Accelerated depreciation/CB | LC'000 |
| R808 | Check fully depreciated | T/F |
| R810 | Tax Depreciation | LC'000 |

---

#### 5.1.9 VAT (R812-R882)

| Row | Label | Data Type |
|-----|-------|-----------|
| R812 | **VAT** (section header) | -- |
| R814 | VAT rate | % |
| R816 | **VAT paid on CapEx** | -- |
| R817-R820 | VAT on CapEx items (refs to PI!C516-C519) | LC'000 |
| R822-R823 | VAT on devex items (refs to PI!C522-C523) | LC'000 |
| R825-R828 | VAT on BESS CapEx / total | LC'000 |
| R830 | VAT rate | % |
| R832-R833 | **VAT received on revenue** / Total revenue | LC'000 |
| R835-R849 | **VAT paid on OpEx** (land lease, O&M, insurance, others, BESS, total) | LC'000 |
| R851 | **Net VAT collected / (paid)** | LC'000 |
| R853 | Cumulative day counter | days |
| R855-R859 | **VAT receivable / (payable)** (OB/received/refunded/CB) | LC'000 |
| R861-R862 | **Net VAT movement** / funding requirement | LC'000 |
| R864-R869 | VAT Loan (OB/Drawdown/Repayment/CB/Quantum) | LC'000 |
| R871-R880 | VAT Loan Interest/Commitment Fee/IDC/Undrawn/Fee | LC'000 |
| R882 | **Net VAT movement after VAT facility** | LC'000 |

---

#### 5.1.10 Working Capital (R884-R927)

| Row | Label | Data Type |
|-----|-------|-----------|
| R884 | **Working capital** (section header) | -- |
| R886 | **Receivables** (sub-header) | -- |
| R888-R891 | Contracted receivables (OB/Recognised/Received/CB) | LC'000 |
| R893-R896 | Merchant receivables (OB/Recognised/Received/CB) | LC'000 |
| R898-R901 | BESS contracted receivables | LC'000 |
| R903-R906 | BESS non-contracted receivables | LC'000 |
| R908-R911 | Total receivables (OB/Recognised/Received/CB) | LC'000 |
| R913 | **Payables** (sub-header) | -- |
| R914-R917 | Total payables (OB/Recognised/Paid/CB) | LC'000 |
| R919-R922 | Solar OpEx payables (OB/Recognised/Paid/CB) | LC'000 |
| R924-R927 | BESS OpEx payables (OB/Recognised/Paid/CB) | LC'000 |

---

#### 5.1.11 Senior Debt (R929-R1033)

| Row | Label | Data Type |
|-----|-------|-----------|
| R929 | **Senior Debt** (section header) | -- |
| R931-R939 | Sizing (quantum, final drawdown, term, grace, periodic/final repayment, BESS sizing) | LC'000, date, months, x |
| R942-R951 | Revenue inputs for DSCR (contracted/non-contracted, Solar/BESS) | LC'000 |
| R953-R959 | Revenues received, OpEx paid, CFADS, CFADS semi-annual | LC'000 |
| R961-R964 | Sizing DSCRs (Solar/BESS, contracted/non-contracted) | x |
| R966-R977 | Sizing DSCR, DS live/paste, delta, debt quantum live/paste, delta | x, LC'000 |
| R979-R980 | Gearing Cap, Uses for Gearing Calculation | %, LC'000 |
| R983-R987 | Senior Debt schedule - fixed for FC (OB/Drawdown/Repayment/Sweep/CB) | LC'000 |
| R989-R993 | Senior Debt schedule - live (OB/Drawdown/Repayment/Sweep/CB) | LC'000 |
| R995-R1018 | **Interest - Senior Debt** (flags, EURIBOR/SONIA, hedging, post-hedging rate, periodic interest) | %, LC'000 |
| R1020 | Periodic Interest Rate | % |
| R1023-R1026 | Interest payable (OB/Recognised/Paid/CB) | LC'000 |
| R1029-R1033 | **DSCR calc** (CFADS, DS, DSCR) | LC'000, x |

---

#### 5.1.12 Reserve Security (R1035-R1084)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1035 | **Reserve Security** (section header) | -- |
| R1039-R1052 | **DSRA / DSRF** (target, OB/funding/top-up/release/CB) | LC'000 |
| R1054-R1066 | **MRA** (period, end, years, requirement, OB/funding/release/CB) | LC'000, months |
| R1068-R1084 | **Repowering CapEx** (active flag, pre-repower, drawdown, funding, OB/funding/release/CB, checks) | LC'000, T/F |

---

#### 5.1.13 LCs / Bonds (R1087-R1092)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1087 | **LCs / Bonds** (section header) | -- |
| R1090 | LC amount | LC'000 |
| R1091 | LC commitment fee | % |
| R1092 | LC commitment fee paid | LC'000 |

---

#### 5.1.14 Tax (R1095-R1157)

##### Corporation Tax (R1097-R1124)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1095 | **Tax** (section header) | -- |
| R1097 | **Corporation Tax** (sub-header) | -- |
| R1098-R1102 | Revenue, OpEx, EBITDA, D&A, EBIT | LC'000 |
| R1104-R1107 | Bonds & guarantees, Interest (senior/SHL/total) | LC'000 |
| R1109-R1113 | Interest deductibility (% of EBITDA, max, applied) | %, LC'000 |
| R1115-R1117 | Taxable income, pre-tax credit yearly | LC'000 |
| R1119-R1122 | Tax loss (use %, generated, used, balance) | %, LC'000 |
| R1124 | Tributable base after tax credit | LC'000 |

##### Local Tax (R1126-R1157)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1126 | **Local Tax** (sub-header) | -- |
| R1127-R1132 | Add-backs (lease, interest, exemption, non-deductible, total) | LC'000 |
| R1134-R1140 | Local tax base, losses (generated/OB/used/generation/CB) | LC'000 |
| R1142 | Local tax base | LC'000 |
| R1144-R1148 | Corporate Tax Rates 1-4, Local Tax rate | % |
| R1149-R1153 | Corporate Tax Basis Caps 1-4, Local Tax Basis | LC'000 |
| R1154 | Taxes | LC'000 |
| R1157 | Taxes (live) | LC'000 |

---

#### 5.1.15 Funding Calcs (R1160-R1234)

##### Sources & Uses (R1160-R1214)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1160 | **Funding calcs** (section header) | -- |
| R1162 | Funding period Flag | 0/1 flag |
| R1163-R1181 | Uses (CapEx, VAT, fees, MRA, bonds, DSRA, senior debt fees, IDC, commitment, DSRA movement, total) | LC'000 |
| R1182 | Total Uses - PASTED | LC'000 |
| R1184-R1190 | Sources (Senior Debt, SHL, Common Equity, Total, Check) | LC'000 |
| R1192-R1214 | Capital Structure (Debt CB formula, equity, base rate, periodic rates, commitment, undrawn, drawdown, OB/Drawdown/CB, arrangement fee, upfront fees, IDC, commitment fee) | LC'000, % |

##### SHL (R1217-R1234)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1217 | **SHL** (sub-header) | -- |
| R1218-R1219 | SHL Interest Rate, SHL IDC | % |
| R1221-R1223 | SHL OB/Drawdown/CB | LC'000 |
| R1225-R1228 | SHL Interest payable (OB/Recognised/Paid/CB) | LC'000 |
| R1230-R1234 | SHL detailed (OB/Drawdown/IDC/Principal Repayment/CB) | LC'000 |

---

#### 5.1.16 Financing Assets (R1236-R1256)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1236 | **Financing assets** (section header) | -- |
| R1238-R1243 | Capitalised costs during construction (VAT fees, senior debt fees/commitment/IDC, total) | LC'000 |
| R1245-R1249 | Financing Assets (OB/Capitalise Construction/Depreciate/CB) | LC'000 |
| R1251-R1254 | SHL IDC (OB/Capitalise/Depreciate/CB) | LC'000 |
| R1256 | Check fully depreciated | T/F |

---

#### 5.1.17 Distribution Waterfall (R1258-R1385)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1258 | **Distribution calcs (Waterfall)** (section header) | -- |
| R1260-R1267 | Flags (repayment, lock-up, post-debt, WC, sweep start/flag, first stage) | 0/1 flag |
| R1269 | Cash OB | LC'000 |
| R1271-R1278 | CapEx, Revenue, OpEx, bonds, VAT, Tax, Cash available for Senior Debt | LC'000 |
| R1280-R1296 | Senior Debt waterfall (fees/IDC/commitment, interest payable, debt schedule OB/Drawdown/Scheduled/Sweep/CB) | LC'000 |
| R1299-R1305 | Balloon repayment test (debt repaid, evaluation date, hurdle, outstanding, check) | LC'000, date, T/F |
| R1309-R1338 | Reserve Security waterfall (DSRF, DSRA, CapEx account, VAT Loan, Cash available for distributions) | LC'000 |
| R1341-R1355 | **Covenants** (CFADS, DS, DSCR, LLCR annual/NPV) | LC'000, x |
| R1358 | Distribution Flag | 0/1 flag |
| R1360 | Last operations flag | 0/1 flag |
| R1363-R1377 | SHL/Equity waterfall (SHL drawdown/interest/principal, WHT, dividends, equity OB/new/buyback/CB) | LC'000 |
| R1380 | Common stock flag | 0/1 flag |
| R1382 | Special Dividend Paid | LC'000 |
| R1384 | WHT on dividend payments | LC'000 |
| R1385 | Cash closing balance | LC'000 |

---

#### 5.1.18 Valuation (R1388-R1418)

##### Project IRR (R1390-R1397)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1388 | **Valuation** (section header) | -- |
| R1390 | **Project IRR** (sub-header) | -- |
| R1391-R1396 | Revenues, OpEx, VAT movement, Tax, CapEx, FCFF | LC'000 |
| R1397 | XIRR | % |

##### Equity IRR -- Distributions (R1400-R1409)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1400 | **Equity IRR (Distributions)** (sub-header) | -- |
| R1401-R1408 | Common stock (new/buyback), Dividend, Special Dividend, SHL (drawdown/repayment/interest), FCFE | LC'000 |
| R1409 | XIRR | % |

##### Equity IRR -- CF for Equity (R1412-R1418)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1412 | **Equity IRR (CF for Equity)** (sub-header) | -- |
| R1414-R1417 | CFO, CFI, Senior Debt net, FCFE (CF for Equity) | LC'000 |
| R1418 | XIRR | % |

---

#### 5.1.19 Quarterly Consolidation (R1420-R1606)

**Purpose**: Macro paste block. `SPV_ConsolidatedCashflows` named range covers R1428:PQ1607.

| Row | Label | Data Type |
|-----|-------|-----------|
| R1420 | **Quarterly consolidation - used for macro** (section header) | -- |
| R1427 | Do not move, for consolidation macro | anchor |
| R1428 | =Project_View (project name) | text |

##### Consolidated Income Statement (R1429-R1450)

Mirrors R34-R55 at quarterly granularity:

| Row | Label |
|-----|-------|
| R1429 | Revenue |
| R1430 | Contracted revenue |
| R1431 | Non-contracted revenue |
| R1432 | OpEx |
| R1433 | EBITDA |
| R1434 | D&A |
| R1435 | EBIT |
| R1436 | Interest |
| R1437 | Agency fees |
| R1438 | Reserve Facilities Commitment Fees |
| R1439 | PBT |
| R1441 | Taxes |
| R1442 | Net Income |
| R1444-R1450 | Retained earnings (OB/NI/Revaluation/Dividend/Special/WHT/CB) |

##### Consolidated Cash Flow (R1453-R1489)

Mirrors R59-R95 at quarterly granularity:

| Row | Label |
|-----|-------|
| R1453 | =Project_View |
| R1454-R1458 | Revenue received, OpEx paid, VAT movement, Tax, CFO |
| R1460-R1461 | CapEx, CFI |
| R1463-R1466 | VAT Loan fees/interest, Agency fees, Bonds |
| R1468-R1472 | Senior Debt (Drawdown/Principal/Sweep/Interest/Fees) |
| R1473-R1475 | SHL (Drawdown/Principal/Interest) |
| R1476-R1485 | Reserve fees, DSRA/DSRF/MRA movements, Equity, Dividends, WHT, CFF |
| R1487-R1489 | Cash (OB/Movement/CB) |

##### Consolidated KPIs & Valuation (R1491-R1543)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1491 | NPV of CFADS | LC'000 |
| R1492 | DSRF available | LC'000 |
| R1494-R1495 | FCFF, Project IRR | LC'000, % |
| R1497-R1498 | FCFE (Distributions), Equity IRR (Distributions) | LC'000, % |
| R1500-R1501 | FCFE (CF for Equity), Equity IRR (CF for Equity) | LC'000, % |
| R1503-R1504 | Solar revenue, BESS Revenue | LC'000 |
| R1506-R1507 | Solar CapEx, BESS CapEx | LC'000 |
| R1509-R1511 | Solar OpEx, BESS OpEx, Total OpEx | LC'000 |
| R1513-R1517 | Opex/MW 10yr, Opex 10yr, Revenue/MW 10yr, Revenue 10yr | LC'000/MW |
| R1519-R1523 | KPI, Senior/Total Gearing, Debt quantum | %, LC'000 |
| R1524-R1529 | LCOE (discount rate, costs NPV, production NPV, LCOE) | %, LC'000, LC/MWh |
| R1531 | Debt / MW | LC'000/MW |
| R1533-R1534 | DSRF, VAT | LC'000 |
| R1536-R1543 | Financing asset references (SHL interest, Senior interest, depreciation) | LC'000 |

##### Consolidated Construction & Volume (R1545-R1566)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1545-R1549 | Construction Costs, Devex breakdown, LCs | LC'000 |
| R1551-R1554 | Contracted/Non-contracted volume, Total production, BESS volume | MWh |
| R1556-R1557 | Solar contracted/merchant volume % | % |
| R1559-R1560 | Solar contracted/merchant revenues recognized | LC'000 |
| R1562-R1563 | BESS contracted/merchant revenues % | % |
| R1565-R1566 | Contracted/Merchant revenues received | LC'000 |

##### Consolidated Balance Sheet (R1568-R1597)

Mirrors R102-R131 at quarterly granularity:

| Row | Label |
|-----|-------|
| R1568-R1575 | ASSETS (Receivables/VAT/Cash/Collateral/DSRA/MRA/Total Current) |
| R1577-R1581 | Fixed Assets/Financial/Revaluation/Repower/Total Assets |
| R1583-R1592 | LIABILITIES (Payables/Interest Payable/Total Current/Senior Debt/DSRF/SHL/Total) |
| R1594-R1597 | SHAREHOLDERS EQUITY (Common Stock/Retained/Total) |

---

#### 5.1.20 Checks (R1600-R1608)

| Row | Label | Data Type |
|-----|-------|-----------|
| R1600 | **Checks** (section header) | -- |
| R1601 | BS check | T/F |
| R1602 | Statements check | T/F |
| R1603 | Senior Debt optimal | T/F |
| R1604 | Master check | T/F |
| R1605 | Balloon capacity test | T/F |
| R1606 | SHL repaid | T/F |
| R1608 | END marker (C1=*, C4=END) | marker |

---

## 6. ASE Output Sheets

### 6.1 Quarterly Output (35999 rows x 433 cols)

**Structure**: 200 repeating blocks, one per project.
Each block is ~180 rows, mirroring the Quarterly Consolidation block (R1428-R1607 of PLW).

**Control cells**:
- R6 C5: Pasted area indicator
- R6 C7: Starting column
- R7 C7: Starting row

**Per-block structure** (using Block 1 starting at R10 as example):

| Offset | Row | Label | Data Type |
|--------|-----|-------|-----------|
| +0 | R10 | Project header (INDEX into ProjectList) | text |
| +1 | R11 | Revenue | LC'000 |
| +2 | R12 | Contracted revenue | LC'000 |
| +3 | R13 | Non-contracted revenue | LC'000 |
| +4 | R14 | OpEx | LC'000 |
| +5 | R15 | EBITDA | LC'000 |
| +6 | R16 | D&A | LC'000 |
| +7 | R17 | EBIT | LC'000 |
| +8 | R18 | Interest | LC'000 |
| +9 | R19 | Agency fees | LC'000 |
| +10 | R20 | Reserve Facilities Commitment Fees | LC'000 |
| +11 | R21 | PBT | LC'000 |
| +13 | R23 | Taxes | LC'000 |
| +14 | R24 | Net Income | LC'000 |
| +16 | R26 | Retained earnings OB | LC'000 |
| +17 | R27 | Net Income | LC'000 |
| +18 | R28 | Revaluation | LC'000 |
| +19 | R29 | Dividend | LC'000 |
| +20 | R30 | Special Dividend | LC'000 |
| +21 | R31 | WHT | LC'000 |
| +22 | R32 | Retained earnings CB | LC'000 |
| +25 | R35 | CF header | -- |
| +26 | R36 | Revenue received | LC'000 |
| +27 | R37 | OpEx paid | LC'000 |
| ... | ... | (continues through full CF, BS, KPI blocks) | ... |
| +170 | R180 | End of block / spacer | -- |

**Column structure**:
- C1: INDEX formula (project lookup from ProjectList)
- C2: Concatenation key (project + line item)
- C3: Country (INDEX from PI row 6)
- C4: Technology (INDEX from PI row 7)
- C5: Project name carry-forward
- C6: Project cluster name
- C7: Line item label
- C8+: Quarterly time series data (433 quarterly columns)

---

### 6.2 Project Level Macro Paste (654 rows x 424 cols)

Intermediate staging area used by the VBA consolidation macro.
Mirrors PLW quarterly consolidation structure for the active project.

---

## 7. ASE HoldCo Sheets

### 7.1 HoldCo CFs & Valuation

**Purpose**: Consolidates all project-level cash flows into HoldCo view.

| Row Range | Section | Data Type |
|-----------|---------|-----------|
| R1-R5 | Headers | text |
| R6-R10 | HoldCo parameters (ownership %, management fee) | % |
| R13-R132 | Per-project rows (120 slots): project name, capacity, dates, devex allocation | various |
| R134-R200 | Cash flow consolidation (dividends received, SHL flows, equity flows) | LC'000 |
| R202-R250 | Valuation (NPV, IRR, equity value per project) | LC'000, % |

### 7.2 HoldCo income

**Purpose**: HoldCo-level income statement (management fees, interest income, G&A).

### 7.3 HoldCo_Facility

**Purpose**: HoldCo-level revolving credit facility / corporate debt.

---

## 8. ASE Presentation Sheets

### 8.1 Dashboard

Key metrics display for the active project. Links to PLW header rows (R4-R13).

### 8.2 ProjectSummary

Summary view per project: capacity, COD, revenue split, gearing, IRR.

### 8.3 HoldCo_Summary

HoldCo-level summary: total AUM, portfolio IRR, debt capacity.

### 8.4 Charts

Chart data source sheet for presentation charts.

### 8.5 Checks

Consolidated check flags across all projects (BS, statements, debt optimal, master).

---

## 9. GTC Sheets (21 sheets)

### 9.1 Instruction tab

**Rows**: ~40. **Purpose**: Instructions for completing the GTC template.

| Row | Content |
|-----|---------|
| R2 | Instructions for completing the template |
| R4-R7 | Format guidance (FY26A+F = 9 months actuals + 3 months forecast) |
| R9-R12 | Tab color coding (Green=slides, Blue=teams, Orange=FX) |
| R15-R16 | Data formatting instructions |
| R17-R18 | FX link |
| R20-R22 | Commentary requirements (FY26A+F vs FY27B, variances >USD100k) |
| R25-R32 | Remapping instructions (P&L, BS, CF links to mapping tabs) |

### 9.2 FX

FX rates tab (EUR to USD, GBP to EUR). Referenced by all GTC financial tabs.

### 9.3 FY26B PL Mapping

Mapping of prior year P&L line items to current format. Used by PnL projection formulas.

### 9.4 FY26B BS Mapping

Mapping of prior year BS line items to current format.

### 9.5 PnL projection - aggregate

**Rows**: ~100. **Cols**: 16.
**Structure**: Platform Total (USD'k) and AGP Proportionate Share (USD'k).
Periods: FY26B, FY26A+F, FY27B, FY28F, FY29F.

| Row | Label | Data Type |
|-----|-------|-----------|
| R5 | Operating Asset Business (subtotal) | USD'000 |
| R6 | Contracted Revenue - Energy | USD'000 |
| R7 | Non-contracted Revenue - Energy | USD'000 |
| R8 | Lease Income - Logistics | USD'000 |
| R9 | Retirement Lifestyle - Sales/Resales | USD'000 |
| R10 | Retirement Lifestyle - Services/DAMA | USD'000 |
| R11-R15 | Spare revenue lines 1-5 | USD'000 |
| R17 | Developer Co Business (subtotal) | USD'000 |
| R18 | Development Management Fee | USD'000 |
| R19 | Development Premium/Fee | USD'000 |
| R20 | Leasing Fee | USD'000 |
| R21-R25 | Spare developer lines 1-5 | USD'000 |
| R27 | Asset Mgmt Business (subtotal) | USD'000 |
| R28 | Asset / Fund Management Fee Income | USD'000 |
| R29 | Acquisition Fee | USD'000 |
| R30-R34 | Spare AM lines 1-5 | USD'000 |

Formulas reference `'FY26B PL Mapping'` and `'FY26A-29F PL Breakdown'` via SUMIF.

### 9.6 PnL projection - by asset

Same structure as aggregate but broken out per-asset.

### 9.7 BS projection - aggregate

**Rows**: ~60. **Structure**: Standard balance sheet format.

| Row Range | Section |
|-----------|---------|
| R110-R133 | ASSETS (Current + Non-current) |
| R134-R160 | LIABILITIES (Non-current + Current) |
| R162 | SHAREHOLDERS EQUITY |
| R164 | BS Check |

Cross-references `'BS projection - aggregate'` for line item labels and `'FY26B BS Mapping'` for prior year data.

### 9.8 BS projection - by asset

Same structure as aggregate but per-asset breakdown.

### 9.9 CF Capital Strategy

**Rows**: ~120. **Structure**: Cash flow statement with Capital Strategy overlay.

| Row Range | Section |
|-----------|---------|
| R166-R212 | Cash Flow Statement (CFO, CFI, CFF, Cash movement) |
| R215-R256 | Capital Strategy section |
| R258-R268 | Debt details |
| R270-R281 | Equity details |
| R283-R286 | Net Cash Flow, Opening/Closing Cash |

### 9.10 Valuation (2)

**Rows**: ~50.

| Row Range | Section |
|-----------|---------|
| R289-R306 | Enterprise Value by asset stage (COD, Under Construction, RTB, First/Second Permit, Pipeline) |
| R308-R316 | Net Debt bridge to Equity Value |
| R318-R336 | IRR analysis, discount rates, multiples |

### 9.11 FY26A-29F PL Breakdown

Detailed P&L breakdown feeding aggregate. One row per line item, columns for each FY.

### 9.12 FY26A-29F BS Breakdown

Detailed BS breakdown feeding aggregate.

### 9.13 Dev Pipeline

Development pipeline tracker (project stages, MW, expected COD).

### 9.14 OpCo Performance

Operating asset performance metrics.

### 9.15 Capacity MW tracker

MW capacity tracking across periods.

### 9.16-9.21 Additional GTC sheets

| Sheet | Purpose |
|-------|---------|
| GTC Summary | Combined 336-row summary (P&L R1-R100, BS R101-R165, CF R166-R288, Valuation R289-R336) |
| QREP Consolidated | Quarterly reporting template |
| Commentary | Text commentary per section |
| Revenue Bridge | Revenue waterfall analysis |
| CapEx Bridge | CapEx waterfall analysis |
| Adjustments | Manual adjustments / overrides |

---

## 10. Cross-Sheet Reference Map

### Primary Data Flows

```
Project Info ──────────────> Project Level Workings
    (project params)              (calc engine)
                                       │
Country Inputs ────────────>           │
    (country-specific rates)           │
                                       │
Financing Inputs ──────────>           │
    (termsheet params)                 │
                                       │
Time Inputs (A) ───────────>           │
    (annual curves)                    │
                                       │
Time Inputs (M) ───────────>           │
    (monthly CapEx/DevEx)              │
                                       │
Sensis ────────────────────>           │
    (scenario levers)                  │
                                       │
                                       v
                              Quarterly Output
                              (35999 rows, all projects)
                                       │
                                       v
                              HoldCo CFs & Valuation
                              (portfolio consolidation)
                                       │
                                       v
                              GTC Sheets
                              (external reporting format)
```

### Key Cross-Sheet References

| Source Sheet | Target Sheet | Reference Pattern |
|-------------|-------------|-------------------|
| Project Info R5 (ProjectList) | Quarterly Output (INDEX/MATCH) | Project name lookup |
| Project Info R6-R7 | PLW R18, R16 | Country, Technology |
| Project Info R569 (modelStartDate) | PLW, Time Inputs | Date anchoring |
| Country Inputs C45, C51 | PLW R218, R228 | VAT on CapEx/devex flags |
| Country Inputs D65-D91 | PLW R837-R848 | VAT on OpEx rates |
| Financing Inputs (termsheet) | PLW R931-R977 | Debt sizing parameters |
| Time Inputs (A) R48-R49 | PLW R1002-R1003 | EURIBOR/SONIA curves |
| Time Inputs (A) R24-R43 | PLW R146-R165 | Inflation profiles |
| Time Inputs (A) R68-R83 | PLW R464 | Solar market price curves |
| Time Inputs (A) R86-R97 | PLW R558-R560 | BESS market price curves |
| Time Inputs (M) R13-R132 | HoldCo CFs & Valuation | Historical devex per project |
| Time Inputs (M) R387+ | PLW R219-R221 | Monthly CapEx profiles |
| Sensis R15-R57 | PLW (via sensitivity) | Scenario lever application |
| Sensis R43 | Time Inputs (A) R48-R49 | Interest rate sensitivity |
| PLW R1428-R1607 | Quarterly Output | SPV_ConsolidatedCashflows paste target |
| PLW R34-R55 | GTC Summary (P&L) | Income statement |
| PLW R57-R95 | GTC Summary (CF) | Cash flow statement |
| PLW R100-R134 | GTC Summary (BS) | Balance sheet |
| Dashboard E100 | Time Inputs (M) R11 | Historical devex link |
| FX!$C$6 | Time Inputs (A) R56-R57 | Spot FX rate |
| BS projection - aggregate | GTC Summary R110-R164 | Balance sheet layout |
| CF Capital Strategy | GTC Summary R166-R288 | Cash flow layout |
| Valuation (2) | GTC Summary R289-R336 | Valuation layout |
| FY26B PL Mapping | PnL projection (SUMIF) | Prior year P&L mapping |

---

## 11. Structural Fingerprint

Use these values for quick structural validation of any new quarterly file.

### Row Count Anchors

| Sheet | Expected Rows | Key Anchor |
|-------|--------------|------------|
| Project Info | 595 | R569 = modelStartDate |
| Country Inputs | 138 | R117 = Seasonality section start |
| Financing Inputs | 77 | R70 = Securities section start |
| Time Inputs (A) | 653 | R203 = MRA section start |
| Time Inputs (M) | 1581 | R258 = Acquisition Cost section start |
| Sensis | 140 | R60 = Results matrix start |
| Project Level Workings | 1609 | R1608 = END marker (C1=*, C4=END) |
| Quarterly Output | 35999 | ~200 blocks of ~180 rows |
| GTC Summary | 336 | R289 = Valuation section start |

### Column Count Anchors

| Sheet | Expected Cols | Key Column |
|-------|--------------|------------|
| Project Info | I-DX (233 data cols) | I5 = first project |
| Project Level Workings | H-PQ (time series) | H15 = Project_View |
| Quarterly Output | H-PQ (433 data cols) | H10 = first data column |
| Time Inputs (A) | H-LU (325 data cols) | H = first annual period |
| Time Inputs (M) | J-XH (634 data cols) | J = first monthly period |

### Section Boundary Markers

| Sheet | Section | Start Row | End Row | Identifier |
|-------|---------|-----------|---------|------------|
| PLW | Headers/Checks | R1 | R28 | C2="Project Parthenon" at R1 |
| PLW | Financial Statements | R30 | R134 | C5="Financial Statements" at R30 |
| PLW | Flags | R136 | R214 | C5="Flags" at R138 |
| PLW | CapEx | R216 | R245 | C5="CapEx" at R216 |
| PLW | Production | R247 | R272 | C5="Production" at R247 |
| PLW | Revenues | R274 | R622 | C5="Revenues" at R274 |
| PLW | OpEx | R624 | R713 | C5="OpEx" at R624 |
| PLW | Depreciation | R715 | R810 | C5="Depreciation" at R715 |
| PLW | VAT | R812 | R882 | C5="VAT" at R812 |
| PLW | Working Capital | R884 | R927 | C5="Working capital" at R884 |
| PLW | Senior Debt | R929 | R1033 | C5="Senior Debt" at R929 |
| PLW | Reserve Security | R1035 | R1084 | C5="Reserve Security" at R1035 |
| PLW | LCs/Bonds | R1087 | R1092 | C5="LCs / Bonds" at R1087 |
| PLW | Tax | R1095 | R1157 | C5="Tax" at R1095 |
| PLW | Funding Calcs | R1160 | R1234 | C5="Funding calcs" at R1160 |
| PLW | Financing Assets | R1236 | R1256 | C5="Financing assets" at R1236 |
| PLW | Distribution Waterfall | R1258 | R1385 | C5="Distribution calcs (Waterfall)" at R1258 |
| PLW | Valuation | R1388 | R1418 | C5="Valuation" at R1388 |
| PLW | Quarterly Consolidation | R1420 | R1606 | C5="Quarterly consolidation - used for macro" at R1420 |
| PLW | Checks | R1600 | R1608 | C5="Checks" at R1600 |

### Formula Pattern Signatures

| Pattern | Location | Purpose |
|---------|----------|---------|
| `INDEX(ProjectList,MATCH(...))` | Quarterly Output C1 | Project lookup |
| `SUMIF('FY26B PL Mapping'!C:C,...)` | PnL projection R6+ | GTC P&L aggregation |
| `='Project Level Workings'!$H$15` | Project_View references | Active project selection |
| `=EOMONTH('HoldCo CFs & Valuation'!I__,0)` | Time Inputs (M) C8 | Monthly date anchoring |
| `='CF Capital Strategy'!C__` | GTC Summary R169+ | CF pass-through |
| `='BS projection - aggregate'!B__` | GTC Summary R110+ | BS pass-through |
| `='Valuation (2)'!B__` | GTC Summary R292+ | Valuation pass-through |

---

*End of Structural Map*
*Generated: 2026-03-12*
*Source: F2 (GTC Enhanced) baseline extraction*
