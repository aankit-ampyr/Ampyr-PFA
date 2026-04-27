import sys
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
ws = wb.active
ws.title = "Project Parthenon Timeline"

# Colors
DARK_BLUE = "1B4F72"
LIGHT_BLUE = "D6EAF8"
DARK_GRAY = "2C3E50"
LIGHT_GRAY = "F2F3F4"
WHITE = "FFFFFF"
BUCKET1_BG = "EBF5FB"
BUCKET2_BG = "FEF9E7"
BUCKET3_BG = "F5EEF8"
SUBTOTAL_BG = "D5F5E3"
GRAND_BG = "FADBD8"

thin = Side(style='thin', color='CCCCCC')
borders = Border(top=thin, bottom=thin, left=thin, right=thin)
thick_bottom = Border(top=thin, bottom=Side(style='medium', color='333333'), left=thin, right=thin)

ws.column_dimensions['A'].width = 8
ws.column_dimensions['B'].width = 52
ws.column_dimensions['C'].width = 65
ws.column_dimensions['D'].width = 14


def write_bucket_header(ws, row, text, fill_color):
    ws.merge_cells(f'A{row}:D{row}')
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name='Arial', size=12, bold=True, color=DARK_BLUE)
    c.fill = PatternFill('solid', fgColor=fill_color)
    c.alignment = Alignment(horizontal='left', vertical='center')
    for col in range(1, 5):
        ws.cell(row=row, column=col).border = borders
        ws.cell(row=row, column=col).fill = PatternFill('solid', fgColor=fill_color)
    ws.row_dimensions[row].height = 30


def write_row(ws, row, num, item, desc, weeks, bg=None):
    vals = [num, item, desc, weeks]
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name='Arial', size=10, color='333333')
        c.alignment = Alignment(
            horizontal='center' if i in [1, 4] else 'left',
            vertical='center', wrap_text=True
        )
        c.border = borders
        if bg:
            c.fill = PatternFill('solid', fgColor=bg)
    ws.row_dimensions[row].height = 45


def write_subtotal(ws, row, text, weeks, bg):
    ws.merge_cells(f'A{row}:C{row}')
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name='Arial', size=11, bold=True, color=DARK_BLUE)
    c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal='right', vertical='center')
    for col in range(1, 5):
        ws.cell(row=row, column=col).border = thick_bottom
        ws.cell(row=row, column=col).fill = PatternFill('solid', fgColor=bg)
    c = ws.cell(row=row, column=4, value=weeks)
    c.font = Font(name='Arial', size=11, bold=True, color=DARK_BLUE)
    c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 30


# ── TITLE ──
ws.merge_cells('A1:D1')
c = ws['A1']
c.value = "Project Parthenon - Development Timeline (v1.1, re-baselined 2026-04-27)"
c.font = Font(name='Arial', size=16, bold=True, color=WHITE)
c.fill = PatternFill('solid', fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 40
for col in ['B1', 'C1', 'D1']:
    ws[col].fill = PatternFill('solid', fgColor=DARK_BLUE)

ws.merge_cells('A2:D2')
c = ws['A2']
c.value = "Basis: 25 hrs/week (~65% dedication) | Solo developer + AI assistance | External workbook deps: out of scope"
c.font = Font(name='Arial', size=10, italic=True, color='7F8C8D')
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[2].height = 25

# ── HEADER ──
row = 4
headers = ['#', 'Work Item', 'Work Involved', 'Timeline\n(Weeks)']
header_fill = PatternFill('solid', fgColor=DARK_GRAY)
header_font = Font(name='Arial', size=11, bold=True, color=WHITE)
for i, h in enumerate(headers, 1):
    c = ws.cell(row=row, column=i, value=h)
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal='center' if i in [1, 4] else 'left', vertical='center', wrap_text=True)
    c.border = borders
ws.row_dimensions[row].height = 30

r = 5

# ── BUCKET 0: PHASE 1 PREREQUISITES (deferred investigations from 23-Apr re-analysis) ──
write_bucket_header(ws, r, "BUCKET 0: PHASE 1 PREREQUISITES (gating deferred investigations - v1.1)", SUBTOTAL_BG)
r += 1

data_b0 = [
    ('P1', "Per-asset parameter inventory",
     "Horizontal scan of Project Info x 120 slots. Confirms 590+ parameter claim and produces the parameter list for schema design.",
     0.3),
    ('P2', "78-asset list by country/tech",
     "Enumerate the 78 active assets in F1 with country (DE/UK/NL) and technology (Solar / BESS / Solar+BESS). Feeds asset registry.",
     0.05),
    ('P3', "Time Inputs (M) disagg mechanic (partial)",
     "Decode annual->monthly disaggregation pattern (just the mechanic). Full deep-dive of 603k formulas can defer until before Revenue (#4).",
     0.5),
    ('P4', "75 PLW edge-case rows - per-row inspection",
     "Catalogue what makes each row's formula differ across time cols. Informs translator override-map design (#1).",
     0.4),
    ('P5', "184 changed PLW rows - F3 vs F1 formula comparison",
     "Side-by-side diff to understand what ASE actively edits each quarter. Informs translator and re-validation strategy.",
     0.2),
]

for num, item, desc, weeks in data_b0:
    write_row(ws, r, num, item, desc, weeks, bg=LIGHT_GRAY)
    r += 1

write_subtotal(ws, r, "Bucket 0 Subtotal (must complete before Phase 1 / translator)", 1.45, SUBTOTAL_BG)
r += 2

# ── BUCKET 1: CORE ENGINE ──
write_bucket_header(ws, r, "BUCKET 1: CORE ENGINE (Replicating the Excel calculation logic in Python)", BUCKET1_BG)
r += 1

data_b1 = [
    (1, "Formula Extraction & Translation Framework",
     "Parse 915 PLW formula rows (cols AB-QF) from .xlsm. Excel-to-Python translator covering 26 functions. Per-column override map for 75 edge-case rows whose formulas differ across the 421 time cols. Re-runnable each quarter with formula-delta report.",
     2.5),
    (2, "Flags & Timing Block",
     "~100 formula rows: construction dates, COD flags, monthly/quarterly period flags, inflation index lookups from Time Inputs (A) (64,413 PLW refs).",
     0.8),
    (3, "Production Engine",
     "~20 formula rows: Solar yield (P50/P90 selection), degradation curves, curtailment, availability, BESS dispatch & repower logic. Smallest calc block.",
     0.6),
    (4, "Revenue Engine (LARGEST BLOCK)",
     "~200 formula rows: PPA, CfD, FiT/FiP, merchant pricing, GoO, tolling, floor revenue. Country-specific logic DE/UK/NL. XLOOKUP into Time Inputs (A) for pricing curves; depends on Time Inputs (M) annual->monthly disagg.",
     2.8),
    (5, "OpEx Block",
     "~60 formula rows: O&M Solar/BESS, insurance, land lease, asset management fees, revenue-dependent lease, community benefit.",
     1.0),
    (6, "Depreciation & Fixed Assets",
     "~70 formula rows: 7 depreciation categories with fully-depreciated checks, BESS repower depreciation, Fixed Assets closing balance.",
     1.0),
    (7, "VAT",
     "~50 formula rows: VAT on CapEx/OpEx/Revenue, receivable/payable tracking, VAT facility drawdown/repayment.",
     0.6),
    (8, "Senior Debt & DSCR Solver",
     "~100 formula rows: debt drawdown, repayment, cash sweep, interest calc, DSCR sculpting. Fixed-point solver with TWO criteria (debt_delta<0.2 AND Use_delta<0.2) plus max_iter cap (VBA has none). Includes DSRA + CapEx repower account; depends on HoldCo_Facility decode.",
     1.6),
    (9, "Tax",
     "~80 formula rows: corporate tax, loss carry-forward, local tax basis, interest deduction limitation. Country-specific rates for DE/UK.",
     1.2),
    (10, "SHL & Distribution Waterfall",
     "~110 formula rows: shareholder loan drawdown/repayment/IDC, distribution calcs, lock-up DSCR test, dividends, equity issuance/buyback, withholding tax.",
     1.4),
    (11, "IRR & Financial Statements",
     "~140 formula rows: XIRR (project & equity level), quarterly consolidation via SUMIFS, income statement, balance sheet, cashflow statement. Depends on HoldCo CFs (664k formulas) + HoldCo income decode.",
     1.0),
    (12, "Scenario Engine",
     "Replicate Sensis mechanism: 24 scenario slots, 15+ levers across 4 sections (Operations / Production / Uncontracted revs / Contracted revs). Mixed types (boolean/enum/percentage). Lever_type discriminator in schema. VBA flag-comparison subtlety (string 'True' vs boolean True) replicated exactly.",
     1.2),
    (13, "Sensitivity Analysis",
     "Single-variable & multi-variable sweeps, tornado chart data generation. Vectorized across all 78 assets - replaces 1,872 sequential Excel recalcs with one pass.",
     0.8),
    (14, "Validation & Debugging",
     "Cross-check Python output vs Excel for all 78 assets x 170 metrics x 421 periods. Oracle confirmed available (PLW + Quarterly Output cached). Explicit per-row tests for all 75 edge-case rows + country-specific branches.",
     1.6),
]

for num, item, desc, weeks in data_b1:
    bg = LIGHT_GRAY if num % 2 == 0 else None
    write_row(ws, r, num, item, desc, weeks, bg=bg)
    r += 1

write_subtotal(ws, r, "Bucket 1 Subtotal (915 PLW formula rows, 75 edge cases, 26 Excel functions)", 18.1, SUBTOTAL_BG)
r += 2

# ── BUCKET 2: DATA & INFRASTRUCTURE ──
write_bucket_header(ws, r, "BUCKET 2: DATA & INFRASTRUCTURE (Making it a working application)", BUCKET2_BG)
r += 1

data_b2 = [
    (15, "Database Schema & Migrations",
     "PostgreSQL tables: assets, parameters, time series, metrics, scenarios (with lever_type discriminator), versions. Alembic migrations. Table partitioning for 3M+ cells.",
     1.0),
    (16, "Excel Ingestion Pipeline",
     "Parse .xlsm into normalized DB tables. 7 input sheets (228K cells, 25K input levers). Read hidden rows (Time Inputs A 69%, HoldCo income 80%). Filter ~490 add-in named-range noise (Capital IQ/Bloomberg/Smartview/Access). Log #REF! errors (40 in F2). Snapshot external SharePoint workbook values inline (Project Canopy, CIP v7) - external files OUT OF SCOPE.",
     1.7),
    (17, "Quarterly File Comparison Engine",
     "Structural fingerprint diff with row-shift alignment (not equality - 13 named ranges shifted +2 rows F3->F1). Parameter-level change detection. PLW formula-row delta report (184 rows changed F3->F1). Automated change report.",
     1.0),
    (18, "FastAPI Backend",
     "REST API: CRUD endpoints, scenario management API, quarterly diff API, report generation trigger, calc engine invocation.",
     1.2),
    (19, "GTC Reporting (21 sheets)",
     "Asset workings (1.19M cells) is a reshape collapsing to one groupby/pivot, not 21 bespoke implementations. PnL/BS projections are generic ROUND(SUMIF(...)) aggregators. Single aggregator + template registry.",
     1.0),
    (20, "Excel Report Export",
     "Generate formatted .xlsx matching audit requirements. Reproduce GTC sheet layouts using openpyxl/xlsxwriter.",
     0.6),
]

for num, item, desc, weeks in data_b2:
    bg = LIGHT_GRAY if num % 2 == 0 else None
    write_row(ws, r, num, item, desc, weeks, bg=bg)
    r += 1

write_subtotal(ws, r, "Bucket 2 Subtotal", 6.5, SUBTOTAL_BG)
r += 2

# ── BUCKET 3: APPLICATION LAYER ──
write_bucket_header(ws, r, "BUCKET 3: APPLICATION LAYER (Login, UI, deployment)", BUCKET3_BG)
r += 1

data_b3 = [
    (21, "Authentication & Security",
     "Streamlit built-in auth, basic session handling. Lean for prototype.",
     0.3),
    (22, "Dashboard & UI",
     "Asset browser, portfolio view, scenario comparison screens, charts, filters. Functional over beautiful for prototype.",
     1.2),
    (23, "Error Handling & Logging",
     "Input validation, calc engine error trapping, audit logging.",
     0.3),
    (24, "Deployment",
     "Docker setup, environment config, basic CI/CD pipeline.",
     0.3),
    (25, "Documentation & Handover",
     "User guide, API documentation, architecture docs for future developers.",
     0.3),
]

for num, item, desc, weeks in data_b3:
    bg = LIGHT_GRAY if num % 2 == 0 else None
    write_row(ws, r, num, item, desc, weeks, bg=bg)
    r += 1

write_subtotal(ws, r, "Bucket 3 Subtotal", 2.5, SUBTOTAL_BG)
r += 2

# ── BUCKET 0B: GATED DEEP-DIVE INVESTIGATIONS (concurrent with Bucket 1) ──
write_bucket_header(ws, r, "BUCKET 0B: GATED DEEP-DIVES (run concurrently with Bucket 1, before specific blocks)", SUBTOTAL_BG)
r += 1

data_b0b = [
    ('G1', "Time Inputs (M) full deep-dive",
     "Complete the 603k-formula monthly disagg analysis. Required before Revenue (#4).",
     0.4),
    ('G2', "HoldCo_Facility decode",
     "92k formulas, facility-level debt model. Required before Senior Debt (#8).",
     0.3),
    ('G3', "HoldCo CFs & Valuation decode",
     "664k formulas, 3,379 rows. Second-largest formula sheet. Required before IRR (#11).",
     0.7),
    ('G4', "HoldCo income decode",
     "155k formulas, 80% rows hidden. Confirm what's alive vs legacy. Required before IRR (#11).",
     0.2),
]

for num, item, desc, weeks in data_b0b:
    write_row(ws, r, num, item, desc, weeks, bg=LIGHT_GRAY)
    r += 1

write_subtotal(ws, r, "Bucket 0B Subtotal (gated deep-dives, concurrent with Bucket 1)", 1.6, SUBTOTAL_BG)
r += 2

# ── BUFFER ──
write_bucket_header(ws, r, "BUFFER", GRAND_BG)
r += 1
write_row(ws, r, '', "Debugging Surprises & SME Response Delays",
          "Edge cases in country-specific logic (DE/UK), Finance team validation turnaround, unexpected formula behaviour, integration issues. Reduced from v1.0 (2.2w) - net effort moved into bucket subtotals where measurable.",
          1.4, bg=GRAND_BG)
r += 2

# ── GRAND TOTAL ──
ws.merge_cells(f'A{r}:C{r}')
c = ws.cell(row=r, column=1, value="GRAND TOTAL")
c.font = Font(name='Arial', size=13, bold=True, color=WHITE)
c.fill = PatternFill('solid', fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal='right', vertical='center')
for col in range(1, 5):
    ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=DARK_BLUE)
    ws.cell(row=r, column=col).border = Border(
        top=Side(style='medium', color='333333'),
        bottom=Side(style='medium', color='333333'),
        left=thin, right=thin
    )
c = ws.cell(row=r, column=4, value="~30 weeks")
c.font = Font(name='Arial', size=13, bold=True, color=WHITE)
c.fill = PatternFill('solid', fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[r].height = 35
r += 1

ws.merge_cells(f'A{r}:C{r}')
c = ws.cell(row=r, column=1, value="ESTIMATED DURATION (Bucket 0B concurrent with Bucket 1)")
c.font = Font(name='Arial', size=13, bold=True, color=DARK_BLUE)
c.fill = PatternFill('solid', fgColor=LIGHT_BLUE)
c.alignment = Alignment(horizontal='right', vertical='center')
for col in range(1, 5):
    ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=LIGHT_BLUE)
    ws.cell(row=r, column=col).border = borders
c = ws.cell(row=r, column=4, value="~7.5 Months")
c.font = Font(name='Arial', size=13, bold=True, color=DARK_BLUE)
c.fill = PatternFill('solid', fgColor=LIGHT_BLUE)
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[r].height = 35
r += 2

# ── NOTES ──
ws.merge_cells(f'A{r}:D{r}')
c = ws.cell(row=r, column=1, value="NOTES & ASSUMPTIONS")
c.font = Font(name='Arial', size=11, bold=True, color=DARK_BLUE)
r += 1

notes = [
    "v1.1 NOTES (re-baselined 2026-04-27 against 23-Apr deep re-analysis):",
    "1. PLW has 915 rows with time-axis formulas (was 877 in v1.0) over cols AB-QF (not M-PQ). 26 Excel functions confirmed.",
    "2. 75 PLW edge-case rows have formulas that differ across time cols (was 13 in v1.0) - per-column override map required in translator.",
    "3. PLW is a single-asset-at-a-time model (421 time columns). Python engine vectorizes across all 78 assets simultaneously.",
    "4. DSCR solver: TWO convergence criteria (debt_delta<0.2 AND Use_delta<0.2). VBA has no max-iter cap; Python adds one with hard error.",
    "5. Asset workings (1.19M cells) is a reshape, not 21 bespoke implementations - collapses to one groupby/pivot. GTC reporting reduced 1.4w -> 1.0w.",
    "6. Revenue engine is the single largest risk item: 200 formula rows with country-specific branching (DE/UK/NL).",
    "7. Sensis: 15+ levers across 4 sections (Operations / Production / Uncontracted revs / Contracted revs), mixed types (boolean/enum/percentage). Schema needs lever_type discriminator.",
    "8. PLW formulas not frozen between quarters: 184 rows changed F3->F1 (11.5%). Translator re-runnable each quarter with formula-delta report.",
    "9. Project Info shifts row indices between quarters (+2 rows F3->F1). Structural fingerprint comparator must align by section label / named-range identity, not equality.",
    "10. External SharePoint workbooks (Project Canopy v14/v24/v30, CIP v7 Capacity) are OUT OF SCOPE - ingestion snapshots cached values inline only.",
    "11. Validation oracle confirmed available (probed 2026-04-27): F1 PLW deep-region 64% cached, Quarterly Output 62%. No manual F9-and-save needed.",
    "12. Bucket 0 (1.45w) is NEW v1.1 work - prerequisites that gate Phase 1 schema/ingestion. Bucket 0B (1.6w) deep-dives run concurrently with Bucket 1.",
    "13. Buffer reduced 2.2w -> 1.4w; net effort moved into measurable bucket items (translator +0.5, debt +0.2, scenario +0.2, ingestion +0.3, GTC -0.4).",
    "14. Assumes Finance team SME availability for validation queries within 2-3 business days.",
]

for note in notes:
    ws.merge_cells(f'A{r}:D{r}')
    c = ws.cell(row=r, column=1, value=note)
    c.font = Font(name='Arial', size=9, color='666666')
    ws.row_dimensions[r].height = 18
    r += 1

ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

out = r"C:\repos\Ampyr-PFA\docs\Ampyr Financial Model Digitisation timeline.xlsx"
wb.save(out)
print(f"Saved to {out}")
