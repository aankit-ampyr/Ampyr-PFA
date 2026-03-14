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
c.value = "Project Parthenon - Development Timeline"
c.font = Font(name='Arial', size=16, bold=True, color=WHITE)
c.fill = PatternFill('solid', fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 40
for col in ['B1', 'C1', 'D1']:
    ws[col].fill = PatternFill('solid', fgColor=DARK_BLUE)

ws.merge_cells('A2:D2')
c = ws['A2']
c.value = "Basis: 25 hrs/week (~65% dedication) | Solo developer + AI assistance"
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

# ── BUCKET 1: CORE ENGINE ──
write_bucket_header(ws, r, "BUCKET 1: CORE ENGINE (Replicating the Excel calculation logic in Python)", BUCKET1_BG)
r += 1

data_b1 = [
    (1, "Formula Extraction & Translation Framework",
     "Parse 877 formula templates from PLW .xlsm. Build Excel-to-Python translator covering 26 Excel functions (SUM, IF, INDEX/MATCH, XLOOKUP, SUMIFS, etc.)",
     2.0),
    (2, "Flags & Timing Block",
     "~100 formula rows: construction dates, COD flags, monthly/quarterly period flags, inflation index lookups from Time Inputs. 43% simple arithmetic.",
     0.8),
    (3, "Production Engine",
     "~20 formula rows: Solar yield (P50/P90 selection), degradation curves, curtailment, availability, BESS dispatch & repower logic. Smallest calc block.",
     0.6),
    (4, "Revenue Engine (LARGEST BLOCK)",
     "~200 formula rows: PPA, CfD, FiT/FiP, merchant pricing, GoO, tolling, floor revenue. Country-specific logic for DE/UK/NL. XLOOKUP into Time Inputs for pricing curves.",
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
     "~100 formula rows: debt drawdown, repayment, cash sweep, interest calc, DSCR sculpting. Solver is simple fixed-point iteration (delta < 0.2 convergence). Includes DSRA and CapEx repower account.",
     1.4),
    (9, "Tax",
     "~80 formula rows: corporate tax, loss carry-forward, local tax basis, interest deduction limitation. Country-specific rates for DE/UK/NL.",
     1.2),
    (10, "SHL & Distribution Waterfall",
     "~110 formula rows: shareholder loan drawdown/repayment/IDC, distribution calcs, lock-up DSCR test, dividends, equity issuance/buyback, withholding tax.",
     1.4),
    (11, "IRR & Financial Statements",
     "~140 formula rows: XIRR (project & equity level), quarterly consolidation via SUMIFS, income statement, balance sheet, cashflow statement. Mostly aggregation.",
     1.0),
    (12, "Scenario Engine",
     "Replicate Sensis mechanism: 24 scenario slots, 6 levers (Devex, Capex, Opex, Production, Revenue, Financing). VBA confirms trivial orchestration: set Live_case, apply overrides, run calc.",
     1.0),
    (13, "Sensitivity Analysis",
     "Single-variable & multi-variable sweeps, tornado chart data generation. Vectorized across all 78 assets - replaces 1,872 sequential Excel recalcs with one pass.",
     0.8),
    (14, "Validation & Debugging",
     "Cross-check Python output vs Excel for all 78 assets x 170 metrics. 98.5% formula reuse limits edge cases. Focus on 13 exception rows + country-specific branches.",
     1.6),
]

for num, item, desc, weeks in data_b1:
    bg = LIGHT_GRAY if num % 2 == 0 else None
    write_row(ws, r, num, item, desc, weeks, bg=bg)
    r += 1

write_subtotal(ws, r, "Bucket 1 Subtotal (877 formula templates, 26 functions, 286 calc blocks)", 17.2, SUBTOTAL_BG)
r += 2

# ── BUCKET 2: DATA & INFRASTRUCTURE ──
write_bucket_header(ws, r, "BUCKET 2: DATA & INFRASTRUCTURE (Making it a working application)", BUCKET2_BG)
r += 1

data_b2 = [
    (15, "Database Schema & Migrations",
     "PostgreSQL tables: assets, parameters, time series, metrics, scenarios, versions. Alembic migrations. Table partitioning for 3M+ cells.",
     1.0),
    (16, "Excel Ingestion Pipeline",
     "Parse .xlsm/.xlsb into normalized DB tables. 7 input sheets (228K cells, 25K input levers). Version tagging per quarterly upload.",
     1.4),
    (17, "Quarterly File Comparison Engine",
     "Structural fingerprint diff against STRUCTURAL_MAP baseline. Parameter-level change detection. Automated change report generation.",
     1.0),
    (18, "FastAPI Backend",
     "REST API: CRUD endpoints, scenario management API, quarterly diff API, report generation trigger, calc engine invocation.",
     1.2),
    (19, "GTC Reporting (21 sheets)",
     "142K formula cells but only 16 complex formulas. Mostly SUM/SUMIF/SUMIFS aggregation. PL/BS Mapping pairs structurally identical. Asset workings = SUMIFS from Quarterly Output.",
     1.4),
    (20, "Excel Report Export",
     "Generate formatted .xlsx matching audit requirements. Reproduce GTC sheet layouts using openpyxl/xlsxwriter.",
     0.6),
]

for num, item, desc, weeks in data_b2:
    bg = LIGHT_GRAY if num % 2 == 0 else None
    write_row(ws, r, num, item, desc, weeks, bg=bg)
    r += 1

write_subtotal(ws, r, "Bucket 2 Subtotal", 6.6, SUBTOTAL_BG)
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

# ── BUFFER ──
write_bucket_header(ws, r, "BUFFER", GRAND_BG)
r += 1
write_row(ws, r, '', "Debugging Surprises & SME Response Delays",
          "Edge cases in country-specific logic, Finance team validation turnaround, unexpected formula behaviour, integration issues.",
          2.2, bg=GRAND_BG)
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
c = ws.cell(row=r, column=4, value="~28.5 weeks")
c.font = Font(name='Arial', size=13, bold=True, color=WHITE)
c.fill = PatternFill('solid', fgColor=DARK_BLUE)
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[r].height = 35
r += 1

ws.merge_cells(f'A{r}:C{r}')
c = ws.cell(row=r, column=1, value="ESTIMATED DURATION")
c.font = Font(name='Arial', size=13, bold=True, color=DARK_BLUE)
c.fill = PatternFill('solid', fgColor=LIGHT_BLUE)
c.alignment = Alignment(horizontal='right', vertical='center')
for col in range(1, 5):
    ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=LIGHT_BLUE)
    ws.cell(row=r, column=col).border = borders
c = ws.cell(row=r, column=4, value="~7 Months")
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
    "1. Estimates based on deep formula analysis of actual .xlsm file: 877 formula templates, 26 Excel functions, 286 calc blocks.",
    "2. 98.5% formula reuse across time periods - only 13 edge-case rows with unique logic.",
    "3. PLW is a single-asset-at-a-time model (421 time columns). Python engine vectorizes across all 78 assets simultaneously.",
    "4. DSCR solver confirmed as simple fixed-point iteration (not Newton-Raphson) - reduces debt block complexity.",
    "5. GTC sheets are low complexity: 142K formula cells but only 16 complex formulas. Mostly SUMIF/SUMIFS aggregation.",
    "6. Revenue engine is the single largest risk item: 200 formula rows with country-specific branching (DE/UK/NL).",
    "7. Assumes Finance team SME availability for validation queries within 2-3 business days.",
    "8. Buffer covers: country-specific edge cases, circular reference convergence tuning, integration testing.",
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

out = r"C:\repos\Ampyr-PFA\Ref Docs\Project_Parthenon_Timeline.xlsx"
wb.save(out)
print(f"Saved to {out}")
