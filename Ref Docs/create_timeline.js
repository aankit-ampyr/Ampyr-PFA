const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");

// ── Colors ──
const AMPYR_BLUE = "1B4F72";
const DARK_BG = "2C3E50";
const MONTH_COLORS = {
  1: "27AE60", 2: "2ECC71",  // Green - Foundation
  3: "2980B9", 4: "3498DB",  // Blue - Reporting + Calc start
  5: "8E44AD", 6: "9B59B6",  // Purple - Calc Engine + Scenario
};
const RISK_RED = "E74C3C";
const RISK_AMBER = "E67E22";
const RISK_GREEN = "27AE60";

// ── Helpers ──
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorderSide = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorderSide, bottom: noBorderSide, left: noBorderSide, right: noBorderSide };
const cellPad = { top: 60, bottom: 60, left: 100, right: 100 };

function hCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || DARK_BG, type: ShadingType.CLEAR },
    margins: cellPad,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 18 })]
    })]
  });
}

function c(text, width, opts = {}) {
  return new TableCell({
    borders: opts.nb ? noBorders : borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellPad,
    verticalAlign: "center",
    columnSpan: opts.span || undefined,
    children: Array.isArray(opts.children) ? opts.children : [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text,
        bold: opts.bold || false,
        font: "Arial",
        size: opts.size || 18,
        color: opts.color || "333333",
        italics: opts.italics || false,
      })]
    })]
  });
}

function spacer(h = 200) {
  return new Paragraph({ spacing: { before: h } });
}

function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 20, ...opts })]
  });
}

function bullet(text, ref = "bullets", level = 0, opts = {}) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20, ...opts })]
  });
}

// ── Gantt-style row ──
function ganttRow(label, months, color) {
  // months is array of 6 booleans
  const cells = [c(label, 2800, { bold: true, size: 16 })];
  for (let i = 0; i < 6; i++) {
    cells.push(c(months[i] ? "" : "", 1080, {
      fill: months[i] ? color : "F8F8F8",
      align: AlignmentType.CENTER,
      size: 16,
      color: months[i] ? "FFFFFF" : "CCCCCC",
      bold: months[i],
    }));
  }
  return new TableRow({ children: cells });
}

// ── Month detail table ──
function monthTable(monthNum, title, color, weeks) {
  const rows = [
    new TableRow({ children: [
      hCell("Week", 1200, { fill: color }),
      hCell("Focus Area", 2500, { fill: color }),
      hCell("Key Deliverables", 3500, { fill: color }),
      hCell("Status Gate", 2160, { fill: color }),
    ]}),
  ];
  for (const w of weeks) {
    rows.push(new TableRow({ children: [
      c(w.week, 1200, { align: AlignmentType.CENTER, bold: true }),
      c(w.focus, 2500),
      c(w.deliverable, 3500),
      c(w.gate, 2160, { italics: true, color: "7F8C8D" }),
    ]}));
  }
  return [
    new Paragraph({
      spacing: { before: 300, after: 120 },
      children: [
        new TextRun({ text: `Month ${monthNum}: `, font: "Arial", size: 24, bold: true, color }),
        new TextRun({ text: title, font: "Arial", size: 24, bold: true, color: "333333" }),
      ]
    }),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1200, 2500, 3500, 2160],
      rows,
    }),
  ];
}

// ── Risk row ──
function riskRow(risk, impact, likelihood, mitigation, color) {
  return new TableRow({ children: [
    c(risk, 2800),
    c(impact, 1200, { align: AlignmentType.CENTER, bold: true, color }),
    c(likelihood, 1200, { align: AlignmentType.CENTER }),
    c(mitigation, 4160),
  ]});
}

// ══════════════════════════════════════════
// DOCUMENT
// ══════════════════════════════════════════

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: AMPYR_BLUE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "34495E" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        ] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [
    // ── COVER PAGE ──
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: [
        spacer(2400),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "PROJECT PARTHENON", font: "Arial", size: 52, bold: true, color: AMPYR_BLUE })]
        }),
        spacer(200),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "6-Month Development Timeline", font: "Arial", size: 28, color: "555555" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Production-Ready Prototype", font: "Arial", size: 24, color: "7F8C8D" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Financial Modeling Platform for Scenario & Sensitivity Analysis", font: "Arial", size: 20, color: "7F8C8D" })]
        }),
        spacer(1200),
        new Table({
          width: { size: 5000, type: WidthType.DXA },
          columnWidths: [2200, 2800],
          rows: [
            new TableRow({ children: [
              c("Prepared for:", 2200, { bold: true, color: "7F8C8D", nb: true }),
              c("Ampyr Energy Tech Solutions", 2800, { bold: true, nb: true }),
            ]}),
            new TableRow({ children: [
              c("Prepared by:", 2200, { bold: true, color: "7F8C8D", nb: true }),
              c("Ankit Agarwal, GM Product & Tech", 2800, { nb: true }),
            ]}),
            new TableRow({ children: [
              c("Date:", 2200, { bold: true, color: "7F8C8D", nb: true }),
              c("March 2026", 2800, { nb: true }),
            ]}),
            new TableRow({ children: [
              c("Duration:", 2200, { bold: true, color: "7F8C8D", nb: true }),
              c("6 Months (April - September 2026)", 2800, { nb: true }),
            ]}),
            new TableRow({ children: [
              c("Classification:", 2200, { bold: true, color: "7F8C8D", nb: true }),
              c("Internal / Confidential", 2800, { color: RISK_RED, nb: true }),
            ]}),
          ],
        }),
      ]
    },

    // ── MAIN CONTENT ──
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AMPYR_BLUE, space: 1 } },
            children: [
              new TextRun({ text: "Project Parthenon", font: "Arial", size: 16, color: AMPYR_BLUE, bold: true }),
              new TextRun({ text: "\t6-Month Timeline", font: "Arial", size: 16, color: "999999" }),
            ],
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Ampyr Energy Tech Solutions  |  Confidential  |  Page ", font: "Arial", size: 14, color: "999999" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 14, color: "999999" }),
            ]
          })]
        })
      },
      children: [
        // ── EXECUTIVE SUMMARY ──
        heading("Executive Summary"),
        body("Project Parthenon replaces a 2-3 hour Excel-based scenario analysis workflow with a Python-powered web application. The platform ingests the quarterly ASE financial model, replicates the full calculation engine in Python, and enables real-time scenario and sensitivity analysis across the entire 78-asset renewable energy portfolio."),
        spacer(100),

        // Key metrics table
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 2340, 2340, 2340],
          rows: [
            new TableRow({ children: [
              hCell("Metric", 2340), hCell("Current State", 2340),
              hCell("Target (Month 6)", 2340), hCell("Improvement", 2340),
            ]}),
            new TableRow({ children: [
              c("Scenario Run Time", 2340, { bold: true }),
              c("2-3 hours", 2340, { color: RISK_RED }),
              c("< 30 seconds", 2340, { color: RISK_GREEN, bold: true }),
              c("360x faster", 2340, { bold: true, color: RISK_GREEN }),
            ]}),
            new TableRow({ children: [
              c("Quarterly Update", 2340, { bold: true }),
              c("2-3 days manual", 2340, { color: RISK_RED }),
              c("< 4 hours", 2340, { color: RISK_GREEN, bold: true }),
              c("90% reduction", 2340, { bold: true, color: RISK_GREEN }),
            ]}),
            new TableRow({ children: [
              c("Concurrent Scenarios", 2340, { bold: true }),
              c("1 at a time", 2340),
              c("Unlimited parallel", 2340, { color: RISK_GREEN, bold: true }),
              c("N/A", 2340),
            ]}),
            new TableRow({ children: [
              c("Asset Coverage", 2340, { bold: true }),
              c("78 assets (all recalc)", 2340),
              c("78 assets (selective)", 2340, { bold: true }),
              c("Only changed assets recalculate", 2340, { size: 16 }),
            ]}),
          ]
        }),

        spacer(100),
        heading("Scope & Assumptions", HeadingLevel.HEADING_2),
        bullet("Solo developer (GM Product & Tech) with AI-assisted development, ~20 hrs/week dedicated"),
        bullet("Tech stack: Python FastAPI + PostgreSQL + Streamlit"),
        bullet("Full calculation engine replication (all 1,609 rows of Project Level Workings)"),
        bullet("All 6 scenario levers: Devex, Capex, Opex, Production, Revenue, Financing"),
        bullet("All 21 GTC reporting sheets included"),
        bullet("Quarterly file comparison and update workflow"),
        bullet("Total estimated effort: ~480 development hours across 24 weeks"),

        new Paragraph({ children: [new PageBreak()] }),

        // ── VISUAL TIMELINE ──
        heading("Visual Timeline Overview"),
        body("The 6-month plan is organized into 3 phases with overlapping workstreams:"),
        spacer(100),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2880, 1080, 1080, 1080, 1080, 1080, 1080],
          rows: [
            new TableRow({ children: [
              hCell("Workstream", 2880),
              hCell("M1", 1080, { align: AlignmentType.CENTER }),
              hCell("M2", 1080, { align: AlignmentType.CENTER }),
              hCell("M3", 1080, { align: AlignmentType.CENTER }),
              hCell("M4", 1080, { align: AlignmentType.CENTER }),
              hCell("M5", 1080, { align: AlignmentType.CENTER }),
              hCell("M6", 1080, { align: AlignmentType.CENTER }),
            ]}),
            // Phase labels
            new TableRow({ children: [
              c("", 2880, { fill: "F0F0F0" }),
              c("Phase 1: Foundation", 2160, { span: 2, fill: "D5F5E3", bold: true, align: AlignmentType.CENTER, size: 16 }),
              c("Phase 2: Calc Engine", 2160, { span: 2, fill: "D6EAF8", bold: true, align: AlignmentType.CENTER, size: 16 }),
              c("Phase 3: Scenario + Polish", 2160, { span: 2, fill: "E8DAEF", bold: true, align: AlignmentType.CENTER, size: 16 }),
            ]}),
            ganttRow("Database & Ingestion",        [true,  true,  false, false, false, false], "27AE60"),
            ganttRow("Quarterly Diff Engine",        [false, true,  true,  false, false, false], "2ECC71"),
            ganttRow("Streamlit UI Shell",           [true,  true,  false, false, false, true ], "27AE60"),
            ganttRow("Calc: Production & Revenue",   [false, false, true,  true,  false, false], "2980B9"),
            ganttRow("Calc: OpEx / Tax / Depr",      [false, false, false, true,  false, false], "3498DB"),
            ganttRow("Calc: Debt & DSCR Solver",     [false, false, false, true,  true,  false], "8E44AD"),
            ganttRow("Calc: IRR & Waterfall",        [false, false, false, false, true,  false], "9B59B6"),
            ganttRow("Scenario Engine (6 levers)",   [false, false, false, false, true,  true ], "8E44AD"),
            ganttRow("GTC Reporting Sheets (21)",    [false, false, false, false, false, true ], "9B59B6"),
            ganttRow("Validation & UAT",             [false, false, false, false, true,  true ], "E67E22"),
          ]
        }),

        spacer(100),
        body("M1 = Month 1 (April 2026), M6 = Month 6 (September 2026)", { size: 16, color: "999999", italics: true }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── MONTH-BY-MONTH DETAIL ──
        heading("Month-by-Month Breakdown"),

        // ── MONTH 1 ──
        ...monthTable(1, "Foundation & Data Layer", MONTH_COLORS[1], [
          { week: "W1-W2", focus: "Database Design", deliverable: "PostgreSQL schema: assets, parameters, time_series, quarterly_metrics, model_versions, scenarios", gate: "Schema review" },
          { week: "W2-W3", focus: "Excel Ingestion", deliverable: "Ingest .xlsm/.xlsb into DB: Project Info, Country Inputs, Financing Inputs, Time Inputs (A/M/Q)", gate: "78 assets loaded" },
          { week: "W3-W4", focus: "Streamlit Shell", deliverable: "Asset browser, metric explorer, file upload UI, basic authentication", gate: "Demo to Finance" },
        ]),
        spacer(80),
        body("Key milestone: Upload the ASE Excel file and browse all 78 assets with their parameters in the web UI.", { bold: true }),

        spacer(200),

        // ── MONTH 2 ──
        ...monthTable(2, "Quarterly Output Ingestion & Diff Engine", MONTH_COLORS[2], [
          { week: "W5-W6", focus: "Output Ingestion", deliverable: "Parse Quarterly Output (36K rows x 433 cols), normalize 170 metrics per asset into quarterly_metrics table", gate: "3.1M cells loaded" },
          { week: "W6-W7", focus: "Structural Fingerprint", deliverable: "Automated structural comparison engine using STRUCTURAL_MAP baseline. Detect row insertions, parameter changes, new assets", gate: "Diff report generated" },
          { week: "W7-W8", focus: "Version Management", deliverable: "Quarter-over-quarter comparison dashboard, change heatmaps, approval workflow for new baseline", gate: "Compare 2 quarters" },
        ]),
        spacer(80),
        body("Key milestone: Upload a new quarterly file, see exactly what changed vs. the previous quarter, approve as new baseline.", { bold: true }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── MONTH 3 ──
        ...monthTable(3, "Calc Engine - Production & Revenue", MONTH_COLORS[3], [
          { week: "W9-W10", focus: "Formula Extraction", deliverable: "Extract all formulas from Project Level Workings .xlsm, map dependency graph, identify calculation blocks", gate: "Block map complete" },
          { week: "W10-W11", focus: "Flags & Production", deliverable: "Python modules: project flags/timings, construction schedule, net production (P50/P90), degradation, curtailment", gate: "<0.01% vs Excel" },
          { week: "W11-W12", focus: "Revenue Engine", deliverable: "PPA revenue, merchant revenue, RoC/CfD, FiT/FiP, indexation logic, country-specific contract types (348 formula rows)", gate: "<0.01% vs Excel" },
        ]),
        spacer(80),
        body("Key milestone: Python-calculated Production and Revenue match Excel output within 0.01% for all 78 assets.", { bold: true }),
        body("Risk flag: Revenue engine is the most complex block (348 rows of country-specific logic). May extend into Month 4.", { color: RISK_AMBER, italics: true }),

        spacer(200),

        // ── MONTH 4 ──
        ...monthTable(4, "Calc Engine - Costs, Tax & Debt", MONTH_COLORS[4], [
          { week: "W13-W14", focus: "OpEx & EBITDA", deliverable: "Operating expenditure module, land lease, insurance, O&M, management fees, EBITDA calculation", gate: "<0.01% vs Excel" },
          { week: "W14-W15", focus: "Capex & Tax", deliverable: "Devex/Capex schedule, depreciation (country-specific), corporate tax, VAT, working capital", gate: "<0.01% vs Excel" },
          { week: "W15-W16", focus: "Debt & DSCR", deliverable: "Senior debt drawdown/repayment, DSCR sculpting with Newton-Raphson iterative solver, reserve accounts", gate: "Circular ref solved" },
        ]),
        spacer(80),
        body("Key milestone: Complete financial cascade from Production through to Debt service for all assets.", { bold: true }),
        body("Risk flag: DSCR circular reference solver is technically the hardest component. Fallback: use fixed-point iteration if Newton-Raphson convergence is problematic.", { color: RISK_AMBER, italics: true }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── MONTH 5 ──
        ...monthTable(5, "Scenario Engine & Validation", MONTH_COLORS[5], [
          { week: "W17-W18", focus: "IRR & Distribution", deliverable: "Distribution waterfall, equity returns, XIRR calculation, HoldCo aggregation, financial statements assembly", gate: "Full cascade validated" },
          { week: "W18-W19", focus: "Scenario Framework", deliverable: "6 scenario levers (Devex/Capex/Opex/Production/Revenue/Financing), parameter override system, selective recalculation via dependency graph", gate: "Scenario runs <30s" },
          { week: "W19-W20", focus: "Sensitivity Analysis", deliverable: "Single-variable sensitivity tables, multi-variable tornado charts, scenario comparison side-by-side, parallel scenario execution", gate: "Demo to Finance" },
        ]),
        spacer(80),
        body("Key milestone: Run a complete scenario changing 3 asset production assumptions and see updated IRR/DSCR/NPV in under 30 seconds.", { bold: true }),

        spacer(200),

        // ── MONTH 6 ──
        ...monthTable(6, "GTC Reporting, Polish & Launch", MONTH_COLORS[6], [
          { week: "W21-W22", focus: "GTC Reporting Sheets", deliverable: "All 21 GTC sheets: PnL/BS projections, CF Capital Strategy, Valuation, OpCo Performance, Dev Pipeline, QREP, GTC Summary", gate: "Reports match Excel" },
          { week: "W22-W23", focus: "Export & Integration", deliverable: "Excel report generation (audit-compliant), scenario comparison exports, PDF summary reports", gate: "Export validated" },
          { week: "W23-W24", focus: "Production Readiness", deliverable: "Error handling, logging, deployment config, user documentation, UAT with Finance team, bug fixes", gate: "Finance sign-off" },
        ]),
        spacer(80),
        body("Key milestone: Finance team runs a full quarterly update + scenario analysis cycle end-to-end without touching Excel.", { bold: true }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── DELIVERABLES SUMMARY ──
        heading("Deliverables Summary by Month"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [1400, 3200, 2600, 2160],
          rows: [
            new TableRow({ children: [
              hCell("Month", 1400), hCell("Core Deliverable", 3200),
              hCell("Demo-able Feature", 2600), hCell("Cumulative %", 2160),
            ]}),
            new TableRow({ children: [
              c("Month 1", 1400, { bold: true, fill: "D5F5E3" }),
              c("Database + Ingestion + UI Shell", 3200),
              c("Browse 78 assets in web UI", 2600),
              c("15%", 2160, { align: AlignmentType.CENTER, bold: true }),
            ]}),
            new TableRow({ children: [
              c("Month 2", 1400, { bold: true, fill: "D5F5E3" }),
              c("Quarterly Output + Diff Engine", 3200),
              c("Upload new file, see changes", 2600),
              c("30%", 2160, { align: AlignmentType.CENTER, bold: true }),
            ]}),
            new TableRow({ children: [
              c("Month 3", 1400, { bold: true, fill: "D6EAF8" }),
              c("Calc: Production + Revenue", 3200),
              c("Python matches Excel for P&R", 2600),
              c("50%", 2160, { align: AlignmentType.CENTER, bold: true }),
            ]}),
            new TableRow({ children: [
              c("Month 4", 1400, { bold: true, fill: "D6EAF8" }),
              c("Calc: Costs + Tax + Debt/DSCR", 3200),
              c("Full financial cascade works", 2600),
              c("70%", 2160, { align: AlignmentType.CENTER, bold: true }),
            ]}),
            new TableRow({ children: [
              c("Month 5", 1400, { bold: true, fill: "E8DAEF" }),
              c("Scenario Engine + Sensitivity", 3200),
              c("Run scenarios in <30 seconds", 2600),
              c("85%", 2160, { align: AlignmentType.CENTER, bold: true }),
            ]}),
            new TableRow({ children: [
              c("Month 6", 1400, { bold: true, fill: "E8DAEF" }),
              c("GTC Reports + Production Launch", 3200),
              c("Full quarterly cycle without Excel", 2600),
              c("100%", 2160, { align: AlignmentType.CENTER, bold: true }),
            ]}),
          ]
        }),

        spacer(200),

        // ── RISK REGISTER ──
        heading("Risk Register"),
        body("The following risks have been identified with the aggressive 6-month timeline. Each includes a mitigation strategy."),
        spacer(100),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 1200, 1200, 4160],
          rows: [
            new TableRow({ children: [
              hCell("Risk", 2800), hCell("Impact", 1200, { align: AlignmentType.CENTER }),
              hCell("Likelihood", 1200, { align: AlignmentType.CENTER }), hCell("Mitigation", 4160),
            ]}),
            riskRow(
              "Revenue engine complexity exceeds estimate (348 rows, country-specific logic)",
              "HIGH", "Medium",
              "Start revenue block first in Month 3. If delayed, defer 2-3 edge-case contract types to post-launch update.",
              RISK_RED
            ),
            riskRow(
              "DSCR circular reference solver convergence issues",
              "HIGH", "Low-Med",
              "Implement both Newton-Raphson and fixed-point iteration. Use Excel COM (xlwings) as validation oracle during development.",
              RISK_RED
            ),
            riskRow(
              "Part-time bandwidth insufficient for full calc engine",
              "HIGH", "Medium",
              "Months 3-4 are the critical path. If behind schedule by Week 14, recommend adding a contract Python developer for calc engine blocks.",
              RISK_RED
            ),
            riskRow(
              "Numerical divergence between Python and Excel (>0.01%)",
              "Medium", "Medium",
              "Block-by-block validation against Excel at each stage. Accept <0.1% divergence for edge cases. Log all divergences for review.",
              RISK_AMBER
            ),
            riskRow(
              "Quarterly file structure changes break ingestion pipeline",
              "Medium", "Low",
              "STRUCTURAL_MAP baseline with automated drift detection. Alert on any structural change before processing.",
              RISK_AMBER
            ),
            riskRow(
              "Streamlit performance with large datasets (3.1M cells)",
              "Low", "Medium",
              "Pre-aggregate data at ingestion time. Use server-side pagination. Cache frequently accessed queries in Redis if needed.",
              RISK_GREEN
            ),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── CRITICAL PATH ──
        heading("Critical Path & Decision Points"),
        body("The following checkpoints require management decisions:"),
        spacer(100),

        heading("End of Month 2 - Go/No-Go for Calc Engine", HeadingLevel.HEADING_2),
        bullet("Foundation platform is working: data ingestion, quarterly comparison, basic UI"),
        bullet("Decision: Proceed with in-house calc engine build vs. exploring Excel-as-a-Service alternative"),
        bullet("If calc engine build is approved, Months 3-4 become the critical path"),

        spacer(100),
        heading("End of Month 4 - Resource Assessment", HeadingLevel.HEADING_2),
        bullet("Full financial cascade (Production through Debt) should be complete"),
        bullet("If behind schedule: recommend hiring a contract Python developer for remaining blocks"),
        bullet("If on track: continue solo for scenario engine and GTC reporting"),

        spacer(100),
        heading("End of Month 5 - UAT Readiness", HeadingLevel.HEADING_2),
        bullet("Scenario engine functional: Finance team begins parallel testing (Excel vs. App)"),
        bullet("Decision: Launch timeline - immediate rollout vs. extended parallel-run period"),
        bullet("GTC reporting sheets may shift to Month 7 if core scenarios take priority"),

        spacer(200),

        // ── TECH STACK ──
        heading("Technology Stack"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2200, 3200, 3960],
          rows: [
            new TableRow({ children: [
              hCell("Component", 2200), hCell("Technology", 3200), hCell("Rationale", 3960),
            ]}),
            new TableRow({ children: [
              c("Backend API", 2200, { bold: true }),
              c("Python FastAPI", 3200),
              c("Async, high performance, auto-docs, type safety", 3960),
            ]}),
            new TableRow({ children: [
              c("Database", 2200, { bold: true }),
              c("PostgreSQL", 3200),
              c("JSONB for flexible params, table partitioning for 3.1M+ cells", 3960),
            ]}),
            new TableRow({ children: [
              c("Calc Engine", 2200, { bold: true }),
              c("Python + NumPy", 3200),
              c("Vectorized computation, selective recalc via dependency graph", 3960),
            ]}),
            new TableRow({ children: [
              c("Frontend", 2200, { bold: true }),
              c("Streamlit", 3200),
              c("Rapid prototyping, Python-native, low frontend overhead", 3960),
            ]}),
            new TableRow({ children: [
              c("Excel Parsing", 2200, { bold: true }),
              c("openpyxl + pyxlsb", 3200),
              c("Formula extraction (.xlsm) + cached values (.xlsb)", 3960),
            ]}),
            new TableRow({ children: [
              c("Deployment", 2200, { bold: true }),
              c("Docker + internal hosting", 3200),
              c("Containerized for consistency, internal deployment", 3960),
            ]}),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ── RESOURCE REQUIREMENTS ──
        heading("Resource Requirements"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2400, 2400, 2400, 2160],
          rows: [
            new TableRow({ children: [
              hCell("Resource", 2400), hCell("Commitment", 2400),
              hCell("Period", 2400), hCell("Status", 2160),
            ]}),
            new TableRow({ children: [
              c("GM Product & Tech", 2400, { bold: true }),
              c("20 hrs/week dedicated", 2400),
              c("Months 1-6 (full duration)", 2400),
              c("Confirmed", 2160, { color: RISK_GREEN, bold: true }),
            ]}),
            new TableRow({ children: [
              c("AI Development Assistant", 2400, { bold: true }),
              c("Claude Code (AI pair programming)", 2400),
              c("Months 1-6 (full duration)", 2400),
              c("Confirmed", 2160, { color: RISK_GREEN, bold: true }),
            ]}),
            new TableRow({ children: [
              c("Finance Team (SME)", 2400, { bold: true }),
              c("4-6 hrs/week for validation", 2400),
              c("Months 3-6 (calc validation)", 2400),
              c("To be confirmed", 2160, { color: RISK_AMBER }),
            ]}),
            new TableRow({ children: [
              c("Contract Python Dev (contingency)", 2400, { bold: true }),
              c("Full-time if triggered", 2400),
              c("Month 5-6 (if behind schedule)", 2400),
              c("Contingency only", 2160, { color: "7F8C8D", italics: true }),
            ]}),
            new TableRow({ children: [
              c("Infrastructure", 2400, { bold: true }),
              c("PostgreSQL server + app hosting", 2400),
              c("Month 1 onwards", 2400),
              c("To be provisioned", 2160, { color: RISK_AMBER }),
            ]}),
          ]
        }),

        spacer(200),

        // ── SUCCESS CRITERIA ──
        heading("Success Criteria"),
        body("The prototype will be considered successful when:"),
        spacer(50),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [600, 5760, 3000],
          rows: [
            new TableRow({ children: [
              hCell("#", 600, { align: AlignmentType.CENTER }), hCell("Criterion", 5760), hCell("Measurement", 3000),
            ]}),
            ...[
              ["1", "Scenario analysis completes in <30 seconds for any single lever change", "Timed test: change Production for 3 assets"],
              ["2", "Python calc engine matches Excel output within 0.01% for all 170 metrics", "Automated validation suite across all 78 assets"],
              ["3", "Quarterly file upload + comparison completes in <4 hours (vs. 2-3 days manual)", "End-to-end quarterly update cycle"],
              ["4", "All 6 scenario levers functional: Devex, Capex, Opex, Production, Revenue, Financing", "Finance team runs each lever successfully"],
              ["5", "21 GTC reporting sheets generated from app match Excel baseline", "Side-by-side comparison report"],
              ["6", "Finance team can run parallel analysis (Excel vs App) for one full quarter", "No blockers during parallel run"],
              ["7", "Selective recalculation: changing 3 assets does not recalculate remaining 75", "Performance log shows targeted recalc"],
              ["8", "Quarterly structural drift detection catches >95% of file changes automatically", "Test with known-changed quarterly files"],
            ].map(([n, criterion, measure]) => new TableRow({ children: [
              c(n, 600, { align: AlignmentType.CENTER, bold: true }),
              c(criterion, 5760),
              c(measure, 3000, { size: 16, color: "555555" }),
            ]}))
          ]
        }),

        spacer(200),

        // ── NEXT STEPS ──
        heading("Immediate Next Steps"),
        body("Upon approval of this timeline:"),
        spacer(50),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [new TextRun({ text: "Provision PostgreSQL database and application hosting environment", font: "Arial", size: 20 })]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [new TextRun({ text: "Confirm Finance team availability for validation support (Months 3-6)", font: "Arial", size: 20 })]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [new TextRun({ text: "Begin Month 1: Database schema design and Excel ingestion pipeline", font: "Arial", size: 20 })]
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [new TextRun({ text: "Schedule bi-weekly progress demos with stakeholders", font: "Arial", size: 20 })]
        }),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Project_Parthenon_6Month_Timeline.docx", buffer);
  console.log("Document created: Project_Parthenon_6Month_Timeline.docx");
});
