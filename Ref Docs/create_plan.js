const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");

// ── Colors ──
const AMPYR_BLUE = "1B4F72";
const AMPYR_LIGHT = "D6EAF8";
const AMPYR_MID = "85C1E9";
const HEADER_BG = "2C3E50";
const PHASE_COLORS = {
  1: "27AE60", // Green - Foundation
  2: "2980B9", // Blue - Reporting
  3: "8E44AD", // Purple - Calc Engine Core
  4: "E67E22", // Orange - Calc Engine Finance
  5: "E74C3C", // Red - Scenario Engine
  6: "16A085", // Teal - Satellite
};

// ── Helpers ──
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

function headerCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || HEADER_BG, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 18 })]
    })]
  });
}

function cell(text, width, opts = {}) {
  return new TableCell({
    borders: opts.noBorders ? noBorders : borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
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

function sectionTitle(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 20, ...opts })]
  });
}

function bulletItem(text, ref = "bullets", level = 0) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })]
  });
}

function phaseHeader(phase, title, months, color) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [400, 7760, 1200],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: 400, type: WidthType.DXA },
          shading: { fill: color, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [] })]
        }),
        new TableCell({
          borders: noBorders,
          width: { size: 7760, type: WidthType.DXA },
          shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 100 },
          children: [new Paragraph({
            children: [
              new TextRun({ text: `PHASE ${phase}: `, bold: true, font: "Arial", size: 28, color }),
              new TextRun({ text: title, bold: true, font: "Arial", size: 28, color: "2C3E50" }),
            ]
          })]
        }),
        new TableCell({
          borders: noBorders,
          width: { size: 1200, type: WidthType.DXA },
          shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: months, font: "Arial", size: 20, color: "7F8C8D", italics: true })]
          })]
        }),
      ]
    })]
  });
}

function monthTable(rows) {
  const colWidths = [1200, 3660, 2800, 1700];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: [
          headerCell("Week", colWidths[0]),
          headerCell("Deliverable", colWidths[1]),
          headerCell("Key Files / Components", colWidths[2]),
          headerCell("Milestone", colWidths[3]),
        ]
      }),
      ...rows.map((r, i) => new TableRow({
        children: [
          cell(r[0], colWidths[0], { bold: true, fill: i % 2 === 0 ? "F8F9FA" : undefined }),
          cell(r[1], colWidths[1], { fill: i % 2 === 0 ? "F8F9FA" : undefined }),
          cell(r[2], colWidths[2], { size: 16, color: "7F8C8D", fill: i % 2 === 0 ? "F8F9FA" : undefined }),
          cell(r[3], colWidths[3], { bold: true, color: "27AE60", fill: i % 2 === 0 ? "F8F9FA" : undefined }),
        ]
      }))
    ]
  });
}

function spacer(pts = 200) {
  return new Paragraph({ spacing: { after: pts }, children: [] });
}

// ── Build Document ──
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: AMPYR_BLUE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AMPYR_MID, space: 4 } } }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "34495E" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        ]
      },
      {
        reference: "numbers",
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ]
      },
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
          spacing: { after: 200 },
          children: [new TextRun({ text: "PROJECT PARTHENON", font: "Arial", size: 52, bold: true, color: AMPYR_BLUE })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: AMPYR_MID, space: 8 } },
          children: [new TextRun({ text: "Development Plan", font: "Arial", size: 36, color: "2C3E50" })]
        }),
        spacer(200),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Financial Modeling Platform", font: "Arial", size: 24, color: "7F8C8D" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Month-by-Month Implementation Roadmap", font: "Arial", size: 24, color: "7F8C8D" })]
        }),
        spacer(1600),
        new Table({
          width: { size: 5000, type: WidthType.DXA },
          columnWidths: [2200, 2800],
          rows: [
            new TableRow({ children: [
              cell("Prepared for:", 2200, { bold: true, color: "7F8C8D", noBorders: true }),
              cell("Ampyr Energy Tech Solutions", 2800, { bold: true, noBorders: true }),
            ]}),
            new TableRow({ children: [
              cell("Prepared by:", 2200, { bold: true, color: "7F8C8D", noBorders: true }),
              cell("GTC Product & Technology", 2800, { noBorders: true }),
            ]}),
            new TableRow({ children: [
              cell("Date:", 2200, { bold: true, color: "7F8C8D", noBorders: true }),
              cell("March 2026", 2800, { noBorders: true }),
            ]}),
            new TableRow({ children: [
              cell("Version:", 2200, { bold: true, color: "7F8C8D", noBorders: true }),
              cell("1.0", 2800, { noBorders: true }),
            ]}),
            new TableRow({ children: [
              cell("Classification:", 2200, { bold: true, color: "7F8C8D", noBorders: true }),
              cell("Internal / Confidential", 2800, { color: "E74C3C", noBorders: true }),
            ]}),
          ]
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
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: AMPYR_MID, space: 4 } },
            children: [
              new TextRun({ text: "Project Parthenon", font: "Arial", size: 16, color: "7F8C8D", italics: true }),
              new TextRun({ text: "\tDevelopment Plan", font: "Arial", size: 16, color: "7F8C8D", italics: true }),
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
              new TextRun({ text: "Ampyr Energy Tech Solutions | Confidential | Page ", font: "Arial", size: 14, color: "95A5A6" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 14, color: "95A5A6" }),
            ]
          })]
        })
      },
      children: [
        // ── EXECUTIVE SUMMARY ──
        sectionTitle("Executive Summary"),
        bodyText("Project Parthenon transforms Ampyr GTC\u2019s financial modeling workflow from a 2\u20133 hour Excel-based process into a real-time web application. The platform ingests the core ASE financial model (120 asset slots, 170 metrics, 35-year horizon), enables instant scenario and sensitivity analysis, and automates quarterly model updates."),
        spacer(100),

        // Key metrics table
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2340, 2340, 2340, 2340],
          rows: [
            new TableRow({ children: [
              headerCell("Assets", 2340, { align: AlignmentType.CENTER }),
              headerCell("Current Runtime", 2340, { align: AlignmentType.CENTER }),
              headerCell("Target Runtime", 2340, { align: AlignmentType.CENTER }),
              headerCell("Timeline", 2340, { align: AlignmentType.CENTER }),
            ]}),
            new TableRow({ children: [
              cell("78 real + 42 placeholder", 2340, { align: AlignmentType.CENTER }),
              cell("2\u20133 hours / scenario", 2340, { align: AlignmentType.CENTER, color: "E74C3C", bold: true }),
              cell("< 5 seconds / scenario", 2340, { align: AlignmentType.CENTER, color: "27AE60", bold: true }),
              cell("12 months (6 phases)", 2340, { align: AlignmentType.CENTER }),
            ]}),
          ]
        }),
        spacer(100),
        bodyText("Tech Stack: Python FastAPI + PostgreSQL + Streamlit", { bold: true, color: AMPYR_BLUE }),
        bodyText("Approach: Hybrid (Option C) \u2014 Ingest Excel output values immediately for reporting; build Python calculation engine incrementally, validated against Excel at every step."),

        // ── TIMELINE OVERVIEW ──
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Timeline Overview"),
        bodyText("The project is divided into 6 phases, each delivering independent value. The first usable product ships at the end of Month 2."),
        spacer(100),

        // Timeline summary table
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [500, 2760, 2000, 2000, 2100],
          rows: [
            new TableRow({ children: [
              headerCell("#", 500, { align: AlignmentType.CENTER }),
              headerCell("Phase", 2760),
              headerCell("Duration", 2000),
              headerCell("Months", 2000),
              headerCell("Key Outcome", 2100),
            ]}),
            ...[
              ["1", "Foundation", "8 weeks", "M1\u2013M2", "Read-only platform"],
              ["2", "Reporting Engine", "6 weeks", "M3\u2013M4", "Auto-generated reports"],
              ["3", "Calc Engine: Core", "12 weeks", "M4\u2013M7", "P&L recalculation"],
              ["4", "Calc Engine: Finance", "8 weeks", "M7\u2013M9", "Full model parity"],
              ["5", "Scenario Engine", "6 weeks", "M9\u2013M10", "Real-time scenarios"],
              ["6", "Satellite Frontends", "8 weeks", "M11\u2013M12", "Multi-team platform"],
            ].map((r, i) => new TableRow({
              children: [
                cell(r[0], 500, { align: AlignmentType.CENTER, bold: true, color: PHASE_COLORS[parseInt(r[0])] }),
                cell(r[1], 2760, { bold: true }),
                cell(r[2], 2000),
                cell(r[3], 2000, { bold: true }),
                cell(r[4], 2100, { italics: true }),
              ].map(c => { if (i % 2 === 0) c.shading = { fill: "F8F9FA", type: ShadingType.CLEAR }; return c; })
            }))
          ]
        }),

        // ════════════════════════════════════════════
        // PHASE 1
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Phase 1: Foundation"),
        phaseHeader(1, "Read-Only Platform", "M1\u2013M2", PHASE_COLORS[1]),
        spacer(100),
        bodyText("Goal: Ingest the core Excel model into PostgreSQL and build a browsable, queryable platform with quarterly diff capabilities. This phase delivers the first usable product."),
        spacer(100),

        sectionTitle("Month 1: Database & Ingestion", HeadingLevel.HEADING_2),
        bodyText("Stand up the project infrastructure, define the database schema, and build the Excel ingestion pipeline."),
        spacer(80),
        monthTable([
          ["W1\u2013W2", "Project setup: FastAPI skeleton, PostgreSQL, Alembic migrations, Docker Compose", "docker-compose.yml, alembic/, app/models/", "Dev env ready"],
          ["W2\u2013W3", "Database schema: assets, asset_parameters, time_series_inputs, country_parameters, financing_terms, model_versions", "app/models/*.py, migrations/", "Schema deployed"],
          ["W3\u2013W4", "Excel ingestion pipeline: pyxlsb/openpyxl reader for all 6 input sheets + Quarterly Output", "app/ingestion/xlsb_reader.py, normalizer.py", "Full ingestion"],
          ["W4", "Version management: upload tracking, file hashing (SHA-256), baseline locking", "app/api/versions.py", "Version system live"],
        ]),
        spacer(100),

        sectionTitle("Month 2: Diff Engine & Dashboard", HeadingLevel.HEADING_2),
        bodyText("Build the quarterly comparison engine and the first Streamlit UI for browsing assets and viewing changes."),
        spacer(80),
        monthTable([
          ["W5\u2013W6", "Diff engine: structural diff (sheet counts, row labels, section boundaries) + value diff (parameter deltas, metric changes)", "app/ingestion/differ.py", "Auto-diff works"],
          ["W6\u2013W7", "Streamlit dashboard: asset browser (filter by country/tech), metric explorer, quarter-over-quarter comparison", "app/ui/dashboard.py, asset_browser.py", "UI live"],
          ["W7\u2013W8", "Diff visualization: change heatmaps, drill-down to individual cells, anomaly flagging (changes > threshold)", "app/ui/diff_viewer.py", ""],
          ["W8", "Integration testing, structural map validation against both quarterly files, performance tuning", "tests/", "Phase 1 complete"],
        ]),
        spacer(80),
        sectionTitle("Phase 1 Deliverables", HeadingLevel.HEADING_3),
        bulletItem("Upload a quarterly .xlsb/.xlsm and have it fully ingested into PostgreSQL in under 5 minutes"),
        bulletItem("Browse all 78 real assets with their 590+ parameters, filterable by country and technology"),
        bulletItem("View 170 quarterly metrics for any asset across all time periods"),
        bulletItem("Compare two quarterly versions: see exactly what changed (parameters, curves, metrics)"),
        bulletItem("Structural validation: flag new sheets, row insertions, named range shifts"),

        // ════════════════════════════════════════════
        // PHASE 2
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Phase 2: Reporting Engine"),
        phaseHeader(2, "Automated Report Generation", "M3\u2013M4", PHASE_COLORS[2]),
        spacer(100),
        bodyText("Goal: Replace the 21 manually maintained GTC Excel sheets with on-demand report generation from the database. Eliminates the GTC team\u2019s manual sheet maintenance work."),
        spacer(100),

        sectionTitle("Month 3: Report Templates & FX Engine", HeadingLevel.HEADING_2),
        monthTable([
          ["W9\u2013W10", "GTC report templates: match formatting of PnL projections, BS projections, CF Capital Strategy, Quarterly Reports (USD/EUR)", "app/reports/gtc_generator.py, templates/", "Templates ready"],
          ["W10\u2013W11", "FX conversion engine: EUR-to-USD using FX rates from Time Inputs (A). AGP proportionate share calculation (ownership %)", "app/reports/fx_engine.py", "FX engine live"],
          ["W11\u2013W12", "Fiscal year mapping: FY26A+F, FY26B, FY27B, FY28F, FY29F. Map quarterly metrics to fiscal year buckets", "app/reports/fiscal_mapper.py", "Fiscal mapping"],
        ]),
        spacer(100),

        sectionTitle("Month 4: Report Generation API & Comparison Reports", HeadingLevel.HEADING_2),
        monthTable([
          ["W13", "Report generation API: generate all 21 GTC-equivalent sheets as .xlsx on demand via FastAPI endpoint", "app/api/reports.py", "API live"],
          ["W14", "Comparison reports: side-by-side version comparison with delta highlighting. Multi-scenario comparison template", "app/reports/comparison.py", "Comparison reports"],
          ["W14", "Streamlit report builder: select version, scenario, output format. Download generated Excel with audit metadata", "app/ui/report_builder.py", "Phase 2 complete"],
        ]),
        spacer(80),
        sectionTitle("Phase 2 Deliverables", HeadingLevel.HEADING_3),
        bulletItem("Generate all 21 GTC reporting sheets (PnL, BS, CF, QRep, Valuation, Cash Breakeven) on demand"),
        bulletItem("Reports include audit trail: version_id, generation timestamp, source data hash"),
        bulletItem("FX conversion (EUR/GBP to USD) applied automatically"),
        bulletItem("Side-by-side comparison reports between any two quarterly versions"),

        // ════════════════════════════════════════════
        // PHASE 3
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Phase 3: Calculation Engine \u2014 Core"),
        phaseHeader(3, "Python Financial Model (P&L Blocks)", "M4\u2013M7", PHASE_COLORS[3]),
        spacer(100),
        bodyText("Goal: Rebuild the core P&L calculation blocks from Project Level Workings in Python. Each block is validated against Excel output for all 120 assets with a <0.01% divergence target."),
        spacer(100),
        bodyText("This phase covers rows 136\u2013927 of Project Level Workings: Flags, Production, Revenue (the most complex block at 348 rows), OpEx, Depreciation, VAT, and Working Capital."),
        spacer(100),

        sectionTitle("Month 5: Flags, Production & Revenue Foundation", HeadingLevel.HEADING_2),
        monthTable([
          ["W15\u2013W16", "Flags & Timings engine (PLW R136\u2013R214): Solar/BESS on/off, inflation flags, construction/operation periods, FC dates", "app/engine/flags.py", "Flags validated"],
          ["W17\u2013W18", "Production engine (PLW R247\u2013R272): Solar degradation, BESS capacity/degradation, availability, seasonality", "app/engine/production.py", "Production validated"],
          ["W18\u2013W19", "Revenue foundation: Generation waterfall (PLW R276\u2013R310), PPA Contract 1 with CfD logic (PLW R312\u2013R342)", "app/engine/revenue.py", "Revenue started"],
        ]),
        spacer(100),

        sectionTitle("Month 6: Revenue Engine (Country-Specific Contracts)", HeadingLevel.HEADING_2),
        bodyText("The revenue engine is the most complex block \u2014 348 rows covering 6 contract types across 3 countries. This month completes the full revenue waterfall."),
        spacer(80),
        monthTable([
          ["W19\u2013W20", "PPA Contract 2: Floor PPA, SDE subsidy, Fixed Price PPA with EEG switch. SDE++ scheme (PLW R388\u2013R423)", "app/engine/revenue.py", ""],
          ["W20\u2013W21", "EEG contracts (PLW R425\u2013R447), GoO/REGO certificates: contracted GoO 1&2 + merchant GoO (PLW R472\u2013R499)", "app/engine/revenue.py", ""],
          ["W21\u2013W22", "BESS Revenue: Tolling, Floor, Capacity Market, CIDT, Merchant (PLW R515\u2013R591). Revenue totals & 10-year realized prices", "app/engine/revenue_bess.py", "Revenue validated"],
        ]),
        spacer(100),

        sectionTitle("Month 7: OpEx, Depreciation, VAT & Working Capital", HeadingLevel.HEADING_2),
        monthTable([
          ["W22\u2013W23", "OpEx engine (PLW R624\u2013R713): Solar O&M (period 1/2), land lease (fixed + revenue-dependent), insurance, BESS OpEx", "app/engine/opex.py", "OpEx validated"],
          ["W23\u2013W24", "Depreciation engine (PLW R715\u2013R810): 6 Solar asset categories + BESS, tax depreciation, accelerated depreciation", "app/engine/depreciation.py", "D&A validated"],
          ["W24\u2013W25", "VAT engine (PLW R812\u2013R882): VAT on CapEx/Revenue/OpEx, VAT facility, refund cycle. Working Capital (PLW R884\u2013R927)", "app/engine/vat.py, working_cap.py", ""],
          ["W25\u2013W26", "Integration: assemble Income Statement (Revenue \u2212 OpEx = EBITDA \u2212 D&A = EBIT). Validate full P&L against Excel for all 120 assets", "app/engine/statements.py", "Phase 3 complete"],
        ]),
        spacer(80),
        sectionTitle("Phase 3 Deliverables", HeadingLevel.HEADING_3),
        bulletItem("Python calculation engine covering Flags \u2192 Production \u2192 Revenue \u2192 OpEx \u2192 D&A \u2192 VAT \u2192 Working Capital"),
        bulletItem("Each block validated against Excel output for all 120 asset slots"),
        bulletItem("Divergence target: <0.01% for all metrics across all assets"),
        bulletItem("Country-specific logic implemented: UK (CfD, 11kV), DE (EEG, SDE++), NL (SDE++)"),
        bulletItem("BESS revenue fully modeled: Tolling, Floor, Capacity Market, CIDT, Merchant"),

        // ════════════════════════════════════════════
        // PHASE 4
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Phase 4: Calculation Engine \u2014 Financing & Valuation"),
        phaseHeader(4, "Debt, Tax, Distribution & IRR", "M7\u2013M9", PHASE_COLORS[4]),
        spacer(100),
        bodyText("Goal: Complete the calculation engine with the financing stack: Senior Debt (including DSCR sculpting with iterative solver), Tax, Funding, Distribution Waterfall, and Valuation (IRR via XIRR)."),
        spacer(100),

        sectionTitle("Month 8: Senior Debt & Tax", HeadingLevel.HEADING_2),
        monthTable([
          ["W27\u2013W28", "Senior Debt engine (PLW R929\u20131033): DSCR sculpting with iterative solver (Newton-Raphson), drawdown, scheduled repayment, cash sweep, interest hedging", "app/engine/debt.py", "Debt validated"],
          ["W28\u2013W29", "Reserve Security (PLW R1035\u20131084): DSRA/DSRF sizing, MRA, Repowering CapEx account", "app/engine/reserves.py", "Reserves validated"],
          ["W29\u2013W30", "Tax engine (PLW R1095\u20131157): Corporation Tax (tiered rates DE/UK/NL), Local Tax (DE add-backs), tax loss carry-forward, tax depreciation adjustments", "app/engine/tax.py", "Tax validated"],
        ]),
        spacer(100),

        sectionTitle("Month 9: Funding, Distribution & Valuation", HeadingLevel.HEADING_2),
        monthTable([
          ["W30\u2013W31", "Funding Calcs (PLW R1160\u20131234): Sources & Uses, SHL (drawdown, IDC, principal, interest), commitment fees", "app/engine/funding.py", "Funding validated"],
          ["W31\u2013W32", "Distribution Waterfall (PLW R1258\u20131385): Covenants (DSCR/LLCR), SHL repayment, dividends, cash sweep, equity buy-back", "app/engine/distribution.py", "Waterfall validated"],
          ["W33\u2013W34", "Valuation (PLW R1388\u20131418): Project IRR (FCFF \u2192 XIRR), Equity IRR (Distributions), Equity IRR (CF for Equity). Full Financial Statements assembly", "app/engine/valuation.py", "Phase 4 complete"],
        ]),
        spacer(80),
        sectionTitle("Phase 4 Deliverables", HeadingLevel.HEADING_3),
        bulletItem("Complete Python financial model: all 1,609 rows of Project Level Workings replicated"),
        bulletItem("DSCR sculpting with iterative solver handles the circular reference (Debt \u2194 DSCR)"),
        bulletItem("Full Financial Statements: Income Statement, Cash Flow, Balance Sheet assembled from calculation blocks"),
        bulletItem("IRR calculation via XIRR matches Excel within 1 basis point"),
        bulletItem("Quarterly consolidation block generates output matching the Quarterly Output sheet"),

        // ════════════════════════════════════════════
        // PHASE 5
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Phase 5: Scenario Engine"),
        phaseHeader(5, "Real-Time Sensitivity Analysis", "M9\u2013M10", PHASE_COLORS[5]),
        spacer(100),
        bodyText("Goal: Build the scenario/sensitivity engine that makes the entire project worthwhile \u2014 drop analysis time from 2\u20133 hours to under 5 seconds by implementing selective recalculation via a dependency graph."),
        spacer(100),

        sectionTitle("Month 10: Dependency Graph & Selective Recalculation", HeadingLevel.HEADING_2),
        monthTable([
          ["W35\u2013W36", "Dependency graph: map which calculation blocks depend on which inputs. Mirror the 24 Sensis lever slots", "app/engine/dependency_graph.py", "Graph built"],
          ["W36\u2013W37", "Scenario override system: parameter overrides stored as deltas from baseline. Asset-level and portfolio-level targeting", "app/api/scenarios.py, models/scenario.py", "Overrides work"],
          ["W37\u2013W38", "Selective recalculation: only recalculate affected assets and affected blocks. Example: changing Net Production for 3 DE assets recalculates only those 3", "app/engine/selective_calc.py", "Selective recalc"],
          ["W39\u2013W40", "Parallel execution: run N scenarios concurrently (multiprocessing). Sensitivity table generation. Scenario comparison UI in Streamlit", "app/ui/scenario_builder.py", "Phase 5 complete"],
        ]),
        spacer(80),
        sectionTitle("Phase 5 Deliverables", HeadingLevel.HEADING_3),
        bulletItem("Change any of the 24 Sensis levers for any subset of assets and get results in < 5 seconds"),
        bulletItem("24-scenario sensitivity run across 3 assets completes in < 30 seconds"),
        bulletItem("Full portfolio recalculation (all 120 assets) completes in < 2 minutes"),
        bulletItem("Scenario comparison: side-by-side results with delta highlighting"),
        bulletItem("Sensitivity tables: matrix view of IRR/NPV/DSCR across parameter ranges"),

        // ════════════════════════════════════════════
        // PHASE 6
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Phase 6: Satellite Frontends & Polish"),
        phaseHeader(6, "Multi-Team Platform", "M11\u2013M12", PHASE_COLORS[6]),
        spacer(100),
        bodyText("Goal: Extend the platform with dedicated frontends for BD, Investment, and Debt teams. Each team gets their own ingestion, analysis dashboard, and reporting tailored to their workflow."),
        spacer(100),

        sectionTitle("Month 11: BD & Investment Team Frontends", HeadingLevel.HEADING_2),
        monthTable([
          ["W41\u2013W42", "BD team frontend: pipeline valuation dashboard, deal screening (filter by IRR/capacity/country), development stage tracking", "app/ui/bd_dashboard.py", "BD frontend live"],
          ["W43\u2013W44", "Investment team frontend: portfolio analytics, IRR attribution (by country, technology, vintage), waterfall charts, concentration analysis", "app/ui/investment_dashboard.py", "Investment frontend"],
        ]),
        spacer(100),

        sectionTitle("Month 12: Debt Team, Integration & Launch", HeadingLevel.HEADING_2),
        monthTable([
          ["W45\u2013W46", "Debt team frontend: covenant monitoring (DSCR/LLCR thresholds), refinancing scenario builder, debt maturity profile", "app/ui/debt_dashboard.py", "Debt frontend live"],
          ["W47", "Integration testing: end-to-end validation across all modules. Performance optimization (database query tuning, caching)", "tests/integration/", "Full test pass"],
          ["W48", "Documentation, user training materials, deployment to production environment. Handover and launch", "docs/", "Production launch"],
        ]),
        spacer(80),
        sectionTitle("Phase 6 Deliverables", HeadingLevel.HEADING_3),
        bulletItem("BD team: pipeline valuation dashboard with deal screening and development tracking"),
        bulletItem("Investment team: portfolio analytics with IRR attribution and concentration analysis"),
        bulletItem("Debt team: covenant monitoring with DSCR/LLCR alerts and refinancing scenarios"),
        bulletItem("Full platform deployed to production with documentation and training"),

        // ════════════════════════════════════════════
        // RISK REGISTER
        // ════════════════════════════════════════════
        new Paragraph({ children: [new PageBreak()] }),
        sectionTitle("Risk Register"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 1200, 1200, 4160],
          rows: [
            new TableRow({ children: [
              headerCell("Risk", 2800),
              headerCell("Likelihood", 1200, { align: AlignmentType.CENTER }),
              headerCell("Impact", 1200, { align: AlignmentType.CENTER }),
              headerCell("Mitigation", 4160),
            ]}),
            ...[
              ["Revenue engine complexity (348 rows, 3 countries, 6 contract types)", "High", "High", "Build incrementally with per-block validation. Revenue is allocated 6 weeks (the longest single block)."],
              ["DSCR circular reference divergence", "Medium", "High", "Iterative solver with configurable tolerance. Cross-validate against Excel for every asset."],
              ["Quarterly Excel structural changes beyond row offsets", "Medium", "Medium", "Structural Map fingerprint detects any drift. Human-in-the-loop approval for structural changes."],
              ["VBA macro logic not fully captured", "Medium", "Medium", "Reverse-engineer macro behavior from input/output comparison. Keep Excel as validation oracle."],
              ["Performance at portfolio scale (120 assets x 420 months)", "Low", "Medium", "NumPy vectorization. Selective recalculation via dependency graph. Profile and optimize in Phase 5."],
              ["New asset types or contract schemes added by ASE", "Low", "Low", "Modular engine design allows adding new revenue/contract modules without rewriting existing code."],
            ].map((r, i) => new TableRow({
              children: [
                cell(r[0], 2800, { fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[1], 1200, { align: AlignmentType.CENTER, bold: true,
                  color: r[1] === "High" ? "E74C3C" : r[1] === "Medium" ? "E67E22" : "27AE60",
                  fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[2], 1200, { align: AlignmentType.CENTER, bold: true,
                  color: r[2] === "High" ? "E74C3C" : r[2] === "Medium" ? "E67E22" : "27AE60",
                  fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[3], 4160, { size: 16, fill: i % 2 === 0 ? "F8F9FA" : undefined }),
              ]
            }))
          ]
        }),

        // ════════════════════════════════════════════
        // TEAM & RESOURCES
        // ════════════════════════════════════════════
        spacer(200),
        sectionTitle("Team & Resource Requirements"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2500, 1800, 5060],
          rows: [
            new TableRow({ children: [
              headerCell("Role", 2500),
              headerCell("Phases", 1800),
              headerCell("Responsibilities", 5060),
            ]}),
            ...[
              ["Backend Engineer (Senior)", "1\u20136", "FastAPI, PostgreSQL schema, ingestion pipeline, API design"],
              ["Financial Engineer", "3\u20135", "Translate Excel formulas to Python. Deep PF domain knowledge required."],
              ["Frontend / Streamlit Dev", "1\u20136", "Dashboard, scenario builder, report UI, satellite team frontends"],
              ["Data Engineer", "1\u20132", "Database optimization, migration scripts, data validation framework"],
              ["QA / Validation", "3\u20136", "Cross-validate Python engine vs Excel. Regression testing."],
              ["Product Manager (GM)", "1\u20136", "Requirements, prioritization, stakeholder alignment with ASE/BD/Investment/Debt"],
            ].map((r, i) => new TableRow({
              children: [
                cell(r[0], 2500, { bold: true, fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[1], 1800, { align: AlignmentType.CENTER, fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[2], 5060, { fill: i % 2 === 0 ? "F8F9FA" : undefined }),
              ]
            }))
          ]
        }),

        // ════════════════════════════════════════════
        // SUCCESS CRITERIA
        // ════════════════════════════════════════════
        spacer(200),
        sectionTitle("Success Criteria"),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [500, 3860, 2500, 2500],
          rows: [
            new TableRow({ children: [
              headerCell("#", 500, { align: AlignmentType.CENTER }),
              headerCell("Criterion", 3860),
              headerCell("Target", 2500),
              headerCell("Phase", 2500),
            ]}),
            ...[
              ["1", "Excel ingestion time (full quarterly file)", "< 5 minutes", "Phase 1"],
              ["2", "Quarterly diff generation", "< 30 seconds", "Phase 1"],
              ["3", "Report generation (all 21 GTC sheets)", "< 2 minutes", "Phase 2"],
              ["4", "Calculation engine vs Excel divergence", "< 0.01% all metrics", "Phase 3\u20134"],
              ["5", "Single-asset scenario recalculation", "< 1 second", "Phase 5"],
              ["6", "3-asset targeted sensitivity run", "< 5 seconds", "Phase 5"],
              ["7", "24-scenario sensitivity matrix", "< 30 seconds", "Phase 5"],
              ["8", "Full portfolio recalculation (120 assets)", "< 2 minutes", "Phase 5"],
              ["9", "Quarterly update cycle (ingest + diff + approve)", "< 1 hour total", "Phase 1\u20132"],
              ["10", "Platform uptime", "99.5%", "Phase 6"],
            ].map((r, i) => new TableRow({
              children: [
                cell(r[0], 500, { align: AlignmentType.CENTER, bold: true, fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[1], 3860, { fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[2], 2500, { bold: true, color: "27AE60", fill: i % 2 === 0 ? "F8F9FA" : undefined }),
                cell(r[3], 2500, { fill: i % 2 === 0 ? "F8F9FA" : undefined }),
              ]
            }))
          ]
        }),

        spacer(400),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: AMPYR_MID, space: 8 } },
          spacing: { before: 200 },
          children: [new TextRun({ text: "End of Document", font: "Arial", size: 18, color: "95A5A6", italics: true })]
        }),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:\\repos\\Ampyr-PFA\\Ref Docs\\Project_Parthenon_Development_Plan.docx", buffer);
  console.log("Document created successfully.");
});
