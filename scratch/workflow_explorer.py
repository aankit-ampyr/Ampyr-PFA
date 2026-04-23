"""
Project Parthenon — ASE ↔ GTC Workflow Explorer
Temporary Streamlit page to explore the financial model's functional structure.

Run:
    streamlit run scratch/workflow_explorer.py

This is a discussion aid, not production UI. All numbers come from the deep
analysis artefacts under .claude/analysis_2026_04/.
"""
from __future__ import annotations
import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

ANALYSIS = Path(__file__).parent.parent / ".claude" / "analysis_2026_04"

st.set_page_config(
    page_title="Parthenon · ASE↔GTC Workflow",
    page_icon="🏛",
    layout="wide",
)

# ---------- Data loading ---------------------------------------------------

@st.cache_data
def load_json(p: Path) -> dict:
    return json.loads(p.read_text()) if p.exists() else {}

inv = load_json(ANALYSIS / "01_inventory.json")
named = load_json(ANALYSIS / "02_named_ranges.json")
breakdown_f1 = load_json(ANALYSIS / "03_sheet_breakdown_F1.json")
breakdown_f2 = load_json(ANALYSIS / "03_sheet_breakdown_F2.json")
breakdown_f3 = load_json(ANALYSIS / "03_sheet_breakdown_F3.json")
plw_summary = load_json(ANALYSIS / "04_plw_summary.json")
plw_fns = load_json(ANALYSIS / "04_plw_fns_crosssheet.json")
gtc = load_json(ANALYSIS / "06_gtc.json")
diff = load_json(ANALYSIS / "07_quarterly_diff.json")
struct = load_json(ANALYSIS / "08_structure.json")

F1_SHEETS = {s["name"] for s in inv.get("F1_ASE_Latest", {}).get("sheets", [])}
F2_SHEETS = {s["name"] for s in inv.get("F2_GTC_Enhanced", {}).get("sheets", [])}

GTC_ADDED_SHEETS = sorted(F2_SHEETS - F1_SHEETS)
F1_ONLY_SHEETS = sorted(F1_SHEETS - F2_SHEETS)  # GTC removed these

# ---------- Header ---------------------------------------------------------

st.title("🏛 Project Parthenon — ASE ↔ GTC Workflow")
st.caption(
    "Temporary exploration app · generated 2026-04-23 from 3 source workbooks in "
    "`data/converted/`. See [Reference/2026-04-23_Excel_Deep_Analysis.md]."
)

# ---------- Top-level workflow diagram ------------------------------------

st.header("The Workflow, in one picture")

st.markdown(
    """
```
  ┌─────────────────────┐                                ┌────────────────────┐
  │  ASE                │    quarterly .xlsb             │  GTC               │
  │  (Ampyr Solar       │ ───────────────────────────▶   │  (Ampyr Energy     │
  │   Europe,           │                                │   Tech Solutions,  │
  │   London)           │    previous quarter .xlsb      │   New Delhi)       │
  │                     │ ───────────────────────────▶   │                    │
  │  Owns core model    │                                │  Adds 21 reporting │
  │  Maintains formulas │                                │  sheets + runs     │
  │  Updates quarterly  │                                │  scenarios/sensi's │
  └─────────────────────┘                                └─────────┬──────────┘
                                                                   │
                                                                   ▼
                                                        2-3 hours per run
                                                        (VBA loops 120 assets)
                                                                   │
                                                                   ▼
                                                        Outputs → board reports,
                                                        investment decisions,
                                                        debt covenant checks
```
"""
)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Assets in model", "120 slots", help="78 real + 42 placeholder")
c2.metric("Time horizon", "35 years", help="421 monthly periods")
c3.metric("Metrics per asset", "170")
c4.metric("Output cells", "~3.1 M", help="Quarterly Output paste target")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Countries", "3", help="UK, DE, NL")
c2.metric("Excel functions used", "26", help="SUM, IF, SUMIFS, MATCH dominate")
c3.metric("Formula rows in PLW", f"{plw_summary.get('rows_with_any_timeaxis_formula', 915):,}")
c4.metric("Full vectorisation target", f"{plw_summary.get('full_421_span_rows', 876):,}", help="Rows that span all 421 time cols")

st.divider()

# ---------- Tabs -----------------------------------------------------------

tab_inputs, tab_calc, tab_outputs, tab_gtc, tab_quarter, tab_risks = st.tabs(
    ["📥 Inputs", "⚙️ Calculation workflow", "📤 Outputs (F1)", "📊 GTC additions (F2)",
     "🔄 Quarterly change profile", "⚠️ Risks & gotchas"]
)

# ==========================================================================
# INPUTS
# ==========================================================================
with tab_inputs:
    st.subheader("What ASE fills in — the 7 input sheets")
    st.caption("These are the 'levers' that change between quarters. 25,063 hardcoded numeric values in total.")

    INPUT_SHEETS = [
        {
            "Sheet": "Project Info",
            "Purpose": "Per-asset parameters (120 slots × 590+ params)",
            "Key contents": "Asset names (row 5), tech (row 6), active flags (row 7), COD dates, capacities, contract terms, pricing assumptions",
            "Rows": 597, "Sections": 21,
            "Changes/Q": "🔴 High (36% of cells)",
        },
        {
            "Sheet": "Country Inputs",
            "Purpose": "UK vs DE vs NL tax/regulatory",
            "Key contents": "Corp tax, VAT, local tax (Gewerbesteuer), interest deduction caps, SHL terms, WHT, VAT refund periods",
            "Rows": 133, "Sections": 3,
            "Changes/Q": "🟢 Stable (<0.2%)",
        },
        {
            "Sheet": "Financing Inputs",
            "Purpose": "Senior debt + VAT loan + SHL terms",
            "Key contents": "Margin 1.9%, SONIA base, swap 4.1%, 20y tenor, DSCRs (1.15/1.30/1.20/1.80), hedge ratios, P90 sizing",
            "Rows": 77, "Sections": 3,
            "Changes/Q": "🟢 Stable — only date shifts (<0.2%)",
        },
        {
            "Sheet": "Time Inputs (A)",
            "Purpose": "Annual curves — inflation, rates, FX, prices, reserves",
            "Key contents": "16 sections: Inflation, Financing curves, FX, Market prices, Clawback, Balancing, Tax, Sources, MRA (165 rows), SLA (166 rows), LC amounts",
            "Rows": 651, "Sections": 16,
            "Changes/Q": "🟡 Low (~0.5%)",
        },
        {
            "Sheet": "Time Inputs (M)",
            "Purpose": "Monthly disaggregation of annual curves",
            "Key contents": "Monthly breakdowns of inflation/rate/price series (574 of 1,581 rows hidden)",
            "Rows": 1581, "Sections": "—",
            "Changes/Q": "🟡 Low (derived from (A))",
        },
        {
            "Sheet": "Time Inputs (Q)",
            "Purpose": "Quarterly reserve/facility schedules",
            "Key contents": "Cluster-level financing arrangements, existing facility schedules",
            "Rows": 254, "Sections": "—",
            "Changes/Q": "🔴 Very high (42%, can be fully restructured)",
        },
        {
            "Sheet": "Sensis",
            "Purpose": "Scenario control (24 slots)",
            "Key contents": "15+ levers across 4 sections: Operations (Devex/Capex/Opex/O&M), Production (yield/curtailment), Uncontracted rev, Contracted rev",
            "Rows": 140, "Sections": 4,
            "Changes/Q": "🟢 Frozen (<0.03%)",
        },
    ]

    st.dataframe(
        pd.DataFrame(INPUT_SHEETS),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("##### Input-sheet volume overview (F1)")
    rows = []
    for sheet in [d["Sheet"] for d in INPUT_SHEETS]:
        b = breakdown_f1.get(sheet, {})
        rows.append({
            "Sheet": sheet,
            "Cells populated": b.get("cells", 0),
            "Formula cells": b.get("formulas", 0),
            "Hardcoded numeric": b.get("hardcoded_numeric", 0),
            "Hardcoded text": b.get("hardcoded_text", 0),
        })
    df_inp = pd.DataFrame(rows).set_index("Sheet")
    st.bar_chart(df_inp, height=280)

    st.markdown("##### Validation + conditional-formatting coverage (input-integrity enforcement)")
    struct_rows = []
    for sheet in [d["Sheet"] for d in INPUT_SHEETS]:
        s = struct.get(sheet, {})
        struct_rows.append({
            "Sheet": sheet,
            "Data validations": s.get("data_validations", 0),
            "Conditional-formatting rules": s.get("conditional_formatting_rules", 0),
            "Hidden rows": s.get("hidden_rows", 0),
            "Hidden columns": s.get("hidden_cols", 0),
        })
    st.dataframe(pd.DataFrame(struct_rows), use_container_width=True, hide_index=True)

    with st.expander("📋 Project Info — the 21 named sections (120 assets span horizontally)"):
        sections = [
            (18, "Timings"),
            (20, "Solar+BESS"),
            (49, "BESS"),
            (61, "Asset sales"),
            (79, "Production"),
            (119, "Construction"),
            (173, "Revenues — PPA / Subsidy contracts"),
            (267, "Merchant"),
            (280, "Solar — Embedded Benefits (11kV)"),
            (291, "Q1/Q2/Q3/Q4 quarterly share"),
            (296, "BESS offtake"),
            (298, "Merchant (BESS)"),
            (305, "Continuous Intraday Trading (CIDT)"),
            (313, "Contracted (BESS)"),
            (523, "Devex — Solar"),
            (527, "Capex — BESS"),
            (532, "O&M"),
            (548, "D&A rates"),
            (550, "Solar depreciation"),
            (559, "BESS depreciation"),
            (566, "Success Factor"),
            (570, "Other"),
            (584, "Inputs (DO NOT DELETE)"),
            (596, "End"),
        ]
        st.dataframe(pd.DataFrame(sections, columns=["Row", "Section"]),
                     use_container_width=True, hide_index=True)

    with st.expander("🎛 Sensis — scenario levers (15+ across 4 sections, not 6)"):
        levers = [
            ("Operations", "Devex %", "numeric adj"),
            ("Operations", "Capex %", "numeric adj"),
            ("Operations", "Opex %", "numeric adj"),
            ("Operations", "O&M %", "numeric adj"),
            ("Production", "Net production %", "numeric adj"),
            ("Production", "Yield selector", "P50 / P90"),
            ("Production", "Curtailment %", "numeric adj"),
            ("Production", "Generation cadence", "Annual / Quarterly"),
            ("Uncontracted rev", "Sensitivity price curve active?", "bool"),
            ("Uncontracted rev", "Price curve", "Low / Base / High"),
            ("Uncontracted rev", "Indexed price curve?", "bool"),
            ("Uncontracted rev", "Merchant power prices", "numeric adj"),
            ("Uncontracted rev", "Breakeven — UK Solar", "numeric adj"),
            ("Uncontracted rev", "Breakeven — DE Solar", "numeric adj"),
            ("Uncontracted rev", "Breakeven — NL Solar", "numeric adj"),
            ("Contracted rev", "PPA on?", "bool"),
            ("Contracted rev", "Contracted share %", "numeric adj"),
            ("Contracted rev", "Contracted offtake prices", "numeric adj"),
        ]
        st.dataframe(pd.DataFrame(levers, columns=["Section", "Lever", "Type"]),
                     use_container_width=True, hide_index=True)


# ==========================================================================
# CALCULATION WORKFLOW
# ==========================================================================
with tab_calc:
    st.subheader("How F1 actually computes results")

    st.markdown("##### The VBA loop — why it takes 2-3 hours")
    st.code(
        """
ASE's `PlatformConsolidation` macro (Consolidation.bas, 200 LOC):

    FOR each project in ProjectList (120 slots):
        IF project is active:
            Set `Project_View` = project name       ◀── triggers PLW recalc
            Call Debt_sizing                         ◀── DSCR solver (3-8 iterations)
                Loop until debt_delta < 0.2 AND Use_delta < 0.2:
                    copy Use_live → Use_paste
                    copy d_service_live → d_service_paste
                    Calculate
            Paste `SPV_ConsolidatedCashflows.Value` → quarterly output
        startingrow += rangeRowsCF
""",
        language="text",
    )

    st.caption(
        "The template model is **single-asset-at-a-time** — one asset's formulas in "
        "cols AB..QF (421 monthly periods), swapped via `Project_View` selector. "
        "This sequential loop is what the Python engine replaces with a `(78, 421)` "
        "NumPy matrix pass."
    )

    st.markdown("##### PLW — the calculation engine (1,597 rows × 16,384 cols)")

    BLOCKS = [
        ("Header & Checks", "1–30", 25, "Low", "Project selection, validation flags"),
        ("Financial statements (summary)", "32–134", 80, "Low", "Income/CF/BS — SUMIFS aggregation of lower blocks"),
        ("Flags & Timing", "136–245", 100, "Low-Med", "COD flags, period flags, inflation XLOOKUP"),
        ("Production", "247–271", 20, "Medium", "Solar yield, degradation, BESS dispatch"),
        ("Generation waterfall", "273–309", 30, "Medium", "UK CfD metering, net generation"),
        ("Revenue ⚠", "310–619", 200, "HIGH", "PPA, CfD, FiT/FiP, EEG, merchant, GoO, tolling, capacity market — DE/UK/NL branching"),
        ("OpEx", "621–710", 60, "Medium", "O&M, insurance, land lease, community benefit"),
        ("Depreciation", "712–807", 70, "Medium", "7 categories, BESS repower"),
        ("VAT", "809–924", 50, "Medium", "VAT on CapEx/OpEx/Revenue, receivable/payable"),
        ("Senior Debt + DSCR ⚠", "926–1090", 100, "HIGH", "Drawdown, sculpting, iterative solver, DSRA/MRA"),
        ("Tax", "1092–1211", 80, "Med-High", "Corp tax, loss carry-forward, interest deduction"),
        ("SHL & Distribution", "1214–1385", 110, "Med-High", "Shareholder loans, waterfall, lock-up DSCR, WHT"),
        ("IRR & Consolidation", "1387–1596", 140, "Medium", "XIRR, quarterly SUMIFS rollup, BS check"),
    ]
    df_blocks = pd.DataFrame(BLOCKS, columns=["Block", "Row range", "Formula rows", "Complexity", "Key logic"])
    st.dataframe(df_blocks, use_container_width=True, hide_index=True)

    st.markdown("##### Dependency chain")
    st.markdown(
        """
```
Flags/Timings ──▶ Production ──▶ Revenue ──▶ OpEx ──▶ EBITDA
                                                       │
                       Working capital ◀── Tax ◀── Depreciation
                                                       │
                              Senior Debt ◀──▶ DSCR sculpting (iterative)
                                                       │
                     Reserve security ──▶ Distribution waterfall ──▶ IRR
                                                       │
                              Quarterly consolidation ──▶ HoldCo aggregation
```
"""
    )

    st.markdown("##### Function usage in PLW time-axis formulas (real counts)")
    fns = plw_fns.get("fn_usage", [])[:15]
    df_fn = pd.DataFrame(fns, columns=["Function", "Calls"])
    fig = px.bar(df_fn, x="Function", y="Calls", text="Calls",
                 title=f"Top Excel functions in PLW · 26 unique total · {sum(n for _,n in fns):,} calls shown")
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### Where PLW reads from (cross-sheet references)")
    cs = plw_fns.get("cross_sheet", [])
    df_cs = pd.DataFrame(cs, columns=["Source sheet", "Refs from PLW"])
    st.dataframe(df_cs, use_container_width=True, hide_index=True)
    st.caption("`Time Inputs (A)` is ~5× more load-bearing than (M) for PLW — prioritise it in the Python ingestion schema.")

    with st.expander("🔬 The DSCR solver in detail (DebtSizing.bas)"):
        st.code(
            """
Sub Debt_sizing()
    If Range("debtapplicable") = True Then
        Range("Sizing_Active") = 1
        Calculate
        projectnumber = Range("ProjectNumber")
        Do
            Range("Use_paste").Offset(projectnumber - 1, 0).Value = Range("Use_live").Value
            Range("d_service_paste").Offset(projectnumber - 1, 0).Value = Range("d_service_live").Value
            Calculate
        Loop While Range("debt_delta") > 0.2 Or Range("Use_delta") > 0.2

        Range("SeniorDebtOptimalValue").Value = Range("SeniorDebtOptimalLive").Value
        Range("Sizing_Active") = 0
        Calculate
    End If
End Sub
""",
            language="vb",
        )
        st.warning(
            "**Two convergence criteria**, both `< 0.2`: `debt_delta` AND `Use_delta`. "
            "No max-iteration cap. Current `DEVELOPMENT_SPEC.md` mentions only one — needs updating."
        )


# ==========================================================================
# OUTPUTS (F1)
# ==========================================================================
with tab_outputs:
    st.subheader("What F1 produces — the outputs ASE ships to GTC")

    F1_OUTPUT_SHEETS = [
        ("Quarterly Output", "THE paste target — VBA dumps every asset's SPV cashflows here", "3,122,453 cells (2.9M hardcoded, 180k formulas)"),
        ("Project Level Macro Paste", "Intermediate paste buffer", "170,110 cells"),
        ("HoldCo CFs & Valuation", "Group-level cashflows & valuation (XIRR/XNPV)", "663,939 formulas"),
        ("HoldCo income", "Group income statement aggregated", "155,117 formulas (609 hidden rows)"),
        ("HoldCo_Facility", "Facility-level debt view", "92,387 formulas, 49 CF rules"),
        ("HoldCo_Summary", "HoldCo exec summary", "2,570 formulas (141 hidden columns!)"),
        ("ProjectSummary", "Per-asset summary", "1,926 formulas"),
        ("Dashboard", "User control panel + headline numbers", "133 formulas"),
        ("Charts", "Chart data staging", "304,729 formulas — mostly chart source arrays"),
        ("Checks", "Balance-sheet / integrity checks", "1,206 formulas"),
        ("BESS DCF Multiple Analysis", "New in F1 — BESS valuation multiples", "382 formulas"),
        ("Pltfrm Costs Devex Analysis", "New in F1 — platform cost model", "325 formulas (GTC REMOVES this in F2)"),
        ("Sensis", "Scenario results paste area (Sens_live → paste at row 65)", "49 formulas"),
    ]
    st.dataframe(
        pd.DataFrame(F1_OUTPUT_SHEETS, columns=["Sheet", "Purpose", "Scale"]),
        use_container_width=True, hide_index=True,
    )

    st.markdown("##### Output-sheet volume (F1)")
    rows = []
    for name, _, _ in F1_OUTPUT_SHEETS:
        b = breakdown_f1.get(name, {})
        if b:
            rows.append({
                "Sheet": name,
                "Formula cells": b.get("formulas", 0),
                "Hardcoded cells": b.get("hardcoded_numeric", 0) + b.get("hardcoded_text", 0),
            })
    df_out = pd.DataFrame(rows).set_index("Sheet")
    st.bar_chart(df_out, height=350)

    st.info(
        "**Quarterly Output is 93% hardcoded paste values.** The VBA consolidation "
        "macro loops through every active asset and pastes `SPV_ConsolidatedCashflows` "
        "into this sheet. GTC then reshapes it in Asset workings."
    )


# ==========================================================================
# GTC ADDITIONS (F2)
# ==========================================================================
with tab_gtc:
    st.subheader("What GTC adds on top — 21 new sheets in F2")

    st.caption(
        f"F2 has **{len(GTC_ADDED_SHEETS)} sheets not in F1**. "
        f"It also removes {len(F1_ONLY_SHEETS)} sheet from F1: `{F1_ONLY_SHEETS}`."
    )

    GTC_TIERS = {
        "Tier 1 · Raw bridge (read from ASE outputs)": [
            ("Asset workings", "Reshapes Quarterly Output by (asset × metric × period). Mostly IFERROR(SUMIFS(...))", "1,191,376 formulas"),
            ("Valuation", "XNPV + XLOOKUP from HoldCo_Facility + Time Inputs (Q) + Dashboard", "65 formulas"),
            ("Valuation (2)", "Valuation rollup", "394 formulas"),
            ("QRep(EUR)", "Quarterly reporting — EUR view, pulls from Valuation", "1,867 formulas"),
            ("Q Rep (USD)", "Quarterly reporting — USD view (FX-converted from EUR)", "4,032 formulas (hidden)"),
        ],
        "Tier 2 · Breakdown (restructure by period)": [
            ("FY26A-29F PL Breakdown", "PL items × quarterly periods", "14,985 formulas"),
            ("FY26A-29F BS Breakdown", "BS items × quarterly periods", "6,867 formulas"),
            ("CF Capital Strategy Breakdown", "CF items × periods × capital categories", "1,459 formulas"),
        ],
        "Tier 3 · Projection aggregates (ROUND+SUMIF pattern)": [
            ("PnL projection - aggregate", "Aggregated PL — pulls from PL Breakdown + PL Mapping", "2,018 formulas (10 #REF!)"),
            ("BS projection - aggregate", "Aggregated BS", "1,127 formulas"),
            ("CF Capital Strategy", "Aggregated CF capital strategy", "1,903 formulas"),
            ("FY26B PL Mapping", "Budget PL mapping layer", "31 formulas"),
            ("FY26B BS Mapping", "Budget BS mapping layer", "39 formulas"),
        ],
        "Tier 4 · Summary & visualisation": [
            ("Summary sheet", "Exec summary — pulls from PnL + CF + BS + Valuation (2) [hidden]", "2,199 formulas (30 #REF!)"),
            ("Cash Breakeven", "Pulls from HoldCo_Summary + PL Breakdown", "146 formulas"),
            ("Graphs", "Graph data staging", "187 formulas"),
            ("Tables", "Table data staging", "209 formulas"),
            ("Net income bridge", "NI reconciliation bridge", "16 formulas"),
        ],
        "Tier 0 · Meta (no formula refs)": [
            ("Instruction tab", "User instructions", "0 formulas"),
            ("FX", "FX rate lookup table", "0 formulas"),
            ("Log", "Usage log", "1 formula"),
        ],
    }
    for tier, sheets in GTC_TIERS.items():
        st.markdown(f"**{tier}**")
        st.dataframe(
            pd.DataFrame(sheets, columns=["Sheet", "Purpose", "Scale"]),
            use_container_width=True, hide_index=True,
        )

    st.markdown("##### Formula volume per GTC sheet")
    rows = []
    for tier_sheets in GTC_TIERS.values():
        for name, _, _ in tier_sheets:
            b = breakdown_f2.get(name, {})
            if b and b.get("formulas", 0):
                rows.append({"Sheet": name, "Formula cells": b.get("formulas", 0)})
    df_gtc = pd.DataFrame(rows).sort_values("Formula cells", ascending=True)
    fig = px.bar(df_gtc, x="Formula cells", y="Sheet", orientation="h",
                 title="GTC sheet sizes (log scale)", log_x=True,
                 height=600)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### GTC dependency graph (who pulls from whom)")
    st.markdown(
        """
```
  ASE outputs  ─────────────────────────────────────▶  Asset workings
       │                                                   │    (2.5M SUMIFS refs
       │                                                   │     into Quarterly Output)
       │                                                   │
       ├──▶ Quarterly Output ──────▶ Asset workings ───────┘
       │                                │
       ├──▶ HoldCo CFs & Valuation ─────┤
       │                                ▼
       │                         FY26A-29F Breakdown sheets (PL/BS/CF Capital)
       │                                │
       │                                ▼
       │                         Projection-aggregate sheets (ROUND(SUMIF(...)))
       │                                │
       └──▶ Time Inputs (Q) ──▶ Valuation (XNPV) ──▶ Valuation (2)
                                                        │
                                                        ▼
                                                    Summary sheet
                                                        │
                                                        ▼
                                                    Cash Breakeven, Graphs, Tables
```
"""
    )

    st.warning(
        "**Live #REF! errors in F2** — these will propagate when the file is refreshed:\n"
        "- `PnL projection - aggregate`: 10 broken refs\n"
        "- `Summary sheet`: 30 broken refs\n\n"
        "Current state of F2 as delivered has these errors unresolved."
    )

    with st.expander("⚠️ Why GTC removed `Pltfrm Costs Devex Analysis` from F1"):
        st.write(
            "This sheet is present in F1 (42 rows × 325 cols, 5 conditional-formatting rules) "
            "but absent from F2. Either GTC's template overwrites it, or GTC consciously drops "
            "it during the upload workflow. **Worth asking ASE** — it might contain info they "
            "consider part of their output but GTC isn't using."
        )


# ==========================================================================
# QUARTERLY CHANGE PROFILE
# ==========================================================================
with tab_quarter:
    st.subheader("What actually changes when partners resend the file")

    st.caption(
        "Comparison: **F1 (2026-01-08, current quarter)** vs **F3 (previous quarter)**. "
        "Input sheets compared cell-by-cell; PLW compared by row-level formula hash."
    )

    rows = []
    for sheet in ["Project Info", "Country Inputs", "Financing Inputs",
                  "Time Inputs (A)", "Time Inputs (Q)", "Sensis",
                  "Project Level Workings"]:
        d = diff.get(sheet, {})
        if "changes_count" in d:
            rows.append({
                "Sheet": sheet,
                "Cells scanned": d["cells_scanned"],
                "Cells changed": d["changes_count"],
                "% changed": d["pct_changed"],
                "Grade": "🔴 High" if d["pct_changed"] > 10 else "🟡 Moderate" if d["pct_changed"] > 0.3 else "🟢 Stable",
            })
        elif "changed_formula_rows_count" in d:
            rows.append({
                "Sheet": sheet + " (formula rows)",
                "Cells scanned": d["f1_rows"],
                "Cells changed": d["changed_formula_rows_count"],
                "% changed": round(100 * d["changed_formula_rows_count"] / d["f1_rows"], 2),
                "Grade": "🔴 High",
            })
    df_diff = pd.DataFrame(rows)
    st.dataframe(df_diff, use_container_width=True, hide_index=True)

    fig = px.bar(df_diff, x="Sheet", y="% changed",
                 color="% changed",
                 color_continuous_scale="Reds",
                 title="Quarterly change % by sheet")
    fig.update_layout(showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### 🔴 High-churn areas (always revalidate)")
        st.markdown(
            """
- **Project Info (36%)** — asset renames, tech switches, flag flips, placeholder → real asset conversions
- **Time Inputs (Q) (42%)** — whole sections can be restructured (happened F3→F1)
- **PLW formula rows (11.5%)** — ASE edits live formulas in Flags/Timing/Production/Revenue blocks
"""
        )

    with col2:
        st.markdown("##### 🟢 Stable areas (safe to skip unless flagged)")
        st.markdown(
            """
- **Country Inputs (0.14%)** — tax/VAT/SHL parameters frozen this Q
- **Financing Inputs (0.12%)** — only date shifts (+2 years on facility maturity)
- **Time Inputs (A) (0.53%)** — margin/inflation tweaks only
- **Sensis (0.03%)** — scenario engine frozen
"""
        )

    with st.expander("📖 Concrete F3 → F1 changes observed"):
        st.markdown(
            """
**Assets renamed (Project Info row 5):**
- DE cluster: Adamshoffnung 1 → Arendsee, Parchim → Dalum Chicken Farm, Bückwitz → Dalum East, Dalum → Dalum West
- UK typo fix: Whitney → Witney
- 3 placeholders became real assets: Jerriestown, Harker 2, North Newton Phase 2

**Tech changes (row 6):** 3 slots flipped Solar → Solar & BESS

**Active flags (row 7):** 6 flipped False→True (new actives), 2 flipped True→False (retired)

**Financing Inputs:** all 29 changes are **date shifts of exactly +2 years**
- Senior Debt Period-1 end: 2037-10-01 → 2039-10-01
- Facility maturities: 2047 → 2049, 2048 → 2050

**Time Inputs (A) margins:** 0.021 → 0.02, 0.024 → 0.02 (margin compression); reference year 2029 → 2030

**Time Inputs (Q):** the entire "Assets with Existing Financing" section (KfW Term Loan cluster layout) was **removed and rebuilt** with a simpler NL cluster format — this drove ~17k formulas disappearing

**PLW:** 184 formula rows changed, concentrated in rows 3, 17-22, 139-141, 190-210, 242-327 (Flags/Timing + Production + early Revenue)

**External workbook links (new!):** F3 had 0, F1 has 4 — SharePoint refs to Project Canopy (v14/v24/v30) + CIP v7 Capacity Analysis
"""
        )

    st.markdown("##### Named-range row shifts (the 'alignment' problem)")
    st.warning(
        "**13 named ranges in F1 point to rows 2 higher than in F3** — ASE inserted 2 rows "
        "in Project Info. The ingestion differ can't compare by row number, it must align "
        "by section label."
    )
    st.code(
        """
modelStartDate:     F1='Project Info'!$F$571   F3='Project Info'!$F$569   shift +2
months:             F1='Project Info'!$F$572   F3='Project Info'!$F$570   shift +2
mths_per:           F1='Project Info'!$F$575   F3='Project Info'!$F$573   shift +2
P50_P90:            F1='Project Info'!$E$85    F3='Project Info'!$E$83    shift +2
PPA_type:           F1='Project Info'!$D$590:$D$596   F3='Project Info'!$D$588:$D$594   shift +2
LL_revenue:         F1='Project Info'!$F$590:$F$591   F3='Project Info'!$F$588:$F$589   shift +2
outputSheetName:    F1='Project Info'!$F$576   F3='Project Info'!$F$574   shift +2
round, round_2, mthyr, mths_yr, modelStartDate: all +2
""",
        language="text",
    )


# ==========================================================================
# RISKS & GOTCHAS
# ==========================================================================
with tab_risks:
    st.subheader("Implications for the Python build")

    st.markdown(
        """
### 1. External workbook dependencies are new this quarter
F1 has 4 external SharePoint links (F3 had 0):
- `Project Canopy FM _v30 — Firsfield standalone.xlsm`
- `Project Canopy FM _v14 (wip).xlsm`
- `Project Canopy FM _v24.xlsm`
- `Project Parthenon - CIP v7 - Capacity Analysis.xlsb`

These files are **not in our data/ folder**. Before Phase 1: decide to ingest (needs SharePoint access), snapshot values inline, or declare out of scope.

### 2. Formulas are stable within a quarter, not between quarters
184 PLW rows changed formula between F3 and F1. The translator must be re-runnable each quarter with a formula-delta report identifying blocks needing re-validation.

### 3. Asset workings is a reshape, not financial logic
Don't translate 1.19M SUMIFS into 1.19M Python lines. In the DB it's a single groupby/pivot on `quarterly_metrics`. Similarly all GTC projection/breakdown sheets are generic `ROUND(SUMIF(...))` — **one aggregator**, not 21 bespoke reports.

### 4. Structural fingerprint must handle row-shift alignment
Between F3 and F1, 13 named ranges shifted by +2 rows. Align by section label, not row number.

### 5. DSCR solver has TWO convergence criteria
Both `debt_delta < 0.2` AND `Use_delta < 0.2` must be satisfied. Current `docs/DEVELOPMENT_SPEC.md` mentions only one.

### 6. Scenario engine is bigger than documented
15+ levers across 4 sections (not 6). Mixed types (boolean, enum, percentage). Schema needs lever-type discriminator.

### 7. Named-range noise must be filtered
F2 has ~490 garbage names from Capital IQ / Bloomberg / Smartview / Access add-ins. Keep only those referencing ASE sheets.

### 8. Hidden rows/cols still hold live data
Time Inputs (A) has 448 hidden rows; HoldCo income has 609 (80%). Read hidden cells; ignore `sheet_state` for ingestion.

### 9. Live #REF! errors exist in F2
`PnL projection - aggregate` (10) and `Summary sheet` (30). Ingestion should log these, not silently treat as zero.

### 10. Input-integrity is enforced via DV + CF rules
Project Info has 15 data validations + 52 CF rules. Replicate via Pydantic validators.

### 11. GTC removes an F1 sheet
`Pltfrm Costs Devex Analysis` exists in F1 but not F2. Worth asking ASE why.

### 12. VBA semantic quirk
`PlatformConsolidation` uses `projectactiveflag(i) = "True"` (string). `Sens_platformconsol` uses `= True` (boolean). Replicate exactly to match outputs.
"""
    )

    st.info(
        "See `Reference/2026-04-23_Excel_Deep_Analysis.md` for the full reference. "
        "Open items are tracked in §13 of that document and should be revisited here."
    )


st.divider()
st.caption(
    "This app reads live from `.claude/analysis_2026_04/`. "
    "Re-run the analysis scripts there to refresh. "
    "Not part of the eventual product UI — this is a scratch-pad for discussion."
)
