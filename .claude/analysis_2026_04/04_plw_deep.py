"""04 - PLW deep dive:
- Identify formula-template rows (where row has a time-series of formulas across cols M..)
- For each row: label, formula template, distinct-template count, unique fns
- Cross-sheet dependency count
- Find edge-case rows (where across-column formulas differ)
"""
import json
import re
from collections import Counter, defaultdict
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")

PATH = DATA / "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"

FN_RE = re.compile(r"\b([A-Z][A-Z0-9._]*)\s*\(")
SHEET_REF_RE = re.compile(r"'([^']+)'!|\b([A-Za-z_][\w]*)!")
CELL_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")

# Time axis columns in PLW: M (13) to PQ = ? From docs: 421 monthly cols starting at M
TIME_START_COL = 13  # M
TIME_END_COL = TIME_START_COL + 420  # inclusive

def normalise_formula(f: str) -> str:
    """Replace all column refs in time-axis cols with 'TX' so same-shape formulas collapse."""
    if not f.startswith("="):
        return f
    def repl(m):
        col, row = m.group(1), m.group(2)
        try:
            cidx = column_index_from_string(col)
        except ValueError:
            return m.group(0)
        if TIME_START_COL <= cidx <= TIME_END_COL:
            return f"T{row}"
        return f"{col}{row}"
    return CELL_RE.sub(repl, f)

wb = load_workbook(PATH, read_only=True, data_only=False, keep_links=False)
ws = wb["Project Level Workings"]

row_templates = {}   # row -> canonical formula from first time col
row_distinct = defaultdict(set)   # row -> set of distinct normalised formulas across time cols
row_label = {}       # row -> label from col B/C
row_section = {}     # row -> section header last seen
row_fn_usage = defaultdict(Counter)  # row -> Counter of functions
row_cross_sheet = defaultdict(Counter)
row_formula_count = Counter()  # row -> # formula cells in time cols
col_counter = Counter()

sections = []   # rows that look like section headers (bold/heading rows — approximated by col B filled and no formulas)

print("Scanning PLW rows...", flush=True)
for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), start=1):
    r = r_idx
    label_cells = []
    for c in row[:5]:  # first 5 cols for label
        v = getattr(c, "value", None)
        if v and isinstance(v, str):
            label_cells.append(v)
    label = " | ".join(label_cells)[:120]
    if label:
        row_label[r] = label

    # formulas across time cols
    time_cells = row[TIME_START_COL-1:TIME_END_COL]  # 0-indexed slice
    formulas_in_row = 0
    for c_offset, c in enumerate(time_cells):
        v = getattr(c, "value", None)
        if isinstance(v, str) and v.startswith("="):
            formulas_in_row += 1
            col_counter[TIME_START_COL + c_offset] += 1
            norm = normalise_formula(v)
            row_distinct[r].add(norm)
            body = v[1:]
            for fn in FN_RE.findall(body):
                row_fn_usage[r][fn] += 1
            for m in SHEET_REF_RE.findall(body):
                s = m[0] if m[0] else m[1]
                if s:
                    row_cross_sheet[r][s] += 1
    row_formula_count[r] = formulas_in_row
    if formulas_in_row > 0:
        # use smallest-col formula as template
        first = next((getattr(c, "value", None) for c in time_cells if isinstance(getattr(c, "value", None), str) and getattr(c, "value").startswith("=")), None)
        row_templates[r] = {
            "label": label,
            "template": first[:500] if first else None,
            "n_formulas_across_time": formulas_in_row,
            "distinct_templates": len(row_distinct[r]),
            "edge_case": len(row_distinct[r]) > 1,
            "top_fns": row_fn_usage[r].most_common(5),
            "cross_sheet": row_cross_sheet[r].most_common(5),
        }

wb.close()

# aggregate stats
n_rows = len(row_templates)
formula_rows = sum(1 for r in row_templates.values() if r["n_formulas_across_time"] > 0)
edge_case_rows = [r for r, v in row_templates.items() if v["edge_case"]]
full_span_rows = [r for r, v in row_templates.items() if v["n_formulas_across_time"] == 421]
partial_span_rows = [r for r, v in row_templates.items() if 0 < v["n_formulas_across_time"] < 421]

summary = {
    "rows_with_any_timeaxis_formula": formula_rows,
    "edge_case_rows_count": len(edge_case_rows),
    "edge_case_rows": sorted(edge_case_rows)[:30],
    "full_421_span_rows": len(full_span_rows),
    "partial_span_rows": len(partial_span_rows),
    "partial_span_sample": [{"row": r, "n": row_templates[r]["n_formulas_across_time"], "label": row_templates[r]["label"]}
                             for r in sorted(partial_span_rows)[:30]],
    "col_coverage_first10": col_counter.most_common(10),
    "col_coverage_last10": sorted(col_counter.items())[-10:],
}

print(f"\nPLW summary:")
for k, v in summary.items():
    if isinstance(v, list) and len(v) > 5:
        print(f"  {k}: {len(v)} items (first 3: {v[:3]})")
    else:
        print(f"  {k}: {v}")

(OUT / "04_plw_summary.json").write_text(json.dumps(summary, indent=2, default=str))
(OUT / "04_plw_rows.json").write_text(json.dumps(row_templates, indent=2, default=str))

# Cross-sheet dependency aggregate
all_cross = Counter()
for r, c in row_cross_sheet.items():
    all_cross.update(c)
print(f"\nCross-sheet references from PLW (top 15):")
for s, n in all_cross.most_common(15):
    print(f"  {s:40} {n}")

# Function usage aggregate (across time-axis formulas only)
all_fn = Counter()
for r, c in row_fn_usage.items():
    all_fn.update(c)
print(f"\nFunction usage in PLW time-axis formulas (top 25):")
for fn, n in all_fn.most_common(25):
    print(f"  {fn:25} {n}")

(OUT / "04_plw_fns_crosssheet.json").write_text(json.dumps({
    "fn_usage": all_fn.most_common(),
    "cross_sheet": all_cross.most_common(),
}, indent=2))
