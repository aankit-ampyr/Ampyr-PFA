import sys
sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import re
from collections import defaultdict, Counter

FILE_PATH = r"C:\repos\Ampyr-PFA\Ref Docs\Converted\Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"
SHEET_NAME = "Project Level Workings"
OUTPUT_PATH = r"C:\repos\Ampyr-PFA\Ref Docs\analysis_plw_formulas.txt"

out_lines = []
def log(msg=""):
    out_lines.append(str(msg))
    print(msg)

log("=" * 80)
log("PROJECT LEVEL WORKINGS - FORMULA ANALYSIS")
log("=" * 80)
log()

# ═══════════════════════════════════════════════════════════════════════════
# STEP 1: COLUMN LAYOUT DISCOVERY
# ═══════════════════════════════════════════════════════════════════════════
log("STEP 1: COLUMN LAYOUT DISCOVERY")
log("-" * 60)

wb = openpyxl.load_workbook(FILE_PATH, data_only=False, read_only=True)
ws = wb[SHEET_NAME]

# Read first 15 rows, up to col 500
header_data = {}
crow = 0
for row in ws.iter_rows(min_row=1, max_row=15, min_col=1, max_col=500):
    crow += 1
    header_data[crow] = []
    for idx, cell in enumerate(row):
        if cell.value is not None:
            header_data[crow].append((idx + 1, cell.value))
wb.close()

# Find last data column
max_data_col = 0
for r, vals in header_data.items():
    for col_idx, val in vals:
        if col_idx > max_data_col:
            max_data_col = col_idx

log(f"  Last data column in headers: {max_data_col} ({get_column_letter(max_data_col)})")

# Print header info
for r in range(1, 16):
    vals = header_data.get(r, [])
    if vals:
        non_formula = [(c, v) for c, v in vals if not (isinstance(v, str) and v.startswith('='))]
        formula = [(c, v) for c, v in vals if isinstance(v, str) and v.startswith('=')]
        log(f"  Row {r:2d}: {len(vals)} cells ({len(non_formula)} labels/const, {len(formula)} formulas)")
        # Show label columns (A-L)
        for col_idx, val in vals:
            if col_idx <= 12:
                cl = get_column_letter(col_idx)
                val_str = str(val)[:80]
                log(f"         {cl}{r} (col {col_idx}): {val_str}")
        # Show first few formula columns
        formula_in_range = [(c, v) for c, v in vals if c > 12]
        if formula_in_range:
            first_c = formula_in_range[0][0]
            last_c = formula_in_range[-1][0]
            log(f"         Data cols: {get_column_letter(first_c)}-{get_column_letter(last_c)} ({len(formula_in_range)} cols)")
            # Show first formula as sample
            log(f"         Sample ({get_column_letter(first_c)}{r}): {str(formula_in_range[0][1])[:80]}")
    else:
        log(f"  Row {r:2d}: (empty)")

# Determine time period structure
# Row 5 = Start date, Row 6 = End Date, Row 7 = Days in period
log()
log("  LAYOUT SUMMARY:")
log(f"    Label columns: A-L (cols 1-12)")
log(f"    Time period columns: M-{get_column_letter(max_data_col)} (cols 13-{max_data_col})")
log(f"    Number of time periods: {max_data_col - 12}")
log(f"    Row 5: Start date per period")
log(f"    Row 6: End date per period")
log(f"    Row 7: Days in period")
log(f"    Row 9: Period active flag")
log(f"    Row 15: Selected project name")

# ═══════════════════════════════════════════════════════════════════════════
# STEP 2: ROW LABEL EXTRACTION & CALC BLOCKS
# ═══════════════════════════════════════════════════════════════════════════
log()
log("=" * 80)
log("STEP 2: ROW LABEL EXTRACTION & CALC BLOCKS")
log("-" * 60)

wb = openpyxl.load_workbook(FILE_PATH, data_only=False, read_only=True)
ws = wb[SHEET_NAME]

row_labels = {}
label_cols_data = {}

crow = 0
for row in ws.iter_rows(min_row=1, max_row=1609, min_col=1, max_col=12):
    crow += 1
    vals = {}
    for idx, cell in enumerate(row):
        col = idx + 1
        if cell.value is not None:
            vals[col] = cell.value
    label_cols_data[crow] = vals
    parts = []
    for c in sorted(vals.keys()):
        v = str(vals[c]).strip()
        if len(v) > 80:
            v = v[:80] + "..."
        parts.append(v)
    row_labels[crow] = " | ".join(parts) if parts else ""
wb.close()

# Build calc blocks - use column A/B/C labels to detect sections
# In financial models, section headers are typically in col A or B, often uppercase
blocks = []
current_block_name = "Header"
current_block_start = 1
current_block_rows = []

# Known section keywords for financial models
SECTION_KEYWORDS = {
    'revenue', 'opex', 'operating', 'capex', 'capital', 'tax', 'debt', 'senior',
    'equity', 'cash', 'balance', 'p&l', 'profit', 'loss', 'flag', 'production',
    'generation', 'inflation', 'interest', 'depreciation', 'amortisation',
    'dividend', 'working capital', 'vat', 'dscr', 'ratio', 'sizing',
    'solar', 'bess', 'battery', 'wind', 'funding', 'construction',
    'shareholders', 'liabilities', 'assets', 'checks', 'irr', 'npv',
    'summary', 'workings', 'project', 'repower', 'refinancing'
}

for r in range(1, 1610):
    label = row_labels.get(r, "")
    is_empty = (label == "")

    if is_empty:
        if current_block_rows:
            blocks.append({
                'name': current_block_name,
                'start': current_block_start,
                'end': current_block_rows[-1],
                'rows': current_block_rows,
                'row_count': len(current_block_rows)
            })
            current_block_rows = []
            current_block_name = f"(unnamed block after row {r})"
            current_block_start = r + 1
    else:
        # Check if this row is a section header
        vals = label_cols_data.get(r, {})
        # Section header: value only in col A or B, short text
        label_cols_present = [c for c in vals.keys() if c <= 8]
        data_cols_present = [c for c in vals.keys() if c > 8]

        if len(label_cols_present) <= 2 and len(data_cols_present) == 0:
            # Only label columns, no data - could be section header
            first_val = str(vals.get(min(vals.keys()), "")).strip() if vals else ""
            words = first_val.lower().split()
            is_header = False
            if first_val.isupper() and len(first_val) > 2:
                is_header = True
            elif any(w in SECTION_KEYWORDS for w in words):
                is_header = True
            elif first_val.endswith(':'):
                is_header = True
            elif len(words) <= 4 and len(first_val) > 2:
                is_header = True

            if is_header:
                if current_block_rows:
                    blocks.append({
                        'name': current_block_name,
                        'start': current_block_start,
                        'end': current_block_rows[-1],
                        'rows': current_block_rows,
                        'row_count': len(current_block_rows)
                    })
                    current_block_rows = []
                current_block_name = first_val
                current_block_start = r

        current_block_rows.append(r)

if current_block_rows:
    blocks.append({
        'name': current_block_name,
        'start': current_block_start,
        'end': current_block_rows[-1],
        'rows': current_block_rows,
        'row_count': len(current_block_rows)
    })

# Merge very small blocks (1-2 rows) with next block
merged_blocks = []
skip_next = False
for i, block in enumerate(blocks):
    if skip_next:
        skip_next = False
        continue
    if block['row_count'] <= 2 and i + 1 < len(blocks):
        next_block = blocks[i + 1]
        merged = {
            'name': block['name'],
            'start': block['start'],
            'end': next_block['end'],
            'rows': block['rows'] + next_block['rows'],
            'row_count': len(block['rows']) + len(next_block['rows'])
        }
        merged_blocks.append(merged)
        skip_next = True
    else:
        merged_blocks.append(block)

blocks = merged_blocks

log(f"  Total rows with labels: {sum(1 for r, l in row_labels.items() if l)}")
log(f"  Total empty rows: {sum(1 for r, l in row_labels.items() if not l)}")
log(f"  Identified {len(blocks)} calc blocks:")
log()

for i, block in enumerate(blocks):
    log(f"  Block {i+1:3d}: {block['name'][:60]:<60s} rows {block['start']:4d}-{block['end']:4d} ({block['row_count']} rows)")

# Print full row label map
log()
log("  FULL ROW LABEL MAP (non-empty rows):")
log("  " + "-" * 100)
for r in range(1, 1610):
    label = row_labels.get(r, "")
    if label:
        log(f"    Row {r:4d}: {label[:140]}")

# ═══════════════════════════════════════════════════════════════════════════
# STEP 3: FORMULA EXTRACTION (first 3 time-period columns = M, N, O)
# ═══════════════════════════════════════════════════════════════════════════
log()
log("=" * 80)
log("STEP 3: FORMULA EXTRACTION FOR FIRST 3 TIME PERIODS")
log("-" * 60)

# In this model, "assets" = time period columns. Each column is one period.
# We extract formulas for columns M (13), N (14), O (15) as "Period 1, 2, 3"
# Plus columns G-L for label/check formulas

period1_col = 13  # M
period2_col = 14  # N
period3_col = 15  # O

# Read columns G through O (7-15) for all rows
wb = openpyxl.load_workbook(FILE_PATH, data_only=False, read_only=True)
ws = wb[SHEET_NAME]

all_row_data = {}  # row_num -> {col: value}

crow = 0
for row in ws.iter_rows(min_row=1, max_row=1609, min_col=7, max_col=15):
    crow += 1
    row_data = {}
    for idx, cell in enumerate(row):
        col = 7 + idx
        val = cell.value
        if val is not None:
            if isinstance(val, str) and val.startswith('='):
                row_data[col] = val
            else:
                row_data[col] = f"CONST:{val}"
        else:
            row_data[col] = None
    all_row_data[crow] = row_data
wb.close()

# Also read a wider range for the full formula picture
# Read cols M through PP (13 to about 433) for sampled rows to check formula consistency
# But first, let's get the first 3 periods fully
log(f"  Extracted formulas for cols G-O (7-15) for all 1609 rows")

# Count formula types per column
for col in [13, 14, 15]:
    formula_count = 0
    const_count = 0
    empty_count = 0
    for r in range(1, 1610):
        val = all_row_data.get(r, {}).get(col)
        if val is None:
            empty_count += 1
        elif val.startswith('='):
            formula_count += 1
        else:
            const_count += 1
    cl = get_column_letter(col)
    log(f"  Column {cl} (Period {col-12}): {formula_count} formulas, {const_count} constants, {empty_count} empty")

# Extract helper/check column stats
for col in [7, 8]:
    formula_count = sum(1 for r in range(1, 1610) if all_row_data.get(r, {}).get(col, "")
                        and str(all_row_data.get(r, {}).get(col, "")).startswith('='))
    const_count = sum(1 for r in range(1, 1610) if all_row_data.get(r, {}).get(col, "")
                      and str(all_row_data.get(r, {}).get(col, "")).startswith('CONST:'))
    cl = get_column_letter(col)
    log(f"  Column {cl} (helper): {formula_count} formulas, {const_count} constants")

# ═══════════════════════════════════════════════════════════════════════════
# STEP 4: FORMULA REUSE CHECK ACROSS TIME PERIODS
# ═══════════════════════════════════════════════════════════════════════════
log()
log("=" * 80)
log("STEP 4: FORMULA REUSE CHECK ACROSS TIME PERIODS")
log("-" * 60)

log("  Comparing Period 2 (col N) vs Period 1 (col M) - checking if formulas")
log("  are structurally identical (just column-shifted by 1).")
log()

# Also read a few more columns to compare further periods
# Read cols 13-18 (M-R) for thorough comparison
wb = openpyxl.load_workbook(FILE_PATH, data_only=False, read_only=True)
ws = wb[SHEET_NAME]

period_data = {}  # row -> {col: formula}
crow = 0
for row in ws.iter_rows(min_row=1, max_row=1609, min_col=13, max_col=18):
    crow += 1
    rd = {}
    for idx, cell in enumerate(row):
        col = 13 + idx
        val = cell.value
        if val is not None and isinstance(val, str) and val.startswith('='):
            rd[col] = val
        elif val is not None:
            rd[col] = f"CONST:{val}"
    period_data[crow] = rd
wb.close()

def normalize_formula_shift(formula, col_offset):
    """Shift all column references in formula by col_offset positions."""
    if not formula or not formula.startswith('='):
        return formula

    pattern = r'(\$?)([A-Z]{1,3})(\$?)(\d+)'

    def shift_ref(match):
        dollar_col = match.group(1)
        col_letters = match.group(2)
        dollar_row = match.group(3)
        row_num = match.group(4)

        # Don't shift absolute column references
        if dollar_col == '$':
            return match.group(0)

        try:
            col_idx = column_index_from_string(col_letters)
            new_col_idx = col_idx + col_offset
            if new_col_idx >= 1:
                new_col_letters = get_column_letter(new_col_idx)
                return f"{dollar_col}{new_col_letters}{dollar_row}{row_num}"
        except:
            pass
        return match.group(0)

    return re.sub(pattern, shift_ref, formula)

# Compare each period to Period 1 (col 13)
for compare_col in [14, 15, 16, 17, 18]:
    col_offset = compare_col - 13
    identical = 0
    different = 0
    total = 0
    const_match = 0
    const_diff = 0
    diff_examples = []

    for r in range(1, 1610):
        f1 = period_data.get(r, {}).get(13)
        f2 = period_data.get(r, {}).get(compare_col)

        if f1 is not None and f2 is not None:
            if f1.startswith('=') and f2.startswith('='):
                total += 1
                # Shift f1 by col_offset and compare
                f1_shifted = normalize_formula_shift(f1, col_offset)
                if f1_shifted == f2:
                    identical += 1
                else:
                    different += 1
                    if len(diff_examples) < 8:
                        diff_examples.append((r, f1[:80], f2[:80]))
            elif f1.startswith('CONST:') and f2.startswith('CONST:'):
                if f1 == f2:
                    const_match += 1
                else:
                    const_diff += 1

    cl = get_column_letter(compare_col)
    if total > 0:
        pct = identical / total * 100
        log(f"  Period {compare_col-12} (col {cl}) vs Period 1 (col M):")
        log(f"    Formulas: {identical}/{total} identical after shift ({pct:.1f}%), {different} different")
        log(f"    Constants: {const_match} same, {const_diff} different")
        if diff_examples:
            log(f"    Sample differences:")
            for r, f1v, f2v in diff_examples[:5]:
                label = row_labels.get(r, "")[:40]
                log(f"      Row {r} ({label}):")
                log(f"        P1: {f1v}")
                log(f"        P{compare_col-12}: {f2v}")
    else:
        log(f"  Period {compare_col-12} (col {cl}) vs Period 1: no formula comparisons possible")
    log()

# ═══════════════════════════════════════════════════════════════════════════
# STEP 5: FORMULA COMPLEXITY CLASSIFICATION
# ═══════════════════════════════════════════════════════════════════════════
log()
log("=" * 80)
log("STEP 5: FORMULA COMPLEXITY CLASSIFICATION")
log("-" * 60)

COMPLEX_FUNCTIONS = {'INDEX', 'MATCH', 'VLOOKUP', 'HLOOKUP', 'SUMPRODUCT', 'OFFSET',
                     'INDIRECT', 'CHOOSE', 'LOOKUP', 'SUMIFS', 'COUNTIFS'}
MEDIUM_FUNCTIONS = {'SUM', 'IF', 'MIN', 'MAX', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN',
                    'AVERAGE', 'COUNT', 'COUNTA', 'SUMIF', 'COUNTIF',
                    'AND', 'OR', 'NOT', 'IFERROR', 'ISERROR', 'ISBLANK',
                    'LEFT', 'RIGHT', 'MID', 'LEN',
                    'DATE', 'YEAR', 'MONTH', 'DAY', 'EOMONTH', 'EDATE', 'DAYS360',
                    'NPV', 'IRR', 'PMT', 'PPMT', 'IPMT', 'PV', 'FV', 'RATE', 'NPER',
                    'XNPV', 'XIRR', 'YEARFRAC', 'MOD', 'INT', 'CEILING', 'FLOOR'}

def classify_formula(formula):
    if not formula or not formula.startswith('='):
        return None

    f = formula.upper()
    is_cross_sheet = '!' in formula

    func_pattern = r'([A-Z][A-Z0-9_.]+)\s*\('
    functions_used = re.findall(func_pattern, f)
    functions_set = set(functions_used)

    nested_ifs = f.count('IF(')
    max_depth = 0
    depth = 0
    for ch in f:
        if ch == '(':
            depth += 1
            max_depth = max(max_depth, depth)
        elif ch == ')':
            depth -= 1

    is_potentially_circular = 'DSCR' in f or 'CIRCULAR' in f

    if is_cross_sheet:
        category = 'cross-sheet'
    elif functions_set & COMPLEX_FUNCTIONS or nested_ifs >= 3 or max_depth >= 4:
        category = 'complex'
    elif functions_used:
        if len(functions_set) >= 3 or max_depth >= 3:
            category = 'complex'
        else:
            category = 'medium'
    else:
        category = 'simple'

    return {
        'category': category,
        'functions': functions_set,
        'is_cross_sheet': is_cross_sheet,
        'is_potentially_circular': is_potentially_circular,
        'nesting_depth': max_depth,
        'formula_length': len(formula)
    }

# Classify formulas from Period 1 (col M = 13) for all rows
# Also classify helper columns G, H
all_classifications = {}  # row -> {'period1': cls, 'helper_G': cls, 'helper_H': cls}
all_functions_used = Counter()
category_counts = Counter()
cross_sheet_formulas = []

# Classify column M (period 1) formulas
for r in range(1, 1610):
    cls_dict = {}
    # Period 1
    val = all_row_data.get(r, {}).get(13)
    if val and val.startswith('='):
        cls = classify_formula(val)
        if cls:
            cls_dict['period1'] = cls
            category_counts[cls['category']] += 1
            for func in cls['functions']:
                all_functions_used[func] += 1
            if cls['is_cross_sheet']:
                cross_sheet_formulas.append((r, val[:100]))

    # Helper col G
    val_g = all_row_data.get(r, {}).get(7)
    if val_g and val_g.startswith('='):
        cls_g = classify_formula(val_g)
        if cls_g:
            cls_dict['helper_G'] = cls_g

    # Helper col H
    val_h = all_row_data.get(r, {}).get(8)
    if val_h and val_h.startswith('='):
        cls_h = classify_formula(val_h)
        if cls_h:
            cls_dict['helper_H'] = cls_h

    all_classifications[r] = cls_dict

total_formulas = sum(category_counts.values())
log(f"  Total formulas classified (Period 1 / col M): {total_formulas}")
log()
log(f"  Complexity Distribution:")
for cat in ['simple', 'medium', 'complex', 'cross-sheet']:
    cnt = category_counts.get(cat, 0)
    pct = cnt / total_formulas * 100 if total_formulas > 0 else 0
    bar = '#' * int(pct / 2)
    log(f"    {cat:<12s}: {cnt:5d} ({pct:5.1f}%) {bar}")

log()
log(f"  Top 30 functions used (in Period 1 formulas):")
for func, count in all_functions_used.most_common(30):
    log(f"    {func:<20s}: {count:5d}")

# Cross-sheet references
log()
log(f"  Cross-sheet formula examples ({len(cross_sheet_formulas)} total):")
for r, formula in cross_sheet_formulas[:20]:
    label = row_labels.get(r, "")[:40]
    log(f"    Row {r:4d} ({label}): {formula}")
if len(cross_sheet_formulas) > 20:
    log(f"    ... and {len(cross_sheet_formulas) - 20} more")

# Circular reference check
circular_rows = []
for r in range(1, 1610):
    cls_dict = all_classifications.get(r, {})
    for key, cls in cls_dict.items():
        if cls.get('is_potentially_circular'):
            circular_rows.append(r)
            break

log()
if circular_rows:
    log(f"  Potentially circular reference rows ({len(circular_rows)}):")
    for r in circular_rows[:20]:
        label = row_labels.get(r, "")[:60]
        log(f"    Row {r}: {label}")
else:
    log(f"  No rows with 'DSCR' or 'CIRCULAR' keyword detected in formulas.")
    log(f"  (Note: actual circular refs may exist without these keywords)")

# Show ALL formulas for Period 1 (col M)
log()
log("  ALL PERIOD 1 (col M) FORMULAS:")
log("  " + "-" * 100)
formula_rows = 0
for r in range(1, 1610):
    val = all_row_data.get(r, {}).get(13)
    if val and val.startswith('='):
        label = row_labels.get(r, "")[:50]
        cls = all_classifications.get(r, {}).get('period1', {})
        cat = cls.get('category', '?') if cls else '?'
        log(f"    Row {r:4d} [{cat:>11s}] ({label}): {val[:140]}")
        formula_rows += 1
log(f"  Total formula rows in col M: {formula_rows}")

# Show helper column formulas
log()
log("  HELPER COLUMN FORMULAS (cols G-H):")
log("  " + "-" * 100)
for r in range(1, 1610):
    val_g = all_row_data.get(r, {}).get(7)
    val_h = all_row_data.get(r, {}).get(8)
    label = row_labels.get(r, "")[:40]
    if val_g or val_h:
        parts = []
        if val_g:
            parts.append(f"G={str(val_g)[:60]}")
        if val_h:
            parts.append(f"H={str(val_h)[:60]}")
        log(f"    Row {r:4d} ({label}): {' | '.join(parts)}")

# ═══════════════════════════════════════════════════════════════════════════
# STEP 6: BLOCK-LEVEL SUMMARY
# ═══════════════════════════════════════════════════════════════════════════
log()
log("=" * 80)
log("STEP 6: BLOCK-LEVEL SUMMARY")
log("-" * 60)

for i, block in enumerate(blocks):
    block_formulas = 0
    block_constants = 0
    block_empty = 0
    block_categories = Counter()
    block_functions = Counter()
    block_cross_sheet = 0
    block_max_depth = 0
    sample_formulas = []

    for r in block['rows']:
        val = all_row_data.get(r, {}).get(13)  # Period 1
        if val is None:
            block_empty += 1
        elif val.startswith('='):
            block_formulas += 1
            if len(sample_formulas) < 3:
                sample_formulas.append((r, val[:80]))
        elif val.startswith('CONST:'):
            block_constants += 1

        cls_dict = all_classifications.get(r, {})
        p1_cls = cls_dict.get('period1')
        if p1_cls:
            block_categories[p1_cls['category']] += 1
            for func in p1_cls['functions']:
                block_functions[func] += 1
            if p1_cls['is_cross_sheet']:
                block_cross_sheet += 1
            block_max_depth = max(block_max_depth, p1_cls['nesting_depth'])

    total_cells = block_formulas + block_constants + block_empty

    log(f"  Block {i+1}: {block['name'][:55]}")
    log(f"    Rows: {block['start']}-{block['end']} ({block['row_count']} rows)")
    log(f"    Col M: {block_formulas} formulas, {block_constants} constants, {block_empty} empty")

    if block_formulas > 0:
        dist_parts = []
        for cat in ['simple', 'medium', 'complex', 'cross-sheet']:
            cnt = block_categories.get(cat, 0)
            if cnt > 0:
                pct = cnt / block_formulas * 100
                dist_parts.append(f"{cat}:{cnt}({pct:.0f}%)")
        log(f"    Complexity: {', '.join(dist_parts)}")
        log(f"    Max nesting depth: {block_max_depth}")

        top_funcs = block_functions.most_common(8)
        if top_funcs:
            func_str = ", ".join(f"{f}({c})" for f, c in top_funcs)
            log(f"    Key functions: {func_str}")

        if sample_formulas:
            log(f"    Sample formulas:")
            for r, f in sample_formulas:
                label = row_labels.get(r, "")[:30]
                log(f"      Row {r} ({label}): {f}")
    log()

# ═══════════════════════════════════════════════════════════════════════════
# SUMMARY STATISTICS
# ═══════════════════════════════════════════════════════════════════════════
log()
log("=" * 80)
log("OVERALL SUMMARY")
log("-" * 60)
log(f"  Sheet: {SHEET_NAME}")
log(f"  Rows analyzed: 1-1609")
log(f"  Label columns: A-L (1-12)")
log(f"  Time period columns: M-{get_column_letter(max_data_col)} ({max_data_col - 12} periods)")
log(f"  Calc blocks identified: {len(blocks)}")
log(f"  Rows with labels: {sum(1 for r, l in row_labels.items() if l)}")
log(f"  Empty separator rows: {sum(1 for r, l in row_labels.items() if not l)}")
log(f"  Total Period 1 formulas: {total_formulas}")
log(f"  Cross-sheet references: {len(cross_sheet_formulas)}")
log(f"  Unique functions used: {len(all_functions_used)}")
log(f"  Most common function: {all_functions_used.most_common(1)[0] if all_functions_used else 'N/A'}")

# ─── Write output ──────────────────────────────────────────────────────────
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    f.write('\n'.join(out_lines))

print(f"\nAnalysis written to: {OUTPUT_PATH}")
print(f"Total lines: {len(out_lines)}")
