import sys
sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
import re
from collections import defaultdict, Counter
import os
import time

F1_PATH = r"C:\repos\Ampyr-PFA\Ref Docs\converted\Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"
F2_PATH = r"C:\repos\Ampyr-PFA\Ref Docs\converted\Converted FY27 Business Planv8 copy - file with GTC sheets added.xlsm"
OUTPUT_PATH = r"C:\repos\Ampyr-PFA\Ref Docs\analysis_gtc_inputs.txt"

MAX_ROWS = 500

# Regex patterns
FUNC_PATTERN = re.compile(r'([A-Z][A-Z0-9_.]+)\s*\(')
SHEET_REF_PATTERN = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_]+))!")

def classify_complexity(formula):
    """Classify formula complexity."""
    if not formula or not isinstance(formula, str):
        return 'simple'
    f = formula.upper()
    nested = f.count('(')
    funcs = set(FUNC_PATTERN.findall(f))
    has_if = 'IF(' in f or 'IFS(' in f
    has_lookup = any(x in f for x in ['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'OFFSET'])
    has_array = '{' in formula and '}' in formula

    if nested >= 4 or len(funcs) >= 4 or has_array or (has_if and has_lookup):
        return 'complex'
    elif nested >= 2 or len(funcs) >= 2 or has_if or has_lookup:
        return 'medium'
    else:
        return 'simple'

def extract_functions(formula):
    """Extract function names from formula."""
    if not formula or not isinstance(formula, str):
        return set()
    return set(FUNC_PATTERN.findall(formula.upper()))

def extract_sheet_refs(formula):
    """Extract referenced sheet names from formula."""
    if not formula or not isinstance(formula, str):
        return set()
    refs = set()
    for match in SHEET_REF_PATTERN.finditer(formula):
        sheet = match.group(1) or match.group(2)
        if sheet:
            refs.add(sheet)
    return refs

def analyze_sheet(ws, sheet_name, max_rows=MAX_ROWS):
    """Analyze a single worksheet."""
    result = {
        'name': sheet_name,
        'total_cells': 0,
        'formula_cells': 0,
        'hardcoded_cells': 0,
        'functions_used': Counter(),
        'cross_sheet_refs': set(),
        'referenced_sheets': set(),
        'complexity': Counter(),
        'sample_formulas': [],
        'input_levers': 0,
        'rows_analyzed': 0,
        'max_col': 0,
    }

    row_count = 0
    for row in ws.iter_rows(max_row=max_rows):
        row_count += 1
        for cell in row:
            if cell.value is not None and cell.value != '':
                result['total_cells'] += 1
                col_idx = cell.column
                if col_idx > result['max_col']:
                    result['max_col'] = col_idx

                val = cell.value
                if isinstance(val, str) and val.startswith('='):
                    result['formula_cells'] += 1
                    funcs = extract_functions(val)
                    for f in funcs:
                        result['functions_used'][f] += 1

                    sheets = extract_sheet_refs(val)
                    for s in sheets:
                        if s != sheet_name:
                            result['referenced_sheets'].add(s)
                    if sheets:
                        result['cross_sheet_refs'].add(cell.coordinate)

                    complexity = classify_complexity(val)
                    result['complexity'][complexity] += 1

                    if len(result['sample_formulas']) < 5:
                        result['sample_formulas'].append((cell.coordinate, val[:200]))
                else:
                    result['hardcoded_cells'] += 1
                    # Check if it's a numeric input lever
                    if isinstance(val, (int, float)):
                        result['input_levers'] += 1

    result['rows_analyzed'] = row_count
    return result

def format_sheet_analysis(result, show_levers=False):
    """Format analysis result as text."""
    lines = []
    lines.append(f"  Sheet: {result['name']}")
    lines.append(f"    Rows analyzed: {result['rows_analyzed']}, Max column: {result['max_col']}")
    lines.append(f"    Total cells with data: {result['total_cells']}")
    lines.append(f"    Formula cells: {result['formula_cells']}")
    lines.append(f"    Hardcoded value cells: {result['hardcoded_cells']}")

    if show_levers:
        lines.append(f"    Numeric input levers (hardcoded numbers): {result['input_levers']}")

    if result['complexity']:
        lines.append(f"    Formula complexity: simple={result['complexity'].get('simple',0)}, "
                     f"medium={result['complexity'].get('medium',0)}, "
                     f"complex={result['complexity'].get('complex',0)}")

    if result['functions_used']:
        top_funcs = result['functions_used'].most_common(30)
        func_str = ', '.join(f"{f}({c})" for f, c in top_funcs)
        lines.append(f"    Unique functions ({len(result['functions_used'])}): {func_str}")

    if result['referenced_sheets']:
        lines.append(f"    Cross-sheet references to: {', '.join(sorted(result['referenced_sheets']))}")
    else:
        lines.append(f"    Cross-sheet references: None")

    if result['sample_formulas']:
        lines.append(f"    Sample formulas:")
        for coord, formula in result['sample_formulas']:
            lines.append(f"      {coord}: {formula}")

    return '\n'.join(lines)

def main():
    output_lines = []
    output_lines.append("=" * 100)
    output_lines.append("ANALYSIS OF GTC SHEETS AND INPUT SHEETS")
    output_lines.append("=" * 100)
    output_lines.append("")

    # =========================================================================
    # PART A: Input Sheets from F1
    # =========================================================================
    output_lines.append("=" * 100)
    output_lines.append("PART A: INPUT SHEETS FROM F1 (ASE Original)")
    output_lines.append("=" * 100)
    output_lines.append("")

    input_sheet_names = [
        'Project Info', 'Country Inputs', 'Financing Inputs',
        'Time Inputs (A)', 'Time Inputs (M)', 'Time Inputs (Q)', 'Sensis'
    ]

    print("Loading F1 (ASE Original)...")
    t0 = time.time()
    wb1 = openpyxl.load_workbook(F1_PATH, read_only=True, data_only=False, keep_links=False)
    print(f"  Loaded in {time.time()-t0:.1f}s. Sheets: {len(wb1.sheetnames)}")

    f1_sheets = set(wb1.sheetnames)
    output_lines.append(f"F1 total sheets: {len(wb1.sheetnames)}")
    output_lines.append(f"F1 sheet names: {', '.join(wb1.sheetnames[:20])}{'...' if len(wb1.sheetnames) > 20 else ''}")
    output_lines.append("")

    f1_input_results = {}
    for sname in input_sheet_names:
        if sname in wb1.sheetnames:
            print(f"  Analyzing '{sname}'...")
            ws = wb1[sname]
            result = analyze_sheet(ws, sname, max_rows=MAX_ROWS)
            f1_input_results[sname] = result
            output_lines.append(format_sheet_analysis(result, show_levers=True))
            output_lines.append("")
        else:
            output_lines.append(f"  Sheet '{sname}' NOT FOUND in F1")
            output_lines.append("")

    wb1.close()
    print(f"F1 analysis complete.")

    # =========================================================================
    # PART B: GTC Sheets from F2
    # =========================================================================
    output_lines.append("")
    output_lines.append("=" * 100)
    output_lines.append("PART B: GTC SHEETS FROM F2 (GTC Enhanced)")
    output_lines.append("=" * 100)
    output_lines.append("")

    print("Loading F2 (GTC Enhanced)...")
    t0 = time.time()
    wb2 = openpyxl.load_workbook(F2_PATH, read_only=True, data_only=False, keep_links=False)
    print(f"  Loaded in {time.time()-t0:.1f}s. Sheets: {len(wb2.sheetnames)}")

    f2_sheets = set(wb2.sheetnames)
    gtc_sheets = [s for s in wb2.sheetnames if s not in f1_sheets]

    output_lines.append(f"F2 total sheets: {len(wb2.sheetnames)}")
    output_lines.append(f"Sheets in F2 but NOT in F1 (GTC sheets): {len(gtc_sheets)}")
    output_lines.append(f"GTC sheet names:")
    for i, s in enumerate(gtc_sheets):
        output_lines.append(f"  {i+1}. {s}")
    output_lines.append("")

    gtc_results = {}
    all_f2_results = {}

    for sname in gtc_sheets:
        print(f"  Analyzing GTC sheet '{sname}'...")
        ws = wb2[sname]
        result = analyze_sheet(ws, sname, max_rows=MAX_ROWS)
        gtc_results[sname] = result
        all_f2_results[sname] = result
        output_lines.append(format_sheet_analysis(result, show_levers=False))

        # Check if formulas reference ASE sheets
        ase_refs = result['referenced_sheets'] & f1_sheets
        if ase_refs:
            output_lines.append(f"    ** References ASE sheets: {', '.join(sorted(ase_refs))}")
        else:
            output_lines.append(f"    ** Does NOT reference any ASE sheets")
        output_lines.append("")

    # Also analyze common sheets in F2 for cross-ref mapping
    print("  Analyzing common sheets in F2 for dependency mapping...")
    common_sheets = [s for s in wb2.sheetnames if s in f1_sheets]
    for sname in common_sheets:
        print(f"    Analyzing common sheet '{sname}'...")
        ws = wb2[sname]
        result = analyze_sheet(ws, sname, max_rows=MAX_ROWS)
        all_f2_results[sname] = result

    wb2.close()
    print(f"F2 analysis complete.")

    # =========================================================================
    # Structural similarity among GTC sheets
    # =========================================================================
    output_lines.append("")
    output_lines.append("-" * 80)
    output_lines.append("STRUCTURAL SIMILARITY AMONG GTC SHEETS")
    output_lines.append("-" * 80)
    output_lines.append("")

    # Group by similar characteristics
    # Look for naming patterns
    name_groups = defaultdict(list)
    for sname in gtc_sheets:
        # Normalize name to find groups
        base = sname.lower()
        # Remove common suffixes/prefixes
        for pattern in ['by asset', 'aggregate', 'agg', 'summary', 'detail']:
            base = base.replace(pattern, '').strip()
        base = re.sub(r'\s+', ' ', base).strip()
        name_groups[base].append(sname)

    # Report groups with more than one member
    for base, members in sorted(name_groups.items()):
        if len(members) > 1:
            output_lines.append(f"  Similar group (base='{base}'): {', '.join(members)}")
            for m in members:
                r = gtc_results[m]
                output_lines.append(f"    {m}: {r['total_cells']} cells, {r['formula_cells']} formulas, "
                                   f"{r['hardcoded_cells']} hardcoded, max_col={r['max_col']}")
            output_lines.append("")

    # Also group by structural similarity (cell counts, column counts)
    output_lines.append("  Structural similarity by dimensions and formula patterns:")
    struct_groups = defaultdict(list)
    for sname in gtc_sheets:
        r = gtc_results[sname]
        # Create a structural fingerprint
        func_set = frozenset(r['functions_used'].keys())
        key = (r['max_col'], len(r['functions_used']), r['complexity'].get('complex', 0) > 0)
        struct_groups[key].append(sname)

    for key, members in sorted(struct_groups.items()):
        if len(members) > 1:
            output_lines.append(f"    Group (max_col={key[0]}, num_funcs={key[1]}, has_complex={key[2]}):")
            for m in members:
                r = gtc_results[m]
                output_lines.append(f"      {m}: {r['total_cells']} cells, {r['formula_cells']} formulas")
            output_lines.append("")

    # =========================================================================
    # PART C: Cross-Sheet Dependency Map
    # =========================================================================
    output_lines.append("")
    output_lines.append("=" * 100)
    output_lines.append("PART C: CROSS-SHEET DEPENDENCY MAP")
    output_lines.append("=" * 100)
    output_lines.append("")

    output_lines.append("--- Input Sheets (F1) Dependencies ---")
    output_lines.append("")
    for sname, result in f1_input_results.items():
        refs = sorted(result['referenced_sheets'])
        if refs:
            output_lines.append(f"  {sname} --> {', '.join(refs)}")
        else:
            output_lines.append(f"  {sname} --> (no cross-sheet references)")

    output_lines.append("")
    output_lines.append("--- GTC Sheets (F2) Dependencies ---")
    output_lines.append("")
    for sname in gtc_sheets:
        result = gtc_results[sname]
        refs = sorted(result['referenced_sheets'])
        if refs:
            # Separate into ASE refs and GTC refs
            ase_refs = sorted(result['referenced_sheets'] & f1_sheets)
            gtc_refs = sorted(result['referenced_sheets'] - f1_sheets)
            parts = []
            if ase_refs:
                parts.append(f"ASE sheets: {', '.join(ase_refs)}")
            if gtc_refs:
                parts.append(f"GTC/Other sheets: {', '.join(gtc_refs)}")
            output_lines.append(f"  {sname} --> {'; '.join(parts)}")
        else:
            output_lines.append(f"  {sname} --> (no cross-sheet references)")

    output_lines.append("")
    output_lines.append("--- All F2 Sheets Dependencies (common sheets included) ---")
    output_lines.append("")
    for sname in wb2.sheetnames if hasattr(wb2, 'sheetnames') else sorted(all_f2_results.keys()):
        if sname in all_f2_results:
            result = all_f2_results[sname]
            refs = sorted(result['referenced_sheets'])
            if refs:
                output_lines.append(f"  {sname} --> {', '.join(refs)}")
            else:
                output_lines.append(f"  {sname} --> (no cross-sheet references)")

    # Fix: we closed wb2, need to use saved sheet order
    # Re-sort by all_f2_results keys
    output_lines.append("")

    # =========================================================================
    # Summary Statistics
    # =========================================================================
    output_lines.append("=" * 100)
    output_lines.append("SUMMARY STATISTICS")
    output_lines.append("=" * 100)
    output_lines.append("")

    output_lines.append("--- Input Sheets Summary ---")
    total_input_cells = sum(r['total_cells'] for r in f1_input_results.values())
    total_input_formulas = sum(r['formula_cells'] for r in f1_input_results.values())
    total_input_hardcoded = sum(r['hardcoded_cells'] for r in f1_input_results.values())
    total_input_levers = sum(r['input_levers'] for r in f1_input_results.values())
    all_input_funcs = set()
    for r in f1_input_results.values():
        all_input_funcs.update(r['functions_used'].keys())

    output_lines.append(f"  Total cells: {total_input_cells}")
    output_lines.append(f"  Formula cells: {total_input_formulas}")
    output_lines.append(f"  Hardcoded cells: {total_input_hardcoded}")
    output_lines.append(f"  Numeric input levers: {total_input_levers}")
    output_lines.append(f"  Unique functions across all input sheets: {len(all_input_funcs)}")
    output_lines.append(f"  Functions: {', '.join(sorted(all_input_funcs))}")
    output_lines.append("")

    output_lines.append("--- GTC Sheets Summary ---")
    total_gtc_cells = sum(r['total_cells'] for r in gtc_results.values())
    total_gtc_formulas = sum(r['formula_cells'] for r in gtc_results.values())
    total_gtc_hardcoded = sum(r['hardcoded_cells'] for r in gtc_results.values())
    all_gtc_funcs = set()
    for r in gtc_results.values():
        all_gtc_funcs.update(r['functions_used'].keys())
    total_simple = sum(r['complexity'].get('simple', 0) for r in gtc_results.values())
    total_medium = sum(r['complexity'].get('medium', 0) for r in gtc_results.values())
    total_complex = sum(r['complexity'].get('complex', 0) for r in gtc_results.values())

    output_lines.append(f"  Total GTC sheets: {len(gtc_sheets)}")
    output_lines.append(f"  Total cells: {total_gtc_cells}")
    output_lines.append(f"  Formula cells: {total_gtc_formulas}")
    output_lines.append(f"  Hardcoded cells: {total_gtc_hardcoded}")
    output_lines.append(f"  Formula complexity: simple={total_simple}, medium={total_medium}, complex={total_complex}")
    output_lines.append(f"  Unique functions across all GTC sheets: {len(all_gtc_funcs)}")
    output_lines.append(f"  Functions: {', '.join(sorted(all_gtc_funcs))}")
    output_lines.append("")

    # GTC sheets that reference ASE sheets
    gtc_ref_ase = [s for s in gtc_sheets if gtc_results[s]['referenced_sheets'] & f1_sheets]
    output_lines.append(f"  GTC sheets referencing ASE sheets: {len(gtc_ref_ase)}")
    for s in gtc_ref_ase:
        ase_refs = sorted(gtc_results[s]['referenced_sheets'] & f1_sheets)
        output_lines.append(f"    {s} --> {', '.join(ase_refs)}")

    output_lines.append("")
    output_lines.append("=" * 100)
    output_lines.append("END OF ANALYSIS")
    output_lines.append("=" * 100)

    # Write output
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_lines))

    print(f"\nAnalysis written to: {OUTPUT_PATH}")
    print(f"Total lines: {len(output_lines)}")

if __name__ == '__main__':
    main()
