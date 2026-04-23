"""07 - Quarterly diff F1 (2026-01-08) vs F3 (previous quarter).
- For shared sheets: per-cell diff on input sheets, hash-based diff on large sheets
- Detect structural changes: row insertions/deletions, new columns
"""
import json
import hashlib
from collections import Counter
from openpyxl import load_workbook
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")

F1 = DATA / "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"
F3 = DATA / "converted Project Parthenon - FM - raw file previous quarter.xlsm"

# Input sheets — small enough for cell-level diff on VALUES (data_only=True)
INPUT_SHEETS = ["Project Info", "Country Inputs", "Financing Inputs", "Time Inputs (A)", "Time Inputs (Q)", "Sensis"]

def hash_cell(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return f"f:{round(v, 8)}"
    return f"{type(v).__name__}:{v}"

def diff_sheet(ws1, ws3, max_rows=None):
    """Return dict: changed, added_rows, removed_rows, changes_sample"""
    changes = []
    seen_cells = 0
    m1, m3 = ws1.max_row, ws3.max_row
    max_r = max(m1, m3)
    if max_rows: max_r = min(max_r, max_rows)
    max_c = max(ws1.max_column, ws3.max_column)
    # fast iteration
    it1 = ws1.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)
    it3 = ws3.iter_rows(min_row=1, max_row=max_r, max_col=max_c, values_only=True)
    for r_idx, (r1, r3) in enumerate(zip(it1, it3), start=1):
        r1_padded = list(r1) + [None]*(max_c - len(r1))
        r3_padded = list(r3) + [None]*(max_c - len(r3))
        for c_idx, (v1, v3) in enumerate(zip(r1_padded, r3_padded), start=1):
            h1, h3 = hash_cell(v1), hash_cell(v3)
            if h1 != h3:
                changes.append((r_idx, c_idx, v1, v3))
            seen_cells += 1
    return changes, seen_cells

# Use data_only=True to compare cached values (input sheets should have values; formula sheets have cache)
print("Loading F1 (data_only)...", flush=True)
wb1 = load_workbook(F1, read_only=True, data_only=True, keep_links=False)
print("Loading F3 (data_only)...", flush=True)
wb3 = load_workbook(F3, read_only=True, data_only=True, keep_links=False)

diff_report = {}

for sheet in INPUT_SHEETS:
    if sheet not in wb1.sheetnames or sheet not in wb3.sheetnames:
        print(f"  skip {sheet}")
        continue
    ws1, ws3 = wb1[sheet], wb3[sheet]
    print(f"\n{sheet}: F1 dim {ws1.max_row}x{ws1.max_column}, F3 dim {ws3.max_row}x{ws3.max_column}", flush=True)
    changes, n = diff_sheet(ws1, ws3)
    print(f"  {n} cells scanned, {len(changes)} differ", flush=True)
    # sample
    sample = [(r, c, str(a)[:60], str(b)[:60]) for r, c, a, b in changes[:30]]
    diff_report[sheet] = {
        "cells_scanned": n,
        "changes_count": len(changes),
        "pct_changed": round(100.0 * len(changes) / max(n, 1), 2),
        "sample_changes": sample,
    }

# For large formula sheets, do hash-of-formulas diff (need formulas — reload without data_only)
print("\n\nReloading with formulas for PLW...")
wb1f = load_workbook(F1, read_only=True, data_only=False, keep_links=False)
wb3f = load_workbook(F3, read_only=True, data_only=False, keep_links=False)

for sheet in ["Project Level Workings"]:
    ws1, ws3 = wb1f[sheet], wb3f[sheet]
    print(f"\n{sheet}: F1 {ws1.max_row}x{ws1.max_column} vs F3 {ws3.max_row}x{ws3.max_column}")
    # row-by-row hash (concat of all formulas in row)
    def row_hash(ws):
        hashes = {}
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            h = hashlib.md5()
            for v in row:
                if isinstance(v, str) and v.startswith("="):
                    h.update(v.encode("utf-8", errors="ignore"))
            hashes[r_idx] = h.hexdigest()[:12]
        return hashes

    h1 = row_hash(ws1)
    h3 = row_hash(ws3)
    diff_rows = [r for r in h1 if r in h3 and h1[r] != h3[r]]
    print(f"  Rows with changed formula hash: {len(diff_rows)} (of {min(len(h1), len(h3))})")
    print(f"  Sample changed rows: {diff_rows[:20]}")
    diff_report[sheet] = {
        "f1_rows": len(h1),
        "f3_rows": len(h3),
        "changed_formula_rows_count": len(diff_rows),
        "changed_rows_sample": diff_rows[:40],
    }

(OUT / "07_quarterly_diff.json").write_text(json.dumps(diff_report, indent=2, default=str))
print(f"\nWritten {OUT/'07_quarterly_diff.json'}")
