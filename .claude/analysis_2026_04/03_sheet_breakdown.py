"""03 - Per-sheet cell/formula/complexity breakdown (read_only streaming)."""
import json
import re
import sys
from collections import Counter, defaultdict
from openpyxl import load_workbook
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")

FILES = {
    "F1": "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm",
    "F2": "Converted FY27 Business Planv8 copy - file with GTC sheets added.xlsm",
    "F3": "converted Project Parthenon - FM - raw file previous quarter.xlsm",
}

FN_RE = re.compile(r"\b([A-Z][A-Z0-9._]*)\s*\(")
SHEET_REF_RE = re.compile(r"'([^']+)'!|([A-Za-z_][\w]*)!")

def classify_formula(f: str) -> str:
    if not f or not f.startswith("="):
        return "other"
    body = f[1:]
    fns = FN_RE.findall(body)
    if not fns:
        # pure arithmetic / reference
        return "simple"
    # complex if has nested IF/INDEX/MATCH with multiple functions, or > 3 unique fns
    uniq = set(fns)
    complex_markers = {"INDEX", "MATCH", "SUMIFS", "SUMIF", "XLOOKUP", "OFFSET", "INDIRECT"}
    if len(uniq) >= 4 or (uniq & complex_markers and len(fns) >= 3):
        return "complex"
    return "medium"

def analyse_sheet(ws):
    n_cells = n_f = n_hc_num = n_hc_txt = n_blank = 0
    fn_counter = Counter()
    sheet_refs = Counter()
    buckets = {"simple": 0, "medium": 0, "complex": 0, "other": 0}
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if v is None:
                continue
            n_cells += 1
            if isinstance(v, str) and v.startswith("="):
                n_f += 1
                body = v[1:]
                for fn in FN_RE.findall(body):
                    fn_counter[fn] += 1
                for m in SHEET_REF_RE.findall(body):
                    sheet = m[0] if m[0] else m[1]
                    if sheet:
                        sheet_refs[sheet] += 1
                buckets[classify_formula(v)] += 1
            elif isinstance(v, (int, float)):
                n_hc_num += 1
            else:
                n_hc_txt += 1
    return {
        "cells": n_cells,
        "formulas": n_f,
        "hardcoded_numeric": n_hc_num,
        "hardcoded_text": n_hc_txt,
        "complexity": buckets,
        "fn_top20": fn_counter.most_common(20),
        "fn_total_calls": sum(fn_counter.values()),
        "cross_sheet_top10": sheet_refs.most_common(10),
    }

for key, fname in FILES.items():
    print(f"\n=== {key}: {fname} ===", flush=True)
    wb = load_workbook(DATA / fname, read_only=True, data_only=False, keep_links=False)
    out = {}
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        stat = analyse_sheet(ws)
        out[sheet] = stat
        print(f"  {sheet:40} cells={stat['cells']:>7}  f={stat['formulas']:>7}  hc#={stat['hardcoded_numeric']:>6}  "
              f"s/m/c={stat['complexity']['simple']}/{stat['complexity']['medium']}/{stat['complexity']['complex']}", flush=True)
    wb.close()
    (OUT / f"03_sheet_breakdown_{key}.json").write_text(json.dumps(out, indent=2, default=str))
    print(f"  [wrote {OUT}/03_sheet_breakdown_{key}.json]", flush=True)
