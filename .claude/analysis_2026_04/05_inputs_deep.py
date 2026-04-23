"""05 - Input-sheet deep dive for F1.
- Project Info: identify per-asset parameter columns, count levers per asset, find unique parameter names
- Country Inputs, Financing Inputs: extract full structure
- Time Inputs (A/M/Q): curve names and cardinalities
- Sensis: scenario slot structure
"""
import json
import re
from collections import Counter, defaultdict
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")

PATH = DATA / "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"
wb = load_workbook(PATH, read_only=True, data_only=True, keep_links=False)  # data_only so we see cached values

# --- Project Info
ws = wb["Project Info"]
pi_rows = []
for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=15, values_only=True):
    pi_rows.append(r)
print(f"Project Info: {len(pi_rows)} rows scanned (up to col 15)")

# row 1 is usually asset-slot headers; find header row
# from docs there are 120 slots across cols. Find row with most non-null in cols 10+
asset_header_row = None
max_nonnull = 0
for i, r in enumerate(pi_rows[:10]):
    nn = sum(1 for v in r if v)
    if nn > max_nonnull:
        max_nonnull = nn
        asset_header_row = i + 1
# scan only first 15 cols here; asset slot headers are further right. Re-scan row 1 with all cols
ws_full = wb["Project Info"]
row1 = [c.value for c in ws_full[1]]
row2 = [c.value for c in ws_full[2]]
row3 = [c.value for c in ws_full[3]]
row4 = [c.value for c in ws_full[4]]
print(f"  Row1 non-null cells (likely asset slots): {sum(1 for v in row1 if v)} / {len(row1)}")
print(f"  Row1 sample (cols 10-25): {row1[9:25]}")
# Find where asset names start
names_in_row1 = [v for v in row1 if isinstance(v, str) and len(v) > 2]
print(f"  Row1 strings (probably section labels): {names_in_row1[:10]}")

# Count column B labels (parameter names)
param_col = 2  # assume col B = parameter names
params = []
for r in range(1, ws_full.max_row + 1):
    v = ws_full.cell(row=r, column=param_col).value
    if isinstance(v, str) and v.strip():
        params.append((r, v.strip()[:80]))
print(f"  Col B labels: {len(params)} non-empty")
print(f"  First 20 params: {params[:20]}")
print(f"  Last 10 params: {params[-10:]}")

# --- Country Inputs
ws = wb["Country Inputs"]
ci_data = []
for r in ws.iter_rows(values_only=True):
    row = [v for v in r if v is not None]
    if row:
        ci_data.append(r)
print(f"\nCountry Inputs: {len(ci_data)} non-empty rows")
# print first 40 rows labels
print("  First 40 rows:")
for i, r in enumerate(ci_data[:40]):
    print(f"    {i+1}: {list(r)[:8]}")

# --- Financing Inputs
ws = wb["Financing Inputs"]
fi_data = []
for r in ws.iter_rows(values_only=True):
    if any(v is not None for v in r):
        fi_data.append(r)
print(f"\nFinancing Inputs: {len(fi_data)} non-empty rows (total sheet rows {ws.max_row})")
for i, r in enumerate(fi_data[:30]):
    print(f"    {i+1}: {list(r)[:8]}")

# --- Time Inputs (A)
ws = wb["Time Inputs (A)"]
ta_labels = []
for r in range(1, ws.max_row + 1):
    v = ws.cell(row=r, column=2).value  # col B
    if isinstance(v, str) and v.strip():
        ta_labels.append((r, v.strip()[:80]))
print(f"\nTime Inputs (A): {len(ta_labels)} labels in col B")
print(f"  First 30: {ta_labels[:30]}")

# --- Sensis
ws = wb["Sensis"]
sensis = []
for r in ws.iter_rows(min_row=1, max_row=40, values_only=True):
    sensis.append([v for v in r[:40] if v is not None][:15])
print(f"\nSensis first 40 rows (col A-N):")
for i, r in enumerate(sensis):
    if r:
        print(f"    {i+1}: {r}")

wb.close()

# Save
(OUT / "05_inputs_pi_params.json").write_text(json.dumps({"params": params}, indent=2, default=str))
