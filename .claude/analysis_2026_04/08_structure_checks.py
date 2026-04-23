"""08 - Structural checks: data validations, hidden rows/columns, conditional formatting.
Fast: uses non-read-only workbook so we can see data validations."""
import json
from collections import Counter
from openpyxl import load_workbook
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")
PATH = DATA / "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"

print("Loading non-read-only (slower)...", flush=True)
wb = load_workbook(PATH, read_only=False, data_only=False, keep_links=False)

report = {}
for sheet in wb.sheetnames:
    ws = wb[sheet]
    dv = len(ws.data_validations.dataValidation) if hasattr(ws, "data_validations") and ws.data_validations else 0
    cf = len(list(ws.conditional_formatting)) if hasattr(ws, "conditional_formatting") else 0
    # hidden rows / columns
    hidden_rows = sum(1 for d in ws.row_dimensions.values() if d.hidden)
    hidden_cols = sum(1 for d in ws.column_dimensions.values() if d.hidden)
    report[sheet] = {
        "data_validations": dv,
        "conditional_formatting_rules": cf,
        "hidden_rows": hidden_rows,
        "hidden_cols": hidden_cols,
        "sheet_state": ws.sheet_state,
        "protection_enabled": ws.protection.sheet if hasattr(ws, "protection") else False,
    }
    if dv or cf or hidden_rows or hidden_cols:
        print(f"  {sheet:40} dv={dv} cf={cf} hidden_rows={hidden_rows} hidden_cols={hidden_cols} state={ws.sheet_state}")

wb.close()
(OUT / "08_structure.json").write_text(json.dumps(report, indent=2, default=str))
print("\nDone.")
