"""01 - File inventory & sheet map across all 3 .xlsm files."""
import json
import sys
from openpyxl import load_workbook
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")

FILES = {
    "F1_ASE_Latest": "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm",
    "F2_GTC_Enhanced": "Converted FY27 Business Planv8 copy - file with GTC sheets added.xlsm",
    "F3_ASE_Previous": "converted Project Parthenon - FM - raw file previous quarter.xlsm",
}

report = {}

for key, fname in FILES.items():
    path = DATA / fname
    print(f"\n=== {key}: {fname} ({path.stat().st_size/1024/1024:.1f} MB) ===")
    # read_only for speed; keep_vba won't work on .xlsm saved via conversion but try
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        sheets.append({
            "name": name,
            "state": ws.sheet_state,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "dim": ws.calculate_dimension() if hasattr(ws, "calculate_dimension") else None,
        })
        print(f"  [{ws.sheet_state:7}] {name:40} {ws.max_row} x {ws.max_column}")
    report[key] = {
        "file": fname,
        "size_mb": round(path.stat().st_size/1024/1024, 2),
        "n_sheets": len(sheets),
        "sheets": sheets,
    }
    wb.close()

(OUT / "01_inventory.json").write_text(json.dumps(report, indent=2))
print(f"\nWritten {OUT/'01_inventory.json'}")
