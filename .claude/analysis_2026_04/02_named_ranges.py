"""02 - Named ranges across all 3 files (workbook + sheet-scoped)."""
import json
from openpyxl import load_workbook
from pathlib import Path

DATA = Path("c:/repos/Ampyr-PFA/data/converted")
OUT = Path("c:/repos/Ampyr-PFA/.claude/analysis_2026_04")

FILES = {
    "F1": "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm",
    "F2": "Converted FY27 Business Planv8 copy - file with GTC sheets added.xlsm",
    "F3": "converted Project Parthenon - FM - raw file previous quarter.xlsm",
}

report = {}
for key, fname in FILES.items():
    wb = load_workbook(DATA / fname, read_only=False, data_only=False, keep_links=False)
    wb_names = []
    for dn in wb.defined_names.values() if hasattr(wb.defined_names, "values") else []:
        wb_names.append({
            "name": dn.name,
            "value": dn.value,
            "comment": dn.comment,
            "hidden": dn.hidden,
        })
    # newer openpyxl: DefinedNameDict iteration
    try:
        for name, dn in wb.defined_names.items():
            wb_names.append({
                "name": name,
                "value": dn.value if hasattr(dn, "value") else str(dn),
                "hidden": getattr(dn, "hidden", False),
            })
    except Exception as e:
        print(f"  iterate: {e}")

    # sheet-scoped
    sheet_names_map = {}
    for ws in wb.worksheets:
        try:
            sn = list(ws.defined_names.items()) if hasattr(ws, "defined_names") else []
            if sn:
                sheet_names_map[ws.title] = [(n, str(d.value) if hasattr(d, "value") else str(d)) for n, d in sn]
        except Exception:
            pass

    # dedup wb_names by (name,value)
    seen = set()
    dedup = []
    for n in wb_names:
        k = (n.get("name"), n.get("value"))
        if k in seen:
            continue
        seen.add(k)
        dedup.append(n)

    report[key] = {
        "wb_named_range_count": len(dedup),
        "sheet_scoped_count": sum(len(v) for v in sheet_names_map.values()),
        "sheet_scoped_names": sheet_names_map,
        "wb_named_ranges": dedup,
    }
    print(f"{key}: {len(dedup)} workbook-scoped names, {sum(len(v) for v in sheet_names_map.values())} sheet-scoped")
    wb.close()

(OUT / "02_named_ranges.json").write_text(json.dumps(report, indent=2, default=str))

# Compare F1 vs F3 names
f1 = {n["name"]: n.get("value") for n in report["F1"]["wb_named_ranges"]}
f3 = {n["name"]: n.get("value") for n in report["F3"]["wb_named_ranges"]}
f2 = {n["name"]: n.get("value") for n in report["F2"]["wb_named_ranges"]}

print(f"\nF1 only (not in F3): {sorted(set(f1)-set(f3))[:30]}")
print(f"F3 only (not in F1): {sorted(set(f3)-set(f1))[:30]}")
print(f"F2 adds vs F1: {len(set(f2)-set(f1))} new names; removes: {len(set(f1)-set(f2))}")
print(f"F2 new names sample: {sorted(set(f2)-set(f1))[:20]}")

# Value diffs for shared names between F1 & F3
shared_vals_diff = [(n, f1[n], f3[n]) for n in sorted(set(f1)&set(f3)) if f1[n] != f3[n]]
print(f"\nShared names with different value (F1 vs F3): {len(shared_vals_diff)}")
for n, a, b in shared_vals_diff[:15]:
    print(f"  {n}: F1={a}  |  F3={b}")
