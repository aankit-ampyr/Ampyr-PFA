"""One-off probe: what does Project Info row 7 actually contain?

The active flag question (ADR 0001) hinges on whether each cell holds the
text "True" or the boolean TRUE. We can answer that directly by reading
F1 and reporting cell types per slot.
"""
from openpyxl import load_workbook
from pathlib import Path

F1 = Path(__file__).parent.parent / "data" / "converted" / \
    "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"

wb = load_workbook(F1, data_only=True, read_only=True)
ws = wb["Project Info"]

# Row 7 (1-indexed), columns G (7) through DR (122) — slots 1-120
type_counts: dict[str, int] = {}
samples: list[tuple[int, str, str]] = []
true_count = false_count = none_count = 0

for col_idx, cell in enumerate(next(ws.iter_rows(min_row=7, max_row=7,
                                                  min_col=7, max_col=126,
                                                  values_only=False)), start=7):
    v = cell.value
    t = type(v).__name__
    type_counts[t] = type_counts.get(t, 0) + 1
    if v is True or (isinstance(v, str) and v.lower() == "true"):
        true_count += 1
    elif v is False or (isinstance(v, str) and v.lower() == "false"):
        false_count += 1
    else:
        none_count += 1
    if len(samples) < 8:
        samples.append((col_idx, t, repr(v)))

print(f"Row 7 of Project Info — slot active flags (cols G:DR):")
print(f"  type breakdown: {type_counts}")
print(f"  TRUE-equivalent: {true_count}, FALSE-equivalent: {false_count}, other/None: {none_count}")
print(f"  first 8 samples:")
for col, t, val in samples:
    print(f"    col {col}: type={t}  value={val}")

# Also show the row 5 (asset names) for the first 5 actives + 5 placeholders
print()
print("First 10 active-flag cells with their asset names (row 5):")
row5 = list(next(ws.iter_rows(min_row=5, max_row=5, min_col=7, max_col=126, values_only=True)))
row7_vals = list(next(ws.iter_rows(min_row=7, max_row=7, min_col=7, max_col=126, values_only=True)))
for i, (name, flag) in enumerate(zip(row5, row7_vals)):
    if i < 10:
        print(f"  slot {i+1}: name={name!r}  flag={flag!r} (type={type(flag).__name__})")

wb.close()
