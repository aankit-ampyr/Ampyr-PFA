"""Build the per-asset parameter inventory from F1 Project Info.

P1 deliverable per ADR 0004 (hybrid naming) + 0005 (full metadata).

Usage:
    python scripts/build_parameter_inventory.py
    python scripts/build_parameter_inventory.py --source path/to/F1.xlsm --out path/to/out.json
    python scripts/build_parameter_inventory.py --skip-formula-pass   # faster, marks all input

Reads:
    Project Info rows 4-596 across cols I:DX (120 slots horizontal). Two passes:
      - data_only=True  -> cached values
      - data_only=False -> formula text (to mark is_input_vs_derived)

Writes:
    data/parameter_inventory.json — {meta: ..., params: [...]} shape.

Per-row record (ADR 0005 option C — full metadata):
    {
      "asset_id": int,                  # slot_index 1..120
      "param_id": str,                  # canonical snake_case (display_name -> id)
      "display_name": str,              # raw Excel label (verbatim)
      "section": str,                   # most-recent col B section header
      "section_path": str,              # "Section > Subsection" (best-effort)
      "row": int,
      "source_cell": str,               # e.g. "I87" — the cell holding this value
      "sheet": "Project Info",
      "units": str | None,              # heuristic, often null
      "scale": float,                   # default 1.0
      "data_type": str,                 # numeric | text | boolean | date | enum | empty
      "is_input_vs_derived": str,       # input | derived | unknown
      "value": Any,                     # cached value
      "unit_hint": str | None,          # raw col E content (selector / unit cue)
    }
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, date, timezone
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent.parent
DEFAULT_SOURCE = ROOT / "data" / "converted" / \
    "Converted 2026-01-08 Project Parthenon - CIP_ASE_Internal_v18 - original file latest quarter.xlsm"
DEFAULT_OUT = ROOT / "data" / "parameter_inventory.json"

SHEET = "Project Info"
ROW_START = 4
ROW_END = 596
SLOT_COL_START = 9    # I
SLOT_COL_END = 128    # DX
LABEL_COL_B = 2
LABEL_COL_C = 3
LABEL_COL_D = 4
LABEL_COL_E = 5       # often a unit / selector hint


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def to_param_id(label: str) -> str:
    """Canonical snake_case from Excel label.

    Rules: lowercase; non-alphanumeric runs -> '_'; trim leading/trailing '_';
    collapse repeats. Preserves embedded numbers and any typos (per ADR 0004
    hybrid: display_name is verbatim, param_id is mechanical).
    """
    s = label.lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def detect_data_type(v) -> str:
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return "empty"
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, (int, float)):
        return "numeric"
    if isinstance(v, (datetime, date)):
        return "date"
    if isinstance(v, str):
        return "text"
    return "unknown"


def is_formula(formula_pass_value) -> bool:
    """A cell is a formula if openpyxl returns a string starting with '='
    when read with data_only=False."""
    return isinstance(formula_pass_value, str) and formula_pass_value.startswith("=")


# -----------------------------------------------------------------------------
# Build
# -----------------------------------------------------------------------------

def build(source: Path, out: Path, skip_formula_pass: bool = False) -> dict:
    if not source.exists():
        raise FileNotFoundError(source)

    print(f"Loading {source.name} (data_only=True, pass 1/2)...", flush=True)
    wb_v = load_workbook(source, data_only=True, read_only=True)
    ws_v = wb_v[SHEET]

    # Read all rows and slot columns in one pass into a dict
    # row_data[r] = (label_b, label_c, label_d, label_e, [slot1_value, slot2_value, ...])
    print("  scanning rows...", flush=True)
    row_data: dict[int, tuple] = {}
    for r in range(ROW_START, ROW_END + 1):
        row = next(ws_v.iter_rows(min_row=r, max_row=r,
                                   min_col=1, max_col=SLOT_COL_END,
                                   values_only=True))
        label_b = row[LABEL_COL_B - 1]
        label_c = row[LABEL_COL_C - 1]
        label_d = row[LABEL_COL_D - 1]
        label_e = row[LABEL_COL_E - 1]
        slot_values = list(row[SLOT_COL_START - 1:SLOT_COL_END])
        row_data[r] = (label_b, label_c, label_d, label_e, slot_values)
    wb_v.close()

    # Second pass: read formulas (data_only=False) to detect input vs derived
    formula_pass: dict[int, list] = {}
    if not skip_formula_pass:
        print(f"Loading {source.name} (data_only=False, pass 2/2)...", flush=True)
        wb_f = load_workbook(source, data_only=False, read_only=True)
        ws_f = wb_f[SHEET]
        print("  scanning formulas...", flush=True)
        for r in range(ROW_START, ROW_END + 1):
            row = next(ws_f.iter_rows(min_row=r, max_row=r,
                                       min_col=SLOT_COL_START,
                                       max_col=SLOT_COL_END,
                                       values_only=True))
            formula_pass[r] = list(row)
        wb_f.close()
    else:
        print("Skipping formula pass (per --skip-formula-pass)", flush=True)

    # Walk the rows: maintain section context, emit param entries
    print("Building entries...", flush=True)
    entries: list[dict] = []
    current_section: str | None = None
    current_subsection: str | None = None

    for r in range(ROW_START, ROW_END + 1):
        label_b, label_c, label_d, label_e, slot_values = row_data[r]

        # Section header detection: col B has a string AND col C is empty
        # OR col B has a string and is a top-level label.
        if isinstance(label_b, str) and label_b.strip() and (label_c is None or
                                                              str(label_c).strip() == ""):
            # Top-level section header (e.g. "Timings", "Production")
            current_section = label_b.strip()
            current_subsection = None
            continue

        # Sub-section: col B has a number/string AND col C has the sub-section name
        # Heuristic: if col B is just a number (1, 2, 3...) it's an instance index
        # like "PPA contract 1, 2, 3" — keep the col C as label, no subsection change.
        if isinstance(label_b, (int, float)) and isinstance(label_c, str) and label_c.strip():
            # Instance row: col B=index, col C=name. Treat as a labeled sub-block start.
            current_subsection = f"{label_c.strip()}"
            # Don't emit; this is a header row for the sub-block

            # But sub-block headers can also have values (uncommon). If any slot has
            # a value here, fall through and emit it. Most don't, so this is rare.
            if any(v is not None for v in slot_values):
                pass  # fall through to emission below
            else:
                continue

        # Param row: col C has a label AND any slot has a value (or label_b is non-empty)
        label_text = None
        if isinstance(label_c, str) and label_c.strip():
            label_text = label_c.strip()
        elif isinstance(label_b, str) and label_b.strip():
            # Edge case: param-like row with only col B label
            label_text = label_b.strip()

        if not label_text:
            continue

        # Skip rows that are entirely empty (no slot has a value AND label is generic)
        if all(v is None for v in slot_values) and not label_text:
            continue

        # Emit one entry per slot that has a non-None value (or all 120, depending)
        # We emit ALL 120 slots (even None values) so the schema can be uniform.
        section = current_section or "Unknown"
        section_path = section if not current_subsection else f"{section} > {current_subsection}"
        param_id = to_param_id(label_text)
        unit_hint = None
        if isinstance(label_e, str) and label_e.strip():
            unit_hint = label_e.strip()

        for slot_idx, value in enumerate(slot_values, start=1):
            # Determine input vs derived using the formula pass for THIS cell
            iv_d = "unknown"
            if not skip_formula_pass:
                fp_row = formula_pass.get(r, [])
                if slot_idx - 1 < len(fp_row):
                    fv = fp_row[slot_idx - 1]
                    if fv is None:
                        # both passes None -> empty cell
                        iv_d = "input"  # default; effectively N/A
                    elif is_formula(fv):
                        iv_d = "derived"
                    else:
                        iv_d = "input"

            entry = {
                "asset_id": slot_idx,
                "param_id": param_id,
                "display_name": label_text,
                "section": section,
                "section_path": section_path,
                "row": r,
                "source_cell": f"{get_column_letter(SLOT_COL_START + slot_idx - 1)}{r}",
                "sheet": SHEET,
                "units": None,             # heuristic — left null for now
                "scale": 1.0,
                "data_type": detect_data_type(value),
                "is_input_vs_derived": iv_d,
                "value": value,
                "unit_hint": unit_hint,
            }
            entries.append(entry)

    # Build counts
    type_counts: dict[str, int] = {}
    iv_counts: dict[str, int] = {}
    section_counts: dict[str, int] = {}
    for e in entries:
        type_counts[e["data_type"]] = type_counts.get(e["data_type"], 0) + 1
        iv_counts[e["is_input_vs_derived"]] = iv_counts.get(e["is_input_vs_derived"], 0) + 1
        section_counts[e["section"]] = section_counts.get(e["section"], 0) + 1

    payload = {
        "meta": {
            "version_source": source.name,
            "extracted_at": datetime.now(timezone.utc).isoformat(),
            "extracted_by": "scripts/build_parameter_inventory.py",
            "row_range": [ROW_START, ROW_END],
            "slot_range": [1, 120],
            "schema_version": 1,
            "decisions_applied": ["ADR-0004", "ADR-0005"],
            "entry_count": len(entries),
            "data_type_counts": type_counts,
            "input_vs_derived_counts": iv_counts,
            "distinct_sections": len(section_counts),
        },
        "params": entries,
    }

    out.parent.mkdir(parents=True, exist_ok=True)
    print(f"Writing {out.relative_to(ROOT)} ({len(entries):,} entries)...", flush=True)
    out.write_text(json.dumps(payload, default=str), encoding="utf-8")
    return payload


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    ap.add_argument("--skip-formula-pass", action="store_true",
                    help="Skip the data_only=False pass; marks all is_input_vs_derived=unknown.")
    args = ap.parse_args()

    payload = build(args.source, args.out, args.skip_formula_pass)
    meta = payload["meta"]

    # Summary
    file_kb = args.out.stat().st_size / 1024
    print()
    print(f"DONE — {args.out.relative_to(ROOT)} ({file_kb:,.0f} KB)")
    print(f"  source: {meta['version_source']}")
    print(f"  entries: {meta['entry_count']:,}")
    print(f"  data type counts: {meta['data_type_counts']}")
    print(f"  input vs derived: {meta['input_vs_derived_counts']}")
    print(f"  distinct sections: {meta['distinct_sections']}")
    print()
    print("Top 10 sections by entry count:")
    sec_counts = {e["section"]: 0 for e in payload["params"]}
    for e in payload["params"]:
        sec_counts[e["section"]] += 1
    for sec, n in sorted(sec_counts.items(), key=lambda kv: -kv[1])[:10]:
        print(f"  {sec[:50]:50} {n:>6,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
