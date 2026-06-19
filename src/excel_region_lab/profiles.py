from __future__ import annotations

from typing import Any

from .schema import Box
from .signals import SheetSignals, row_merge_signature


def build_row_profiles(signals: SheetSignals) -> list[dict[str, Any]]:
    if signals.bounds is None:
        return []
    box = signals.bounds
    rows: list[dict[str, Any]] = []
    for row in range(box.min_row, box.max_row + 1):
        coords = [(row, col) for col in range(box.min_col, box.max_col + 1)]
        sigs = [signals.cells.get(coord) for coord in coords]
        sigs = [s for s in sigs if s is not None]
        non_empty = [s for s in sigs if s.has_value]
        occupied = [coord for coord in coords if coord in signals.occupied]
        numeric = [s for s in non_empty if s.is_numeric]
        formulas = [s for s in non_empty if s.is_formula]
        bold = [s for s in sigs if s.bold]
        bordered = [s for s in sigs if s.border_score > 0]
        signature = row_merge_signature(signals, row, box.min_col, box.max_col)
        text_preview = []
        for s in non_empty[:6]:
            text_preview.append(str(s.value)[:80])
        rows.append({
            "row": row,
            "non_empty": len(non_empty),
            "occupied": len(occupied),
            "numeric": len(numeric),
            "formula": len(formulas),
            "bold": len(bold),
            "bordered": len(bordered),
            "merge_signature": signature,
            "text_preview": " | ".join(text_preview),
        })
    return rows


def build_col_profiles(signals: SheetSignals) -> list[dict[str, Any]]:
    if signals.bounds is None:
        return []
    box = signals.bounds
    cols: list[dict[str, Any]] = []
    for col in range(box.min_col, box.max_col + 1):
        coords = [(row, col) for row in range(box.min_row, box.max_row + 1)]
        sigs = [signals.cells.get(coord) for coord in coords]
        sigs = [s for s in sigs if s is not None]
        non_empty = [s for s in sigs if s.has_value]
        occupied = [coord for coord in coords if coord in signals.occupied]
        numeric = [s for s in non_empty if s.is_numeric]
        formulas = [s for s in non_empty if s.is_formula]
        bold = [s for s in sigs if s.bold]
        bordered = [s for s in sigs if s.border_score > 0]
        cols.append({
            "col": col,
            "non_empty": len(non_empty),
            "occupied": len(occupied),
            "numeric": len(numeric),
            "formula": len(formulas),
            "bold": len(bold),
            "bordered": len(bordered),
        })
    return cols


def blank_runs(values: list[dict[str, Any]], key: str, index_key: str, min_gap: int) -> list[dict[str, int]]:
    runs: list[dict[str, int]] = []
    start: int | None = None
    prev: int | None = None
    for item in values:
        idx = int(item[index_key])
        blank = int(item.get(key, 0)) == 0
        if blank and start is None:
            start = idx
        if not blank and start is not None:
            end = prev if prev is not None else idx - 1
            if end - start + 1 >= min_gap:
                runs.append({"start": start, "end": end, "length": end - start + 1})
            start = None
        prev = idx
    if start is not None and prev is not None:
        if prev - start + 1 >= min_gap:
            runs.append({"start": start, "end": prev, "length": prev - start + 1})
    return runs


def projection_profile(signals: SheetSignals, config: dict[str, Any]) -> dict[str, Any]:
    row_profiles = build_row_profiles(signals)
    col_profiles = build_col_profiles(signals)
    return {
        "sheet_name": signals.sheet_name,
        "bounds": signals.bounds.to_dict() if signals.bounds else None,
        "row_profiles": row_profiles,
        "col_profiles": col_profiles,
        "blank_row_runs": blank_runs(row_profiles, "occupied", "row", int(config.get("min_blank_gap_rows", 2))),
        "blank_col_runs": blank_runs(col_profiles, "occupied", "col", int(config.get("min_blank_gap_cols", 1))),
    }
