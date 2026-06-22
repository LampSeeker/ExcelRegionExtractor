from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .schema import Box


def is_non_empty(value: Any) -> bool:
    return value is not None and value != ""


def merged_boxes_with_values(ws: Worksheet) -> list[Box]:
    boxes: list[Box] = []
    for rng in ws.merged_cells.ranges:
        top_left = ws.cell(rng.min_row, rng.min_col)
        if is_non_empty(top_left.value):
            boxes.append(Box(rng.min_row, rng.min_col, rng.max_row, rng.max_col))
    return boxes


def collect_cell_occupied(ws: Worksheet, bounds: Box | None, config: dict[str, Any]) -> set[tuple[int, int]]:
    if bounds is None:
        return set()
    occupied: set[tuple[int, int]] = set()

    if config.get("include_values", True):
        for (row, col), cell in ws._cells.items():
            if bounds.contains(row, col) and is_non_empty(cell.value):
                occupied.add((row, col))

    if config.get("include_merged_cells", True):
        for box in merged_boxes_with_values(ws):
            for row in range(max(bounds.min_row, box.min_row), min(bounds.max_row, box.max_row) + 1):
                for col in range(max(bounds.min_col, box.min_col), min(bounds.max_col, box.max_col) + 1):
                    occupied.add((row, col))

    return occupied
