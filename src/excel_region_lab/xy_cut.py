from __future__ import annotations

from typing import Any

from .schema import Box, Region
from .signals import SheetSignals, region_features


def occupied_count_in_row(signals: SheetSignals, row: int, box: Box) -> int:
    return sum(1 for col in range(box.min_col, box.max_col + 1) if (row, col) in signals.occupied)


def occupied_count_in_col(signals: SheetSignals, col: int, box: Box) -> int:
    return sum(1 for row in range(box.min_row, box.max_row + 1) if (row, col) in signals.occupied)


def blank_runs_in_box(signals: SheetSignals, box: Box, axis: str, min_gap: int) -> list[tuple[int, int, int]]:
    runs: list[tuple[int, int, int]] = []
    start: int | None = None
    last: int | None = None
    if axis == "row":
        iterator = range(box.min_row, box.max_row + 1)
        is_blank = lambda idx: occupied_count_in_row(signals, idx, box) == 0
    else:
        iterator = range(box.min_col, box.max_col + 1)
        is_blank = lambda idx: occupied_count_in_col(signals, idx, box) == 0
    for idx in iterator:
        if is_blank(idx):
            if start is None:
                start = idx
            last = idx
        else:
            if start is not None and last is not None and last - start + 1 >= min_gap:
                runs.append((start, last, last - start + 1))
            start = None
            last = None
    if start is not None and last is not None and last - start + 1 >= min_gap:
        runs.append((start, last, last - start + 1))
    return runs


def choose_cut(signals: SheetSignals, box: Box, config: dict[str, Any]) -> tuple[str, int, int] | None:
    row_runs = blank_runs_in_box(signals, box, "row", int(config.get("min_blank_gap_rows", 2)))
    col_runs = blank_runs_in_box(signals, box, "col", int(config.get("min_blank_gap_cols", 1)))
    candidates: list[tuple[str, int, int, int]] = []
    for start, end, length in row_runs:
        candidates.append(("row", start, end, length * box.width))
    for start, end, length in col_runs:
        candidates.append(("col", start, end, length * box.height))
    if not candidates:
        return None
    axis, start, end, _ = max(candidates, key=lambda item: item[3])
    return axis, start, end


def split_box(box: Box, cut: tuple[str, int, int]) -> list[Box]:
    axis, start, end = cut
    result: list[Box] = []
    if axis == "row":
        if start > box.min_row:
            result.append(Box(box.min_row, box.min_col, start - 1, box.max_col))
        if end < box.max_row:
            result.append(Box(end + 1, box.min_col, box.max_row, box.max_col))
    else:
        if start > box.min_col:
            result.append(Box(box.min_row, box.min_col, box.max_row, start - 1))
        if end < box.max_col:
            result.append(Box(box.min_row, end + 1, box.max_row, box.max_col))
    return [b for b in result if b.area > 0]


def has_enough_occupied(signals: SheetSignals, box: Box, min_region_cells: int) -> bool:
    return sum(1 for coord in signals.occupied if box.contains(*coord)) >= min_region_cells


def xy_cut(signals: SheetSignals, config: dict[str, Any]) -> list[Region]:
    if signals.bounds is None:
        return []
    max_depth = int(config.get("max_depth", 8))
    min_region_cells = int(config.get("min_region_cells", 6))
    leaves: list[Box] = []

    def recurse(box: Box, depth: int) -> None:
        if depth >= max_depth or not has_enough_occupied(signals, box, min_region_cells):
            if has_enough_occupied(signals, box, min_region_cells):
                leaves.append(box)
            return
        cut = choose_cut(signals, box, config)
        if cut is None:
            leaves.append(box)
            return
        children = [child for child in split_box(box, cut) if has_enough_occupied(signals, child, min_region_cells)]
        if len(children) <= 1:
            leaves.append(box)
            return
        for child in children:
            recurse(child, depth + 1)

    recurse(signals.bounds, 0)
    regions: list[Region] = []
    for idx, box in enumerate(sorted(leaves, key=lambda b: (b.min_row, b.min_col)), 1):
        regions.append(Region(
            id=f"R{idx:03d}",
            sheet_name=signals.sheet_name,
            box=box,
            algorithm="xy_cut",
            features=region_features(signals, box),
            members=[],
        ))
    return regions
