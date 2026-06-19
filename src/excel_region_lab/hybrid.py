from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .merge import graph_union_find_merge
from .schema import Box, Region, RegionEdge
from .signals import SheetSignals, region_features


@dataclass(frozen=True)
class SplitDecision:
    axis: str
    start: int
    end: int
    score: float


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


def occupied_count(signals: SheetSignals, box: Box) -> int:
    return sum(1 for coord in signals.occupied if box.contains(*coord))


def split_box(box: Box, axis: str, start: int, end: int) -> list[Box]:
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


def choose_projection_split(signals: SheetSignals, box: Box, config: dict[str, Any]) -> SplitDecision | None:
    min_blank_gap_rows = int(config.get("min_blank_gap_rows", 2))
    min_blank_gap_cols = int(config.get("min_blank_gap_cols", 1))
    row_runs = blank_runs_in_box(signals, box, "row", min_blank_gap_rows)
    col_runs = blank_runs_in_box(signals, box, "col", min_blank_gap_cols)
    min_child_occupied = int(config.get("min_child_occupied", 3))

    candidates: list[SplitDecision] = []
    for axis, runs in (("row", row_runs), ("col", col_runs)):
        for start, end, length in runs:
            children = split_box(box, axis, start, end)
            if len(children) < 2:
                continue
            child_counts = [occupied_count(signals, child) for child in children]
            if min(child_counts) < min_child_occupied:
                continue
            blank_area = length * (box.width if axis == "row" else box.height)
            score = blank_area / max(box.area, 1)
            candidates.append(SplitDecision(axis, start, end, score))

    if not candidates:
        return None
    return max(candidates, key=lambda item: item.score)


def recursively_split_root(signals: SheetSignals, box: Box, config: dict[str, Any], depth: int = 0) -> list[Box]:
    max_depth = int(config.get("max_depth", 4))
    min_region_occupied = int(config.get("min_region_occupied", 6))
    if depth >= max_depth or occupied_count(signals, box) < min_region_occupied:
        return [box]

    decision = choose_projection_split(signals, box, config)
    if decision is None:
        return [box]

    children = split_box(box, decision.axis, decision.start, decision.end)
    valid_children = [child for child in children if occupied_count(signals, child) >= min_region_occupied]
    if len(valid_children) < 2:
        return [box]

    leaves: list[Box] = []
    for child in valid_children:
        leaves.extend(recursively_split_root(signals, child, config, depth + 1))
    return leaves


def hybrid_grouping(
    signals: SheetSignals,
    primitive_regions: list[Region],
    pair_score_config: dict[str, Any],
    graph_config: dict[str, Any],
    hybrid_config: dict[str, Any],
) -> tuple[list[Region], list[Region], list[RegionEdge]]:
    """
    Hybrid pipeline:
    1) Connected components create primitive regions.
    2) Graph/Union-Find merges primitives into coarse root regions.
    3) Projection-profile splitting optionally subdivides each root into candidate regions.
    """
    roots, edges = graph_union_find_merge(
        signals,
        primitive_regions,
        pair_score_config,
        float(graph_config.get("merge_threshold", 0.62)),
        algorithm="hybrid_root",
    )

    candidate_regions: list[Region] = []
    split_cfg = hybrid_config.get("split", {})
    next_id = 1
    for root_idx, root in enumerate(roots, 1):
        leaves = recursively_split_root(signals, root.box, split_cfg)
        if not leaves:
            leaves = [root.box]
        for leaf in sorted(leaves, key=lambda b: (b.min_row, b.min_col)):
            candidate_regions.append(
                Region(
                    id=f"C{next_id:03d}",
                    sheet_name=signals.sheet_name,
                    box=leaf,
                    algorithm="hybrid_candidate",
                    features=region_features(signals, leaf),
                    members=root.members,
                    score=root.score,
                )
            )
            next_id += 1

    normalized_roots: list[Region] = []
    for idx, root in enumerate(sorted(roots, key=lambda r: (r.box.min_row, r.box.min_col)), 1):
        normalized_roots.append(
            Region(
                id=f"R{idx:03d}",
                sheet_name=root.sheet_name,
                box=root.box,
                algorithm="hybrid_root",
                features=root.features,
                members=root.members,
                score=root.score,
            )
        )

    return normalized_roots, candidate_regions, edges
