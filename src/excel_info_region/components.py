from __future__ import annotations

from collections import deque
from typing import Any

from .schema import Box


def intersection_area(a: Box, b: Box) -> int:
    min_row = max(a.min_row, b.min_row)
    max_row = min(a.max_row, b.max_row)
    min_col = max(a.min_col, b.min_col)
    max_col = min(a.max_col, b.max_col)
    if min_row > max_row or min_col > max_col:
        return 0
    return (max_row - min_row + 1) * (max_col - min_col + 1)


def overlap_len_1d(a_min: int, a_max: int, b_min: int, b_max: int) -> int:
    return max(0, min(a_max, b_max) - max(a_min, b_min) + 1)


def overlap_ratio_on_axis(a: Box, b: Box, *, axis: str) -> float:
    if axis == "col":
        overlap = overlap_len_1d(a.min_col, a.max_col, b.min_col, b.max_col)
        denom = max(1, min(a.width, b.width))
        return overlap / denom
    overlap = overlap_len_1d(a.min_row, a.max_row, b.min_row, b.max_row)
    denom = max(1, min(a.height, b.height))
    return overlap / denom


def union_find_groups(indices: list[int], pairs: list[tuple[int, int]]) -> list[list[int]]:
    parent = {i: i for i in indices}

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for a, b in pairs:
        union(a, b)

    groups: dict[int, list[int]] = {}
    for i in indices:
        groups.setdefault(find(i), []).append(i)
    return list(groups.values())


def neighbor_offsets(connectivity: int) -> list[tuple[int, int]]:
    if connectivity == 4:
        return [(-1, 0), (1, 0), (0, -1), (0, 1)]
    return [
        (-1, -1), (-1, 0), (-1, 1),
        (0, -1),           (0, 1),
        (1, -1),  (1, 0),  (1, 1),
    ]


def connected_components_from_cells(
    occupied: set[tuple[int, int]],
    *,
    connectivity: int = 8,
    min_occupied_cells: int = 1,
) -> list[Box]:
    visited: set[tuple[int, int]] = set()
    offsets = neighbor_offsets(connectivity)
    boxes: list[Box] = []

    for start in sorted(occupied):
        if start in visited:
            continue
        q = deque([start])
        visited.add(start)
        coords: list[tuple[int, int]] = []

        while q:
            row, col = q.popleft()
            coords.append((row, col))
            for dr, dc in offsets:
                nxt = (row + dr, col + dc)
                if nxt in occupied and nxt not in visited:
                    visited.add(nxt)
                    q.append(nxt)

        if len(coords) < min_occupied_cells:
            continue
        boxes.append(Box(
            min(r for r, _ in coords),
            min(c for _, c in coords),
            max(r for r, _ in coords),
            max(c for _, c in coords),
        ))

    return boxes


def dedupe_boxes(boxes: list[Box]) -> list[Box]:
    seen: set[tuple[int, int, int, int]] = set()
    result: list[Box] = []
    for box in boxes:
        key = (box.min_row, box.min_col, box.max_row, box.max_col)
        if key in seen:
            continue
        seen.add(key)
        result.append(box)
    return result


def clip_box_to_bounds(box: Box, bounds: Box) -> Box | None:
    min_row = max(box.min_row, bounds.min_row)
    min_col = max(box.min_col, bounds.min_col)
    max_row = min(box.max_row, bounds.max_row)
    max_col = min(box.max_col, bounds.max_col)
    if min_row > max_row or min_col > max_col:
        return None
    return Box(min_row, min_col, max_row, max_col)


def remove_exact_or_contained_duplicates(boxes: list[Box], config: dict[str, Any]) -> list[Box]:
    boxes = dedupe_boxes(boxes)
    if not config.get("remove_contained_duplicates", False):
        return boxes

    result: list[Box] = []
    for i, box in enumerate(boxes):
        contained = False
        for j, other in enumerate(boxes):
            if i == j:
                continue
            if other.contains_box(box) and other.area > box.area:
                contained = True
                break
        if not contained:
            result.append(box)
    return result
