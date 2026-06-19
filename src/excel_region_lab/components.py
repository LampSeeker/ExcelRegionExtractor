from __future__ import annotations

from collections import deque
from typing import Any

from .schema import Box, Region
from .signals import SheetSignals, region_features


def neighbor_offsets(connectivity: int) -> list[tuple[int, int]]:
    if connectivity == 4:
        return [(-1, 0), (1, 0), (0, -1), (0, 1)]
    return [
        (-1, -1), (-1, 0), (-1, 1),
        (0, -1),           (0, 1),
        (1, -1),  (1, 0),  (1, 1),
    ]


def box_cells(box: Box) -> set[tuple[int, int]]:
    return {
        (row, col)
        for row in range(box.min_row, box.max_row + 1)
        for col in range(box.min_col, box.max_col + 1)
    }


def image_cells(signals: SheetSignals) -> set[tuple[int, int]]:
    cells: set[tuple[int, int]] = set()
    for box in signals.image_ranges:
        cells.update(box_cells(box))
    return cells


def boxes_touch_or_overlap(a: Box, b: Box, row_gap: int = 0, col_gap: int = 0) -> bool:
    row_touch = not (a.max_row + row_gap + 1 < b.min_row or b.max_row + row_gap + 1 < a.min_row)
    col_touch = not (a.max_col + col_gap + 1 < b.min_col or b.max_col + col_gap + 1 < a.min_col)
    return row_touch and col_touch


def group_image_boxes(boxes: list[Box], row_gap: int = 0, col_gap: int = 0) -> list[Box]:
    if not boxes:
        return []
    remaining = list(boxes)
    grouped: list[Box] = []
    while remaining:
        current = remaining.pop(0)
        changed = True
        while changed:
            changed = False
            rest: list[Box] = []
            for other in remaining:
                if boxes_touch_or_overlap(current, other, row_gap=row_gap, col_gap=col_gap):
                    current = current.union(other)
                    changed = True
                else:
                    rest.append(other)
            remaining = rest
        grouped.append(current)
    return sorted(grouped, key=lambda b: (b.min_row, b.min_col))


def regions_from_occupied(
    signals: SheetSignals,
    occupied: set[tuple[int, int]],
    config: dict[str, Any],
    algorithm: str,
    id_prefix: str = "B",
) -> list[Region]:
    if signals.bounds is None:
        return []

    visited: set[tuple[int, int]] = set()
    offsets = neighbor_offsets(int(config.get("connectivity", 8)))
    regions: list[Region] = []
    min_non_empty = int(config.get("min_non_empty_cells", 1))

    for start in sorted(occupied):
        if start in visited:
            continue
        queue = deque([start])
        visited.add(start)
        coords: list[tuple[int, int]] = []
        while queue:
            row, col = queue.popleft()
            coords.append((row, col))
            for dr, dc in offsets:
                nxt = (row + dr, col + dc)
                if nxt in occupied and nxt not in visited:
                    visited.add(nxt)
                    queue.append(nxt)

        box = Box(
            min(r for r, _ in coords),
            min(c for _, c in coords),
            max(r for r, _ in coords),
            max(c for _, c in coords),
        )
        features = region_features(signals, box)
        if (
            features.get("non_empty_count", 0) < min_non_empty
            and features.get("image_count", 0) == 0
            and features.get("occupied_count", 0) < min_non_empty
        ):
            continue
        region_id = f"{id_prefix}{len(regions) + 1:03d}"
        regions.append(
            Region(
                id=region_id,
                sheet_name=signals.sheet_name,
                box=box,
                algorithm=algorithm,
                features=features,
                members=[region_id],
            )
        )
    return regions


def connected_components(signals: SheetSignals, config: dict[str, Any], algorithm: str = "connected_components") -> list[Region]:
    """Build primitive regions.

    Default behavior keeps inserted images separate from cell-value/table components.
    This avoids a common failure mode where a large figure rectangle touches a table below
    and the two become one huge connected component.
    """
    if signals.bounds is None:
        return []

    mode = str(config.get("image_mode", "separate"))
    occupied = set(signals.occupied)

    if mode == "connect":
        return regions_from_occupied(signals, occupied, config, algorithm, "B")

    img_cells = image_cells(signals)
    cell_occupied = occupied - img_cells
    cell_regions = regions_from_occupied(signals, cell_occupied, config, algorithm, "B")

    image_regions: list[Region] = []
    grouped_images = group_image_boxes(
        list(signals.image_ranges),
        row_gap=int(config.get("image_group_row_gap", 0)),
        col_gap=int(config.get("image_group_col_gap", 0)),
    )
    for idx, box in enumerate(grouped_images, 1):
        features = region_features(signals, box)
        features = {**features, "role_hint": "figure_or_diagram", "source": "image"}
        image_regions.append(
            Region(
                id=f"I{idx:03d}",
                sheet_name=signals.sheet_name,
                box=box,
                algorithm=algorithm,
                features=features,
                members=[f"I{idx:03d}"],
            )
        )

    regions = sorted(cell_regions + image_regions, key=lambda r: (r.box.min_row, r.box.min_col, r.box.max_row, r.box.max_col))
    for idx, region in enumerate(regions, 1):
        region.id = f"B{idx:03d}"
        region.members = [region.id]
    return regions
