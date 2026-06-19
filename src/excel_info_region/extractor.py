from __future__ import annotations

from collections import deque
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .raw_drawing import drawing_image_boxes
from .schema import Box, InfoRegion


def open_workbook(path: str | Path, *, data_only: bool = False):
    return load_workbook(path, read_only=False, data_only=data_only)


def iter_target_sheets(workbook, sheet_name: str | None = None):
    if sheet_name:
        yield workbook[sheet_name]
    else:
        for ws in workbook.worksheets:
            yield ws


def is_non_empty(value: Any) -> bool:
    return value is not None and value != ""


def merged_boxes_with_values(ws: Worksheet) -> list[Box]:
    boxes: list[Box] = []
    for rng in ws.merged_cells.ranges:
        top_left = ws.cell(rng.min_row, rng.min_col)
        if is_non_empty(top_left.value):
            boxes.append(Box(rng.min_row, rng.min_col, rng.max_row, rng.max_col))
    return boxes


def image_boxes(ws: Worksheet, workbook_path: str | Path | None, config: dict[str, Any]) -> list[Box]:
    if not config.get("include_images", True):
        return []

    raw_boxes: list[Box] = []
    if workbook_path and config.get("include_grouped_drawing_images", True):
        try:
            raw_boxes = drawing_image_boxes(workbook_path, ws.title, ws)
        except Exception:
            raw_boxes = []

    # raw DrawingML parser is preferred because openpyxl may expose grouped images as one object.
    if raw_boxes:
        return dedupe_boxes(raw_boxes)

    # Fallback: approximate from openpyxl image anchors.
    boxes: list[Box] = []
    for img in getattr(ws, "_images", []):
        anchor = getattr(img, "anchor", None)
        if anchor is None or not hasattr(anchor, "_from"):
            continue
        start = anchor._from
        min_row = int(start.row) + 1
        min_col = int(start.col) + 1
        end_marker = None
        if hasattr(anchor, "_to") and getattr(anchor, "_to", None) is not None:
            end_marker = anchor._to
        elif hasattr(anchor, "to") and getattr(anchor, "to", None) is not None:
            end_marker = anchor.to

        if end_marker is not None:
            max_row = max(min_row, int(end_marker.row) + 1)
            max_col = max(min_col, int(end_marker.col) + 1)
        else:
            # Approximate one-cell anchor image size in cells.
            width_cells = max(1, int((getattr(img, "width", 64) or 64) / 64))
            height_cells = max(1, int((getattr(img, "height", 20) or 20) / 20))
            max_row = min_row + height_cells - 1
            max_col = min_col + width_cells - 1
        boxes.append(Box(min_row, min_col, max_row, max_col))
    return dedupe_boxes(boxes)


def effective_bounds(ws: Worksheet, workbook_path: str | Path | None, config: dict[str, Any]) -> Box | None:
    rows: list[int] = []
    cols: list[int] = []

    for (row, col), cell in ws._cells.items():
        if is_non_empty(cell.value):
            rows.append(row)
            cols.append(col)

    for box in merged_boxes_with_values(ws):
        rows.extend([box.min_row, box.max_row])
        cols.extend([box.min_col, box.max_col])

    for box in image_boxes(ws, workbook_path, config):
        rows.extend([box.min_row, box.max_row])
        cols.extend([box.min_col, box.max_col])

    if not rows or not cols:
        return None

    pad_r = int(config.get("bounds_padding_rows", 0))
    pad_c = int(config.get("bounds_padding_cols", 0))
    return Box(
        max(1, min(rows) - pad_r),
        max(1, min(cols) - pad_c),
        min(1048576, max(rows) + pad_r),
        min(16384, max(cols) + pad_c),
    )


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




def _side_has_style(side) -> bool:
    return side is not None and getattr(side, "style", None) is not None


def cell_has_border(cell, *, strong_only: bool = False) -> bool:
    border = cell.border
    sides = [border.left, border.right, border.top, border.bottom]

    if not strong_only:
        return any(_side_has_style(side) for side in sides)

    strong_styles = {
        "medium",
        "thick",
        "double",
        "mediumDashed",
        "mediumDashDot",
        "mediumDashDotDot",
        "slantDashDot",
    }
    return any(getattr(side, "style", None) in strong_styles for side in sides)


def collect_border_occupied(ws: Worksheet, bounds: Box | None, config: dict[str, Any]) -> set[tuple[int, int]]:
    if bounds is None or not config.get("use_borders", True):
        return set()

    strong_only = bool(config.get("strong_borders_only", False))
    occupied: set[tuple[int, int]] = set()

    for (row, col), cell in ws._cells.items():
        if not bounds.contains(row, col):
            continue
        if cell_has_border(cell, strong_only=strong_only):
            occupied.add((row, col))

    # Merged cells often store border information only on edge cells.
    # If any edge cell of a merged range has a border, mark the whole merged range
    # as part of the border shell so its bbox can expand correctly.
    if config.get("include_merged_cells", True):
        for rng in ws.merged_cells.ranges:
            box = Box(rng.min_row, rng.min_col, rng.max_row, rng.max_col)
            if not box.intersects(bounds):
                continue

            edge_cells = []
            for col in range(rng.min_col, rng.max_col + 1):
                edge_cells.append(ws.cell(rng.min_row, col))
                edge_cells.append(ws.cell(rng.max_row, col))
            for row in range(rng.min_row, rng.max_row + 1):
                edge_cells.append(ws.cell(row, rng.min_col))
                edge_cells.append(ws.cell(row, rng.max_col))

            if any(cell_has_border(cell, strong_only=strong_only) for cell in edge_cells):
                for row in range(max(bounds.min_row, rng.min_row), min(bounds.max_row, rng.max_row) + 1):
                    for col in range(max(bounds.min_col, rng.min_col), min(bounds.max_col, rng.max_col) + 1):
                        occupied.add((row, col))

    return occupied


def intersection_area(a: Box, b: Box) -> int:
    min_row = max(a.min_row, b.min_row)
    max_row = min(a.max_row, b.max_row)
    min_col = max(a.min_col, b.min_col)
    max_col = min(a.max_col, b.max_col)
    if min_row > max_row or min_col > max_col:
        return 0
    return (max_row - min_row + 1) * (max_col - min_col + 1)


def box_gap(a: Box, b: Box) -> int:
    row_gap = 0
    if a.max_row < b.min_row:
        row_gap = b.min_row - a.max_row - 1
    elif b.max_row < a.min_row:
        row_gap = a.min_row - b.max_row - 1

    col_gap = 0
    if a.max_col < b.min_col:
        col_gap = b.min_col - a.max_col - 1
    elif b.max_col < a.min_col:
        col_gap = a.min_col - b.max_col - 1

    return max(row_gap, col_gap)


def should_expand_to_border_shell(value_box: Box, border_box: Box, config: dict[str, Any]) -> bool:
    inter = intersection_area(value_box, border_box)
    if inter <= 0:
        return False

    value_overlap = inter / max(1, value_box.area)
    border_overlap = inter / max(1, border_box.area)

    if value_overlap < float(config.get("border_expand_min_value_overlap", 0.80)):
        return False

    # Border expansion is bbox correction, not section grouping.
    # Reject huge shells that would swallow titles, diagrams, and nearby tables.
    max_area_ratio = float(config.get("border_expand_max_area_ratio", 3.0))
    if border_box.area > value_box.area * max_area_ratio:
        return False

    max_extra_rows = int(config.get("border_expand_max_extra_rows", 3))
    max_extra_cols = int(config.get("border_expand_max_extra_cols", 3))
    extra_top = max(0, value_box.min_row - border_box.min_row)
    extra_bottom = max(0, border_box.max_row - value_box.max_row)
    extra_left = max(0, value_box.min_col - border_box.min_col)
    extra_right = max(0, border_box.max_col - value_box.max_col)

    if max(extra_top, extra_bottom) > max_extra_rows:
        return False
    if max(extra_left, extra_right) > max_extra_cols:
        return False

    # Need some meaningful shared footprint, otherwise a large outlined region touching
    # a tiny value region can still over-expand.
    if border_overlap < float(config.get("border_expand_min_border_overlap", 0.10)):
        return False

    return True


def expand_cell_boxes_with_borders(
    cell_boxes: list[Box],
    border_boxes: list[Box],
    config: dict[str, Any],
) -> list[Box]:
    if not config.get("use_borders", True) or not border_boxes:
        return cell_boxes

    expanded: list[Box] = []
    for value_box in cell_boxes:
        current = value_box
        for border_box in border_boxes:
            if should_expand_to_border_shell(current, border_box, config):
                current = current.union(border_box)
        expanded.append(current)

    if config.get("add_border_only_regions", False):
        for border_box in border_boxes:
            has_value = any(
                border_box.contains_box(value_box) or intersection_area(border_box, value_box) > 0
                for value_box in cell_boxes
            )
            if not has_value:
                expanded.append(border_box)

    return dedupe_boxes(expanded)


# ---------------------------------------------------------------------------
# Final correction: border-contact merge
#
# This is not semantic grouping.
# It merges value regions only when they touch the same real border edge
# component and are spatially adjacent enough to be considered one table shell.
# Images are excluded before this step.
# ---------------------------------------------------------------------------

BorderEdge = tuple[str, int, int]  # ("h", row_line, col) or ("v", row, col_line)


def collect_border_edges(ws: Worksheet, bounds: Box | None, config: dict[str, Any]) -> set[BorderEdge]:
    if bounds is None or not config.get("use_border_contact_merge", False):
        return set()

    strong_only = bool(config.get("border_contact_strong_only", False))
    edges: set[BorderEdge] = set()

    def add_cell_edges(row: int, col: int, cell) -> None:
        b = cell.border
        if _side_has_style(b.top) and (not strong_only or cell_has_border_side(b.top, strong_only=True)):
            edges.add(("h", row, col))
        if _side_has_style(b.bottom) and (not strong_only or cell_has_border_side(b.bottom, strong_only=True)):
            edges.add(("h", row + 1, col))
        if _side_has_style(b.left) and (not strong_only or cell_has_border_side(b.left, strong_only=True)):
            edges.add(("v", row, col))
        if _side_has_style(b.right) and (not strong_only or cell_has_border_side(b.right, strong_only=True)):
            edges.add(("v", row, col + 1))

    for (row, col), cell in ws._cells.items():
        if bounds.contains(row, col):
            add_cell_edges(row, col, cell)

    # Merged ranges can store visible borders only on the edge cells.
    # Add the merged perimeter explicitly when edge cells carry borders.
    if config.get("include_merged_cells", True):
        for rng in ws.merged_cells.ranges:
            box = Box(rng.min_row, rng.min_col, rng.max_row, rng.max_col)
            if not box.intersects(bounds):
                continue

            # top / bottom
            for col in range(rng.min_col, rng.max_col + 1):
                top_cell = ws.cell(rng.min_row, col)
                bottom_cell = ws.cell(rng.max_row, col)
                if _side_has_style(top_cell.border.top) and (not strong_only or cell_has_border_side(top_cell.border.top, strong_only=True)):
                    edges.add(("h", rng.min_row, col))
                if _side_has_style(bottom_cell.border.bottom) and (not strong_only or cell_has_border_side(bottom_cell.border.bottom, strong_only=True)):
                    edges.add(("h", rng.max_row + 1, col))

            # left / right
            for row in range(rng.min_row, rng.max_row + 1):
                left_cell = ws.cell(row, rng.min_col)
                right_cell = ws.cell(row, rng.max_col)
                if _side_has_style(left_cell.border.left) and (not strong_only or cell_has_border_side(left_cell.border.left, strong_only=True)):
                    edges.add(("v", row, rng.min_col))
                if _side_has_style(right_cell.border.right) and (not strong_only or cell_has_border_side(right_cell.border.right, strong_only=True)):
                    edges.add(("v", row, rng.max_col + 1))

    return edges


def cell_has_border_side(side, *, strong_only: bool = False) -> bool:
    if not _side_has_style(side):
        return False
    if not strong_only:
        return True
    return getattr(side, "style", None) in {
        "medium",
        "thick",
        "double",
        "mediumDashed",
        "mediumDashDot",
        "mediumDashDotDot",
        "slantDashDot",
    }


def edge_endpoints(edge: BorderEdge) -> tuple[tuple[int, int], tuple[int, int]]:
    kind, a, b = edge
    if kind == "h":
        row_line = a
        col = b
        return (row_line, col), (row_line, col + 1)
    row = a
    col_line = b
    return (row, col_line), (row + 1, col_line)


def border_edge_components(edges: set[BorderEdge]) -> dict[BorderEdge, int]:
    endpoint_to_edges: dict[tuple[int, int], list[BorderEdge]] = {}
    for edge in edges:
        p1, p2 = edge_endpoints(edge)
        endpoint_to_edges.setdefault(p1, []).append(edge)
        endpoint_to_edges.setdefault(p2, []).append(edge)

    edge_to_component: dict[BorderEdge, int] = {}
    visited: set[BorderEdge] = set()
    component_id = 0

    for start in sorted(edges):
        if start in visited:
            continue
        component_id += 1
        q = deque([start])
        visited.add(start)
        edge_to_component[start] = component_id

        while q:
            edge = q.popleft()
            for point in edge_endpoints(edge):
                for nxt in endpoint_to_edges.get(point, []):
                    if nxt not in visited:
                        visited.add(nxt)
                        edge_to_component[nxt] = component_id
                        q.append(nxt)

    return edge_to_component


def perimeter_edges(box: Box, *, tolerance: int = 0) -> set[BorderEdge]:
    edges: set[BorderEdge] = set()

    min_row = max(1, box.min_row - tolerance)
    min_col = max(1, box.min_col - tolerance)
    max_row = box.max_row + tolerance
    max_col = box.max_col + tolerance

    for col in range(min_col, max_col + 1):
        edges.add(("h", min_row, col))
        edges.add(("h", max_row + 1, col))
    for row in range(min_row, max_row + 1):
        edges.add(("v", row, min_col))
        edges.add(("v", row, max_col + 1))

    return edges



def perimeter_edges_by_side(box: Box, *, tolerance: int = 0) -> dict[str, set[BorderEdge]]:
    min_row = max(1, box.min_row - tolerance)
    min_col = max(1, box.min_col - tolerance)
    max_row = box.max_row + tolerance
    max_col = box.max_col + tolerance

    top = {("h", min_row, col) for col in range(min_col, max_col + 1)}
    bottom = {("h", max_row + 1, col) for col in range(min_col, max_col + 1)}
    left = {("v", row, min_col) for row in range(min_row, max_row + 1)}
    right = {("v", row, max_col + 1) for row in range(min_row, max_row + 1)}

    return {
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right,
    }


def touched_border_component_sides(
    box: Box,
    edge_to_component: dict[BorderEdge, int],
    config: dict[str, Any],
) -> dict[int, set[str]]:
    """Return which sides of the region touch each border component.

    This is stricter than checking whether the box is inside a border bbox.
    A region qualifies for contact merge only if it touches the same physical
    border edge component on enough sides, e.g. left+right or top+bottom.
    """
    tolerance = int(config.get("border_contact_tolerance_cells", 0))
    min_edges_per_side = int(
        config.get(
            "border_contact_min_edges_per_side",
            config.get("border_contact_min_edges", 1),
        )
    )

    component_side_counts: dict[int, dict[str, int]] = {}
    for side_name, edges in perimeter_edges_by_side(box, tolerance=tolerance).items():
        for edge in edges:
            component_id = edge_to_component.get(edge)
            if component_id is None:
                continue
            component_side_counts.setdefault(component_id, {})
            component_side_counts[component_id][side_name] = (
                component_side_counts[component_id].get(side_name, 0) + 1
            )

    result: dict[int, set[str]] = {}
    for component_id, side_counts in component_side_counts.items():
        sides = {
            side_name
            for side_name, count in side_counts.items()
            if count >= min_edges_per_side
        }
        if sides:
            result[component_id] = sides
    return result


def touched_side_count_for_component(
    box: Box,
    component_id: int,
    edge_to_component: dict[BorderEdge, int],
    config: dict[str, Any],
) -> int:
    return len(touched_border_component_sides(box, edge_to_component, config).get(component_id, set()))


def touched_border_components(
    box: Box,
    edge_to_component: dict[BorderEdge, int],
    config: dict[str, Any],
) -> set[int]:
    tolerance = int(config.get("border_contact_tolerance_cells", 0))
    min_edges = int(config.get("border_contact_min_edges", 1))

    counts: dict[int, int] = {}
    for edge in perimeter_edges(box, tolerance=tolerance):
        component_id = edge_to_component.get(edge)
        if component_id is not None:
            counts[component_id] = counts.get(component_id, 0) + 1

    return {component_id for component_id, count in counts.items() if count >= min_edges}


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


def boxes_are_contact_merge_neighbors(a: Box, b: Box, config: dict[str, Any]) -> bool:
    max_gap = int(config.get("border_contact_merge_max_gap", 1))
    min_axis_overlap = float(config.get("border_contact_merge_min_axis_overlap", 0.80))

    # vertical stack
    if a.max_row < b.min_row:
        row_gap = b.min_row - a.max_row - 1
        return row_gap <= max_gap and overlap_ratio_on_axis(a, b, axis="col") >= min_axis_overlap
    if b.max_row < a.min_row:
        row_gap = a.min_row - b.max_row - 1
        return row_gap <= max_gap and overlap_ratio_on_axis(a, b, axis="col") >= min_axis_overlap

    # horizontal stack
    if a.max_col < b.min_col:
        col_gap = b.min_col - a.max_col - 1
        return col_gap <= max_gap and overlap_ratio_on_axis(a, b, axis="row") >= min_axis_overlap
    if b.max_col < a.min_col:
        col_gap = a.min_col - b.max_col - 1
        return col_gap <= max_gap and overlap_ratio_on_axis(a, b, axis="row") >= min_axis_overlap

    # overlapping boxes
    return True


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


def merge_boxes_by_border_contact(
    cell_boxes: list[Box],
    ws: Worksheet,
    bounds: Box | None,
    config: dict[str, Any],
) -> list[Box]:
    if not config.get("use_border_contact_merge", False) or len(cell_boxes) < 2:
        return cell_boxes

    edges = collect_border_edges(ws, bounds, config)
    if not edges:
        return cell_boxes

    edge_to_component = border_edge_components(edges)
    min_touched_sides = int(config.get("border_contact_min_touched_sides", 2))
    component_to_indices: dict[int, list[int]] = {}

    for idx, box in enumerate(cell_boxes):
        touched_sides_by_component = touched_border_component_sides(box, edge_to_component, config)
        for component_id, sides in touched_sides_by_component.items():
            if len(sides) >= min_touched_sides:
                component_to_indices.setdefault(component_id, []).append(idx)

    candidate_pairs: set[tuple[int, int]] = set()
    for component_id, indices in component_to_indices.items():
        unique_indices = sorted(set(indices))
        if len(unique_indices) < 2:
            continue
        for pos, a_idx in enumerate(unique_indices):
            for b_idx in unique_indices[pos + 1:]:
                # Both regions have already satisfied min_touched_sides for this
                # same border component. Now check spatial adjacency.
                if boxes_are_contact_merge_neighbors(cell_boxes[a_idx], cell_boxes[b_idx], config):
                    candidate_pairs.add((a_idx, b_idx))

    if not candidate_pairs:
        return cell_boxes

    all_indices = list(range(len(cell_boxes)))
    merged_groups = union_find_groups(all_indices, sorted(candidate_pairs))

    output: list[Box] = []
    for group in merged_groups:
        if len(group) == 1:
            output.append(cell_boxes[group[0]])
            continue

        merged = cell_boxes[group[0]]
        for idx in group[1:]:
            merged = merged.union(cell_boxes[idx])

        # Safety: this stage bridges neighboring regions, not page-size sections.
        total_area = sum(cell_boxes[idx].area for idx in group)
        max_area_ratio = float(config.get("border_contact_merge_max_area_ratio", 2.5))
        if merged.area > total_area * max_area_ratio:
            output.extend(cell_boxes[idx] for idx in group)
        else:
            output.append(merged)

    return dedupe_boxes(output)


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


def extract_info_regions_from_sheet(
    ws: Worksheet,
    *,
    workbook_path: str | Path | None = None,
    config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    cfg = config or {}
    bounds = effective_bounds(ws, workbook_path, cfg)
    if bounds is None:
        return {
            "sheet_name": ws.title,
            "info_regions": [],
        }

    cell_occupied = collect_cell_occupied(ws, bounds, cfg)
    cell_boxes = connected_components_from_cells(
        cell_occupied,
        connectivity=int(cfg.get("connectivity", 8)),
        min_occupied_cells=int(cfg.get("min_occupied_cells", 1)),
    )

    border_occupied = collect_border_occupied(ws, bounds, cfg)
    border_boxes = connected_components_from_cells(
        border_occupied,
        connectivity=int(cfg.get("border_connectivity", cfg.get("connectivity", 8))),
        min_occupied_cells=int(cfg.get("min_border_cells", 2)),
    )
    cell_boxes = expand_cell_boxes_with_borders(cell_boxes, border_boxes, cfg)
    cell_boxes = merge_boxes_by_border_contact(cell_boxes, ws, bounds, cfg)

    # Important: images are not mixed into the cell connected-components.
    # This prevents image areas and adjacent tables from being merged into one region.
    img_boxes = image_boxes(ws, workbook_path, cfg)

    boxes = remove_exact_or_contained_duplicates(
        [*cell_boxes, *img_boxes],
        cfg,
    )

    boxes = sorted(boxes, key=lambda b: (b.min_row, b.min_col, b.max_row, b.max_col))
    regions = [box.range_ref for box in boxes]

    return {
        "sheet_name": ws.title,
        "info_regions": regions,
    }


def extract_workbook_info_regions(
    workbook_path: str | Path,
    *,
    sheet_name: str | None = None,
    config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    cfg = dict(config or {})
    cfg["workbook_path"] = str(workbook_path)

    wb = open_workbook(workbook_path, data_only=False)
    result = {
        "workbook": str(workbook_path),
        "sheets": {},
    }
    for ws in iter_target_sheets(wb, sheet_name):
        result["sheets"][ws.title] = extract_info_regions_from_sheet(
            ws,
            workbook_path=workbook_path,
            config=cfg,
        )
    return result


def summarize_workbook_result(result: dict[str, Any]) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    for sheet_name, data in result["sheets"].items():
        rows.append({
            "sheet_name": sheet_name,
            "info_region_count": len(data.get("info_regions", [])),
            "info_regions": data.get("info_regions", []),
        })
    return {
        "workbook": result["workbook"],
        "sheets": rows,
    }
