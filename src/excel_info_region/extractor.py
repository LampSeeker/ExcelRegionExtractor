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
