from __future__ import annotations

from pathlib import Path
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet
from typing import Any

from .config import load_config
from .borders import collect_border_edges, edge_endpoints
from .cells import merged_boxes_with_values
from .components import connected_components_from_cells, dedupe_boxes
from .extractor import extract_workbook_info_regions, summarize_workbook_result, open_workbook
from .chart_export import extract_sheet_charts_to_dir
from .io import ensure_dir, safe_name, write_json
from .image_export import extract_sheet_images_to_dir
from .schema import Box
from .visualize import render_region_overlay


def _overlay_regions_from_ranges(ranges: list[str], sheet_name: str) -> list[dict]:
    regions: list[dict] = []
    for idx, range_ref in enumerate(ranges, 1):
        regions.append(_overlay_region(range_ref, sheet_name, f"R{idx:03d}"))
    return regions


def _overlay_region(range_ref: str, sheet_name: str, region_id: str) -> dict:
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    return {
        "id": region_id,
        "sheet_name": sheet_name,
        "range_ref": range_ref,
        "min_row": min_row,
        "min_col": min_col,
        "max_row": max_row,
        "max_col": max_col,
        "height": max_row - min_row + 1,
        "width": max_col - min_col + 1,
        "area": (max_row - min_row + 1) * (max_col - min_col + 1),
    }


def _overlay_regions_from_tree(region_tree: list[dict[str, Any]], sheet_name: str) -> list[dict]:
    regions: list[dict] = []
    for idx, root in enumerate(region_tree, 1):
        root_id = f"R{idx:03d}"
        regions.append(_overlay_region(root["range_ref"], sheet_name, root_id))
        for child_idx, child in enumerate(root.get("children", []), 1):
            regions.append(_overlay_region(child["range_ref"], sheet_name, f"{root_id}.C{child_idx:02d}"))
    return regions


def _box_from_range(range_ref: str) -> Box:
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    return Box(min_row, min_col, max_row, max_col)


def _root_ranges(ranges: list[str]) -> list[str]:
    boxes = [(range_ref, _box_from_range(range_ref)) for range_ref in ranges]
    return [
        range_ref
        for range_ref, box in boxes
        if not any(other_ref != range_ref and other_box.contains_box(box) for other_ref, other_box in boxes)
    ]


def _has_closed_border(ws: Worksheet, box: Box) -> bool:
    edges = collect_border_edges(ws, box, {"use_border_contact_merge": True, "include_merged_cells": True})
    if not edges:
        return False
    return all(("h", box.min_row, col) in edges and ("h", box.max_row + 1, col) in edges for col in range(box.min_col, box.max_col + 1)) and all(
        ("v", row, box.min_col) in edges and ("v", row, box.max_col + 1) in edges for row in range(box.min_row, box.max_row + 1)
    )


def _edge_component_boxes(edges: set[tuple[str, int, int]]) -> list[Box]:
    endpoint_to_edges: dict[tuple[int, int], list[tuple[str, int, int]]] = {}
    for edge in edges:
        for point in edge_endpoints(edge):
            endpoint_to_edges.setdefault(point, []).append(edge)

    boxes: list[Box] = []
    seen: set[tuple[str, int, int]] = set()
    for start in edges:
        if start in seen:
            continue
        stack = [start]
        seen.add(start)
        points: list[tuple[int, int]] = []
        while stack:
            edge = stack.pop()
            points.extend(edge_endpoints(edge))
            for point in edge_endpoints(edge):
                for nxt in endpoint_to_edges.get(point, []):
                    if nxt not in seen:
                        seen.add(nxt)
                        stack.append(nxt)
        if points:
            min_row = min(row for row, _col in points)
            min_col = min(col for _row, col in points)
            max_row = max(row for row, _col in points) - 1
            max_col = max(col for _row, col in points) - 1
            if max_row >= min_row and max_col >= min_col:
                boxes.append(Box(min_row, min_col, max_row, max_col))
    return dedupe_boxes(boxes)


def _inner_table_boxes(ws: Worksheet, root: Box) -> list[Box]:
    edges = collect_border_edges(ws, root, {"use_border_contact_merge": True, "include_merged_cells": True})
    if not edges:
        return []

    def has_grid_cell(row: int, col: int) -> bool:
        border = ws.cell(row, col).border
        has_horizontal = bool(getattr(border.top, "style", None) or getattr(border.bottom, "style", None))
        has_vertical = bool(getattr(border.left, "style", None) or getattr(border.right, "style", None))
        return has_horizontal and has_vertical

    def closed(box: Box) -> bool:
        return _has_closed_border(ws, box)

    occupied = {
        (row, col)
        for row in range(root.min_row + 1, root.max_row + 1)
        for col in range(root.min_col, root.max_col + 1)
        if has_grid_cell(row, col)
    }
    for merged in merged_boxes_with_values(ws):
        if not root.contains_box(merged):
            continue
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                occupied.add((row, col))

    candidates: list[Box] = []
    for box in connected_components_from_cells(occupied, connectivity=4, min_occupied_cells=4):
        for candidate in (
            Box(box.min_row, box.min_col, min(root.max_row, box.max_row + 1), box.max_col),
            Box(box.min_row, box.min_col, box.max_row, min(root.max_col, box.max_col + 1)),
        ):
            if candidate != box and closed(candidate):
                box = candidate
        if root.contains_box(box) and box != root and box.width < root.width and box.height < root.height and closed(box):
            candidates.append(box)

    kept: list[Box] = []
    for box in sorted(dedupe_boxes(candidates), key=lambda b: b.area, reverse=True):
        if not any(other.contains_box(box) for other in kept):
            kept.append(box)
    return kept


def _topology_root_boxes(root_refs: list[str], ws: Worksheet | None) -> list[Box]:
    roots = [_box_from_range(ref) for ref in root_refs]
    if not roots or ws is None:
        return roots

    bounds = roots[0]
    for box in roots[1:]:
        bounds = bounds.union(box)

    search_bounds = Box(
        max(1, bounds.min_row - 1),
        max(1, bounds.min_col - 1),
        min(ws.max_row, bounds.max_row + 1),
        min(ws.max_column, bounds.max_col + 1),
    )
    edges = collect_border_edges(ws, search_bounds, {"use_border_contact_merge": True, "include_merged_cells": True})
    parents = [
        box
        for box in _edge_component_boxes(edges)
        if sum(1 for root in roots if box.contains_box(root)) >= 2 and _has_closed_border(ws, box)
    ]
    if parents:
        return [min(parents, key=lambda box: box.area)]
    return roots


def _region_tree(ranges: list[str], ws: Worksheet | None = None) -> list[dict[str, Any]]:
    boxes = [(range_ref, _box_from_range(range_ref)) for range_ref in ranges]
    base_roots = _root_ranges(ranges)
    roots = [(box.range_ref, box) for box in _topology_root_boxes(base_roots, ws)]
    children_by_root: dict[str, list[Box]] = {range_ref: [] for range_ref, _ in roots}

    for root_ref, root_box in roots:
        topology_children = _inner_table_boxes(ws, root_box) if ws is not None else []
        if topology_children:
            children_by_root[root_ref].extend(topology_children)
            continue

        for _child_ref, child_box in boxes:
            if root_box.contains_box(child_box) and root_box != child_box:
                children_by_root[root_ref].append(child_box)

    return [
        {
            "range_ref": root_ref,
            "children": [
                {"range_ref": child.range_ref}
                for child in sorted(_keep_outer_boxes(children_by_root[root_ref]), key=lambda box: (box.min_row, box.min_col))
            ],
        }
        for root_ref, _ in roots
    ]


def _keep_outer_boxes(boxes: list[Box]) -> list[Box]:
    deduped = dedupe_boxes(boxes)
    return [
        box
        for box in deduped
        if not any(other != box and other.contains_box(box) for other in deduped)
    ]


def run_and_write(
    workbook_path: str | Path,
    *,
    out_dir: str | Path,
    sheet_name: str | None = None,
    config_path: str | Path | None = None,
    config_overrides: dict[str, Any] | None = None,
    write_images: bool = True,
) -> dict[str, Any]:
    config = load_config(config_path)
    if config_overrides:
        config.update(config_overrides)
    result = extract_workbook_info_regions(workbook_path, sheet_name=sheet_name, config=config)

    out = ensure_dir(out_dir)

    # Use a normal workbook for extracting embedded images because image anchors live
    # in the drawing layer, not in cached cell values.
    wb_drawings = open_workbook(workbook_path, data_only=False)

    # data_only=True is used only for debug overlay PNGs so formula cache values can
    # be displayed when Excel has saved them.
    wb_values = None
    if write_images and config.get("visualization", {}).get("enabled", True):
        wb_values = open_workbook(workbook_path, data_only=True)

    for sheet, data in result["sheets"].items():
        sheet_dir = ensure_dir(out / safe_name(sheet))

        if config.get("extract_embedded_images", True):
            ws_drawing = wb_drawings[sheet]
            images = extract_sheet_images_to_dir(
                workbook_path,
                ws_drawing,
                sheet_dir,
                rel_dir=str(config.get("embedded_image_dir", "images")),
                config=config,
            )
        else:
            images = []

        charts = extract_sheet_charts_to_dir(
            wb_drawings[sheet],
            sheet_dir,
            rel_dir=str(config.get("chart_image_dir", "charts")),
            config=config,
        )

        # Final user-facing schema:
        # {
        #   "sheet_name": "...",
        #   "regions": ["A1:P1", ...],
        #   "images": [{"name": "...", "range_ref": "...", "path": "..."}]
        # }
        root_regions = [box.range_ref for box in _topology_root_boxes(_root_ranges(data.get("info_regions", [])), wb_drawings[sheet])]
        sheet_output = {
            "sheet_name": data["sheet_name"],
            "regions": root_regions,
            "region_tree": _region_tree(data.get("info_regions", []), wb_drawings[sheet]),
            "images": images,
            "charts": charts,
        }

        result["sheets"][sheet] = sheet_output
        write_json(sheet_dir / "info_regions.json", sheet_output)

        if wb_values is not None and sheet_output.get("regions"):
            ws = wb_values[sheet]
            viz_cfg = config.get("visualization", {})
            overlay_regions = _overlay_regions_from_tree(sheet_output["region_tree"], sheet)
            render_region_overlay(
                ws,
                overlay_regions,
                sheet_dir / "info_regions.png",
                title=f"{sheet} - info_regions",
                bounds=None,
                max_rows=int(viz_cfg.get("max_rows", 160)),
                max_cols=int(viz_cfg.get("max_cols", 90)),
                pad_rows=int(viz_cfg.get("pad_rows", 1)),
                pad_cols=int(viz_cfg.get("pad_cols", 1)),
                preserve_dimensions=bool(viz_cfg.get("preserve_dimensions", True)),
                include_images=bool(viz_cfg.get("include_images", True)),
                include_cell_text=bool(viz_cfg.get("include_cell_text", True)),
                include_merged_cells=bool(viz_cfg.get("include_merged_cells", True)),
                scale=float(viz_cfg.get("scale", 1.0)),
                font_path=viz_cfg.get("font_path"),
                workbook_path=str(workbook_path),
            )

    summary = summarize_workbook_result(result)
    write_json(out / "info_regions_full.json", result)
    write_json(out / "info_regions_summary.json", summary)

    return result
