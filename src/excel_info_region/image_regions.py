from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .components import dedupe_boxes
from .raw_drawing import drawing_image_boxes
from .schema import Box


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
