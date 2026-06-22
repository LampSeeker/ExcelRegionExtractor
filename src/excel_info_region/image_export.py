from __future__ import annotations

from pathlib import Path
from typing import Any
import re

from openpyxl.worksheet.worksheet import Worksheet

from .io import ensure_dir
from .cells import apply_print_area_bounds, visible_box
from .components import clip_box_to_bounds
from .raw_drawing import (
    extract_drawing_images,
    drawing_image_pixel_box,
    pixel_box_to_cell_box,
    sheet_pixel_axes,
)
from .schema import Box


def _safe_filename_part(text: str) -> str:
    text = text.replace(":", "_").replace("/", "_").replace("\\", "_")
    text = re.sub(r"[^0-9A-Za-z가-힣_.()-]+", "_", text)
    return text.strip("_") or "image"


def _image_extension(media_path: str, data: bytes) -> str:
    suffix = Path(media_path).suffix.lower()
    if suffix in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff"}:
        return suffix
    if data.startswith(b"\x89PNG"):
        return ".png"
    if data.startswith(b"\xff\xd8"):
        return ".jpg"
    if data.startswith(b"GIF"):
        return ".gif"
    if data.startswith(b"RIFF") and b"WEBP" in data[:16]:
        return ".webp"
    return ".bin"


def extract_sheet_images_to_dir(
    workbook_path: str | Path,
    ws: Worksheet,
    out_dir: str | Path,
    *,
    rel_dir: str = "images",
    config: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Extract embedded drawing images for one sheet.

    Returned JSON intentionally stays small:
    - range_ref: where the image appears on the sheet
    - path: relative file path from the sheet output directory
    """
    workbook_path = Path(workbook_path)
    out_dir = Path(out_dir)
    cfg = config or {}

    images = extract_drawing_images(workbook_path, ws.title)
    if not images:
        return []

    max_row = max(ws.max_row, max((i.to_marker.row if i.to_marker else i.from_marker.row + 50) for i in images))
    max_col = max(ws.max_column, max((i.to_marker.col if i.to_marker else i.from_marker.col + 30) for i in images))
    bounds = apply_print_area_bounds(ws, Box(1, 1, max_row + 2, max_col + 2), cfg)
    col_x, row_y, _col_w, _row_h = sheet_pixel_axes(ws, max_row + 2, max_col + 2)

    output: list[dict[str, Any]] = []
    used_names: set[str] = set()
    pending_files: list[tuple[Path, bytes]] = []

    for idx, img in enumerate(images, 1):
        px_box = drawing_image_pixel_box(img, ws, col_x, row_y)
        box = pixel_box_to_cell_box(px_box, col_x, row_y, max_row + 2, max_col + 2)
        visible = visible_box(ws, box, cfg)
        if visible is None:
            continue
        if bounds is not None:
            clipped = clip_box_to_bounds(visible, bounds)
            if clipped is None:
                continue
            box = clipped
        else:
            box = visible

        ext = _image_extension(img.media_path, img.data)
        range_part = _safe_filename_part(box.range_ref)
        name_part = _safe_filename_part(img.name) if img.name else f"IMG{idx:03d}"
        filename = f"IMG{idx:03d}_{range_part}_{name_part}{ext}"

        # Prevent accidental overwrite when Excel contains duplicated image names.
        base = filename
        counter = 2
        while filename in used_names:
            stem = Path(base).stem
            suffix = Path(base).suffix
            filename = f"{stem}_{counter}{suffix}"
            counter += 1
        used_names.add(filename)

        pending_files.append((Path(filename), img.data))

        output.append({
            "name": img.name or f"IMG{idx:03d}",
            "range_ref": box.range_ref,
            "path": f"{rel_dir}/{filename}",
        })

    if output:
        image_dir = ensure_dir(out_dir / rel_dir)
        for filename, data in pending_files:
            (image_dir / filename).write_bytes(data)

    return output
