from __future__ import annotations

from pathlib import Path
from openpyxl.utils.cell import range_boundaries
from typing import Any

from .config import load_config
from .extractor import extract_workbook_info_regions, summarize_workbook_result, open_workbook
from .io import ensure_dir, safe_name, write_json
from .visualize import render_region_overlay


def _overlay_regions_from_ranges(ranges: list[str], sheet_name: str) -> list[dict]:
    regions: list[dict] = []
    for idx, range_ref in enumerate(ranges, 1):
        min_col, min_row, max_col, max_row = range_boundaries(range_ref)
        regions.append({
            "id": f"R{idx:03d}",
            "sheet_name": sheet_name,
            "range_ref": range_ref,
            "min_row": min_row,
            "min_col": min_col,
            "max_row": max_row,
            "max_col": max_col,
            "height": max_row - min_row + 1,
            "width": max_col - min_col + 1,
            "area": (max_row - min_row + 1) * (max_col - min_col + 1),
        })
    return regions


def run_and_write(
    workbook_path: str | Path,
    *,
    out_dir: str | Path,
    sheet_name: str | None = None,
    config_path: str | Path | None = None,
    write_images: bool = True,
) -> dict[str, Any]:
    config = load_config(config_path)
    result = extract_workbook_info_regions(workbook_path, sheet_name=sheet_name, config=config)
    summary = summarize_workbook_result(result)

    out = ensure_dir(out_dir)
    write_json(out / "info_regions_full.json", result)
    write_json(out / "info_regions_summary.json", summary)

    wb_values = None
    if write_images and config.get("visualization", {}).get("enabled", True):
        # data_only=True shows cached formula results when the workbook contains them.
        wb_values = open_workbook(workbook_path, data_only=True)

    for sheet, data in result["sheets"].items():
        sheet_dir = ensure_dir(out / safe_name(sheet))
        write_json(sheet_dir / "info_regions.json", data)

        if wb_values is not None and data.get("info_regions"):
            ws = wb_values[sheet]
            viz_cfg = config.get("visualization", {})
            overlay_regions = _overlay_regions_from_ranges(data["info_regions"], sheet)
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

    return result
