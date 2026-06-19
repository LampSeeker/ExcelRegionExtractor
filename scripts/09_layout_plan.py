from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from excel_region_lab.config import load_config
from excel_region_lab.io import ensure_dir, write_json
from excel_region_lab.runner import run_workbook
from excel_region_lab.signals import open_workbook
from excel_region_lab.visualize import render_region_overlay


def safe_name(name: str) -> str:
    forbidden = set(r'\/:*?"<>|')
    return "".join("_" if ch in forbidden else ch for ch in name)


def main() -> None:
    p = argparse.ArgumentParser(description="Run conservative layout plan only")
    p.add_argument("--workbook", default=str(PROJECT_ROOT / "examples" / "sample.xlsx"))
    p.add_argument("--sheet", default=None)
    p.add_argument("--config", default=str(PROJECT_ROOT / "config" / "default.json"))
    p.add_argument("--out", default=str(PROJECT_ROOT / "outputs" / "layout_plan"))
    p.add_argument("--no-images", action="store_true")
    args = p.parse_args()

    config = load_config(args.config)
    config.setdefault("visualization", {})["workbook_path"] = str(args.workbook)
    result = run_workbook(args.workbook, sheet_name=args.sheet, config_path=args.config)
    out = ensure_dir(args.out)
    write_json(out / "layout_plan_full.json", {
        "workbook": result["workbook"],
        "sheets": {
            sheet_name: sheet_data["layout_plan"]
            for sheet_name, sheet_data in result["sheets"].items()
        },
    })

    wb_values = None if args.no_images else open_workbook(args.workbook, data_only=True)
    for sheet_name, sheet_data in result["sheets"].items():
        sheet_dir = ensure_dir(out / safe_name(sheet_name))
        layout = sheet_data["layout_plan"]
        write_json(sheet_dir / "layout_plan.json", layout)
        write_json(sheet_dir / "layout_candidates.json", sheet_data["layout_candidates"])
        write_json(sheet_dir / "layout_roots.json", sheet_data["layout_roots"])

        if wb_values is not None:
            ws = wb_values[sheet_name]
            viz_cfg = config.get("visualization", {})
            render_region_overlay(
                ws,
                sheet_data["layout_candidates"]["regions"],
                sheet_dir / "layout_candidates.png",
                title=f"{sheet_name} - layout candidates",
                bounds=sheet_data.get("bounds"),
                max_rows=int(viz_cfg.get("max_rows", 140)),
                max_cols=int(viz_cfg.get("max_cols", 80)),
                pad_rows=int(viz_cfg.get("pad_rows", 1)),
                pad_cols=int(viz_cfg.get("pad_cols", 1)),
                preserve_dimensions=bool(viz_cfg.get("preserve_dimensions", True)),
                include_images=bool(viz_cfg.get("include_images", True)),
                include_cell_text=bool(viz_cfg.get("include_cell_text", True)),
                include_merged_cells=bool(viz_cfg.get("include_merged_cells", True)),
                scale=float(viz_cfg.get("scale", 1.0)),
                font_path=viz_cfg.get("font_path"),
                workbook_path=viz_cfg.get("workbook_path"),
            )
            render_region_overlay(
                ws,
                sheet_data["layout_roots"]["regions"],
                sheet_dir / "layout_roots.png",
                title=f"{sheet_name} - layout roots",
                bounds=sheet_data.get("bounds"),
                max_rows=int(viz_cfg.get("max_rows", 140)),
                max_cols=int(viz_cfg.get("max_cols", 80)),
                pad_rows=int(viz_cfg.get("pad_rows", 1)),
                pad_cols=int(viz_cfg.get("pad_cols", 1)),
                preserve_dimensions=bool(viz_cfg.get("preserve_dimensions", True)),
                include_images=bool(viz_cfg.get("include_images", True)),
                include_cell_text=bool(viz_cfg.get("include_cell_text", True)),
                include_merged_cells=bool(viz_cfg.get("include_merged_cells", True)),
                scale=float(viz_cfg.get("scale", 1.0)),
                font_path=viz_cfg.get("font_path"),
                workbook_path=viz_cfg.get("workbook_path"),
            )

    print(f"[layout_plan] sheets={len(result['sheets'])} -> {out}")


if __name__ == "__main__":
    main()
