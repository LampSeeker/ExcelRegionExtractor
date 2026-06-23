from __future__ import annotations

import argparse

from .runner import run_and_write


def main() -> None:
    p = argparse.ArgumentParser(description="Extract Excel information regions only")
    p.add_argument("--workbook", required=True)
    p.add_argument("--sheet", default=None)
    p.add_argument("--config", default=None)
    p.add_argument("--out", default="outputs/info_regions")
    p.add_argument("--no-overlay", action="store_true", help="Skip PNG overlay generation")
    p.add_argument("--no-images", action="store_true", dest="no_overlay", help=argparse.SUPPRESS)
    p.add_argument("--respect-hidden", action="store_true", help="Exclude hidden rows and columns")
    p.add_argument("--use-print-area", action="store_true", help="Limit extraction to each sheet's print area")
    args = p.parse_args()

    overrides = {}
    if args.respect_hidden:
        overrides["respect_hidden_rows_cols"] = True
        overrides["respect_hidden_sheets"] = True
    if args.use_print_area:
        overrides["use_print_area_bounds"] = True

    result = run_and_write(
        args.workbook,
        out_dir=args.out,
        sheet_name=args.sheet,
        config_path=args.config,
        config_overrides=overrides,
        write_images=not args.no_overlay,
    )

    print(f"[extract_info_regions] sheets={len(result['sheets'])} -> {args.out}")
    for sheet, data in result["sheets"].items():
        regions = data.get("regions", data.get("info_regions", []))
        print(f"  - {sheet}: regions={len(regions)}, ranges={regions}, images={len(data.get('images', []))}")
