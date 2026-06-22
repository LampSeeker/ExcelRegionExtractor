from __future__ import annotations

import argparse
from pathlib import Path

from .runner import run_and_write


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def main() -> None:
    p = argparse.ArgumentParser(description="Extract Excel information regions only")
    p.add_argument("--workbook", default=str(PROJECT_ROOT / "examples" / "sample.xlsx"))
    p.add_argument("--sheet", default=None)
    p.add_argument("--config", default=str(PROJECT_ROOT / "config" / "default.json"))
    p.add_argument("--out", default=str(PROJECT_ROOT / "outputs" / "info_regions"))
    p.add_argument("--no-images", action="store_true", help="Skip PNG overlay generation")
    args = p.parse_args()

    result = run_and_write(
        args.workbook,
        out_dir=args.out,
        sheet_name=args.sheet,
        config_path=args.config,
        write_images=not args.no_images,
    )

    print(f"[extract_info_regions] sheets={len(result['sheets'])} -> {args.out}")
    for sheet, data in result["sheets"].items():
        regions = data.get("regions", data.get("info_regions", []))
        print(f"  - {sheet}: regions={len(regions)}, ranges={regions}, images={len(data.get('images', []))}")
