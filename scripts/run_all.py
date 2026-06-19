from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from excel_info_region.runner import run_and_write


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
        print(
            f"  - {sheet}: "
            f"regions={len(data.get('info_regions', []))}, "
            f"ranges={data.get('info_regions', [])}"
        )


if __name__ == "__main__":
    main()
