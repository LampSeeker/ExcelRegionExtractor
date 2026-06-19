from __future__ import annotations

from _common import parser, sheet_contexts, write_json
from excel_region_lab.xy_cut import xy_cut


def main() -> None:
    args = parser("Test XY-cut Recursive Segmentation").parse_args()
    for config, out, ws, signals, _ in sheet_contexts(args):
        regions = xy_cut(signals, config["xy_cut"])
        data = {
            "sheet_name": ws.title,
            "algorithm": "xy_cut",
            "regions": [r.to_dict() for r in regions],
        }
        path = write_json(out / f"{ws.title}_xy_cut_regions.json", data)
        print(f"[xy_cut] {ws.title}: {len(regions)} regions -> {path}")


if __name__ == "__main__":
    main()
