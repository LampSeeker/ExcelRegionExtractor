from __future__ import annotations

from _common import parser, sheet_contexts, write_json
from excel_region_lab.dbscan_like import dbscan_like_cluster


def main() -> None:
    args = parser("Test DBSCAN-like Spatial Clustering").parse_args()
    for config, out, ws, signals, primitives in sheet_contexts(args):
        regions = dbscan_like_cluster(
            signals,
            primitives,
            float(config["dbscan_like"]["eps"]),
            int(config["dbscan_like"].get("min_samples", 1)),
        )
        data = {
            "sheet_name": ws.title,
            "algorithm": "dbscan_like",
            "primitive_count": len(primitives),
            "regions": [r.to_dict() for r in regions],
        }
        path = write_json(out / f"{ws.title}_dbscan_like_regions.json", data)
        print(f"[dbscan_like] {ws.title}: {len(primitives)} -> {len(regions)} regions -> {path}")


if __name__ == "__main__":
    main()
