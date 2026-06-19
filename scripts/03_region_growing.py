from __future__ import annotations

from _common import parser, sheet_contexts, write_json
from excel_region_lab.merge import region_growing_merge


def main() -> None:
    args = parser("Test Region Growing / Region Merging").parse_args()
    for config, out, ws, signals, primitives in sheet_contexts(args):
        regions, edges = region_growing_merge(
            signals,
            primitives,
            config["pair_score"],
            float(config["region_growing"]["merge_threshold"]),
            int(config["region_growing"].get("max_iterations", 20)),
        )
        data = {
            "sheet_name": ws.title,
            "algorithm": "region_growing",
            "primitive_count": len(primitives),
            "regions": [r.to_dict() for r in regions],
            "top_edges": [e.to_dict() for e in edges[:100]],
        }
        path = write_json(out / f"{ws.title}_region_growing_regions.json", data)
        print(f"[region_growing] {ws.title}: {len(primitives)} -> {len(regions)} regions -> {path}")


if __name__ == "__main__":
    main()
