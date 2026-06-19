from __future__ import annotations

from _common import parser, sheet_contexts, write_json
from excel_region_lab.merge import hierarchical_merge


def main() -> None:
    args = parser("Test Hierarchical Agglomerative Clustering").parse_args()
    for config, out, ws, signals, primitives in sheet_contexts(args):
        regions, edges = hierarchical_merge(
            signals,
            primitives,
            config["pair_score"],
            float(config["hierarchical"]["merge_threshold"]),
            int(config["hierarchical"].get("max_merges", 200)),
        )
        data = {
            "sheet_name": ws.title,
            "algorithm": "hierarchical",
            "primitive_count": len(primitives),
            "regions": [r.to_dict() for r in regions],
            "top_edges": [e.to_dict() for e in edges[:100]],
        }
        path = write_json(out / f"{ws.title}_hierarchical_regions.json", data)
        print(f"[hierarchical] {ws.title}: {len(primitives)} -> {len(regions)} regions -> {path}")


if __name__ == "__main__":
    main()
