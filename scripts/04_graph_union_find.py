from __future__ import annotations

from _common import parser, sheet_contexts, write_json
from excel_region_lab.merge import graph_union_find_merge


def main() -> None:
    args = parser("Test Graph-based Clustering + Union-Find").parse_args()
    for config, out, ws, signals, primitives in sheet_contexts(args):
        regions, edges = graph_union_find_merge(
            signals,
            primitives,
            config["pair_score"],
            float(config["graph_union_find"]["merge_threshold"]),
        )
        data = {
            "sheet_name": ws.title,
            "algorithm": "graph_union_find",
            "primitive_count": len(primitives),
            "regions": [r.to_dict() for r in regions],
            "top_edges": [e.to_dict() for e in edges[:100]],
        }
        path = write_json(out / f"{ws.title}_graph_union_find_regions.json", data)
        print(f"[graph_union_find] {ws.title}: {len(primitives)} -> {len(regions)} regions -> {path}")


if __name__ == "__main__":
    main()
