from __future__ import annotations

from _common import parser, sheet_contexts
from excel_region_lab.hybrid import hybrid_grouping
from excel_region_lab.io import write_json
from excel_region_lab.visualize import render_region_overlay
from excel_region_lab.signals import open_workbook


def main() -> None:
    p = parser("Run hybrid grouping: connected components -> graph merge -> projection split")
    p.add_argument("--no-images", action="store_true")
    args = p.parse_args()

    display_wb = None if args.no_images else open_workbook(args.workbook, data_only=True)

    for config, out, ws, signals, primitives in sheet_contexts(args):
        roots, candidates, edges = hybrid_grouping(
            signals,
            primitives,
            config["pair_score"],
            config["graph_union_find"],
            config.get("hybrid", {}),
        )
        data = {
            "roots": [r.to_dict() for r in roots],
            "regions": [r.to_dict() for r in candidates],
            "edges": [e.to_dict() for e in edges[:200]],
            "description": "connected_components -> graph_union_find root merge -> projection_profile candidate split",
        }
        path = write_json(out / f"{ws.title}_hybrid.json", data)
        if not args.no_images:
            viz = config.get("visualization", {})
            display_ws = display_wb[ws.title] if display_wb is not None else ws
            render_region_overlay(
                display_ws,
                data["regions"],
                out / f"{ws.title}_hybrid.png",
                title=f"{ws.title} - hybrid candidates",
                bounds=signals.bounds.to_dict() if signals.bounds else None,
                max_rows=int(viz.get("max_rows", 120)),
                max_cols=int(viz.get("max_cols", 50)),
                pad_rows=int(viz.get("pad_rows", 2)),
                pad_cols=int(viz.get("pad_cols", 2)),
                preserve_dimensions=bool(viz.get("preserve_dimensions", True)),
                include_images=bool(viz.get("include_images", True)),
                include_cell_text=bool(viz.get("include_cell_text", True)),
                include_merged_cells=bool(viz.get("include_merged_cells", True)),
                scale=float(viz.get("scale", 1.0)),
            )
        print(f"[hybrid] {ws.title}: roots={len(roots)} candidates={len(candidates)} -> {path}")


if __name__ == "__main__":
    main()
