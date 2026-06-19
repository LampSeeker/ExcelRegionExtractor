from __future__ import annotations

from pathlib import Path
from typing import Any

from .components import connected_components
from .config import load_config
from .dbscan_like import dbscan_like_cluster
from .merge import graph_union_find_merge, hierarchical_merge, region_growing_merge
from .hybrid import hybrid_grouping
from .profiles import projection_profile
from .signals import extract_sheet_signals, iter_target_sheets, open_workbook
from .xy_cut import xy_cut
from .layout_plan import build_layout_plan


def as_region_list(regions):
    return [r.to_dict() for r in regions]


def as_edge_list(edges, limit: int = 200):
    return [e.to_dict() for e in edges[:limit]]


def run_algorithms_on_sheet(ws, config: dict[str, Any]) -> dict[str, Any]:
    signals = extract_sheet_signals(ws, config["signals"])
    primitives = connected_components(signals, config["connected_components"])
    profile = projection_profile(signals, config["projection_profile"])
    growing, growing_edges = region_growing_merge(
        signals,
        primitives,
        config["pair_score"],
        float(config["region_growing"]["merge_threshold"]),
        int(config["region_growing"].get("max_iterations", 20)),
    )
    graph_regions, graph_edges = graph_union_find_merge(
        signals,
        primitives,
        config["pair_score"],
        float(config["graph_union_find"]["merge_threshold"]),
    )
    hierarchical, hierarchical_edges = hierarchical_merge(
        signals,
        primitives,
        config["pair_score"],
        float(config["hierarchical"]["merge_threshold"]),
        int(config["hierarchical"].get("max_merges", 200)),
    )
    dbscan = dbscan_like_cluster(
        signals,
        primitives,
        float(config["dbscan_like"]["eps"]),
        int(config["dbscan_like"].get("min_samples", 1)),
    )
    xy = xy_cut(signals, config["xy_cut"])
    hybrid_roots, hybrid_candidates, hybrid_edges = hybrid_grouping(
        signals,
        primitives,
        config["pair_score"],
        config["graph_union_find"],
        config.get("hybrid", {}),
    )
    layout = build_layout_plan(signals, primitives, config.get("layout_plan", {}))

    return {
        "sheet_name": ws.title,
        "bounds": signals.bounds.to_dict() if signals.bounds else None,
        "signals_summary": {
            "cell_signal_count": len(signals.cells),
            "occupied_count": len(signals.occupied),
            "merged_range_count": len(signals.merged_ranges),
            "image_range_count": len(signals.image_ranges),
            "image_ranges": [b.to_dict() for b in signals.image_ranges],
        },
        "connected_components": {"regions": as_region_list(primitives)},
        "projection_profile": profile,
        "region_growing": {"regions": as_region_list(growing), "edges": as_edge_list(growing_edges)},
        "graph_union_find": {"regions": as_region_list(graph_regions), "edges": as_edge_list(graph_edges)},
        "hierarchical": {"regions": as_region_list(hierarchical), "edges": as_edge_list(hierarchical_edges)},
        "dbscan_like": {"regions": as_region_list(dbscan)},
        "xy_cut": {"regions": as_region_list(xy)},
        "layout_plan": layout,
        "layout_candidates": {
            "regions": layout.get("candidate_regions", []),
            "relations": layout.get("relations", []),
            "description": "primitive connected-component islands; default VLM candidate regions",
        },
        "layout_roots": {
            "regions": layout.get("root_regions", []),
            "relations": layout.get("relations", []),
            "description": "conservative root regions; title/body remains relation unless explicitly merged",
        },
        "hybrid": {
            "roots": as_region_list(hybrid_roots),
            "regions": as_region_list(hybrid_candidates),
            "edges": as_edge_list(hybrid_edges),
            "description": "connected_components -> graph_union_find root merge -> projection_profile candidate split",
        },
    }


def run_workbook(workbook_path: str | Path, sheet_name: str | None = None, config_path: str | Path | None = None) -> dict[str, Any]:
    config = load_config(config_path)
    config.setdefault("signals", {})["workbook_path"] = str(workbook_path)
    wb = open_workbook(workbook_path)
    result = {
        "workbook": str(workbook_path),
        "sheets": {},
    }
    for ws in iter_target_sheets(wb, sheet_name):
        result["sheets"][ws.title] = run_algorithms_on_sheet(ws, config)
    return result


def summarize_run(result: dict[str, Any]) -> dict[str, Any]:
    summary = {"workbook": result["workbook"], "sheets": []}
    algorithms = [
        "connected_components",
        "region_growing",
        "graph_union_find",
        "hierarchical",
        "dbscan_like",
        "xy_cut",
        "hybrid",
        "layout_candidates",
        "layout_roots",
    ]
    for sheet_name, sheet_data in result["sheets"].items():
        row = {
            "sheet_name": sheet_name,
            "bounds": sheet_data["bounds"],
            "signals_summary": sheet_data["signals_summary"],
        }
        for algo in algorithms:
            row[f"{algo}_count"] = len(sheet_data[algo]["regions"])
            row[f"{algo}_ranges"] = [r["range_ref"] for r in sheet_data[algo]["regions"][:20]]
        layout = sheet_data.get("layout_plan", {})
        row["layout_candidate_count"] = len(layout.get("candidate_regions", []))
        row["layout_root_count"] = len(layout.get("root_regions", []))
        row["layout_relation_count"] = len(layout.get("relations", []))
        row["layout_candidate_ranges"] = [r["range_ref"] for r in layout.get("candidate_regions", [])[:20]]
        row["layout_root_ranges"] = [r["range_ref"] for r in layout.get("root_regions", [])[:20]]
        row["blank_row_runs"] = sheet_data["projection_profile"].get("blank_row_runs", [])[:20]
        row["blank_col_runs"] = sheet_data["projection_profile"].get("blank_col_runs", [])[:20]
        summary["sheets"].append(row)
    return summary
