
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .merge import graph_union_find_merge
from .schema import Box, Region, RegionEdge
from .scoring import interval_overlap_ratio
from .signals import SheetSignals, region_features


def _feature_float(region: Region, key: str, default: float = 0.0) -> float:
    try:
        return float(region.features.get(key, default))
    except Exception:
        return default


def role_hint(region: Region) -> str:
    """Generic role hint only. It is not a hardcoded sheet-specific classifier."""
    f = region.features
    height = region.box.height
    width = region.box.width
    non_empty = int(f.get("non_empty_count", 0) or 0)
    numeric_density = _feature_float(region, "numeric_density")
    text_density = _feature_float(region, "text_density")
    bold_density = _feature_float(region, "bold_density")
    merged_count = int(f.get("merged_count", 0) or 0)
    image_count = int(f.get("image_count", 0) or 0)
    border = _feature_float(region, "border_coverage")

    if height <= 2 and width >= 4 and text_density >= 0.5 and (bold_density > 0 or merged_count > 0):
        return "title"
    if image_count > 0 and numeric_density < 0.2:
        return "figure_or_diagram"
    if non_empty >= 6 and (numeric_density >= 0.2 or border >= 0.15):
        return "table_or_form"
    if text_density >= 0.6:
        return "note_or_metadata"
    return "unknown"


def _col_overlap_score(a: Region, b: Region) -> float:
    return interval_overlap_ratio(a.box.col_overlap(b.box), a.box.width, b.box.width)


def _row_overlap_score(a: Region, b: Region) -> float:
    return interval_overlap_ratio(a.box.row_overlap(b.box), a.box.height, b.box.height)


def _vertical_order(a: Region, b: Region) -> tuple[Region, Region] | None:
    if a.box.max_row < b.box.min_row:
        return a, b
    if b.box.max_row < a.box.min_row:
        return b, a
    return None


def _relation_score_title_of(title: Region, body: Region, config: dict[str, Any]) -> float:
    max_gap = int(config.get("title_max_row_gap", 3))
    min_col_overlap = float(config.get("title_min_col_overlap", 0.70))
    if role_hint(title) != "title":
        return 0.0
    ordered = _vertical_order(title, body)
    if ordered is None or ordered[0].id != title.id:
        return 0.0
    gap = title.box.row_gap(body.box)
    if gap > max_gap:
        return 0.0
    col_overlap = _col_overlap_score(title, body)
    if col_overlap < min_col_overlap:
        return 0.0
    gap_score = 1.0 - (gap / max(max_gap, 1))
    body_size_score = min(1.0, body.box.area / max(title.box.area, 1))
    return round(0.60 * col_overlap + 0.25 * gap_score + 0.15 * body_size_score, 4)


def _relation_score_same_table_fragment(a: Region, b: Region, config: dict[str, Any]) -> float:
    ordered = _vertical_order(a, b)
    if ordered is None:
        return 0.0
    top, bottom = ordered
    if role_hint(top) == "title" or role_hint(bottom) == "title":
        return 0.0
    max_gap = int(config.get("fragment_max_row_gap", 4))
    gap = top.box.row_gap(bottom.box)
    if gap > max_gap:
        return 0.0

    col_overlap = _col_overlap_score(top, bottom)
    min_col_overlap = float(config.get("fragment_min_col_overlap", 0.85))
    if col_overlap < min_col_overlap:
        return 0.0

    width_ratio = min(top.box.width, bottom.box.width) / max(top.box.width, bottom.box.width, 1)
    numeric_sim = 1.0 - min(1.0, abs(_feature_float(top, "numeric_density") - _feature_float(bottom, "numeric_density")))
    border_sim = min(_feature_float(top, "border_coverage"), _feature_float(bottom, "border_coverage"))
    gap_score = 1.0 - (gap / max(max_gap, 1))

    return round(
        0.35 * col_overlap
        + 0.20 * width_ratio
        + 0.15 * numeric_sim
        + 0.20 * border_sim
        + 0.10 * gap_score,
        4,
    )


def build_region_relations(regions: list[Region], config: dict[str, Any]) -> list[dict[str, Any]]:
    relations: list[dict[str, Any]] = []
    for i, a in enumerate(regions):
        for b in regions[i + 1:]:
            title_score_ab = _relation_score_title_of(a, b, config)
            title_score_ba = _relation_score_title_of(b, a, config)
            if title_score_ab >= float(config.get("title_relation_threshold", 0.70)):
                relations.append({
                    "source": a.id,
                    "target": b.id,
                    "type": "title_of",
                    "score": title_score_ab,
                    "source_range": a.box.range_ref,
                    "target_range": b.box.range_ref,
                })
            if title_score_ba >= float(config.get("title_relation_threshold", 0.70)):
                relations.append({
                    "source": b.id,
                    "target": a.id,
                    "type": "title_of",
                    "score": title_score_ba,
                    "source_range": b.box.range_ref,
                    "target_range": a.box.range_ref,
                })

            fragment_score = _relation_score_same_table_fragment(a, b, config)
            if fragment_score >= float(config.get("fragment_relation_threshold", 0.76)):
                relations.append({
                    "source": a.id,
                    "target": b.id,
                    "type": "same_table_fragment",
                    "score": fragment_score,
                    "source_range": a.box.range_ref,
                    "target_range": b.box.range_ref,
                })
    return sorted(relations, key=lambda r: (-float(r["score"]), r["source"], r["target"]))


class _UF:
    def __init__(self, ids: list[str]):
        self.parent = {i: i for i in ids}

    def find(self, x: str) -> str:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: str, b: str) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra

    def groups(self) -> dict[str, list[str]]:
        out: dict[str, list[str]] = {}
        for x in self.parent:
            out.setdefault(self.find(x), []).append(x)
        return out


def _merge_regions(signals: SheetSignals, members: list[Region], idx: int) -> Region:
    box = members[0].box
    primitive_ids: list[str] = []
    for member in members:
        box = box.union(member.box)
        primitive_ids.extend(member.members or [member.id])
    return Region(
        id=f"R{idx:03d}",
        sheet_name=signals.sheet_name,
        box=box,
        algorithm="layout_root",
        features=region_features(signals, box),
        members=sorted(set(primitive_ids)),
    )


def build_layout_plan(
    signals: SheetSignals,
    primitive_regions: list[Region],
    config: dict[str, Any],
) -> dict[str, Any]:
    """Build a conservative layout plan.

    Important behavior:
    - Primitive/candidate regions remain connected-component islands.
    - Title/body is represented as a relation, not a bbox merge.
    - Root merge is only allowed for high-confidence same_table_fragment relations.
    """
    normalized: list[Region] = []
    for idx, region in enumerate(sorted(primitive_regions, key=lambda r: (r.box.min_row, r.box.min_col)), 1):
        normalized.append(
            Region(
                id=f"C{idx:03d}",
                sheet_name=region.sheet_name,
                box=region.box,
                algorithm="layout_candidate",
                features={**region.features, "role_hint": role_hint(region)},
                members=region.members or [region.id],
                score=region.score,
            )
        )

    relations = build_region_relations(normalized, config.get("relations", {}))

    uf = _UF([r.id for r in normalized])
    if bool(config.get("merge_same_table_fragments", False)):
        merge_threshold = float(config.get("root_merge_relation_threshold", 0.88))
        for rel in relations:
            if rel["type"] == "same_table_fragment" and float(rel["score"]) >= merge_threshold:
                uf.union(str(rel["source"]), str(rel["target"]))

    by_id = {r.id: r for r in normalized}
    root_regions: list[Region] = []
    for idx, ids in enumerate(uf.groups().values(), 1):
        members = [by_id[i] for i in ids]
        root_regions.append(_merge_regions(signals, members, idx))

    root_regions = sorted(root_regions, key=lambda r: (r.box.min_row, r.box.min_col))
    for idx, root in enumerate(root_regions, 1):
        root.id = f"R{idx:03d}"

    return {
        "primitive_regions": [r.to_dict() for r in normalized],
        "candidate_regions": [r.to_dict() for r in normalized],
        "root_regions": [r.to_dict() for r in root_regions],
        "relations": relations,
    }
