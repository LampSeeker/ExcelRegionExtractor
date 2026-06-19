from __future__ import annotations

from collections import defaultdict
from typing import Any

from .schema import Box, Region, RegionEdge
from .signals import SheetSignals, region_features
from .scoring import pair_score


class UnionFind:
    def __init__(self, ids: list[str]):
        self.parent = {i: i for i in ids}
        self.rank = {i: 0 for i in ids}

    def find(self, x: str) -> str:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: str, b: str) -> bool:
        ra = self.find(a)
        rb = self.find(b)
        if ra == rb:
            return False
        if self.rank[ra] < self.rank[rb]:
            ra, rb = rb, ra
        self.parent[rb] = ra
        if self.rank[ra] == self.rank[rb]:
            self.rank[ra] += 1
        return True

    def groups(self) -> dict[str, list[str]]:
        result: dict[str, list[str]] = defaultdict(list)
        for item in self.parent:
            result[self.find(item)].append(item)
        return result


def build_edges(signals: SheetSignals, regions: list[Region], score_config: dict[str, Any]) -> list[RegionEdge]:
    edges: list[RegionEdge] = []
    for i, a in enumerate(regions):
        for b in regions[i + 1:]:
            edge = pair_score(signals, a, b, score_config)
            edges.append(edge)
    return sorted(edges, key=lambda e: e.score, reverse=True)


def merge_group_to_region(
    signals: SheetSignals,
    members: list[Region],
    algorithm: str,
    idx: int,
    score: float | None = None,
) -> Region:
    box = members[0].box
    for m in members[1:]:
        box = box.union(m.box)
    member_ids: list[str] = []
    for m in members:
        member_ids.extend(m.members or [m.id])
    return Region(
        id=f"R{idx:03d}",
        sheet_name=signals.sheet_name,
        box=box,
        algorithm=algorithm,
        features=region_features(signals, box),
        members=sorted(set(member_ids)),
        score=score,
    )


def graph_union_find_merge(
    signals: SheetSignals,
    primitive_regions: list[Region],
    score_config: dict[str, Any],
    threshold: float,
    algorithm: str = "graph_union_find",
) -> tuple[list[Region], list[RegionEdge]]:
    if not primitive_regions:
        return [], []
    edges = build_edges(signals, primitive_regions, score_config)
    uf = UnionFind([r.id for r in primitive_regions])
    for edge in edges:
        if edge.score >= threshold:
            uf.union(edge.source, edge.target)

    region_by_id = {r.id: r for r in primitive_regions}
    merged: list[Region] = []
    for idx, ids in enumerate(uf.groups().values(), 1):
        members = [region_by_id[i] for i in ids]
        score_values = [e.score for e in edges if e.source in ids and e.target in ids]
        group_score = round(sum(score_values) / len(score_values), 4) if score_values else None
        merged.append(merge_group_to_region(signals, members, algorithm, idx, group_score))
    return sorted(merged, key=lambda r: (r.box.min_row, r.box.min_col)), edges


def region_growing_merge(
    signals: SheetSignals,
    primitive_regions: list[Region],
    score_config: dict[str, Any],
    threshold: float,
    max_iterations: int = 20,
) -> tuple[list[Region], list[RegionEdge]]:
    current = list(primitive_regions)
    all_edges: list[RegionEdge] = []
    next_id = 1
    for _ in range(max_iterations):
        if len(current) <= 1:
            break
        edges = build_edges(signals, current, score_config)
        all_edges.extend(edges)
        best = next((e for e in edges if e.score >= threshold), None)
        if best is None:
            break
        a = next(r for r in current if r.id == best.source)
        b = next(r for r in current if r.id == best.target)
        current = [r for r in current if r.id not in {a.id, b.id}]
        merged = merge_group_to_region(signals, [a, b], "region_growing", next_id, best.score)
        merged.id = f"G{next_id:03d}"
        next_id += 1
        current.append(merged)

    final: list[Region] = []
    for idx, r in enumerate(sorted(current, key=lambda r: (r.box.min_row, r.box.min_col)), 1):
        final.append(Region(
            id=f"R{idx:03d}",
            sheet_name=r.sheet_name,
            box=r.box,
            algorithm="region_growing",
            features=r.features,
            members=r.members,
            score=r.score,
        ))
    return final, sorted(all_edges, key=lambda e: e.score, reverse=True)


def hierarchical_merge(
    signals: SheetSignals,
    primitive_regions: list[Region],
    score_config: dict[str, Any],
    threshold: float,
    max_merges: int = 200,
) -> tuple[list[Region], list[RegionEdge]]:
    current = list(primitive_regions)
    all_edges: list[RegionEdge] = []
    merge_count = 0
    next_id = 1
    while len(current) > 1 and merge_count < max_merges:
        edges = build_edges(signals, current, score_config)
        all_edges.extend(edges)
        if not edges or edges[0].score < threshold:
            break
        best = edges[0]
        a = next(r for r in current if r.id == best.source)
        b = next(r for r in current if r.id == best.target)
        current = [r for r in current if r.id not in {a.id, b.id}]
        merged = merge_group_to_region(signals, [a, b], "hierarchical", next_id, best.score)
        merged.id = f"H{next_id:03d}"
        current.append(merged)
        merge_count += 1
        next_id += 1

    final: list[Region] = []
    for idx, r in enumerate(sorted(current, key=lambda r: (r.box.min_row, r.box.min_col)), 1):
        final.append(Region(
            id=f"R{idx:03d}",
            sheet_name=r.sheet_name,
            box=r.box,
            algorithm="hierarchical",
            features=r.features,
            members=r.members,
            score=r.score,
        ))
    return final, sorted(all_edges, key=lambda e: e.score, reverse=True)
