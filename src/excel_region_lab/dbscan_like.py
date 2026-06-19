from __future__ import annotations

from typing import Any

from .schema import Region
from .signals import SheetSignals
from .merge import merge_group_to_region


def block_distance(a: Region, b: Region, signals: SheetSignals) -> float:
    if signals.bounds is None:
        return 1.0
    bounds = signals.bounds
    row_scale = max(bounds.height, 1)
    col_scale = max(bounds.width, 1)
    ac_r = (a.box.min_row + a.box.max_row) / 2
    ac_c = (a.box.min_col + a.box.max_col) / 2
    bc_r = (b.box.min_row + b.box.max_row) / 2
    bc_c = (b.box.min_col + b.box.max_col) / 2
    spatial = ((abs(ac_r - bc_r) / row_scale) ** 2 + (abs(ac_c - bc_c) / col_scale) ** 2) ** 0.5
    col_overlap_bonus = a.box.col_overlap(b.box) / max(min(a.box.width, b.box.width), 1)
    row_overlap_bonus = a.box.row_overlap(b.box) / max(min(a.box.height, b.box.height), 1)
    density_diff = abs(float(a.features.get("non_empty_density", 0)) - float(b.features.get("non_empty_density", 0)))
    numeric_diff = abs(float(a.features.get("numeric_density", 0)) - float(b.features.get("numeric_density", 0)))
    # Lower is closer. Axis overlaps lower the effective distance.
    return max(0.0, spatial + 0.25 * density_diff + 0.15 * numeric_diff - 0.20 * max(col_overlap_bonus, row_overlap_bonus))


def neighbors(regions: list[Region], idx: int, signals: SheetSignals, eps: float) -> list[int]:
    return [j for j, other in enumerate(regions) if j != idx and block_distance(regions[idx], other, signals) <= eps]


def dbscan_like_cluster(
    signals: SheetSignals,
    primitive_regions: list[Region],
    eps: float,
    min_samples: int,
) -> list[Region]:
    n = len(primitive_regions)
    if n == 0:
        return []
    labels = [None] * n
    cluster_id = 0

    for i in range(n):
        if labels[i] is not None:
            continue
        nbrs = neighbors(primitive_regions, i, signals, eps)
        if len(nbrs) + 1 < min_samples:
            labels[i] = -1
            continue
        cluster_id += 1
        labels[i] = cluster_id
        seed_set = list(nbrs)
        while seed_set:
            j = seed_set.pop()
            if labels[j] == -1:
                labels[j] = cluster_id
            if labels[j] is not None:
                continue
            labels[j] = cluster_id
            j_nbrs = neighbors(primitive_regions, j, signals, eps)
            if len(j_nbrs) + 1 >= min_samples:
                for k in j_nbrs:
                    if labels[k] is None:
                        seed_set.append(k)

    grouped: dict[int, list[Region]] = {}
    noise: list[Region] = []
    for label, region in zip(labels, primitive_regions):
        if label == -1 or label is None:
            noise.append(region)
        else:
            grouped.setdefault(int(label), []).append(region)

    result: list[Region] = []
    for idx, members in enumerate(grouped.values(), 1):
        result.append(merge_group_to_region(signals, members, "dbscan_like", idx))
    offset = len(result)
    for i, member in enumerate(noise, 1):
        result.append(merge_group_to_region(signals, [member], "dbscan_like", offset + i))
    return sorted(result, key=lambda r: (r.box.min_row, r.box.min_col))
