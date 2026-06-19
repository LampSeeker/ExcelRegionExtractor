from __future__ import annotations

from collections import Counter
from typing import Any

from .schema import Box, Region, RegionEdge
from .signals import SheetSignals, row_merge_signature


def clamp01(value: float) -> float:
    return max(0.0, min(1.0, value))


def interval_overlap_ratio(overlap: int, len_a: int, len_b: int) -> float:
    return overlap / max(min(len_a, len_b), 1)


def gap_score(gap: int, max_gap: int) -> float:
    if gap <= 0:
        return 1.0
    if gap > max_gap:
        return 0.0
    return 1.0 - (gap / max_gap)


def border_continuity(signals: SheetSignals, a: Box, b: Box) -> float:
    # Look at cells between vertically or horizontally adjacent regions.
    # It is intentionally generic: strong border or normal border in the connecting corridor increases score.
    row_gap = a.row_gap(b)
    col_gap = a.col_gap(b)
    scores: list[float] = []

    if row_gap >= 0 and a.col_overlap(b) > 0 and (a.max_row < b.min_row or b.max_row < a.min_row):
        top = a if a.max_row < b.min_row else b
        bottom = b if top is a else a
        min_col = max(top.min_col, bottom.min_col)
        max_col = min(top.max_col, bottom.max_col)
        rows = range(top.max_row, bottom.min_row + 1)
        for row in rows:
            for col in range(min_col, max_col + 1):
                sig = signals.cells.get((row, col))
                if sig:
                    scores.append(max(sig.border_score, sig.strong_border_score))

    if col_gap >= 0 and a.row_overlap(b) > 0 and (a.max_col < b.min_col or b.max_col < a.min_col):
        left = a if a.max_col < b.min_col else b
        right = b if left is a else a
        min_row = max(left.min_row, right.min_row)
        max_row = min(left.max_row, right.max_row)
        cols = range(left.max_col, right.min_col + 1)
        for row in range(min_row, max_row + 1):
            for col in cols:
                sig = signals.cells.get((row, col))
                if sig:
                    scores.append(max(sig.border_score, sig.strong_border_score))

    if not scores:
        return 0.0
    return clamp01(sum(scores) / len(scores))


def merge_signature_similarity(signals: SheetSignals, a: Box, b: Box) -> float:
    rows_a = range(a.min_row, a.max_row + 1)
    rows_b = range(b.min_row, b.max_row + 1)
    sigs_a = Counter(row_merge_signature(signals, r, a.min_col, a.max_col) for r in rows_a)
    sigs_b = Counter(row_merge_signature(signals, r, b.min_col, b.max_col) for r in rows_b)
    keys = set(sigs_a) | set(sigs_b)
    if not keys:
        return 0.0
    intersection = sum(min(sigs_a[k], sigs_b[k]) for k in keys)
    union = sum(max(sigs_a[k], sigs_b[k]) for k in keys)
    return intersection / max(union, 1)


def style_similarity(a: Region, b: Region) -> float:
    fill_score = 1.0 if a.features.get("dominant_fill") and a.features.get("dominant_fill") == b.features.get("dominant_fill") else 0.0
    align_score = 1.0 if a.features.get("dominant_alignment") and a.features.get("dominant_alignment") == b.features.get("dominant_alignment") else 0.0
    bold_diff = abs(float(a.features.get("bold_density", 0)) - float(b.features.get("bold_density", 0)))
    bold_score = 1.0 - min(1.0, bold_diff)
    return 0.4 * bold_score + 0.3 * fill_score + 0.3 * align_score


def density_similarity(a: Region, b: Region) -> float:
    keys = ["non_empty_density", "numeric_density", "text_density", "formula_density"]
    diffs = [abs(float(a.features.get(k, 0)) - float(b.features.get(k, 0))) for k in keys]
    return 1.0 - min(1.0, sum(diffs) / len(diffs))


def pair_score(signals: SheetSignals, a: Region, b: Region, config: dict[str, Any]) -> RegionEdge:
    box_a = a.box
    box_b = b.box
    max_row_gap = int(config.get("max_row_gap", 8))
    max_col_gap = int(config.get("max_col_gap", 3))
    weights = config.get("weights", {})
    penalties_cfg = config.get("penalties", {})

    col_overlap = interval_overlap_ratio(box_a.col_overlap(box_b), box_a.width, box_b.width)
    row_overlap = interval_overlap_ratio(box_a.row_overlap(box_b), box_a.height, box_b.height)
    row_gap_s = gap_score(box_a.row_gap(box_b), max_row_gap)
    col_gap_s = gap_score(box_a.col_gap(box_b), max_col_gap)
    border_s = border_continuity(signals, box_a, box_b)
    merge_sig_s = merge_signature_similarity(signals, box_a, box_b)
    style_s = style_similarity(a, b)
    density_s = density_similarity(a, b)

    components = {
        "column_overlap": round(col_overlap, 4),
        "row_overlap": round(row_overlap, 4),
        "row_gap": round(row_gap_s, 4),
        "col_gap": round(col_gap_s, 4),
        "border_continuity": round(border_s, 4),
        "merge_signature_similarity": round(merge_sig_s, 4),
        "style_similarity": round(style_s, 4),
        "density_similarity": round(density_s, 4),
    }

    score = sum(float(weights.get(k, 0)) * v for k, v in components.items())

    penalties = 0.0
    if box_a.row_gap(box_b) > max_row_gap or box_a.col_gap(box_b) > max_col_gap:
        penalties += float(penalties_cfg.get("large_blank_gap", 0.0))
    if col_overlap == 0 and row_overlap == 0:
        penalties += float(penalties_cfg.get("low_axis_overlap", 0.0))
    width_ratio = min(box_a.width, box_b.width) / max(box_a.width, box_b.width, 1)
    if width_ratio < 0.5:
        penalties += float(penalties_cfg.get("different_table_width", 0.0))
    score = clamp01(score - penalties)
    components["penalty"] = round(penalties, 4)

    return RegionEdge(a.id, b.id, round(score, 4), components)
