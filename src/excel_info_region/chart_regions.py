from __future__ import annotations

from typing import Any

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from .components import dedupe_boxes
from .schema import Box


def _anchor_box(anchor: Any) -> Box | None:
    if isinstance(anchor, str):
        row, col = coordinate_to_tuple(anchor)
        return Box(row, col, row + 10, col + 6)

    start = getattr(anchor, "_from", None)
    if start is None:
        return None

    min_row = int(start.row) + 1
    min_col = int(start.col) + 1
    end = getattr(anchor, "to", None) or getattr(anchor, "_to", None)
    if end is None:
        return Box(min_row, min_col, min_row + 10, min_col + 6)
    return Box(min_row, min_col, max(min_row, int(end.row) + 1), max(min_col, int(end.col) + 1))


def chart_boxes(ws: Worksheet, config: dict[str, Any]) -> list[Box]:
    if not config.get("include_charts", True):
        return []
    return dedupe_boxes([box for chart in getattr(ws, "_charts", []) if (box := _anchor_box(chart.anchor))])


def _local_range_ref(formula: str | None) -> str | None:
    if not formula:
        return None
    ref = formula.split("!", 1)[-1].replace("$", "").strip().strip("'")
    try:
        range_boundaries(ref)
    except ValueError:
        return None
    return ref


def _range_values(ws: Worksheet, range_ref: str) -> list[list[Any]]:
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    return [
        [ws.cell(row, col).value for col in range(min_col, max_col + 1)]
        for row in range(min_row, max_row + 1)
    ]


def _source_formula(obj: Any, attr: str) -> str | None:
    source = getattr(obj, attr, None)
    if source is None:
        return None
    for ref_name in ("numRef", "strRef"):
        ref = getattr(source, ref_name, None)
        formula = getattr(ref, "f", None)
        if formula:
            return formula
    return None


def _source_cache_values(obj: Any, attr: str) -> list[Any]:
    source = getattr(obj, attr, None)
    if source is None:
        return []
    for ref_name, cache_name in (("numRef", "numCache"), ("strRef", "strCache")):
        ref = getattr(source, ref_name, None)
        cache = getattr(ref, cache_name, None)
        points = getattr(cache, "pt", None)
        if points:
            count = int(getattr(cache, "ptCount", 0) or 0)
            if count:
                values = [None] * count
                for point in points:
                    idx = int(point.idx)
                    if 0 <= idx < count:
                        values[idx] = point.v
                return values
            return [point.v for point in sorted(points, key=lambda p: int(p.idx))]
    return []


def _chart_title(chart: Any) -> str | None:
    title = getattr(chart, "title", None)
    try:
        text = title.tx.rich.p[0].r[0].t
    except Exception:
        return None
    return text if text and text != "None" else None


def chart_metadata(ws: Worksheet, config: dict[str, Any]) -> list[dict[str, Any]]:
    if not config.get("include_charts", True):
        return []

    output: list[dict[str, Any]] = []
    include_values = bool(config.get("include_chart_source_values", True))
    for idx, chart in enumerate(getattr(ws, "_charts", []), 1):
        box = _anchor_box(chart.anchor)
        if box is None:
            continue

        sources: list[dict[str, Any]] = []
        seen: set[str] = set()
        for series in getattr(chart, "series", []):
            for role in ("cat", "val"):
                range_ref = _local_range_ref(_source_formula(series, role))
                if range_ref is None or range_ref in seen:
                    continue
                seen.add(range_ref)
                item: dict[str, Any] = {"role": role, "range_ref": range_ref}
                if include_values:
                    item["values"] = _range_values(ws, range_ref)
                    cache_values = _source_cache_values(series, role)
                    if cache_values:
                        item["cached_values"] = cache_values
                sources.append(item)

        output.append({
            "name": _chart_title(chart) or f"Chart {idx}",
            "kind": type(chart).__name__,
            "range_ref": box.range_ref,
            "sources": sources,
        })

    return output
