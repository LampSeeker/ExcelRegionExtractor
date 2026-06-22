from __future__ import annotations

from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont

from openpyxl.worksheet.worksheet import Worksheet

from .chart_regions import chart_metadata
from .io import ensure_dir


def _remove_empty_dir(path: Path) -> None:
    try:
        path.rmdir()
    except OSError:
        pass


def _safe_filename_part(text: str) -> str:
    return "".join(ch if ch.isalnum() or ch in ".()-_" else "_" for ch in text).strip("_") or "chart"


def _flat(values: Any) -> list[Any]:
    if not values:
        return []
    if isinstance(values, list) and values and isinstance(values[0], list):
        return values[0]
    return list(values)


def render_chart_preview(chart: dict[str, Any]) -> Image.Image | None:
    cat = next((s for s in chart["sources"] if s["role"] == "cat"), None)
    val = next((s for s in chart["sources"] if s["role"] == "val"), None)
    if not cat or not val:
        return None

    labels = [str(v) if v is not None else "" for v in _flat(cat.get("cached_values") or cat.get("values"))]
    values = [float(v or 0) for v in _flat(val.get("cached_values") or val.get("values"))]
    if not labels or not values:
        return None

    width, height = 900, 420
    margin_l, margin_t, margin_r, margin_b = 70, 50, 30, 110
    plot_w = width - margin_l - margin_r
    plot_h = height - margin_t - margin_b
    max_v = max(values) or 1.0

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()

    title = chart.get("name") or "Chart"
    draw.text((margin_l, 18), title, fill=(30, 30, 30), font=font)
    for i in range(6):
        y = margin_t + int(plot_h * i / 5)
        draw.line((margin_l, y, margin_l + plot_w, y), fill=(220, 220, 220), width=1)
    draw.line((margin_l, margin_t + plot_h, margin_l + plot_w, margin_t + plot_h), fill=(80, 80, 80), width=2)
    draw.line((margin_l, margin_t, margin_l, margin_t + plot_h), fill=(80, 80, 80), width=2)

    palette = [
        (237, 125, 49),
        (112, 48, 160),
        (155, 187, 89),
        (255, 192, 0),
        (91, 155, 213),
        (165, 165, 165),
        (68, 114, 196),
        (112, 173, 71),
    ]
    n = len(values)
    slot = plot_w / max(1, n)
    bar_w = max(8, int(slot * 0.62))
    for idx, value in enumerate(values):
        x0 = int(margin_l + idx * slot + (slot - bar_w) / 2)
        x1 = x0 + bar_w
        bar_h = int((value / max_v) * plot_h)
        y0 = margin_t + plot_h - bar_h
        y1 = margin_t + plot_h
        color = palette[idx % len(palette)]
        draw.rectangle((x0, y0, x1, y1), fill=color, outline=(70, 70, 70))
        draw.text((x0, y0 - 14), f"{value:g}", fill=(40, 40, 40), font=font)
        label = labels[idx][:14] if idx < len(labels) else ""
        draw.text((x0, y1 + 8), label, fill=(40, 40, 40), font=font)

    return image


def extract_sheet_charts_to_dir(
    ws: Worksheet,
    out_dir: str | Path,
    *,
    rel_dir: str = "charts",
    config: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    chart_dir = Path(out_dir) / rel_dir
    if config is not None and not config.get("extract_chart_images", True):
        _remove_empty_dir(chart_dir)
        return chart_metadata(ws, config)

    output: list[dict[str, Any]] = []
    for idx, chart in enumerate(chart_metadata(ws, config or {}), 1):
        item = dict(chart)
        filename = f"CHART{idx:03d}_{_safe_filename_part(chart['range_ref'])}_{_safe_filename_part(chart['name'])}.png"
        if chart.get("kind") == "BarChart":
            image = render_chart_preview(chart)
            if image is not None:
                ensure_dir(chart_dir)
                image.save(chart_dir / filename)
                item["path"] = f"{rel_dir}/{filename}"
        output.append(item)
    if not any("path" in item for item in output):
        _remove_empty_dir(chart_dir)
    return output
