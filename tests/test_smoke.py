from __future__ import annotations

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

from excel_info_region.config import load_config
from excel_info_region.extractor import extract_workbook_info_regions


def test_load_packaged_default_config():
    cfg = load_config()
    assert cfg["extract_embedded_images"] is True
    assert cfg["extract_chart_images"] is True


def test_extract_synthetic_demo():
    cfg = load_config(PROJECT_ROOT / "config/default.json")
    result = extract_workbook_info_regions(
        PROJECT_ROOT / "examples/synthetic_demo.xlsx",
        sheet_name="Synthetic Demo",
        config=cfg,
    )
    regions = result["sheets"]["Synthetic Demo"]["info_regions"]
    assert regions
    assert any(r.startswith("A1:") for r in regions)
