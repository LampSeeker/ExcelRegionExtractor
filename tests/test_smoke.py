from __future__ import annotations

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

from excel_info_region.config import load_config
from excel_info_region.extractor import extract_workbook_info_regions


def test_extract_sample_manhole():
    cfg = load_config(PROJECT_ROOT / "config/default.json")
    result = extract_workbook_info_regions(
        PROJECT_ROOT / "examples/sample.xlsx",
        sheet_name="각형맨홀(특2호)",
        config=cfg,
    )
    regions = result["sheets"]["각형맨홀(특2호)"]["info_regions"]
    assert regions
    assert any(r.startswith("A1:") for r in regions)
