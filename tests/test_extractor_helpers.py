from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from excel_info_region.borders import boxes_are_contact_merge_neighbors
from excel_info_region.chart_regions import chart_boxes, chart_metadata
from excel_info_region.components import (
    connected_components_from_cells,
    remove_exact_or_contained_duplicates,
)
from excel_info_region.extractor import extract_info_regions_from_sheet
from excel_info_region.schema import Box


def test_connected_components_respects_connectivity():
    occupied = {(1, 1), (2, 2)}

    assert [b.range_ref for b in connected_components_from_cells(occupied, connectivity=4)] == [
        "A1:A1",
        "B2:B2",
    ]
    assert [b.range_ref for b in connected_components_from_cells(occupied, connectivity=8)] == [
        "A1:B2",
    ]


def test_border_contact_neighbors_need_gap_and_overlap():
    cfg = {
        "border_contact_merge_max_gap": 1,
        "border_contact_merge_min_axis_overlap": 0.8,
    }

    assert boxes_are_contact_merge_neighbors(Box(1, 1, 2, 4), Box(4, 1, 5, 4), cfg)
    assert not boxes_are_contact_merge_neighbors(Box(1, 1, 2, 4), Box(5, 1, 6, 4), cfg)
    assert not boxes_are_contact_merge_neighbors(Box(1, 1, 2, 4), Box(4, 4, 5, 7), cfg)


def test_remove_contained_duplicates_is_config_gated():
    boxes = [Box(1, 1, 4, 4), Box(2, 2, 3, 3), Box(1, 1, 4, 4)]

    assert remove_exact_or_contained_duplicates(boxes, {"remove_contained_duplicates": False}) == [
        Box(1, 1, 4, 4),
        Box(2, 2, 3, 3),
    ]
    assert remove_exact_or_contained_duplicates(boxes, {"remove_contained_duplicates": True}) == [
        Box(1, 1, 4, 4),
    ]


def test_respect_hidden_rows_cols_excludes_hidden_cells():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "visible"
    ws["B1"] = "hidden"
    ws.column_dimensions["B"].hidden = True

    result = extract_info_regions_from_sheet(
        ws,
        config={
            "respect_hidden_rows_cols": True,
            "include_images": False,
            "use_borders": False,
            "use_border_contact_merge": False,
        },
    )

    assert result["info_regions"] == ["A1:A1"]


def test_use_print_area_bounds_limits_regions():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "inside"
    ws["Z1"] = "outside"
    ws.print_area = "A1:B2"

    result = extract_info_regions_from_sheet(
        ws,
        config={
            "use_print_area_bounds": True,
            "include_images": False,
            "use_borders": False,
            "use_border_contact_merge": False,
        },
    )

    assert result["info_regions"] == ["A1:A1"]


def test_chart_metadata_includes_anchor_and_sources():
    wb = Workbook()
    ws = wb.active
    ws.append(["Label", "Score"])
    ws.append(["A", 1])
    ws.append(["B", 2])

    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=3))
    ws.add_chart(chart, "D2")

    charts = chart_metadata(ws, {})

    assert chart_boxes(ws, {})[0].range_ref.startswith("D2:")
    assert charts[0]["kind"] == "BarChart"
    assert charts[0]["range_ref"].startswith("D2:")
    assert {source["role"] for source in charts[0]["sources"]} == {"cat", "val"}
