from __future__ import annotations

from excel_info_region.borders import boxes_are_contact_merge_neighbors
from excel_info_region.components import (
    connected_components_from_cells,
    remove_exact_or_contained_duplicates,
)
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
