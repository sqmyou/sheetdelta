"""Tests for CellIndex, the lookup the diff and audit both build."""

from __future__ import annotations

import pytest

from sheetdelta.model import CellIndex, CellRef, RangeRef


def _cells(*addresses: tuple[str, int, int]) -> list[CellRef]:
    return [CellRef(sheet, col, row) for sheet, col, row in addresses]


def test_single_cell_reference_finds_only_that_cell():
    index = CellIndex(_cells(("S", 1, 1), ("S", 1, 2), ("S", 2, 1)))
    assert index.covered(CellRef("S", 1, 1)) == [CellRef("S", 1, 1)]


def test_single_cell_reference_to_an_empty_address_finds_nothing():
    index = CellIndex(_cells(("S", 1, 1)))
    assert index.covered(CellRef("S", 5, 9)) == []


def test_range_finds_the_block_it_covers():
    index = CellIndex(_cells(("S", 1, 1), ("S", 2, 2), ("S", 3, 3), ("S", 9, 9)))
    covered = index.covered(RangeRef("S", 1, 2, 1, 2))
    assert set(covered) == {CellRef("S", 1, 1), CellRef("S", 2, 2)}


def test_range_is_bounded_on_all_four_sides():
    index = CellIndex(_cells(("S", 1, 1), ("S", 5, 1), ("S", 1, 5)))
    # Columns 1-3, rows 1-3: only the corner cell is inside.
    assert index.covered(RangeRef("S", 1, 3, 1, 3)) == [CellRef("S", 1, 1)]


def test_whole_column_range_only_visits_rows_that_hold_cells():
    # A formula like SUM(A:A) spans a million rows. Walking them all would be
    # the same quadratic cost the index exists to avoid, so the range must cost
    # the cells it actually touches.
    index = CellIndex(_cells(("S", 1, 1), ("S", 1, 7), ("S", 2, 3)))
    covered = index.covered(RangeRef("S", 1, 1, 1, 1_048_576))
    assert set(covered) == {CellRef("S", 1, 1), CellRef("S", 1, 7)}


def test_references_do_not_leak_across_sheets():
    index = CellIndex(_cells(("One", 1, 1), ("Two", 1, 1)))
    assert index.covered(CellRef("One", 1, 1)) == [CellRef("One", 1, 1)]
    assert index.covered(RangeRef("Two", 1, 9, 1, 9)) == [CellRef("Two", 1, 1)]


@pytest.mark.parametrize(
    "reference",
    [
        CellRef("S", 2, 3),
        CellRef("S", 9, 9),
        RangeRef("S", 1, 1, 1, 1),
        RangeRef("S", 1, 4, 1, 4),
        RangeRef("S", 2, 2, 3, 3),
        RangeRef("S", 1, 4, 1, 1_000_000),
    ],
)
def test_matches_a_brute_force_scan(reference):
    """The index must agree with the linear scan it replaced, exactly."""
    cells = _cells(
        ("S", 1, 1), ("S", 2, 1), ("S", 3, 2), ("S", 4, 4), ("S", 2, 3), ("S", 1, 9)
    )
    index = CellIndex(cells)
    expected = {ref for ref in cells if reference.contains(ref)}
    assert set(index.covered(reference)) == expected
