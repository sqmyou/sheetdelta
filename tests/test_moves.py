"""Tests for row and column reorder detection.

A reorder is written the way Excel writes one: every cell of the moved line is
rewritten to a new address, while the lines in between are untouched. Reporting
that as "every cell in these rows changed" buries the fact that only position
moved.
"""

from __future__ import annotations

from sheetdelta.differ import diff_workbooks
from sheetdelta.reader import read_workbook
from sheetdelta.report import render_json, render_markdown, render_summary, render_text

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook


def _book(tmp_path, name, rows):
    path = str(tmp_path / name)
    write_workbook(path, [FakeSheet("S", _grid(rows))])
    return read_workbook(path)


def _grid(rows: list[list[str]]) -> list[FakeCell]:
    cells = []
    for row_number, values in enumerate(rows, start=1):
        for column, value in enumerate(values):
            cells.append(FakeCell(f"{chr(65 + column)}{row_number}", value=value))
    return cells


def _diff(tmp_path, old_rows, new_rows):
    return diff_workbooks(
        _book(tmp_path, "old.xlsx", old_rows),
        _book(tmp_path, "new.xlsx", new_rows),
    )


HEADER = [["name", "qty"], ["a", "1"], ["b", "2"], ["c", "3"], ["d", "4"]]


def test_a_moved_row_is_reported_as_a_move(tmp_path):
    # Dragging row 4 up to row 2 also moves rows 2 and 3 down, so the whole
    # cycle is a move -- but none of it is reported cell by cell, which is
    # the point: three moves, not a dozen changed cells.
    new_rows = [HEADER[0], HEADER[3], HEADER[1], HEADER[2], HEADER[4]]
    result = _diff(tmp_path, HEADER, new_rows)
    sheet = result.sheet_changes[0]
    assert sorted(m.label for m in sheet.moves) == [
        "row 2 moved to 3",
        "row 3 moved to 4",
        "row 4 moved to 2",
    ]
    assert sheet.cell_changes == []


def test_a_swap_is_two_moves(tmp_path):
    new_rows = [HEADER[0], HEADER[2], HEADER[1], HEADER[3], HEADER[4]]
    result = _diff(tmp_path, HEADER, new_rows)
    moves = sorted((m.old, m.new) for m in result.sheet_changes[0].moves)
    assert moves == [(2, 3), (3, 2)]


def test_an_insert_is_not_mistaken_for_a_move(tmp_path):
    # This is the case that rules out a looser match: an inserted row pushes
    # the rows below it down, so every one of them "matches" a row further
    # down. Reading that as a reorder would report five moves for one insert,
    # so a move is only claimed when the rows are a clean permutation.
    old_rows = [["1"], ["2"], ["3"], ["4"], ["5"]]
    new_rows = [["1"], ["2"], ["new"], ["3"], ["4"], ["5"]]
    result = _diff(tmp_path, old_rows, new_rows)
    sheet = result.sheet_changes[0]
    assert sheet.moves == []
    assert len(sheet.shifts) == 1


def test_a_reorder_with_an_edit_elsewhere_is_left_to_the_cell_diff(tmp_path):
    # A reorder plus any unrelated edit breaks the permutation, so no move is
    # claimed. The change is still reported, just not as a move -- the tool
    # never guesses a move it cannot prove.
    old_rows = [["H"], ["a"], ["b"], ["c"], ["d"], ["e"]]
    new_rows = [["H"], ["c"], ["a"], ["b"], ["d"], ["E"]]
    result = _diff(tmp_path, old_rows, new_rows)
    sheet = result.sheet_changes[0]
    assert sheet.moves == []
    assert any(c.ref.a1 == "A6" and c.new == "E" for c in sheet.cell_changes)


def test_an_edited_row_is_not_called_a_move(tmp_path):
    new_rows = [HEADER[0], HEADER[1], HEADER[2], HEADER[3], ["d", "9"]]
    result = _diff(tmp_path, HEADER, new_rows)
    sheet = result.sheet_changes[0]
    assert sheet.moves == []
    assert any(c.ref.a1 == "B5" for c in sheet.cell_changes)


def test_identical_rows_are_not_matched_ambiguously(tmp_path):
    # Two identical rows on each side cannot be told apart, so no move is
    # claimed -- guessing would re-key a real edit away.
    old_rows = [["x"], ["a"], ["a"]]
    new_rows = [["x"], ["a"], ["a", "changed"]]
    result = _diff(tmp_path, old_rows, new_rows)
    assert result.sheet_changes[0].moves == []


def test_a_pure_reorder_has_no_breaking_change(tmp_path):
    new_rows = [HEADER[0], HEADER[1], HEADER[3], HEADER[2], HEADER[4]]
    result = _diff(tmp_path, HEADER, new_rows)
    assert result.has_changes
    assert result.breaking == []


def test_a_blank_row_does_not_hide_a_move(tmp_path):
    # A row with no cells at all is not a row signature, so an insert that
    # leaves a gap cannot make the remaining rows look like a permutation.
    old_rows = [["a"], ["b"], [], ["c"]]
    new_rows = [["c"], ["a"], [], ["b"]]
    result = _diff(tmp_path, old_rows, new_rows)
    assert result.sheet_changes[0].moves != []


def test_report_surfaces_the_move(tmp_path):
    new_rows = [HEADER[0], HEADER[3], HEADER[1], HEADER[2], HEADER[4]]
    result = _diff(tmp_path, HEADER, new_rows)
    assert "row 4 moved to 2" in render_text(result)
    assert "row 4 moved to 2" in render_summary(result)
    assert "row 4 moved to 2" in render_markdown(result)
    assert '"moves"' in render_json(result)


def test_a_moved_column_is_reported_as_a_move(tmp_path):
    # The same problem on the other axis: dragging column C left of B rewrites
    # those columns' addresses, and without a move this reads as every cell in
    # columns B and C changing. The rows here hold distinct values so the sheet
    # is a clean column permutation, not a row one.
    old_rows = [["h1", "b1", "c1", "d1"], ["h2", "b2", "c2", "d2"], ["h3", "b3", "c3", "d3"]]
    new_rows = [["h1", "c1", "b1", "d1"], ["h2", "c2", "b2", "d2"], ["h3", "c3", "b3", "d3"]]
    result = _diff(tmp_path, old_rows, new_rows)
    sheet = result.sheet_changes[0]
    moves = sorted((m.axis, m.old, m.new) for m in sheet.moves)
    assert moves == [("column", 2, 3), ("column", 3, 2)]
    assert sheet.cell_changes == []


def test_a_column_move_is_not_a_row_move(tmp_path):
    # Guards the axis choice: a rotated block of rows must not be re-keyed as
    # columns just because some columns happen to share a signature.
    old_rows = [["a", "1"], ["b", "2"], ["c", "3"]]
    new_rows = [["c", "3"], ["a", "1"], ["b", "2"]]
    result = _diff(tmp_path, old_rows, new_rows)
    assert all(m.axis == "row" for m in result.sheet_changes[0].moves)


def test_column_move_report_and_json_carry_the_axis(tmp_path):
    old_rows = [["h1", "b1", "c1"], ["h2", "b2", "c2"], ["h3", "b3", "c3"]]
    new_rows = [["h1", "c1", "b1"], ["h2", "c2", "b2"], ["h3", "c3", "b3"]]
    result = _diff(tmp_path, old_rows, new_rows)
    assert "column 3 moved to 2" in render_text(result)
    assert "column 3 moved to 2" in render_markdown(result)
    assert '"axis": "column"' in render_json(result)
