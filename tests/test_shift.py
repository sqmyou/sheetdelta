"""Tests for row and column insert/remove detection.

The fixtures build a sheet from a grid of values, so an insert is written the
way Excel writes one: the new line is added and every line below it moves down
with its contents intact.
"""

from __future__ import annotations

from sheetdelta.differ import diff_workbooks
from sheetdelta.reader import read_workbook

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook


def _book(tmp_path, name, sheets, **kwargs):
    path = str(tmp_path / name)
    write_workbook(path, sheets, **kwargs)
    return read_workbook(path)


def _grid(rows: list[list[str]]) -> list[FakeCell]:
    """Place each list of values on its own row, starting at A1."""
    cells = []
    for row_number, values in enumerate(rows, start=1):
        for column, value in enumerate(values):
            cells.append(FakeCell(f"{chr(65 + column)}{row_number}", value=value))
    return cells


def _sheet(rows: list[list[str]]) -> list[FakeSheet]:
    return [FakeSheet("S", _grid(rows))]


def _inserted_row(shift):
    assert shift is not None, "expected a row shift to be detected"
    assert shift.axis == "row"
    assert shift.inserted is True
    return shift


def test_inserted_row_is_reported_once_not_as_every_row_below(tmp_path):
    old = _book(tmp_path, "old.xlsx", _sheet([["1"], ["2"], ["3"], ["4"], ["5"]]))
    new = _book(
        tmp_path, "new.xlsx", _sheet([["1"], ["2"], ["new"], ["3"], ["4"], ["5"]])
    )

    sheet = diff_workbooks(old, new).sheet_changes[0]
    shift = _inserted_row(sheet.shift)

    assert shift.count == 1
    assert shift.at == 3
    # Only the row that was actually added is left over; the rest lined up.
    assert [change.ref.a1 for change in sheet.cell_changes] == ["A3"]


def test_inserted_rows_report_the_count_and_position(tmp_path):
    old = _book(tmp_path, "old.xlsx", _sheet([["1"], ["2"], ["3"], ["4"], ["5"]]))
    new = _book(
        tmp_path, "new.xlsx", _sheet([["1"], ["2"], ["x"], ["y"], ["3"], ["4"], ["5"]])
    )

    shift = _inserted_row(diff_workbooks(old, new).sheet_changes[0].shift)
    assert shift.count == 2
    assert shift.at == 3
    assert shift.label == "2 rows inserted at row 3"


def test_removed_row_is_reported_as_a_removal(tmp_path):
    old = _book(tmp_path, "old.xlsx", _sheet([["1"], ["2"], ["3"], ["4"], ["5"]]))
    new = _book(tmp_path, "new.xlsx", _sheet([["1"], ["2"], ["4"], ["5"]]))

    shift = diff_workbooks(old, new).sheet_changes[0].shift
    assert shift is not None
    assert shift.inserted is False
    assert shift.count == 1
    assert shift.at == 3
    assert shift.label == "1 row removed at row 3"


def test_a_real_edit_below_the_insert_is_still_reported(tmp_path):
    """The insert must not swallow a genuine change that moved with it."""
    old = _book(tmp_path, "old.xlsx", _sheet([["1"], ["2"], ["3"], ["4"]]))
    new = _book(tmp_path, "new.xlsx", _sheet([["1"], ["2"], ["new"], ["3"], ["edited"]]))

    sheet = diff_workbooks(old, new).sheet_changes[0]
    _inserted_row(sheet.shift)
    edits = [c for c in sheet.cell_changes if c.ref.a1 == "A5"]
    assert len(edits) == 1
    assert edits[0].new == "edited"


def test_inserted_column_is_reported_on_the_column_axis(tmp_path):
    old = _book(tmp_path, "old.xlsx", _sheet([["1", "2", "3"], ["4", "5", "6"]]))
    new = _book(tmp_path, "new.xlsx", _sheet([["1", "X", "2", "3"], ["4", "Y", "5", "6"]]))

    shift = diff_workbooks(old, new).sheet_changes[0].shift
    assert shift is not None
    assert shift.axis == "column"
    assert shift.inserted is True
    assert shift.count == 1
    assert shift.at == 2


def test_insert_at_the_very_top(tmp_path):
    old = _book(tmp_path, "old.xlsx", _sheet([["1"], ["2"], ["3"]]))
    new = _book(tmp_path, "new.xlsx", _sheet([["top"], ["1"], ["2"], ["3"]]))

    shift = _inserted_row(diff_workbooks(old, new).sheet_changes[0].shift)
    assert shift.at == 1
    assert shift.count == 1


def test_a_longer_sheet_is_not_called_a_shift(tmp_path):
    """Two sheets with unrelated contents must not be aligned by force."""
    old = _book(tmp_path, "old.xlsx", _sheet([["1"], ["2"], ["3"], ["4"]]))
    new = _book(tmp_path, "new.xlsx", _sheet([["1"], ["2"], ["3"], ["4"], ["5"]]))

    sheet = diff_workbooks(old, new).sheet_changes[0]
    assert sheet.shift is None


def test_an_edit_without_a_shift_is_not_a_shift(tmp_path):
    old = _book(tmp_path, "old.xlsx", _sheet([["a"], ["b"], ["c"], ["d"]]))
    new = _book(tmp_path, "new.xlsx", _sheet([["a"], ["b"], ["X"], ["d"]]))

    assert diff_workbooks(old, new).sheet_changes[0].shift is None


def test_identical_workbooks_have_no_shift(tmp_path):
    sheets = _sheet([["1"], ["2"], ["3"]])
    result = diff_workbooks(
        _book(tmp_path, "a.xlsx", sheets), _book(tmp_path, "b.xlsx", sheets)
    )
    assert not result.sheet_changes
