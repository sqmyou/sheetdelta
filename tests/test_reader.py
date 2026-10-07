"""Tests for reading .xlsx files and pulling references out of formulas."""

from __future__ import annotations

import pytest

from sheetdiff.model import CellKind, CellRef, RangeRef, column_letter, column_number
from sheetdiff.reader import read_workbook, shift_formula
from sheetdiff.references import extract_references

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook


def test_reads_values_formulas_and_shared_strings(tmp_path):
    path = str(tmp_path / "book.xlsx")
    write_workbook(
        path,
        [
            FakeSheet(
                "Sheet1",
                [
                    FakeCell("A1", value="0", is_string=True),
                    FakeCell("B1", value="10"),
                    FakeCell("B2", value="20"),
                    FakeCell("B3", formula="SUM(B1:B2)", value="30"),
                ],
            )
        ],
        shared_strings=["Revenue"],
    )

    workbook = read_workbook(path)
    sheet = workbook.sheet_by_name("Sheet1")
    assert sheet is not None
    assert sheet.cells[CellRef("Sheet1", 1, 1)].cached_value == "Revenue"
    assert sheet.cells[CellRef("Sheet1", 2, 1)].cached_value == "10"
    total = sheet.cells[CellRef("Sheet1", 2, 3)]
    assert total.kind is CellKind.FORMULA
    assert total.formula == "SUM(B1:B2)"
    assert total.cached_value == "30"


def test_reads_inline_string_cells(tmp_path):
    """openpyxl writes strings as <is><t> with no <v>; they must not be dropped."""
    path = str(tmp_path / "book.xlsx")
    write_workbook(
        path,
        [
            FakeSheet(
                "Sheet1",
                [
                    FakeCell("A1", value="Region", inline_string=True),
                    FakeCell("B1", value="42"),
                ],
            )
        ],
    )
    workbook = read_workbook(path)
    sheet = workbook.sheet_by_name("Sheet1")
    assert sheet is not None
    assert sheet.cells[CellRef("Sheet1", 1, 1)].cached_value == "Region"
    assert sheet.cells[CellRef("Sheet1", 2, 1)].cached_value == "42"


def test_reads_defined_names(tmp_path):
    path = str(tmp_path / "book.xlsx")
    write_workbook(
        path,
        [FakeSheet("Sheet1", [FakeCell("A1", value="1")])],
        defined_names={"Revenue": "Sheet1!$A$1"},
    )
    workbook = read_workbook(path)
    assert workbook.defined_names == {"revenue": "Sheet1!$A$1"}


def test_shared_formula_is_shifted_per_cell(tmp_path):
    """A shared formula block must be resolved with relative offsets applied."""
    path = str(tmp_path / "book.xlsx")
    write_workbook(
        path,
        [
            FakeSheet(
                "Sheet1",
                [
                    FakeCell("A1", value="1"),
                    FakeCell("A2", value="2"),
                    FakeCell("B1", formula="A1*2", value="2", shared="0"),
                    FakeCell("B2", value="4", shared="0"),
                ],
            )
        ],
    )
    workbook = read_workbook(path)
    sheet = workbook.sheet_by_name("Sheet1")
    assert sheet is not None
    assert sheet.cells[CellRef("Sheet1", 2, 2)].formula == "A2*2"


def test_date_style_renders_iso(tmp_path):
    path = str(tmp_path / "book.xlsx")
    write_workbook(
        path,
        [FakeSheet("Sheet1", [FakeCell("A1", value="44197", style=0)])],
        date_styles=(0,),
    )
    workbook = read_workbook(path)
    sheet = workbook.sheet_by_name("Sheet1")
    assert sheet is not None
    assert sheet.cells[CellRef("Sheet1", 1, 1)].cached_value == "2021-01-01"


def test_unsupported_format_is_clear(tmp_path):
    path = str(tmp_path / "old.xls")
    path_obj = tmp_path / "old.xls"
    path_obj.write_bytes(b"not really a workbook")
    from sheetdiff.errors import UnsupportedFormatError

    with pytest.raises(UnsupportedFormatError, match="not supported"):
        read_workbook(path)


def test_missing_file_is_clear(tmp_path):
    from sheetdiff.errors import WorkbookReadError

    with pytest.raises(WorkbookReadError, match="no such file"):
        read_workbook(str(tmp_path / "nope.xlsx"))


def test_not_a_zip_is_clear(tmp_path):
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"this is not a zip")
    from sheetdiff.errors import WorkbookReadError

    with pytest.raises(WorkbookReadError, match="not a readable"):
        read_workbook(str(path))


class TestReferences:
    def test_simple_cell_and_range(self):
        refs = extract_references("SUM(B1:B2)", sheet="S", sheets={"s": "S"})
        assert refs == {RangeRef("S", 2, 2, 1, 2)}

    def test_function_name_is_not_a_reference(self):
        """LOG10 looks exactly like column LOG row 10 if you only use a regex."""
        refs = extract_references("LOG10(A1)", sheet="S", sheets={"s": "S"})
        assert refs == {CellRef("S", 1, 1)}

    def test_reference_inside_text_is_ignored(self):
        refs = extract_references('IF(A1>0,"see B2",0)', sheet="S", sheets={"s": "S"})
        assert refs == {CellRef("S", 1, 1)}

    def test_quoted_sheet_name(self):
        refs = extract_references("'My Sheet'!A1", sheet="S", sheets={"s": "S", "my sheet": "My Sheet"})
        assert refs == {CellRef("My Sheet", 1, 1)}

    def test_sheet_name_is_case_insensitive(self):
        refs = extract_references("sheet1!A1", sheet="S", sheets={"sheet1": "Sheet1"})
        assert refs == {CellRef("Sheet1", 1, 1)}

    def test_absolute_markers_are_ignored(self):
        refs = extract_references("$A$1+$B1+C$1", sheet="S", sheets={"s": "S"})
        assert refs == {CellRef("S", 1, 1), CellRef("S", 2, 1), CellRef("S", 3, 1)}

    def test_whole_column_range_is_not_expanded(self):
        refs = extract_references("SUM(A:A)", sheet="S", sheets={"s": "S"})
        assert len(refs) == 1
        assert isinstance(next(iter(refs)), RangeRef)

    def test_defined_name_expands_to_its_range(self):
        refs = extract_references(
            "SUM(Revenue)", sheet="S", sheets={"s": "S"}, defined_names={"revenue": "S!$A$1:$A$3"}
        )
        assert refs == {RangeRef("S", 1, 1, 1, 3)}

    def test_out_of_grid_address_is_rejected(self):
        """A word like ZZZ9999999 is not a cell; it must not enter the graph."""
        refs = extract_references("SUM(ZZZZ1)", sheet="S", sheets={"s": "S"})
        assert refs == frozenset()


class TestShift:
    def test_relative_reference_moves(self):
        assert shift_formula("A1*2", CellRef("S", 2, 1), CellRef("S", 2, 2)) == "A2*2"

    def test_absolute_reference_stays(self):
        assert shift_formula("$A$1*2", CellRef("S", 2, 1), CellRef("S", 2, 5)) == "$A$1*2"

    def test_mixed_reference_moves_only_relative_half(self):
        assert shift_formula("A$1+$B1", CellRef("S", 1, 1), CellRef("S", 3, 4)) == "C$1+$B4"

    def test_text_is_not_shifted(self):
        assert shift_formula('"A1"&B1', CellRef("S", 2, 1), CellRef("S", 2, 2)) == '"A1"&B2'

    def test_shift_past_the_grid_becomes_ref_error(self):
        assert shift_formula("A1", CellRef("S", 1, 1), CellRef("S", 1, 0)) == "#REF!"


def test_column_letters_round_trip():
    for number in (1, 26, 27, 52, 703, 16384):
        assert column_number(column_letter(number)) == number
