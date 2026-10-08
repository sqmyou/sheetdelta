"""Tests for Excel tables and structured references."""

from __future__ import annotations

from sheetdelta.differ import diff_workbooks
from sheetdelta.model import RangeRef, Reference, TableRef
from sheetdelta.reader import read_workbook
from sheetdelta.references import extract_references

from .xlsx_fixtures import FakeCell, FakeSheet, FakeTable, write_workbook

SHEETS = {"s": "S"}


def _refs(formula: str, tables: dict[str, str] | None = None) -> frozenset[Reference]:
    return extract_references(
        formula, sheet="S", sheets=SHEETS, defined_names={}, tables=tables or {}
    )


def test_structured_reference_is_parsed():
    refs = _refs("SUM(Table1[Amount])", {"table1": "Table1"})
    assert TableRef("Table1", "Amount", None) in refs


def test_a_name_that_is_not_a_table_is_not_a_structured_reference():
    # The same text with no table of that name must not produce a table ref,
    # otherwise any function-like word would look like one.
    assert _refs("SUM(Table1[Amount])") == frozenset()


def test_specifier_forms_are_parsed():
    tables = {"table1": "Table1"}
    assert TableRef("Table1", None, "#All") in _refs("SUM(Table1[#All])", tables)
    assert TableRef("Table1", None, "#Headers") in _refs("SUM(Table1[#Headers])", tables)
    assert TableRef("Table1", None, "#Totals") in _refs("SUM(Table1[#Totals])", tables)


def test_combined_specifier_and_column_is_parsed():
    refs = _refs("SUM(Table1[[#Totals],[Amount]])", {"table1": "Table1"})
    assert TableRef("Table1", "Amount", "#Totals") in refs


def test_current_row_form_keeps_the_column():
    refs = _refs("Table1[@Amount]*2", {"table1": "Table1"})
    assert TableRef("Table1", "Amount", None) in refs


def _table_sheet(cells, table):
    return [FakeSheet("S", cells, tables=[table])]


def _read(tmp_path, name, cells, table):
    path = str(tmp_path / name)
    write_workbook(path, _table_sheet(cells, table))
    return read_workbook(path)


TABLE = FakeTable("Table1", "A1:B4", ["Item", "Amount"])
TABLE_CELLS = [
    FakeCell("A1", value="Item"),
    FakeCell("B1", value="Amount"),
    FakeCell("A2", value="a"),
    FakeCell("B2", value="10"),
    FakeCell("A3", value="b"),
    FakeCell("B3", value="20"),
    FakeCell("A4", value="c"),
    FakeCell("B4", value="30"),
]


def test_reader_records_the_table(tmp_path):
    workbook = _read(tmp_path, "t.xlsx", TABLE_CELLS, TABLE)
    table = workbook.tables["table1"]
    assert table.name == "Table1"
    assert table.columns == ("Item", "Amount")
    assert (table.min_col, table.max_col) == (1, 2)
    assert (table.min_row, table.max_row) == (1, 4)


def test_a_column_resolves_to_its_grid_column_excluding_the_header(tmp_path):
    workbook = _read(tmp_path, "t.xlsx", TABLE_CELLS, TABLE)
    resolved = TableRef("Table1", "Amount", None).resolve(workbook.tables)
    assert resolved == RangeRef("S", 2, 2, 2, 4)


def test_column_lookup_is_case_insensitive(tmp_path):
    workbook = _read(tmp_path, "t.xlsx", TABLE_CELLS, TABLE)
    resolved = TableRef("Table1", "amount", None).resolve(workbook.tables)
    assert resolved == RangeRef("S", 2, 2, 2, 4)


def test_totals_row_is_excluded_from_the_data_body(tmp_path):
    table = FakeTable("Table1", "A1:B5", ["Item", "Amount"], totals_row=True)
    cells = TABLE_CELLS + [FakeCell("A5", value="Total"), FakeCell("B5", value="60")]
    workbook = _read(tmp_path, "t.xlsx", cells, table)
    assert TableRef("Table1", "Amount", None).resolve(workbook.tables) == RangeRef(
        "S", 2, 2, 2, 4
    )
    assert TableRef("Table1", None, "#Totals").resolve(workbook.tables) == RangeRef(
        "S", 1, 2, 5, 5
    )


def test_an_unknown_column_does_not_resolve(tmp_path):
    workbook = _read(tmp_path, "t.xlsx", TABLE_CELLS, TABLE)
    assert TableRef("Table1", "Nope", None).resolve(workbook.tables) is None


def test_a_structured_reference_carries_impact_to_the_table_column(tmp_path):
    """The whole point: editing a table cell reaches the formula that reads it."""
    old = _read(tmp_path, "old.xlsx", TABLE_CELLS, TABLE)
    changed = [c for c in TABLE_CELLS if c.ref != "B3"]
    changed.append(FakeCell("B3", value="99"))
    changed.append(FakeCell("D1", formula="SUM(Table1[Amount])", value="60"))
    new = _read(tmp_path, "new.xlsx", changed, TABLE)

    result = diff_workbooks(old, new)
    impacted = [c for c in result.cell_changes if c.ref.a1 == "B3"]
    assert len(impacted) == 1
    assert "D1" in {ref.a1 for ref in impacted[0].affected}
