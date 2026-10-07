"""Read workbooks written by openpyxl, not by our own fixture builder.

The hand-written fixtures in ``xlsx_fixtures`` match what the reader expects.
This file is the check that the reader also matches what a real writer
produces, which is the only thing that matters in practice. It is skipped if
openpyxl is not installed, so the suite still runs with no dependencies.
"""

from __future__ import annotations

import pytest

from sheetdiff.differ import ChangeKind, diff_workbooks
from sheetdiff.model import CellRef, column_number
from sheetdiff.reader import read_workbook

openpyxl = pytest.importorskip("openpyxl")


def ref(sheet: str, a1: str) -> CellRef:
    letters = "".join(ch for ch in a1 if ch.isalpha())
    digits = "".join(ch for ch in a1 if ch.isdigit())
    return CellRef(sheet, column_number(letters), int(digits))


def _build(path, last_row, summary_total):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Revenue"
    sheet["A1"] = "Region"
    sheet["B1"] = "Amount"
    for row in range(2, 12):
        sheet[f"A{row}"] = f"R{row}"
        sheet[f"B{row}"] = 1000 + row
    sheet["D12"] = f"=SUM(B2:B{last_row})"
    sheet["E12"] = "=D12*1.1"
    sheet["F12"] = "=E12-D12"

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Total"
    summary["B4"] = "=Revenue!D12*2"
    summary["B7"] = summary_total
    workbook.save(path)


def test_reads_a_real_workbook(tmp_path):
    path = str(tmp_path / "real.xlsx")
    _build(path, 11, 1200)
    workbook = read_workbook(path)

    assert workbook.sheet_names == ["Revenue", "Summary"]
    revenue = workbook.sheet_by_name("Revenue")
    assert revenue is not None
    assert revenue.cells[ref("Revenue", "D12")].formula == "SUM(B2:B11)"
    assert revenue.cells[ref("Revenue", "A2")].cached_value == "R2"


def test_diffs_a_real_workbook_across_sheets(tmp_path):
    """The scenario from the README, built by a real writer rather than by us."""
    old = str(tmp_path / "old.xlsx")
    new = str(tmp_path / "new.xlsx")
    _build(old, 11, 1200)
    _build(new, 13, 1450)

    result = diff_workbooks(read_workbook(old), read_workbook(new))
    change = next(c for c in result.cell_changes if c.ref.a1 == "D12")

    assert change.kind is ChangeKind.FORMULA
    assert [str(r) for r in change.affected] == ["Revenue!E12", "Revenue!F12", "Summary!B4"]


def test_a_workbook_without_cached_values_is_not_stale(tmp_path):
    """openpyxl writes no cached value; that is uncalculated, not stale."""
    old = str(tmp_path / "old.xlsx")
    new = str(tmp_path / "new.xlsx")
    _build(old, 11, 1200)
    _build(new, 13, 1200)

    result = diff_workbooks(read_workbook(old), read_workbook(new))
    change = next(c for c in result.cell_changes if c.ref.a1 == "D12")
    assert change.kind is ChangeKind.FORMULA
