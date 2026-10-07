"""Tests for the diff engine and its dependency impact."""

from __future__ import annotations

from sheetdelta.differ import ChangeKind, Severity, diff_workbooks
from sheetdelta.reader import read_workbook

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook


def _book(tmp_path, name, sheets, **kwargs):
    path = str(tmp_path / name)
    write_workbook(path, sheets, **kwargs)
    return read_workbook(path)


def test_identical_workbooks_report_nothing(tmp_path):
    sheets = [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("B1", formula="A1*2", value="2")])]
    result = diff_workbooks(_book(tmp_path, "a.xlsx", sheets), _book(tmp_path, "b.xlsx", sheets))
    assert not result.has_changes
    assert result.breaking == []


def test_formula_change_with_dependents_is_breaking(tmp_path):
    """The core promise: an edit that reaches a total is flagged breaking."""
    old = _book(
        tmp_path,
        "old.xlsx",
        [
            FakeSheet(
                "S",
                [
                    FakeCell("A1", value="1"),
                    FakeCell("A2", value="2"),
                    FakeCell("A3", formula="SUM(A1:A2)", value="3"),
                    FakeCell("B1", formula="A3*10", value="30"),
                ],
            )
        ],
    )
    new = _book(
        tmp_path,
        "new.xlsx",
        [
            FakeSheet(
                "S",
                [
                    FakeCell("A1", value="1"),
                    FakeCell("A2", value="2"),
                    FakeCell("A3", formula="SUM(A1:A2)*2", value="6"),
                    FakeCell("B1", formula="A3*10", value="60"),
                ],
            )
        ],
    )
    result = diff_workbooks(old, new)
    a3 = next(c for c in result.cell_changes if c.ref.a1 == "A3")
    assert a3.kind is ChangeKind.FORMULA
    assert a3.severity is Severity.BREAKING
    assert [str(r) for r in a3.affected] == ["S!B1"]


def test_formula_change_with_no_dependents_is_a_warning(tmp_path):
    old = _book(tmp_path, "old.xlsx", [FakeSheet("S", [FakeCell("A1", formula="1+1", value="2")])])
    new = _book(tmp_path, "new.xlsx", [FakeSheet("S", [FakeCell("A1", formula="1+2", value="3")])])
    change = diff_workbooks(old, new).cell_changes[0]
    assert change.severity is Severity.WARNING


def test_value_change_that_feeds_a_formula_is_breaking(tmp_path):
    def build(path, value):
        return _book(
            tmp_path,
            path,
            [FakeSheet("S", [FakeCell("A1", value=value), FakeCell("B1", formula="A1*2", value="4")])],
        )

    result = diff_workbooks(build("old.xlsx", "2"), build("new.xlsx", "5"))
    a1 = next(c for c in result.cell_changes if c.ref.a1 == "A1")
    assert a1.kind is ChangeKind.VALUE
    assert a1.severity is Severity.BREAKING
    assert [str(r) for r in a1.affected] == ["S!B1"]


def test_transitive_dependents_are_followed(tmp_path):
    """A change three links away from the total must still be reported."""
    def build(path, first):
        return _book(
            tmp_path,
            path,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", value=first),
                        FakeCell("B1", formula="A1+1", value="0"),
                        FakeCell("C1", formula="B1+1", value="0"),
                        FakeCell("D1", formula="C1+1", value="0"),
                    ],
                )
            ],
        )

    result = diff_workbooks(build("old.xlsx", "1"), build("new.xlsx", "9"))
    a1 = next(c for c in result.cell_changes if c.ref.a1 == "A1")
    assert [str(r) for r in a1.affected] == ["S!B1", "S!C1", "S!D1"]


def test_deleted_cell_with_dependents_is_breaking(tmp_path):
    old = _book(
        tmp_path,
        "old.xlsx",
        [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("B1", formula="A1*2", value="2")])],
    )
    new = _book(tmp_path, "new.xlsx", [FakeSheet("S", [FakeCell("B1", formula="A1*2", value="2")])])
    a1 = next(c for c in diff_workbooks(old, new).cell_changes if c.ref.a1 == "A1")
    assert a1.kind is ChangeKind.REMOVED
    assert a1.severity is Severity.BREAKING


def test_new_cell_is_info(tmp_path):
    old = _book(tmp_path, "old.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    new = _book(
        tmp_path, "new.xlsx", [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("A2", value="2")])]
    )
    change = next(c for c in diff_workbooks(old, new).cell_changes if c.ref.a1 == "A2")
    assert change.kind is ChangeKind.ADDED
    assert change.severity is Severity.INFO


def test_formula_edit_without_recalculation_is_stale(tmp_path):
    """The signal that needs the cached value: new formula, old number."""
    old = _book(tmp_path, "old.xlsx", [FakeSheet("S", [FakeCell("A1", formula="1+1", value="2")])])
    new = _book(tmp_path, "new.xlsx", [FakeSheet("S", [FakeCell("A1", formula="1+2", value="2")])])
    change = diff_workbooks(old, new).cell_changes[0]
    assert change.kind is ChangeKind.STALE
    assert "not recalculated" in change.detail


def test_renamed_sheet_is_reported_as_a_rename(tmp_path):
    cells = [FakeCell("A1", value="1"), FakeCell("B1", formula="A1*2", value="2")]
    old = _book(tmp_path, "old.xlsx", [FakeSheet("Data", cells)])
    new = _book(tmp_path, "new.xlsx", [FakeSheet("Numbers", cells)])
    result = diff_workbooks(old, new)
    assert len(result.sheet_changes) == 1
    assert result.sheet_changes[0].kind == "renamed"
    assert result.sheet_changes[0].old_name == "Data"
    assert result.cell_changes == []


def test_added_and_removed_sheets(tmp_path):
    """Different contents, so this is a real add and remove, not a rename."""
    old = _book(tmp_path, "old.xlsx", [FakeSheet("Gone", [FakeCell("A1", value="1")])])
    new = _book(tmp_path, "new.xlsx", [FakeSheet("Fresh", [FakeCell("A1", value="2")])])
    kinds = {s.kind for s in diff_workbooks(old, new).sheet_changes}
    assert kinds == {"added", "removed"}


def test_untouched_sheet_is_not_reported(tmp_path):
    sheets = [FakeSheet("Same", [FakeCell("A1", value="1")]), FakeSheet("Other", [FakeCell("A1", value="1")])]
    old = _book(tmp_path, "old.xlsx", sheets)
    new = _book(
        tmp_path,
        "new.xlsx",
        [FakeSheet("Same", [FakeCell("A1", value="1")]), FakeSheet("Other", [FakeCell("A1", value="2")])],
    )
    result = diff_workbooks(old, new)
    assert [s.name for s in result.sheet_changes] == ["Other"]


def test_cross_sheet_impact_is_followed(tmp_path):
    """A change on one sheet must report the readers on another sheet."""
    old = _book(
        tmp_path,
        "old.xlsx",
        [
            FakeSheet("A", [FakeCell("A1", value="1")]),
            FakeSheet("B", [FakeCell("A1", formula="A!A1*2", value="2")]),
        ],
    )
    new = _book(
        tmp_path,
        "new.xlsx",
        [
            FakeSheet("A", [FakeCell("A1", value="5")]),
            FakeSheet("B", [FakeCell("A1", formula="A!A1*2", value="2")]),
        ],
    )
    change = diff_workbooks(old, new).cell_changes[0]
    assert change.ref.a1 == "A1"
    assert change.severity is Severity.BREAKING
    assert [str(r) for r in change.affected] == ["B!A1"]


def test_a_formula_reading_its_own_cell_is_not_a_dependency(tmp_path):
    """SUM(D2:D13) written in D12 includes D12; it must not list itself."""
    old = _book(
        tmp_path,
        "old.xlsx",
        [
            FakeSheet(
                "S",
                [
                    *[FakeCell(f"D{r}", value="1") for r in range(2, 13)],
                    FakeCell("D12", formula="SUM(D2:D13)", value="11"),
                ],
            )
        ],
    )
    new = _book(
        tmp_path,
        "new.xlsx",
        [
            FakeSheet(
                "S",
                [
                    *[FakeCell(f"D{r}", value="1") for r in range(2, 13)],
                    FakeCell("D12", formula="SUM(D2:D13)*2", value="22"),
                ],
            )
        ],
    )
    change = next(c for c in diff_workbooks(old, new).cell_changes if c.ref.a1 == "D12")
    assert change.affected == []
