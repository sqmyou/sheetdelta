"""Tests for volatile references: INDIRECT, OFFSET and their effect on the graph."""

from __future__ import annotations

from pathlib import Path

from sheetdelta.audit import AuditResult, audit_workbook
from sheetdelta.differ import DiffResult, diff_workbooks
from sheetdelta.model import CellRef, Reference, VolatileRef
from sheetdelta.reader import read_workbook
from sheetdelta.references import extract_references
from sheetdelta.report import render_audit_github, render_audit_text

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook

SHEETS = {"s": "S"}


def _refs(
    formula: str, *, row: int | None = None, scope: str = "sheet"
) -> frozenset[Reference]:
    return extract_references(
        formula,
        sheet="S",
        sheets=SHEETS,
        defined_names={},
        tables={},
        row=row,
        volatile_scope=scope,
    )


class TestExtraction:
    def test_indirect_with_a_string_argument_is_volatile(self):
        refs = _refs('INDIRECT("A1")+B1')
        assert VolatileRef("S", "INDIRECT", "sheet", "A1") in refs
        # The ordinary reference is still found alongside it.
        assert CellRef("S", 2, 1) in refs

    def test_indirect_with_computed_argument_has_no_literal(self):
        refs = _refs('INDIRECT("A"&B1)')
        volatile = [r for r in refs if isinstance(r, VolatileRef)]
        assert volatile == [VolatileRef("S", "INDIRECT", "sheet", None)]

    def test_offset_is_volatile_and_keeps_its_anchor(self):
        refs = _refs("OFFSET(C3,0,1)")
        assert VolatileRef("S", "OFFSET", "sheet", None) in refs
        # The anchor cell the offset is relative to is still a real dependency.
        assert CellRef("S", 3, 3) in refs

    def test_case_insensitive_function_name(self):
        assert VolatileRef("S", "INDIRECT", "sheet", None) in _refs("indirect(B1)")

    def test_a_longer_name_is_not_a_volatile_function(self):
        # MYINDIRECT is not INDIRECT; the guard stops a match mid-word.
        assert not [r for r in _refs("MYINDIRECT(1)") if isinstance(r, VolatileRef)]

    def test_indirect_inside_a_string_literal_is_not_detected(self):
        refs = _refs('"see INDIRECT(a1)"')
        assert not [r for r in refs if isinstance(r, VolatileRef)]

    def test_scope_is_carried_through(self):
        refs = _refs('INDIRECT("A1")', scope="workbook")
        assert VolatileRef("S", "INDIRECT", "workbook", "A1") in refs


class TestAudit:
    def _audit(self, tmp_path: Path, formula: str, scope: str = "sheet") -> AuditResult:
        path = str(tmp_path / "book.xlsx")
        write_workbook(
            path,
            [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("B1", formula=formula, value="1")])],
        )
        return audit_workbook(read_workbook(path, volatile_scope=scope))

    def test_audit_reports_a_volatile_reference(self, tmp_path: Path) -> None:
        result = self._audit(tmp_path, 'INDIRECT("A1")')
        assert len(result.volatile) == 1
        finding = result.volatile[0]
        assert finding.kind == "INDIRECT"
        assert finding.literal == "A1"
        assert finding.ref.a1 == "B1"

    def test_audit_stays_sound_with_a_volatile_reference(self, tmp_path):
        # A volatile reference is not a broken one: the workbook is still
        # sound, because Excel will compute it. It is the graph that cannot
        # see it, and that is what the separate finding says.
        result = self._audit(tmp_path, 'INDIRECT("A1")')
        assert result.is_sound
        assert result.issue_count == 0

    def test_audit_of_a_clean_workbook_has_no_volatile_finding(self, tmp_path):
        result = self._audit(tmp_path, "A1+1")
        assert result.volatile == []

    def test_audit_text_names_the_volatile_reference(self, tmp_path):
        result = self._audit(tmp_path, 'INDIRECT("A1")')
        text = render_audit_text(result)
        assert "volatile reference" in text
        assert "INDIRECT(A1)" in text

    def test_audit_github_emits_a_warning(self, tmp_path):
        result = self._audit(tmp_path, "OFFSET(A1,0,1)")
        output = render_audit_github(result)
        assert "::warning" in output
        assert "OFFSET" in output


class TestDiffGraph:
    def _pair(
        self, tmp_path: Path, old_formula: str, new_formula: str, scope: str = "sheet"
    ) -> DiffResult:
        old = str(tmp_path / "old.xlsx")
        new = str(tmp_path / "new.xlsx")
        write_workbook(
            old,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", value="1"),
                        FakeCell("B1", formula=old_formula, value="1"),
                    ],
                )
            ],
        )
        write_workbook(
            new,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", value="2"),
                        FakeCell("B1", formula=new_formula, value="1"),
                    ],
                )
            ],
        )
        return diff_workbooks(
            read_workbook(old, volatile_scope=scope),
            read_workbook(new, volatile_scope=scope),
        )

    def test_a_volatile_reader_is_seen_as_affected(self, tmp_path: Path) -> None:
        # B1 reads A1 through INDIRECT. Without the volatile edge, the change
        # to A1 would not be followed into B1 and the impact understated.
        result = self._pair(tmp_path, 'INDIRECT("A1")', 'INDIRECT("A1")')
        affected = {str(ref) for change in result.cell_changes for ref in change.affected}
        assert "S!B1" in affected

    def test_identical_workbooks_have_no_changes(self, tmp_path):
        path = str(tmp_path / "same.xlsx")
        write_workbook(
            path,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", value="1"),
                        FakeCell("B1", formula='INDIRECT("A1")', value="1"),
                    ],
                )
            ],
        )
        result = diff_workbooks(read_workbook(path), read_workbook(path))
        assert not result.has_changes
