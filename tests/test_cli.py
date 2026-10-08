"""Tests for the CLI, the report rendering and the audit."""

from __future__ import annotations

import json

from sheetdelta.audit import audit_workbook
from sheetdelta.cli import main
from sheetdelta.differ import diff_workbooks
from sheetdelta.reader import read_workbook
from sheetdelta.report import exit_code, render_json, render_text

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook


def _pair(tmp_path):
    """A pair of workbooks with one breaking formula change and one value change."""
    old = str(tmp_path / "old.xlsx")
    new = str(tmp_path / "new.xlsx")
    common = [FakeCell("A1", value="1"), FakeCell("A2", value="2")]
    write_workbook(
        old,
        [FakeSheet("Revenue", [*common, FakeCell("A3", formula="SUM(A1:A2)", value="3"), FakeCell("B1", formula="A3*10", value="30")])],
    )
    write_workbook(
        new,
        [FakeSheet("Revenue", [*common, FakeCell("A3", formula="SUM(A1:A2)*2", value="6"), FakeCell("B1", formula="A3*10", value="60")])],
    )
    return old, new


class TestExitCodes:
    def test_breaking_fails_on_breaking(self, tmp_path):
        old, new = _pair(tmp_path)
        result = diff_workbooks(read_workbook(old), read_workbook(new))
        assert exit_code(result, "breaking") == 1
        assert exit_code(result, "any") == 1

    def test_identical_workbooks_never_fail(self, tmp_path):
        path = str(tmp_path / "same.xlsx")
        write_workbook(path, [FakeSheet("S", [FakeCell("A1", value="1")])])
        result = diff_workbooks(read_workbook(path), read_workbook(path))
        assert exit_code(result, "any") == 0
        assert exit_code(result, "breaking") == 0

    def test_never_fails_regardless(self, tmp_path):
        old, new = _pair(tmp_path)
        result = diff_workbooks(read_workbook(old), read_workbook(new))
        assert exit_code(result, "never") == 0


class TestCli:
    def test_diff_exits_one_on_breaking(self, tmp_path, capsys):
        old, new = _pair(tmp_path)
        code = main(["diff", old, new])
        out = capsys.readouterr().out
        assert code == 1
        assert "affects" in out

    def test_diff_exits_zero_when_nothing_breaking(self, tmp_path):
        old = str(tmp_path / "old.xlsx")
        new = str(tmp_path / "new.xlsx")
        write_workbook(old, [FakeSheet("S", [FakeCell("A1", formula="1+1", value="2")])])
        write_workbook(new, [FakeSheet("S", [FakeCell("A1", formula="1+2", value="3")])])
        assert main(["diff", old, new]) == 0

    def test_json_output_parses(self, tmp_path, capsys):
        old, new = _pair(tmp_path)
        main(["diff", old, new, "--json", "--fail-on", "never"])
        payload = json.loads(capsys.readouterr().out)
        assert payload["summary"]["breaking"] == 1
        cells = [c["a1"] for c in payload["sheets"][0]["changes"]]
        assert "A3" in cells

    def test_missing_file_exits_two_with_a_message(self, tmp_path, capsys):
        code = main(["diff", str(tmp_path / "nope.xlsx"), str(tmp_path / "nope2.xlsx")])
        assert code == 2
        assert "no such file" in capsys.readouterr().err

    def test_audit_of_a_sound_workbook(self, tmp_path, capsys):
        path = str(tmp_path / "ok.xlsx")
        write_workbook(
            path,
            [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("A2", formula="A1*2", value="2")])],
        )
        code = main(["audit", path])
        out = capsys.readouterr().out
        assert code == 0
        assert "No broken references" in out

    def test_audit_finds_a_reference_to_a_missing_sheet(self, tmp_path, capsys):
        path = str(tmp_path / "bad.xlsx")
        write_workbook(
            path,
            [FakeSheet("S", [FakeCell("A1", formula="'Gone'!B2*2", value="0")])],
        )
        code = main(["audit", path])
        out = capsys.readouterr().out
        assert code == 1
        assert "does not exist" in out

    def test_audit_fail_on_breaking_still_fails_on_an_issue(self, tmp_path, capsys):
        """A broken reference is a defect, so --fail-on breaking must not pass."""
        path = str(tmp_path / "bad2.xlsx")
        write_workbook(
            path,
            [FakeSheet("S", [FakeCell("A1", formula="'Gone'!B2*2", value="0")])],
        )
        code = main(["audit", path, "--fail-on", "breaking"])
        capsys.readouterr()
        assert code == 1

    def test_audit_fail_on_never_always_passes(self, tmp_path, capsys):
        path = str(tmp_path / "bad3.xlsx")
        write_workbook(
            path,
            [FakeSheet("S", [FakeCell("A1", formula="'Gone'!B2*2", value="0")])],
        )
        code = main(["audit", path, "--fail-on", "never"])
        capsys.readouterr()
        assert code == 0

    def test_audit_reports_an_unreadable_sheet_and_fails(self, tmp_path, capsys):
        import zipfile

        path = str(tmp_path / "partial.xlsx")
        write_workbook(
            path,
            [FakeSheet("Good", [FakeCell("A1", value="1")]), FakeSheet("Broken")],
        )
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        parts["xl/worksheets/sheet2.xml"] = b"<worksheet><not-closed>"
        with zipfile.ZipFile(path, "w") as out:
            for name, data in parts.items():
                out.writestr(name, data)

        code = main(["audit", path])
        out = capsys.readouterr().out
        assert code == 1
        assert "could not be read" in out
        assert "Broken" in out

    def test_summary_flag_prints_counts_without_cell_detail(self, tmp_path, capsys):
        old, new = _pair(tmp_path)
        code = main(["diff", old, new, "--summary", "--fail-on", "never"])
        out = capsys.readouterr().out
        assert code == 0
        assert "Revenue" in out
        assert "A3" not in out  # no per-cell detail in a summary
        assert "breaking" in out

    def test_summary_reports_a_shift(self, tmp_path, capsys):
        old = str(tmp_path / "old.xlsx")
        new = str(tmp_path / "new.xlsx")
        write_workbook(
            old, [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("A2", value="2")])]
        )
        write_workbook(
            new,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", value="1"),
                        FakeCell("A2", value="new"),
                        FakeCell("A3", value="2"),
                    ],
                )
            ],
        )
        main(["diff", old, new, "--summary", "--fail-on", "never"])
        assert "1 row inserted at row 2" in capsys.readouterr().out

    def test_summary_of_identical_workbooks(self, tmp_path, capsys):
        path = str(tmp_path / "same.xlsx")
        write_workbook(path, [FakeSheet("S", [FakeCell("A1", value="1")])])
        main(["diff", path, path, "--summary"])
        assert "No changes." in capsys.readouterr().out


class TestReport:
    def test_text_report_shows_old_and_new(self, tmp_path):
        old, new = _pair(tmp_path)
        text = render_text(diff_workbooks(read_workbook(old), read_workbook(new)))
        assert "-  =SUM(A1:A2)" in text
        assert "+  =SUM(A1:A2)*2" in text
        assert "1 breaking change" in text

    def test_no_changes_report(self, tmp_path):
        path = str(tmp_path / "same.xlsx")
        write_workbook(path, [FakeSheet("S", [FakeCell("A1", value="1")])])
        text = render_text(diff_workbooks(read_workbook(path), read_workbook(path)))
        assert "No changes." in text

    def test_json_has_stable_keys(self, tmp_path):
        old, new = _pair(tmp_path)
        payload = json.loads(render_json(diff_workbooks(read_workbook(old), read_workbook(new))))
        assert set(payload) == {"old", "new", "has_changes", "summary", "sheets"}


class TestAudit:
    def test_cycle_is_found(self, tmp_path):
        """Two cells that read each other cannot be calculated by Excel."""
        path = str(tmp_path / "cycle.xlsx")
        write_workbook(
            path,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", formula="B1+1", value="0"),
                        FakeCell("B1", formula="A1+1", value="0"),
                    ],
                )
            ],
        )
        result = audit_workbook(read_workbook(path))
        assert not result.is_sound
        assert len(result.cycles) == 1
        assert "S!A1" in result.cycles[0].display

    def test_cross_sheet_cycle_is_found(self, tmp_path):
        path = str(tmp_path / "cycle.xlsx")
        write_workbook(
            path,
            [
                FakeSheet("A", [FakeCell("A1", formula="B!A1+1", value="0")]),
                FakeSheet("B", [FakeCell("A1", formula="A!A1+1", value="0")]),
            ],
        )
        result = audit_workbook(read_workbook(path))
        assert len(result.cycles) == 1

    def test_a_normal_chain_is_not_a_cycle(self, tmp_path):
        path = str(tmp_path / "chain.xlsx")
        write_workbook(
            path,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", value="1"),
                        FakeCell("B1", formula="A1+1", value="2"),
                        FakeCell("C1", formula="B1+1", value="3"),
                    ],
                )
            ],
        )
        result = audit_workbook(read_workbook(path))
        assert result.cycles == []
        assert result.is_sound

    def test_a_total_inside_its_own_range_is_not_a_cycle(self, tmp_path):
        """SUM(D2:D12) written in D12 reads D12; that is not circular."""
        path = str(tmp_path / "total.xlsx")
        write_workbook(
            path,
            [
                FakeSheet(
                    "S",
                    [
                        *[FakeCell(f"D{r}", value="1") for r in range(2, 13)],
                        FakeCell("D12", formula="SUM(D2:D12)", value="11"),
                    ],
                )
            ],
        )
        assert audit_workbook(read_workbook(path)).is_sound

    def test_sound_workbook_has_no_issues(self, tmp_path):
        path = str(tmp_path / "ok.xlsx")
        write_workbook(
            path,
            [
                FakeSheet(
                    "S",
                    [FakeCell("A1", value="1"), FakeCell("A2", formula="SUM(A1:A1)", value="1")],
                )
            ],
        )
        result = audit_workbook(read_workbook(path))
        assert result.is_sound
        assert result.formula_count == 1

    def test_audit_json_includes_cycles(self, tmp_path, capsys):
        path = str(tmp_path / "cycle.xlsx")
        write_workbook(
            path,
            [
                FakeSheet(
                    "S",
                    [
                        FakeCell("A1", formula="B1+1", value="0"),
                        FakeCell("B1", formula="A1+1", value="0"),
                    ],
                )
            ],
        )
        main(["audit", path, "--json"])
        payload = json.loads(capsys.readouterr().out)
        assert payload["sound"] is False
        assert payload["cycles"]
