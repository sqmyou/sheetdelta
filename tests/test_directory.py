"""Tests for diffing two directories of workbooks."""

from __future__ import annotations

import json

import pytest

from sheetdelta.batch import diff_directories
from sheetdelta.cli import main
from sheetdelta.errors import DirectoryError
from sheetdelta.report import directory_exit_code

from .xlsx_fixtures import FakeCell, FakeSheet, write_workbook


def _write(path, sheets):
    write_workbook(str(path), sheets)


def _tree(tmp_path):
    """Two trees: one workbook changed, one only added, one only removed."""
    old = tmp_path / "old"
    new = tmp_path / "new"
    (old / "team").mkdir(parents=True)
    (new / "team").mkdir(parents=True)

    _write(old / "team" / "revenue.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(new / "team" / "revenue.xlsx", [FakeSheet("S", [FakeCell("A1", value="2")])])

    _write(new / "team" / "new.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(old / "team" / "gone.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    return str(old), str(new)


def test_directory_pairs_by_relative_path(tmp_path):
    old, new = _tree(tmp_path)
    result = diff_directories(old, new)
    statuses = {e.name: e.status for e in result.entries}
    assert statuses == {
        "team/revenue.xlsx": "changed",
        "team/new.xlsx": "added",
        "team/gone.xlsx": "removed",
    }


def test_same_name_in_different_dirs_is_two_workbooks(tmp_path):
    # Pairing is by path, not by file name: sales/q1 is not ops/q1.
    old = tmp_path / "old"
    new = tmp_path / "new"
    (old / "sales").mkdir(parents=True)
    (old / "ops").mkdir(parents=True)
    (new / "sales").mkdir(parents=True)
    (new / "ops").mkdir(parents=True)
    _write(old / "sales" / "q1.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(old / "ops" / "q1.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(new / "sales" / "q1.xlsx", [FakeSheet("S", [FakeCell("A1", value="2")])])
    _write(new / "ops" / "q1.xlsx", [FakeSheet("S", [FakeCell("A1", value="2")])])
    result = diff_directories(str(old), str(new))
    assert {e.name for e in result.changed} == {"sales/q1.xlsx", "ops/q1.xlsx"}


def test_a_directory_move_is_a_remove_and_an_add(tmp_path):
    old = tmp_path / "old"
    new = tmp_path / "new"
    (old / "a").mkdir(parents=True)
    (new / "b").mkdir(parents=True)
    _write(old / "a" / "book.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(new / "b" / "book.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    result = diff_directories(str(old), str(new))
    statuses = {e.name: e.status for e in result.entries}
    assert statuses == {"a/book.xlsx": "removed", "b/book.xlsx": "added"}


def test_unchanged_workbooks_are_listed_but_not_changed(tmp_path):
    old = tmp_path / "old"
    new = tmp_path / "new"
    old.mkdir()
    new.mkdir()
    _write(old / "same.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(new / "same.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    result = diff_directories(str(old), str(new))
    assert result.entries[0].status == "unchanged"
    assert result.changed == []
    assert not result.has_changes


def test_breaking_counts_only_breaking_workbooks(tmp_path):
    old = tmp_path / "old"
    new = tmp_path / "new"
    old.mkdir()
    new.mkdir()
    # A formula whose cached value did not move is breaking when read downstream.
    _write(
        old / "book.xlsx",
        [FakeSheet("S", [FakeCell("A1", value="1"), FakeCell("B1", formula="A1*2", value="2")])],
    )
    _write(
        new / "book.xlsx",
        [FakeSheet("S", [FakeCell("A1", value="9"), FakeCell("B1", formula="A1*2", value="2")])],
    )
    result = diff_directories(str(old), str(new))
    assert len(result.breaking) == 1
    assert directory_exit_code(result, "breaking") == 1
    assert directory_exit_code(result, "never") == 0


def test_non_workbooks_in_the_tree_are_ignored(tmp_path):
    old = tmp_path / "old"
    new = tmp_path / "new"
    old.mkdir()
    new.mkdir()
    (old / "notes.txt").write_text("hi")
    (new / "notes.txt").write_text("hi")
    (old / "data.csv").write_text("a,b\n")
    (new / "data.csv").write_text("a,b\n")
    _write(old / "book.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    _write(new / "book.xlsx", [FakeSheet("S", [FakeCell("A1", value="1")])])
    result = diff_directories(str(old), str(new))
    assert [e.name for e in result.entries] == ["book.xlsx"]


def test_a_missing_directory_is_a_clear_error(tmp_path):
    with pytest.raises(DirectoryError):
        diff_directories(str(tmp_path / "nope"), str(tmp_path))


def test_cli_diff_dir_text_and_exit(tmp_path, capsys):
    old, new = _tree(tmp_path)
    code = main(["diff-dir", old, new, "--fail-on", "never"])
    out = capsys.readouterr().out
    assert code == 0
    assert "team/revenue.xlsx" in out
    assert "team/new.xlsx  added" in out
    assert "team/gone.xlsx  removed" in out


def test_cli_diff_dir_json(tmp_path, capsys):
    old, new = _tree(tmp_path)
    main(["diff-dir", old, new, "--json", "--fail-on", "never"])
    payload = json.loads(capsys.readouterr().out)
    names = {w["name"] for w in payload["workbooks"]}
    assert names == {"team/revenue.xlsx", "team/new.xlsx", "team/gone.xlsx"}
    assert payload["summary"]["added"] == 1
    assert payload["summary"]["removed"] == 1


def test_cli_diff_dir_github_annotates_add_and_remove(tmp_path, capsys):
    old, new = _tree(tmp_path)
    main(["diff-dir", old, new, "--format", "github", "--fail-on", "never"])
    out = capsys.readouterr().out
    assert "::notice" in out
    assert "team/new.xlsx%3A added" in out
    assert "team/gone.xlsx%3A removed" in out


def test_cli_diff_dir_markdown(tmp_path, capsys):
    old, new = _tree(tmp_path)
    main(["diff-dir", old, new, "--format", "markdown", "--fail-on", "never"])
    out = capsys.readouterr().out
    assert "| Workbook | Change | Cells | Breaking |" in out
    assert "team/revenue.xlsx" in out


def test_cli_diff_dir_missing_directory_exits_two(tmp_path, capsys):
    code = main(["diff-dir", str(tmp_path / "nope"), str(tmp_path)])
    assert code == 2
    assert "not a directory" in capsys.readouterr().err
