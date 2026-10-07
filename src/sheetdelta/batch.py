"""Diff two directory trees of workbooks, one workbook at a time.

A single ``diff`` answers a question about one file. In a repository that holds
many workbooks -- a folder per team, a model split across files -- the question
is instead "what changed anywhere in here, and did any of it break something".
This walks two trees, pairs workbooks by their path relative to the root, and
runs the ordinary single-file diff on each pair. A workbook present on only one
side is an add or a remove.

Pairing is by relative path, not by name: ``sales/q1.xlsx`` and ``ops/q1.xlsx``
are different workbooks even though the file name matches, and a workbook moved
between folders reads as a remove plus an add rather than a rename, because a
path is what identifies it here.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .differ import DiffResult, diff_workbooks
from .errors import DirectoryError
from .reader import read_workbook

# The formats the reader understands. A directory walk ignores anything else,
# so a stray .csv or .xls in the tree does not abort the run.
WORKBOOK_SUFFIXES = (".xlsx", ".xlsm")


@dataclass
class WorkbookEntry:
    """One workbook within the tree, and how it changed."""

    name: str  # path relative to the root, POSIX-style
    status: str  # "added" | "removed" | "changed" | "unchanged"
    result: DiffResult | None = None

    @property
    def breaking_count(self) -> int:
        if self.result is None:
            return 0
        return len(self.result.breaking)

    @property
    def cell_change_count(self) -> int:
        if self.result is None:
            return 0
        return len(self.result.cell_changes)


@dataclass
class DirectoryDiff:
    """The whole comparison of two trees."""

    old_dir: str
    new_dir: str
    entries: list[WorkbookEntry] = field(default_factory=list)

    @property
    def changed(self) -> list[WorkbookEntry]:
        return [e for e in self.entries if e.status != "unchanged"]

    @property
    def breaking(self) -> list[WorkbookEntry]:
        return [e for e in self.entries if e.breaking_count]

    @property
    def has_changes(self) -> bool:
        return bool(self.changed)


def diff_directories(
    old_dir: str, new_dir: str, *, volatile_scope: str = "sheet"
) -> DirectoryDiff:
    """Compare every workbook under ``new_dir`` with its twin under ``old_dir``."""
    old_root = Path(old_dir)
    new_root = Path(new_dir)
    if not old_root.is_dir():
        raise DirectoryError(f"not a directory: {old_dir}")
    if not new_root.is_dir():
        raise DirectoryError(f"not a directory: {new_dir}")

    old_files = _workbooks(old_root)
    new_files = _workbooks(new_root)
    result = DirectoryDiff(old_dir=str(old_root), new_dir=str(new_root))

    for name in sorted(set(old_files) | set(new_files)):
        in_old, in_new = name in old_files, name in new_files
        if in_old and not in_new:
            result.entries.append(WorkbookEntry(name=name, status="removed"))
        elif in_new and not in_old:
            result.entries.append(WorkbookEntry(name=name, status="added"))
        else:
            diff = diff_workbooks(
                read_workbook(str(old_files[name]), volatile_scope=volatile_scope),
                read_workbook(str(new_files[name]), volatile_scope=volatile_scope),
            )
            status = "changed" if diff.has_changes else "unchanged"
            result.entries.append(WorkbookEntry(name=name, status=status, result=diff))

    return result


def _workbooks(root: Path) -> dict[str, Path]:
    """Every workbook under ``root``, keyed by its POSIX path relative to root."""
    found: dict[str, Path] = {}
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in WORKBOOK_SUFFIXES:
            found[path.relative_to(root).as_posix()] = path
    return found
