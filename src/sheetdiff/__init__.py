"""sheetdiff: diff Excel workbooks without Excel.

Compares two .xlsx files cell by cell, including the formulas, the values
Excel cached, and the dependency graph that says which cells a change
reaches. Nothing is evaluated -- the reader only looks at what Excel already
stored.

Typical use is a CI check: fail the build when a formula edit corrupts a
downstream total, pass it when the change is cosmetic.

    from sheetdiff import read_workbook, diff_workbooks

    result = diff_workbooks(read_workbook("old.xlsx"), read_workbook("new.xlsx"))
    for change in result.breaking:
        print(change.ref, change.detail, change.affected)
"""

from __future__ import annotations

from .differ import CellChange, ChangeKind, DiffResult, Severity, SheetChange, diff_workbooks
from .errors import SheetDiffError, UnsupportedFormatError, WorkbookReadError
from .model import (
    Cell,
    CellKind,
    CellRef,
    RangeRef,
    Reference,
    Sheet,
    Workbook,
    column_letter,
    column_number,
)
from .reader import read_workbook
from .references import extract_references
from .report import exit_code, render_json, render_text, to_dict

__version__ = "0.1.0"

__all__ = [
    "Cell",
    "CellChange",
    "CellKind",
    "CellRef",
    "ChangeKind",
    "DiffResult",
    "RangeRef",
    "Reference",
    "Severity",
    "Sheet",
    "SheetChange",
    "SheetDiffError",
    "UnsupportedFormatError",
    "Workbook",
    "WorkbookReadError",
    "__version__",
    "column_letter",
    "column_number",
    "diff_workbooks",
    "exit_code",
    "extract_references",
    "read_workbook",
    "render_json",
    "render_text",
    "to_dict",
]
