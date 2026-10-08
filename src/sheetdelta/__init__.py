"""sheetdelta: diff Excel workbooks without Excel.

Compares two .xlsx files cell by cell, including the formulas, the values
Excel cached, and the dependency graph that says which cells a change
reaches. Nothing is evaluated -- the reader only looks at what Excel already
stored.

Typical use is a CI check: fail the build when a formula edit corrupts a
downstream total, pass it when the change is cosmetic.

    from sheetdelta import read_workbook, diff_workbooks

    result = diff_workbooks(read_workbook("old.xlsx"), read_workbook("new.xlsx"))
    for change in result.breaking:
        print(change.ref, change.detail, change.affected)
"""

from __future__ import annotations

from .differ import (
    CellChange,
    ChangeKind,
    DiffResult,
    Severity,
    SheetChange,
    Shift,
    diff_workbooks,
)
from .errors import SheetDeltaError, UnsupportedFormatError, WorkbookReadError
from .model import (
    Cell,
    CellIndex,
    CellKind,
    CellRef,
    RangeRef,
    Reference,
    Sheet,
    Table,
    TableRef,
    Workbook,
    column_letter,
    column_number,
)
from .reader import read_workbook
from .references import extract_references
from .report import exit_code, render_json, render_summary, render_text, to_dict

__version__ = "0.2.2"

__all__ = [
    "Cell",
    "CellChange",
    "CellIndex",
    "CellKind",
    "CellRef",
    "ChangeKind",
    "DiffResult",
    "RangeRef",
    "Reference",
    "Severity",
    "Sheet",
    "SheetChange",
    "SheetDeltaError",
    "Shift",
    "Table",
    "TableRef",
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
    "render_summary",
    "render_text",
    "to_dict",
]
