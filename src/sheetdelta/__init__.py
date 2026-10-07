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

from .batch import DirectoryDiff, WorkbookEntry, diff_directories
from .differ import (
    CellChange,
    ChangeKind,
    DiffResult,
    Move,
    RowMove,
    Severity,
    SheetChange,
    Shift,
    diff_workbooks,
)
from .errors import (
    DirectoryError,
    SheetDeltaError,
    UnsupportedFormatError,
    WorkbookReadError,
)
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
    VolatileRef,
    Workbook,
    column_letter,
    column_number,
)
from .reader import read_workbook
from .references import extract_references
from .report import (
    directory_exit_code,
    exit_code,
    render_audit_github,
    render_audit_json,
    render_audit_text,
    render_directory_github,
    render_directory_json,
    render_directory_markdown,
    render_directory_summary,
    render_directory_text,
    render_github,
    render_json,
    render_markdown,
    render_summary,
    render_text,
    to_dict,
)

__version__ = "0.6.0"

__all__ = [
    "Cell",
    "CellChange",
    "CellIndex",
    "CellKind",
    "CellRef",
    "ChangeKind",
    "DiffResult",
    "DirectoryDiff",
    "DirectoryError",
    "Move",
    "RangeRef",
    "RowMove",
    "Reference",
    "Severity",
    "Sheet",
    "SheetChange",
    "SheetDeltaError",
    "Shift",
    "Table",
    "TableRef",
    "UnsupportedFormatError",
    "VolatileRef",
    "Workbook",
    "WorkbookEntry",
    "WorkbookReadError",
    "__version__",
    "column_letter",
    "column_number",
    "diff_directories",
    "diff_workbooks",
    "directory_exit_code",
    "exit_code",
    "extract_references",
    "read_workbook",
    "render_audit_github",
    "render_audit_json",
    "render_audit_text",
    "render_directory_github",
    "render_directory_json",
    "render_directory_markdown",
    "render_directory_summary",
    "render_directory_text",
    "render_github",
    "render_json",
    "render_markdown",
    "render_summary",
    "render_text",
    "to_dict",
]
