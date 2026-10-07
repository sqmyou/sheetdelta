"""Exceptions raised by sheetdiff.

All inherit from :class:`SheetDiffError` so the CLI can catch one type and
print a plain message instead of a traceback.
"""

from __future__ import annotations


class SheetDiffError(Exception):
    """Base class for every error sheetdiff raises on purpose."""


class WorkbookReadError(SheetDiffError):
    """A workbook could not be read as an .xlsx file."""


class UnsupportedFormatError(WorkbookReadError):
    """The file is a workbook format we do not parse (.xls, .xlsb, ...)."""
