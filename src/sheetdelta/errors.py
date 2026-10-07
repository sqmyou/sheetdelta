"""Exceptions raised by sheetdelta.

All inherit from :class:`SheetDeltaError` so the CLI can catch one type and
print a plain message instead of a traceback.
"""

from __future__ import annotations


class SheetDeltaError(Exception):
    """Base class for every error sheetdelta raises on purpose."""


class WorkbookReadError(SheetDeltaError):
    """A workbook could not be read as an .xlsx file."""


class UnsupportedFormatError(WorkbookReadError):
    """The file is a workbook format we do not parse (.xls, .xlsb, ...)."""
