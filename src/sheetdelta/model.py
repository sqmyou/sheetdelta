"""Data model shared by the reader, the diff and the report."""

from __future__ import annotations

import bisect
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from enum import Enum


class CellKind(str, Enum):
    """What a cell holds, which decides how a change to it is described."""

    VALUE = "value"
    FORMULA = "formula"
    BLANK = "blank"


@dataclass(frozen=True)
class CellRef:
    """A single cell address, e.g. ``Sheet1!D12``.

    Kept as sheet plus column/row so callers never parse A1 strings by hand.
    ``col`` and ``row`` are 1-based, matching what a human sees in Excel.
    """

    sheet: str
    col: int
    row: int

    @property
    def a1(self) -> str:
        """The bare A1 address without the sheet, e.g. ``D12``."""
        return f"{column_letter(self.col)}{self.row}"

    def contains(self, ref: CellRef) -> bool:
        return self == ref

    def __str__(self) -> str:
        return f"{self.sheet}!{self.a1}"


@dataclass(frozen=True)
class RangeRef:
    """A rectangular block of cells, e.g. ``Sheet1!D2:D11``.

    Ranges are kept as ranges rather than expanded into individual cells. A
    formula like ``SUM(A:A)`` covers a million cells, and expanding it would
    cost more than the whole diff. The dependency graph asks ``contains``
    instead.
    """

    sheet: str
    min_col: int
    max_col: int
    min_row: int
    max_row: int

    def contains(self, ref: CellRef) -> bool:
        return (
            ref.sheet == self.sheet
            and self.min_col <= ref.col <= self.max_col
            and self.min_row <= ref.row <= self.max_row
        )

    @property
    def a1(self) -> str:
        start = f"{column_letter(self.min_col)}{self.min_row}"
        end = f"{column_letter(self.max_col)}{self.max_row}"
        return start if start == end else f"{start}:{end}"

    def __str__(self) -> str:
        return f"{self.sheet}!{self.a1}"


@dataclass(frozen=True)
class TableRef:
    """A structured reference into an Excel table, e.g. ``Table1[Amount]``.

    The column named is kept rather than resolved to a range, because a table
    can be edited (rows appended, columns renamed) and the reference text stays
    the same. Resolution needs the table's definition, which only the workbook
    has, so :meth:`resolve` takes it.
    """

    table: str
    column: str | None = None
    specifier: str | None = None

    @property
    def a1(self) -> str:
        return f"{self.table}[{self.column or self.specifier or ''}]"

    def __str__(self) -> str:
        return self.a1

    def resolve(self, tables: Mapping[str, Table]) -> RangeRef | None:
        """The range this reference covers, or None if it cannot be pinned down."""
        table = tables.get(self.table.lower())
        if table is None:
            return None

        if self.column is not None:
            column = table.column_number(self.column)
            if column is None:
                return None
            min_col = max_col = column
        else:
            min_col, max_col = table.min_col, table.max_col

        # The specifier picks the rows; without one a structured reference means
        # the data body, which excludes the header and any totals row.
        specifier = (self.specifier or "").lower()
        if specifier in {"", "#data"}:
            min_row, max_row = table.first_data_row, table.last_data_row
        elif specifier == "#all":
            min_row, max_row = table.first_row, table.last_row
        elif specifier == "#headers":
            if not table.has_header_row:
                return None
            min_row = max_row = table.first_row
        elif specifier == "#totals":
            if not table.has_totals_row:
                return None
            min_row = max_row = table.last_row
        else:
            return None

        return RangeRef(table.sheet, min_col, max_col, min_row, max_row)


Reference = CellRef | RangeRef | TableRef


class CellIndex:
    """Finds the cells a reference covers without scanning the whole workbook.

    Both the diff and the audit need "which cells does this formula read", once
    per reference. Asking every reference against every cell in the file is
    quadratic, and on a workbook of a few thousand rows it dominates the run.
    Here a cell lookup goes straight to its address, and a range is walked row by
    row over only the rows that hold cells -- so ``SUM(A:A)`` costs the cells it
    actually touches rather than a million probes.
    """

    def __init__(self, cells: Iterable[CellRef]) -> None:
        self._by_row: dict[tuple[str, int], list[CellRef]] = {}
        rows_by_sheet: dict[str, set[int]] = {}
        for ref in cells:
            self._by_row.setdefault((ref.sheet, ref.row), []).append(ref)
            rows_by_sheet.setdefault(ref.sheet, set()).add(ref.row)
        # A range is answered by binary-searching the rows that actually hold
        # cells, so ``SUM(A:A)`` costs the rows in use, not all 1,048,576.
        self._rows_by_sheet: dict[str, list[int]] = {
            sheet: sorted(rows) for sheet, rows in rows_by_sheet.items()
        }

    def covered(self, reference: Reference) -> list[CellRef]:
        """Every indexed cell that ``reference`` points at."""
        if isinstance(reference, TableRef):
            # A structured reference cannot be resolved without the workbook's
            # table definitions, and this index does not carry them.
            return []

        if isinstance(reference, CellRef):
            row_cells = self._by_row.get((reference.sheet, reference.row))
            if not row_cells:
                return []
            return [ref for ref in row_cells if ref.col == reference.col]

        rows = self._rows_by_sheet.get(reference.sheet)
        if not rows:
            return []
        start = bisect.bisect_left(rows, reference.min_row)
        end = bisect.bisect_right(rows, reference.max_row)
        out: list[CellRef] = []
        for row_number in rows[start:end]:
            for ref in self._by_row[(reference.sheet, row_number)]:
                if reference.min_col <= ref.col <= reference.max_col:
                    out.append(ref)
        return out


def column_letter(col: int) -> str:
    """Convert a 1-based column number to letters (1 -> A, 27 -> AA)."""
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def column_number(letters: str) -> int:
    """Convert column letters to a 1-based number (A -> 1, AA -> 27).

    ``$`` markers are ignored, so ``$AA`` parses the same as ``AA``.
    """
    n = 0
    for ch in letters.upper():
        if not ("A" <= ch <= "Z"):
            continue
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


@dataclass
class Cell:
    """One cell, holding the raw formula and the value Excel last cached.

    The cached value matters as much as the formula. Excel writes both, and
    when a file is edited by something that does not recalculate, the formula
    and its cached value disagree. Comparing the two is how we spot a workbook
    whose numbers are stale.
    """

    ref: CellRef
    kind: CellKind
    formula: str | None = None
    cached_value: str | None = None
    value_type: str | None = None
    refs: frozenset[Reference] = field(default_factory=frozenset)
    raw_value: str | None = None

    @property
    def display(self) -> str:
        """A short human string for the cell's contents."""
        if self.kind is CellKind.FORMULA:
            return f"={self.formula}"
        return self.cached_value if self.cached_value is not None else ""

    @property
    def signature(self) -> str | None:
        """A canonical key for the whole cell, ignoring how it is displayed.

        A formula is keyed by its text; a value by its normalized number. Used
        where two cells must be judged the same cell -- lining up a shift, or
        deciding a value changed.
        """
        if self.kind is CellKind.FORMULA:
            return f"={self.formula}"
        return self.value_signature

    @property
    def value_signature(self) -> str | None:
        """A canonical key for the cached value alone, ignoring its format.

        Numbers are normalized so ``1`` and ``1.0`` read as the same, and a
        bare number is compared as stored rather than as its formatted text:
        flipping a cell from a date format to a plain one changes the displayed
        string while the underlying serial is identical.
        """
        raw = self.raw_value if self.raw_value is not None else self.cached_value
        if raw is None:
            return None
        if self.value_type == "date":
            return raw
        try:
            return format(Decimal(raw).normalize(), "f")
        except (InvalidOperation, ValueError):
            return raw


@dataclass(frozen=True)
class Table:
    """One Excel table: the sheet and block it occupies, plus its column names.

    The header row is not always present and the totals row is optional, so both
    are tracked: a structured reference means different rows depending on which
    part of the table it names.
    """

    name: str
    sheet: str
    min_col: int
    max_col: int
    min_row: int
    max_row: int
    columns: tuple[str, ...] = ()
    has_header_row: bool = True
    has_totals_row: bool = False

    @property
    def first_row(self) -> int:
        return self.min_row

    @property
    def last_row(self) -> int:
        return self.max_row

    @property
    def first_data_row(self) -> int:
        return self.min_row + 1 if self.has_header_row else self.min_row

    @property
    def last_data_row(self) -> int:
        return self.max_row - 1 if self.has_totals_row else self.max_row

    def column_number(self, name: str) -> int | None:
        """The grid column a table column name sits in, case-insensitively."""
        lowered = name.strip().lower()
        for offset, column in enumerate(self.columns):
            if column.strip().lower() == lowered:
                return self.min_col + offset
        return None


@dataclass
class Sheet:
    """One worksheet: its name in file order, plus every non-empty cell."""

    name: str
    index: int
    cells: dict[CellRef, Cell]

    @property
    def cell_count(self) -> int:
        return len(self.cells)


@dataclass
class Workbook:
    """A parsed workbook: the sheets and the defined names.

    Defined names (``Revenue`` -> ``Sheet1!$B$2:$B$10``) are how a formula
    refers to a range by a human name, so they have to travel with the
    dependency graph or a rename would look like a breaking change.
    """

    path: str
    sheets: list[Sheet]
    defined_names: dict[str, str] = field(default_factory=dict)
    tables: dict[str, Table] = field(default_factory=dict)
    date1904: bool = False
    partial_reads: list[str] = field(default_factory=list)

    def sheet_by_name(self, name: str) -> Sheet | None:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None

    @property
    def sheet_names(self) -> list[str]:
        return [sheet.name for sheet in self.sheets]

    def all_cells(self) -> list[Cell]:
        return [cell for sheet in self.sheets for cell in sheet.cells.values()]
