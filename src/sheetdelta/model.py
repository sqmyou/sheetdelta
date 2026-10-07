"""Data model shared by the reader, the diff and the report."""

from __future__ import annotations

from dataclasses import dataclass, field
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


Reference = CellRef | RangeRef


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

    @property
    def display(self) -> str:
        """A short human string for the cell's contents."""
        if self.kind is CellKind.FORMULA:
            return f"={self.formula}"
        return self.cached_value if self.cached_value is not None else ""


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
    date1904: bool = False

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
