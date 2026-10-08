"""Pull cell and range references out of an Excel formula.

This is a scanner, not a formula evaluator. It needs to be accurate enough to
build a dependency graph, which is why it is fussier than a one-line regex:
``LOG10(...)`` looks exactly like a reference to column LOG row 10, and the
text ``"see A1"`` looks like a reference to A1. Both are handled here.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping

from .model import CellRef, RangeRef, Reference, TableRef, column_number

# Excel string literals are double-quoted, and a literal quote is doubled.
_STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')

_REF = re.compile(
    r"""
    (?<![A-Za-z0-9_$])          # not the tail of a longer name
    (?:
        '(?P<qsheet>[^']+)'\s*!\s*          # 'My Sheet'!
      | (?P<sheet>[A-Za-z_][A-Za-z0-9_.]*)\s*!\s*   # Sheet1!
    )?
    (?:
        (?P<cc1>\$?[A-Za-z]{1,3})\s*:\s*(?P<cc2>\$?[A-Za-z]{1,3})
      | (?P<rr1>\d{1,7})\s*:\s*(?P<rr2>\d{1,7})
      | (?P<c1>\$?[A-Za-z]{1,3})(?P<r1>\$?\d{1,7})\s*:\s*
        (?P<c2>\$?[A-Za-z]{1,3})(?P<r2>\$?\d{1,7})
      | (?P<sc>\$?[A-Za-z]{1,3})(?P<sr>\$?\d{1,7})
    )
    (?![A-Za-z0-9_$(])        # a trailing "(" means it is a function name
    """,
    re.VERBOSE,
)

# A bare name that is not part of a function call or a longer token, which is
# how a defined name appears. The trailing guard rejects the "SU" inside "SUM".
_NAME = re.compile(r"(?<![A-Za-z0-9_$.'!])(?P<name>[A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_.(!])")

# A structured reference into a table: the table name, optionally followed by
# one or more [bracketed] parts. Table names never contain spaces. The bracket
# part may nest, because the combined form is [[#Totals],[Amount]].
_TABLE = re.compile(
    r"""
    (?<![A-Za-z0-9_$.'!])       # not the tail of a longer name
    (?P<table>[A-Za-z_][A-Za-z0-9_.]*)
    (?P<brackets>(?:\[(?:[^\[\]]|\[[^\[\]]*\])*\])+)
    """,
    re.VERBOSE,
)

MAX_COL = 16384  # Excel's last column is XFD.
MAX_ROW = 1048576


def _strip_string_literals(formula: str) -> str:
    """Blank out quoted text so references inside it are not matched.

    Replaced with spaces rather than removed so match offsets stay put, which
    keeps debugging sane.
    """
    return _STRING_LITERAL.sub(lambda m: " " * len(m.group(0)), formula)


def _resolve_sheet(name: str | None, current: str, sheets: Mapping[str, str]) -> str:
    """Map a written sheet name to the workbook's spelling, case-insensitively."""
    if name is None:
        return current
    return sheets.get(name.lower(), name)


def _to_int(value: str) -> int:
    return int(value.lstrip("$"))


_SPECIFIER = re.compile(r"^#(all|data|headers|totals)$", re.IGNORECASE)
_ITEM = re.compile(r"\[([^\[\]]*)\]")


def _parse_structured(brackets: str) -> TableRef | None:
    """Turn the bracketed part of a structured reference into a TableRef.

    Handles the shapes Excel writes: ``[Amount]``, ``[#All]``, the combined
    ``[[#Totals],[Amount]]``, and the current-row ``[@Amount]``. The current-row
    form means "this row of the column", which a diff cannot pin to a cell
    without knowing the formula's own row, so only the column is kept.
    """
    items: list[str] = []
    for item in _ITEM.findall(brackets):
        inner = item.strip()
        # [@Amount] and [@[Amount]] both mean the current row of a column.
        if inner.startswith("@"):
            inner = inner[1:].strip().strip("[]").strip()
        if inner:
            items.append(inner)

    if not items:
        return None

    specifier: str | None = None
    column: str | None = None
    for item in items:
        match = _SPECIFIER.match(item)
        if match is not None:
            specifier = item
        elif column is None:
            column = item

    if specifier is None and column is None:
        return None
    return TableRef("", column, specifier)


def extract_references(
    formula: str,
    *,
    sheet: str,
    sheets: Mapping[str, str],
    defined_names: Mapping[str, str] | None = None,
    tables: Mapping[str, str] | None = None,
) -> frozenset[Reference]:
    """Return every cell and range a formula depends on.

    ``sheets`` maps a lowercased sheet name to its canonical spelling, so
    ``sheet1!a1`` and ``Sheet1!A1`` land on the same node. ``defined_names``
    maps a lowercased name to its ``refers_to`` text; a name used in the
    formula is expanded to the range behind it. ``tables`` is the set of known
    table names, lowercased; a bare name is only read as a structured reference
    when it names one, so ``SUM(x)`` is not mistaken for a table.
    """
    return frozenset(
        _scan(
            formula,
            sheet=sheet,
            sheets=sheets,
            defined_names=defined_names or {},
            tables=tables or {},
            seen=set(),
        )
    )


def _scan(
    formula: str,
    *,
    sheet: str,
    sheets: Mapping[str, str],
    defined_names: Mapping[str, str],
    tables: Mapping[str, str],
    seen: set[str],
) -> Iterable[Reference]:
    text = _strip_string_literals(formula)

    for match in _TABLE.finditer(text):
        name = match.group("table")
        if name.lower() not in tables:
            continue  # not a table we know about, so leave it to the name scan
        structured = _parse_structured(match.group("brackets"))
        if structured is not None:
            yield TableRef(tables[name.lower()], structured.column, structured.specifier)

    for match in _REF.finditer(text):
        target = _resolve_sheet(match.group("qsheet") or match.group("sheet"), sheet, sheets)
        if target.startswith("#"):
            continue  # an error literal such as '#REF'!A1
        reference = _build(match, target)
        if reference is not None:
            yield reference

    for match in _NAME.finditer(text):
        name = match.group("name").lower()
        if name in defined_names and name not in seen:
            # Expand a named range into whatever it points at. Excel does allow
            # a name to point at another name, so recurse; ``seen`` breaks the
            # chain if two names ever point at each other.
            seen.add(name)
            yield from _scan(
                defined_names[name],
                sheet=sheet,
                sheets=sheets,
                defined_names=defined_names,
                tables=tables,
                seen=seen,
            )


def _build(match: re.Match[str], sheet: str) -> Reference | None:
    if (cc1 := match.group("cc1")) and (cc2 := match.group("cc2")):
        col1, col2 = column_number(cc1), column_number(cc2)
        if not _valid(col1, 1) or not _valid(col2, MAX_ROW):
            return None
        return RangeRef(sheet, min(col1, col2), max(col1, col2), 1, MAX_ROW)

    if (rr1 := match.group("rr1")) and (rr2 := match.group("rr2")):
        row1, row2 = int(rr1), int(rr2)
        if not _valid(1, row1) or not _valid(MAX_COL, row2):
            return None
        return RangeRef(sheet, 1, MAX_COL, min(row1, row2), max(row1, row2))

    if (c1 := match.group("c1")) and (c2 := match.group("c2")):
        col1, row1 = column_number(c1), _to_int(match.group("r1") or "0")
        col2, row2 = column_number(c2), _to_int(match.group("r2") or "0")
        if not _valid(col1, row1) or not _valid(col2, row2):
            return None
        return RangeRef(sheet, min(col1, col2), max(col1, col2), min(row1, row2), max(row1, row2))

    if (sc := match.group("sc")) and (sr := match.group("sr")):
        col, row = column_number(sc), _to_int(sr)
        if not _valid(col, row):
            return None
        return CellRef(sheet, col, row)

    return None


def _valid(col: int, row: int) -> bool:
    """Reject addresses past Excel's grid, which are usually misread words."""
    return 1 <= col <= MAX_COL and 1 <= row <= MAX_ROW
