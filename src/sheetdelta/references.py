"""Pull cell and range references out of an Excel formula.

This is a scanner, not a formula evaluator. It needs to be accurate enough to
build a dependency graph, which is why it is fussier than a one-line regex:
``LOG10(...)`` looks exactly like a reference to column LOG row 10, and the
text ``"see A1"`` looks like a reference to A1. Both are handled here.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping

from .model import CellRef, RangeRef, Reference, Table, TableRef, VolatileRef, column_number

# Excel string literals are double-quoted, and a literal quote is doubled.
_STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')

# Functions that compute their target at runtime. A static scan cannot say which
# cells they read, so the reference is recorded as a volatile one instead of
# being dropped. The argument is captured only to report *why* the target is
# unknown; a constant string still does not make the target static, because the
# address it names can be edited without touching the formula.
_VOLATILE = re.compile(r"(?<![A-Za-z0-9_.])(?P<kind>INDIRECT|OFFSET)\s*\(", re.IGNORECASE)

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


def _parse_structured(brackets: str) -> tuple[str | None, str | None, bool]:
    """Turn the bracketed part of a structured reference into its parts.

    Handles the shapes Excel writes: ``[Amount]``, ``[#All]``, the combined
    ``[[#Totals],[Amount]]``, and the current-row ``[@Amount]``. Returns the
    specifier, the column, and whether the current-row form was used; the row
    itself is filled in from the cell the formula sits in.
    """
    items: list[str] = []
    current_row = False
    for item in _ITEM.findall(brackets):
        inner = item.strip()
        # [@Amount] and [@[Amount]] both mean the current row of a column.
        if inner.startswith("@"):
            current_row = True
            inner = inner[1:].strip().strip("[]").strip()
        if inner:
            items.append(inner)

    if not items:
        return None, None, current_row

    specifier: str | None = None
    column: str | None = None
    for item in items:
        match = _SPECIFIER.match(item)
        if match is not None:
            specifier = item
        elif column is None:
            column = item

    return specifier, column, current_row


def _validate(reference: Reference, tables: Mapping[str, Table]) -> Reference | None:
    """Drop a structured reference that cannot be resolved against its table.

    A current-row reference is pinned to the row the formula sits in. If that
    row is outside the table -- the table is on a different sheet, or the formula
    sits above or below it -- the reference cannot be placed and is dropped
    rather than pointed at a row of the wrong table. A column that does not exist
    is dropped for the same reason.
    """
    if not isinstance(reference, TableRef):
        return reference
    table = tables.get(reference.table.lower())
    if table is None:
        return None
    if reference.column is not None and table.column_number(reference.column) is None:
        return None
    if reference.row is not None and not (table.first_row <= reference.row <= table.last_row):
        return None
    return reference


def extract_references(
    formula: str,
    *,
    sheet: str,
    sheets: Mapping[str, str],
    defined_names: Mapping[str, str] | None = None,
    tables: Mapping[str, Table] | None = None,
    row: int | None = None,
    volatile_scope: str = "sheet",
) -> frozenset[Reference]:
    """Return every cell and range a formula depends on.

    ``sheets`` maps a lowercased sheet name to its canonical spelling, so
    ``sheet1!a1`` and ``Sheet1!A1`` land on the same node. ``defined_names``
    maps a lowercased name to its ``refers_to`` text; a name used in the
    formula is expanded to the range behind it. ``tables`` maps a lowercased
    table name to its definition; a bare name is only read as a structured
    reference when it names a table there, so ``SUM(x)`` is not mistaken for one.
    ``row`` is the row the formula sits in, which a current-row structured
    reference such as ``[@Amount]`` needs. ``volatile_scope`` says how wide a
    computed reference (``INDIRECT``, ``OFFSET``) is treated as reaching; see
    :class:`~sheetdelta.model.VolatileRef`.
    """
    return frozenset(
        _scan(
            formula,
            sheet=sheet,
            sheets=sheets,
            defined_names=defined_names or {},
            tables=tables or {},
            seen=set(),
            row=row,
            volatile_scope=volatile_scope,
        )
    )


def _first_argument(formula: str, open_paren: int) -> str | None:
    """The first argument if it is a single string literal, else None.

    ``formula`` is the original text; ``open_paren`` is the offset of the ``(``.
    Only a lone quoted argument is returned, because it is the one case where
    the report can show what the formula was reaching for. A computed argument
    such as ``"A"&B1`` is not a literal target, and an address is left to the
    ordinary reference scan, so both come back as None.
    """
    i = open_paren + 1
    while i < len(formula) and formula[i].isspace():
        i += 1
    if i >= len(formula) or formula[i] != '"':
        return None
    j = i + 1
    while j < len(formula):
        if formula[j] == '"':
            if j + 1 < len(formula) and formula[j + 1] == '"':
                j += 2
                continue
            after = j + 1
            while after < len(formula) and formula[after].isspace():
                after += 1
            if after >= len(formula) or formula[after] not in ",)":
                return None  # part of a larger expression, not a lone literal
            return formula[i + 1 : j].replace('""', '"')
        j += 1
    return None


def _scan(
    formula: str,
    *,
    sheet: str,
    sheets: Mapping[str, str],
    defined_names: Mapping[str, str],
    tables: Mapping[str, Table],
    seen: set[str],
    row: int | None,
    volatile_scope: str,
) -> Iterable[Reference]:
    text = _strip_string_literals(formula)

    # A computed reference is recorded against the original text, because the
    # argument that explains it is a string literal the stripped text blanks out.
    # A match inside a string literal, e.g. "see INDIRECT(a1)", is not a call:
    # the stripped text has it blanked to spaces at the same offset, so a
    # differing character there means the name is inside the literal.
    for match in _VOLATILE.finditer(formula):
        if match.start() < len(text) and text[match.start()] != formula[match.start()]:
            continue
        literal = _first_argument(formula, match.end() - 1)
        yield VolatileRef(sheet, match.group("kind").upper(), volatile_scope, literal)

    for match in _TABLE.finditer(text):
        name = match.group("table")
        table = tables.get(name.lower())
        if table is None:
            continue  # not a table we know about, so leave it to the name scan
        specifier, column, current_row = _parse_structured(match.group("brackets"))
        if specifier is None and column is None:
            continue
        # A current-row reference only means a row when the table is on the
        # formula's own sheet; otherwise there is no row to pin it to.
        pinned = row if current_row and table.sheet == sheet else None
        structured: Reference = TableRef(table.name, column, specifier, pinned)
        validated = _validate(structured, tables)
        if validated is not None:
            yield validated

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
                row=row,
                volatile_scope=volatile_scope,
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
