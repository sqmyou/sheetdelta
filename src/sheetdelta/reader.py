"""Read an .xlsx workbook using only the standard library.

An .xlsx file is a zip of XML parts. We read four of them: the workbook
(sheet names and defined names), the relationship map, the shared string
table, and each sheet. Nothing is evaluated -- Excel already stored the value
it last calculated, and that value is what we compare.

The reader is deliberately read-only and forgiving: a workbook that Excel can
open should not stop the diff because of an unusual cell type.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from collections.abc import Iterator
from datetime import datetime, timedelta
from xml.etree import ElementTree

from .errors import UnsupportedFormatError, WorkbookReadError
from .model import (
    Cell,
    CellKind,
    CellRef,
    Reference,
    Sheet,
    Table,
    Workbook,
    column_letter,
    column_number,
)
from .references import MAX_COL, MAX_ROW, extract_references

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_A1 = re.compile(r"([A-Za-z]{1,3})(\d{1,7})")

# Built-in number formats that mean "this number is a date or a time".
_DATE_FORMATS = frozenset(
    {14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47}
)
_DATE_TOKENS = re.compile(r"[ymdhs]", re.IGNORECASE)

_EPOCH_1900 = datetime(1899, 12, 30)
_EPOCH_1904 = datetime(1904, 1, 1)


def read_workbook(path: str) -> Workbook:
    """Parse an .xlsx file into a :class:`Workbook`.

    Raises :class:`UnsupportedFormatError` for the older binary formats and
    :class:`WorkbookReadError` for anything that is not a readable workbook.
    """
    if path.lower().endswith((".xls", ".xlsb")):
        raise UnsupportedFormatError(
            f"{path}: .xls and .xlsb are not supported, only .xlsx and .xlsm"
        )

    try:
        archive = zipfile.ZipFile(path)
    except FileNotFoundError as exc:
        raise WorkbookReadError(f"{path}: no such file") from exc
    except (zipfile.BadZipFile, OSError) as exc:
        raise WorkbookReadError(f"{path}: not a readable .xlsx file ({exc})") from exc

    with archive:
        try:
            workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        except KeyError as exc:
            raise WorkbookReadError(f"{path}: missing xl/workbook.xml, not an .xlsx file") from exc
        except ElementTree.ParseError as exc:
            raise WorkbookReadError(f"{path}: xl/workbook.xml is malformed ({exc})") from exc

        shared = _read_shared_strings(archive)
        date_formats = _read_date_formats(archive)
        targets = _read_sheet_targets(archive)
        date1904 = _uses_1904(workbook_xml)

        sheets: list[Sheet] = []
        sheet_names = [name for name, _ in _iter_sheets(workbook_xml)]
        sheet_map = {name.lower(): name for name in sheet_names}
        defined_names = _read_defined_names(workbook_xml)
        tables = _read_tables(archive, sheet_map)
        table_names = {name.lower(): table.name for name, table in tables.items()}

        for index, (name, rel_id) in enumerate(_iter_sheets(workbook_xml)):
            target = targets.get(rel_id)
            if target is None:
                continue
            sheet_path = _resolve_part(target)
            try:
                sheet_xml = ElementTree.fromstring(archive.read(sheet_path))
            except (KeyError, ElementTree.ParseError):
                # A sheet we cannot read is skipped rather than fatal; the rest
                # of the diff is still useful.
                continue
            sheets.append(
                Sheet(
                    name=name,
                    index=index,
                    cells=_read_cells(
                        sheet_xml,
                        name,
                        shared,
                        date_formats,
                        date1904,
                        sheet_map,
                        defined_names,
                        table_names,
                    ),
                )
            )

        return Workbook(
            path=path,
            sheets=sheets,
            defined_names=defined_names,
            tables=tables,
            date1904=date1904,
        )


def _resolve_part(target: str) -> str:
    """Turn a relationship target into a zip path inside the package."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def _iter_sheets(workbook_xml: ElementTree.Element) -> Iterator[tuple[str, str]]:
    for sheet in workbook_xml.iter(f"{{{_MAIN_NS}}}sheet"):
        name = sheet.get("name")
        rel_id = sheet.get(f"{{{_REL_NS}}}id")
        if name and rel_id:
            yield name, rel_id


def _read_sheet_targets(archive: zipfile.ZipFile) -> dict[str, str]:
    """Map each sheet's relationship id to the part that holds its cells."""
    try:
        rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ElementTree.ParseError):
        return {}
    targets: dict[str, str] = {}
    for rel in rels.iter(f"{{{_PKG_REL_NS}}}Relationship"):
        rel_id, target = rel.get("Id"), rel.get("Target")
        if rel_id and target:
            targets[rel_id] = target
    return targets


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    """Read the shared string table, which most text cells point into."""
    try:
        root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
    except (KeyError, ElementTree.ParseError):
        return []
    strings: list[str] = []
    for si in root.iter(f"{{{_MAIN_NS}}}si"):
        # A string may be split across runs; join every <t> inside it.
        strings.append("".join(node.text or "" for node in si.iter(f"{{{_MAIN_NS}}}t")))
    return strings


def _read_date_formats(archive: zipfile.ZipFile) -> frozenset[int]:
    """Return the style indexes whose number format looks like a date."""
    try:
        root = ElementTree.fromstring(archive.read("xl/styles.xml"))
    except (KeyError, ElementTree.ParseError):
        return frozenset()

    custom: dict[int, str] = {}
    for fmt in root.iter(f"{{{_MAIN_NS}}}numFmt"):
        fmt_id, code = fmt.get("numFmtId"), fmt.get("formatCode")
        if fmt_id and code:
            custom[int(fmt_id)] = code

    date_styles: set[int] = set()
    cell_xfs = root.find(f"{{{_MAIN_NS}}}cellXfs")
    if cell_xfs is None:
        return frozenset()
    for index, xf in enumerate(cell_xfs.findall(f"{{{_MAIN_NS}}}xf")):
        num_fmt = int(xf.get("numFmtId") or 0)
        if num_fmt in _DATE_FORMATS or (
            num_fmt in custom and _looks_like_date_format(custom[num_fmt])
        ):
            date_styles.add(index)
    return frozenset(date_styles)


def _looks_like_date_format(code: str) -> bool:
    """Guess whether a custom format code renders a date.

    Format codes contain quoted literals and colour words, so those are
    stripped before looking for date tokens; otherwise ``"Report"`` would
    count as a date because of its ``t``.
    """
    stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code)
    return bool(_DATE_TOKENS.search(stripped))


def _uses_1904(workbook_xml: ElementTree.Element) -> bool:
    props = workbook_xml.find(f"{{{_MAIN_NS}}}workbookPr")
    return props is not None and props.get("date1904") in {"1", "true"}


def _read_defined_names(workbook_xml: ElementTree.Element) -> dict[str, str]:
    """Map lowercased defined name to the reference text it stands for."""
    names: dict[str, str] = {}
    for node in workbook_xml.iter(f"{{{_MAIN_NS}}}definedName"):
        name = node.get("name")
        if name and node.text:
            names[name.lower()] = node.text.strip()
    return names


def _sheet_table_parts(archive: zipfile.ZipFile) -> dict[str, list[str]]:
    """Map a sheet's part path to the table parts it declares.

    The link runs sheet -> tablePart relationship -> table XML, so the sheet's
    own rels file has to be read before the tables themselves.
    """
    parts: dict[str, list[str]] = {}
    for sheet_path in _sheet_part_paths(archive):
        rels_path = _rels_path_for(sheet_path)
        try:
            rels = ElementTree.fromstring(archive.read(rels_path))
        except (KeyError, ElementTree.ParseError):
            continue
        targets = []
        for rel in rels.iter(f"{{{_PKG_REL_NS}}}Relationship"):
            target = rel.get("Target")
            if target and "table" in (rel.get("Type") or ""):
                targets.append(_resolve_part_for(sheet_path, target))
        if targets:
            parts[sheet_path] = targets
    return parts


def _sheet_part_paths(archive: zipfile.ZipFile) -> list[str]:
    return [
        name
        for name in archive.namelist()
        if name.startswith("xl/worksheets/") and name.endswith(".xml")
    ]


def _rels_path_for(part: str) -> str:
    directory, _, filename = part.rpartition("/")
    return f"{directory}/_rels/{filename}.rels"


def _resolve_part_for(owner: str, target: str) -> str:
    """Resolve a relationship target relative to the part that declares it."""
    if target.startswith("/"):
        return target.lstrip("/")
    directory = owner.rpartition("/")[0]
    return posixpath.normpath(posixpath.join(directory, target))


def _read_tables(archive: zipfile.ZipFile, sheet_map: dict[str, str]) -> dict[str, Table]:
    """Read every Excel table, keyed by its lowercased name.

    A table's ``ref`` is a range like ``A1:C10``, and its columns are listed in
    order, so a structured reference such as ``Table1[Amount]`` can be resolved
    to the grid column that name sits in.
    """
    tables: dict[str, Table] = {}
    part_to_sheet = _sheet_names_by_part(archive, sheet_map)

    for sheet_path, table_parts in _sheet_table_parts(archive).items():
        sheet_name = part_to_sheet.get(sheet_path)
        if sheet_name is None:
            continue
        for table_part in table_parts:
            try:
                root = ElementTree.fromstring(archive.read(table_part))
            except (KeyError, ElementTree.ParseError):
                continue
            for node in root.iter(f"{{{_MAIN_NS}}}table"):
                table = _build_table(node, sheet_name)
                if table is not None:
                    tables[table.name.lower()] = table
    return tables


def _sheet_names_by_part(
    archive: zipfile.ZipFile, sheet_map: dict[str, str]
) -> dict[str, str]:
    """Map a worksheet part path to the sheet name the workbook gives it."""
    try:
        rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ElementTree.ParseError):
        return {}
    rel_to_target = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.iter(f"{{{_PKG_REL_NS}}}Relationship")
        if rel.get("Id") and rel.get("Target")
    }
    out: dict[str, str] = {}
    for name, rel_id in _iter_sheets(_read_workbook_xml(archive)):
        target = rel_to_target.get(rel_id)
        if target:
            out[_resolve_part(target)] = sheet_map.get(name.lower(), name)
    return out


def _read_workbook_xml(archive: zipfile.ZipFile) -> ElementTree.Element:
    try:
        return ElementTree.fromstring(archive.read("xl/workbook.xml"))
    except (KeyError, ElementTree.ParseError):
        return ElementTree.Element("workbook")


def _build_table(node: ElementTree.Element, sheet_name: str) -> Table | None:
    name = node.get("name") or node.get("displayName")
    ref = node.get("ref")
    if not name or not ref:
        return None

    bounds = _range_bounds(ref)
    if bounds is None:
        return None
    min_col, max_col, min_row, max_row = bounds

    columns: list[str] = []
    for column in node.iter(f"{{{_MAIN_NS}}}tableColumn"):
        columns.append(column.get("name") or "")

    return Table(
        name=name,
        sheet=sheet_name,
        min_col=min_col,
        max_col=max_col,
        min_row=min_row,
        max_row=max_row,
        columns=tuple(columns),
        has_header_row=node.get("headerRowCount") != "0",
        has_totals_row=node.get("totalsRowCount") == "1",
    )


def _range_bounds(ref: str) -> tuple[int, int, int, int] | None:
    """Split a table's ``A1:C10`` ref into column and row bounds."""
    start, _, end = ref.partition(":")
    first = _A1.fullmatch(start.strip())
    if first is None:
        return None
    last = _A1.fullmatch(end.strip()) if end else first
    if last is None:
        return None
    return (
        column_number(first.group(1)),
        column_number(last.group(1)),
        int(first.group(2)),
        int(last.group(2)),
    )


def _read_cells(
    sheet_xml: ElementTree.Element,
    sheet_name: str,
    shared: list[str],
    date_formats: frozenset[int],
    date1904: bool,
    sheet_map: dict[str, str],
    defined_names: dict[str, str],
    table_names: dict[str, str],
) -> dict[CellRef, Cell]:
    """Read every non-empty cell of one sheet."""
    cells: dict[CellRef, Cell] = {}
    shared_formulas: dict[str, tuple[str, CellRef]] = {}

    for cell_xml in sheet_xml.iter(f"{{{_MAIN_NS}}}c"):
        ref = _parse_ref(cell_xml.get("r"), sheet_name)
        if ref is None:
            continue

        formula_xml = cell_xml.find(f"{{{_MAIN_NS}}}f")
        formula = _formula_text(formula_xml, ref, shared_formulas)
        cell_type = cell_xml.get("t") or "n"
        raw = _raw_value(cell_xml, cell_type)

        value, is_date = _cell_value(
            raw, cell_type, cell_xml.get("s"), shared, date_formats, date1904
        )

        if formula is not None:
            kind = CellKind.FORMULA
        elif value is None:
            continue  # a truly empty cell carries no signal
        else:
            kind = CellKind.VALUE

        refs: frozenset[Reference] = frozenset()
        if formula is not None:
            refs = extract_references(
                formula,
                sheet=sheet_name,
                sheets=sheet_map,
                defined_names=defined_names,
                tables=table_names,
            )

        cells[ref] = Cell(
            ref=ref,
            kind=kind,
            formula=formula,
            cached_value=value,
            value_type="date" if is_date else cell_type,
            refs=refs,
        )

    return cells


def _raw_value(cell_xml: ElementTree.Element, cell_type: str) -> str | None:
    """Return a cell's raw value text.

    Most values live in ``<v>``, but an inline string is written as
    ``<is><t>text</t></is>`` with no ``<v>`` at all. openpyxl writes strings
    that way by default, so a reader that only looks for ``<v>`` drops every
    text cell in an openpyxl-produced file.
    """
    if cell_type == "inlineStr":
        inline = cell_xml.find(f"{{{_MAIN_NS}}}is")
        if inline is None:
            return None
        return "".join(node.text or "" for node in inline.iter(f"{{{_MAIN_NS}}}t"))

    value_xml = cell_xml.find(f"{{{_MAIN_NS}}}v")
    return value_xml.text if value_xml is not None else None


def _parse_ref(raw: str | None, sheet: str) -> CellRef | None:
    if not raw:
        return None
    match = _A1.fullmatch(raw)
    if match is None:
        return None
    col, row = column_number(match.group(1)), int(match.group(2))
    return CellRef(sheet, col, row)


def _formula_text(
    formula_xml: ElementTree.Element | None,
    ref: CellRef,
    shared_formulas: dict[str, tuple[str, CellRef]],
) -> str | None:
    """Return a formula's text, resolving shared-formula blocks.

    Excel stores a repeated formula once and lets the other cells reference it,
    differing only by relative offsets. Those cells have to be shifted into
    place, or every cell in the block would look like a copy of the first.
    """
    if formula_xml is None:
        return None

    text = formula_xml.text
    kind = formula_xml.get("t")
    shared_id = formula_xml.get("si")

    if kind == "shared" and shared_id is not None:
        if text:
            shared_formulas[shared_id] = (text, ref)
            return text
        master = shared_formulas.get(shared_id)
        if master is None:
            return None
        master_text, master_ref = master
        return shift_formula(master_text, master_ref, ref)

    if kind == "array":
        # An array formula covers a range; its text lives on the anchor cell.
        return text or None

    return text or None


def shift_formula(formula: str, source: CellRef, target: CellRef) -> str:
    """Rewrite a formula's relative references from one cell to another.

    Absolute parts (``$``) stay put, which is exactly how Excel fills a
    shared formula across a block.
    """
    dcol = target.col - source.col
    drow = target.row - source.row
    if dcol == 0 and drow == 0:
        return formula

    scrubbed = _strip_literals(formula)
    out: list[str] = []
    last = 0
    for match in _BARE_REF.finditer(scrubbed):
        out.append(formula[last : match.start()])
        out.append(_shift_one(match.group(1), dcol, drow))
        last = match.end()
    out.append(formula[last:])
    return "".join(out)


def _shift_one(token: str, dcol: int, drow: int) -> str:
    parts = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})", token)
    if parts is None:
        return token
    col_dollar, letters, row_dollar, digits = parts.groups()
    col, row = column_number(letters), int(digits)
    if col_dollar != "$":
        col += dcol
    if row_dollar != "$":
        row += drow
    if not (1 <= col <= MAX_COL and 1 <= row <= MAX_ROW):
        return "#REF!"
    return f"{col_dollar}{column_letter(col)}{row_dollar}{row}"


_STRINGS = re.compile(r'"(?:[^"]|"")*"')
_BARE_REF = re.compile(r"(?<![A-Za-z0-9_$!.])(\$?[A-Za-z]{1,3}\$?\d{1,7})(?![A-Za-z0-9_$(])")


def _strip_literals(formula: str) -> str:
    return _STRINGS.sub(lambda m: " " * len(m.group(0)), formula)


def _cell_value(
    raw: str | None,
    cell_type: str,
    style: str | None,
    shared: list[str],
    date_formats: frozenset[int],
    date1904: bool,
) -> tuple[str | None, bool]:
    """Turn a raw <v> into the string a human would see, and say if it is a date."""
    if raw is None:
        return None, False

    if cell_type == "s":
        try:
            return shared[int(raw)], False
        except (ValueError, IndexError):
            return None, False
    if cell_type == "inlineStr":
        return raw, False
    if cell_type == "b":
        return ("TRUE" if raw == "1" else "FALSE"), False
    if cell_type == "e":
        return raw, False
    if cell_type == "str":
        return raw, False

    if style is not None and style.isdigit() and int(style) in date_formats:
        return _serial_to_iso(raw, date1904), True
    return raw, False


def _serial_to_iso(raw: str, date1904: bool) -> str:
    """Render an Excel serial date as an ISO string, falling back to the raw value."""
    try:
        serial = float(raw)
    except ValueError:
        return raw
    epoch = _EPOCH_1904 if date1904 else _EPOCH_1900
    moment = epoch + timedelta(days=serial)
    if moment.time() == datetime.min.time():
        return moment.date().isoformat()
    return moment.isoformat(sep=" ")
