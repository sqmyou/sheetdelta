"""Build small .xlsx files for tests, using only the standard library.

Tests should not depend on openpyxl. Writing the XML by hand also keeps the
fixtures honest: they contain exactly the parts the reader claims to handle,
including a shared formula and a date-formatted number, so a regression in
the reader shows up here rather than in a real workbook.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"


@dataclass
class FakeCell:
    ref: str
    value: str | None = None
    formula: str | None = None
    shared: str | None = None  # shared formula index
    is_string: bool = False
    inline_string: bool = False  # written as <is><t>, the way openpyxl does
    style: int | None = None


@dataclass
class FakeSheet:
    name: str
    cells: list[FakeCell] = field(default_factory=list)
    tables: list[FakeTable] = field(default_factory=list)


@dataclass
class FakeTable:
    """An Excel table: the block it covers and its column names."""

    name: str
    ref: str  # e.g. "A1:B5"
    columns: list[str]
    header_row: bool = True
    totals_row: bool = False


def write_workbook(
    path: str,
    sheets: list[FakeSheet],
    *,
    shared_strings: list[str] | None = None,
    defined_names: dict[str, str] | None = None,
    date_styles: tuple[int, ...] = (),
    date1904: bool = False,
) -> str:
    """Write a minimal but valid .xlsx file and return its path."""
    shared_strings = shared_strings or []
    defined_names = defined_names or {}

    table_parts: dict[int, list[tuple[str, FakeTable]]] = {}
    for index, sheet in enumerate(sheets, start=1):
        if sheet.tables:
            table_parts[index] = [
                (f"xl/tables/table{index}_{position}.xml", table)
                for position, table in enumerate(sheet.tables, start=1)
            ]

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(table_parts))
        archive.writestr("_rels/.rels", _root_rels())
        archive.writestr("xl/workbook.xml", _workbook(sheets, defined_names, date1904))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(sheets))
        if shared_strings:
            archive.writestr("xl/sharedStrings.xml", _shared_strings(shared_strings))
        archive.writestr("xl/styles.xml", _styles(date_styles))
        for index, sheet in enumerate(sheets, start=1):
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml", _sheet(sheet, table_parts.get(index, []))
            )
            if index in table_parts:
                archive.writestr(
                    f"xl/worksheets/_rels/sheet{index}.xml.rels",
                    _sheet_rels(table_parts[index]),
                )
        for parts in table_parts.values():
            for part, table in parts:
                archive.writestr(part, _table(table))
    return path


def _content_types(table_parts: dict[int, list[tuple[str, FakeTable]]]) -> str:
    overrides = "".join(
        f'<Override PartName="/{part}" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.table+xml"/>'
        for parts in table_parts.values()
        for part, _ in parts
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Types xmlns="{CONTENT_TYPES}">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{overrides}"
        "</Types>"
    )


def _sheet_rels(parts: list[tuple[str, FakeTable]]) -> str:
    rels = "".join(
        f'<Relationship Id="rIdTable{position}" Type="{REL}/table" '
        f'Target="../tables/{part.rsplit("/", 1)[-1]}"/>'
        for position, (part, _) in enumerate(parts, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{PKG_REL}">{rels}</Relationships>'
    )


def _table(table: FakeTable) -> str:
    columns = "".join(
        f'<tableColumn id="{i}" name="{_esc(name)}"/>'
        for i, name in enumerate(table.columns, start=1)
    )
    attrs = f'id="1" name="{_esc(table.name)}" displayName="{_esc(table.name)}" ref="{table.ref}"'
    if not table.header_row:
        attrs += ' headerRowCount="0"'
    if table.totals_row:
        attrs += ' totalsRowCount="1"'
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<table xmlns="{MAIN}" {attrs}>'
        f'<tableColumns count="{len(table.columns)}">{columns}</tableColumns>'
        "</table>"
    )


def _root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def _workbook(sheets: list[FakeSheet], defined_names: dict[str, str], date1904: bool) -> str:
    props = '<workbookPr date1904="1"/>' if date1904 else ""
    sheet_xml = "".join(
        f'<sheet name="{_esc(sheet.name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, sheet in enumerate(sheets, start=1)
    )
    names = "".join(
        f'<definedName name="{_esc(name)}">{_esc(text)}</definedName>'
        for name, text in defined_names.items()
    )
    defined = f"<definedNames>{names}</definedNames>" if names else ""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<workbook xmlns="{MAIN}" xmlns:r="{REL}">'
        f"{props}<sheets>{sheet_xml}</sheets>{defined}"
        "</workbook>"
    )


def _workbook_rels(sheets: list[FakeSheet]) -> str:
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{REL}/worksheet" '
        f'Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(sheets) + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{PKG_REL}">{rels}</Relationships>'
    )


def _shared_strings(strings: list[str]) -> str:
    items = "".join(f"<si><t>{_esc(s)}</t></si>" for s in strings)
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<sst xmlns="{MAIN}" count="{len(strings)}" uniqueCount="{len(strings)}">'
        f"{items}</sst>"
    )


def _styles(date_styles: tuple[int, ...]) -> str:
    xfs = "".join(
        f'<xf numFmtId="{14 if i in date_styles else 0}"/>'
        for i in range(max(date_styles, default=-1) + 1 or 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<styleSheet xmlns="{MAIN}"><cellXfs count="1">{xfs}</cellXfs></styleSheet>'
    )


def _sheet(sheet: FakeSheet, table_parts: list[tuple[str, FakeTable]]) -> str:
    cells = "".join(_cell(cell) for cell in sheet.cells)
    table_parts_xml = ""
    if table_parts:
        parts = "".join(
            f'<tablePart r:id="rIdTable{position}"/>'
            for position, _ in enumerate(table_parts, start=1)
        )
        table_parts_xml = f'<tableParts count="{len(table_parts)}">{parts}</tableParts>'
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
        f"<sheetData>{cells}</sheetData>{table_parts_xml}</worksheet>"
    )


def _cell(cell: FakeCell) -> str:
    attrs = f'r="{cell.ref}"'
    if cell.is_string:
        attrs += ' t="s"'
    if cell.inline_string:
        attrs += ' t="inlineStr"'
    if cell.style is not None:
        attrs += f' s="{cell.style}"'

    body = ""
    if cell.formula is not None:
        if cell.shared is not None:
            body += f'<f t="shared" si="{cell.shared}">{_esc(cell.formula)}</f>'
        else:
            body += f"<f>{_esc(cell.formula)}</f>"
    elif cell.shared is not None:
        body += f'<f t="shared" si="{cell.shared}"/>'

    if cell.inline_string:
        body += f"<is><t>{_esc(cell.value or '')}</t></is>"
    elif cell.value is not None:
        body += f"<v>{_esc(cell.value)}</v>"
    return f"<c {attrs}>{body}</c>"


def _esc(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
