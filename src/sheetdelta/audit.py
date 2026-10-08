"""Inspect a single workbook for defects a diff cannot see.

An audit answers "is this file sound", not "what changed". Two defects are
worth reporting, and both are unambiguous rather than heuristic:

* a formula pointing at a sheet that does not exist, which Excel shows as
  ``#REF!``; and
* a circular reference, a cell that depends on itself through some chain of
  other cells, which Excel refuses to calculate at all.

Deliberately not checked: a reference to a cell that is merely empty. That is
normal in a spreadsheet, and flagging it would make the audit cry wolf on
every real file.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .model import (
    Cell,
    CellIndex,
    CellKind,
    CellRef,
    Reference,
    TableRef,
    VolatileRef,
    Workbook,
)


@dataclass
class BrokenReference:
    """A formula reference to a sheet that is not in the workbook."""

    ref: CellRef
    formula: str
    target: str
    reason: str


@dataclass
class VolatileReference:
    """A formula whose dependencies are computed at runtime.

    ``INDIRECT`` and ``OFFSET`` build their target from a string or an offset,
    so the cells they read cannot be known without evaluating the formula. The
    audit cannot say the workbook is sound when a formula's dependencies are
    partly invisible, so it reports the call site instead of staying silent.
    """

    ref: CellRef
    formula: str
    kind: str
    literal: str | None = None


@dataclass
class CircularReference:
    """A cycle of cells that each depend on the next, ending where it started."""

    cells: list[CellRef]

    @property
    def display(self) -> str:
        return " -> ".join(str(ref) for ref in self.cells)


@dataclass
class AuditResult:
    path: str
    sheets: list[str] = field(default_factory=list)
    formula_count: int = 0
    cell_count: int = 0
    incomplete: list[str] = field(default_factory=list)
    defined_names: dict[str, str] = field(default_factory=dict)
    broken: list[BrokenReference] = field(default_factory=list)
    cycles: list[CircularReference] = field(default_factory=list)
    volatile: list[VolatileReference] = field(default_factory=list)

    @property
    def is_sound(self) -> bool:
        return not self.broken and not self.cycles and not self.incomplete

    @property
    def issue_count(self) -> int:
        return len(self.broken) + len(self.cycles) + (1 if self.incomplete else 0)


def audit_workbook(workbook: Workbook) -> AuditResult:
    """Check a workbook for broken references and circular references."""
    result = AuditResult(
        path=workbook.path,
        sheets=workbook.sheet_names,
        incomplete=list(workbook.partial_reads),
        defined_names=dict(workbook.defined_names),
    )

    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            result.cell_count += 1
            if cell.kind is CellKind.FORMULA:
                result.formula_count += 1
                result.broken.extend(_missing_sheets(cell, workbook))
                result.volatile.extend(_volatile_refs(cell))

    result.broken.sort(key=lambda b: (b.ref.sheet, b.ref.row, b.ref.col))
    result.volatile.sort(key=lambda v: (v.ref.sheet, v.ref.row, v.ref.col, v.kind))
    result.cycles = _find_cycles(workbook)
    return result


def _volatile_refs(cell: Cell) -> list[VolatileReference]:
    """One finding per computed reference in a formula, deduplicated by kind."""
    out: list[VolatileReference] = []
    seen: set[tuple[str, str | None]] = set()
    for reference in sorted(cell.refs, key=str):
        if not isinstance(reference, VolatileRef):
            continue
        key = (reference.kind, reference.literal)
        if key in seen:
            continue
        seen.add(key)
        out.append(
            VolatileReference(
                ref=cell.ref,
                formula=cell.formula or "",
                kind=reference.kind,
                literal=reference.literal,
            )
        )
    return out


def _missing_sheets(cell: Cell, workbook: Workbook) -> list[BrokenReference]:
    """Every reference in one formula that names a sheet or table that is not there."""
    out: list[BrokenReference] = []
    for reference in sorted(cell.refs, key=str):
        if isinstance(reference, VolatileRef):
            # A computed reference names no sheet or table to check; it is
            # reported separately, under volatile references.
            continue
        if isinstance(reference, TableRef):
            if reference.table.lower() not in workbook.tables:
                out.append(
                    BrokenReference(
                        ref=cell.ref,
                        formula=cell.formula or "",
                        target=str(reference),
                        reason=f"table '{reference.table}' is not defined",
                    )
                )
            continue
        if workbook.sheet_by_name(reference.sheet) is None:
            out.append(
                BrokenReference(
                    ref=cell.ref,
                    formula=cell.formula or "",
                    target=str(reference),
                    reason=f"sheet '{reference.sheet}' does not exist",
                )
            )
    return out


def _find_cycles(workbook: Workbook) -> list[CircularReference]:
    """Find circular references across the whole workbook.

    Walks the dependency graph depth-first, tracking the cells on the current
    path. Reaching a cell already on the path closes a cycle, and the slice of
    the path from that cell onwards is the loop to report. The walk keeps its
    own stack rather than recursing: a running-balance column chains thousands
    of cells deep and would otherwise hit Python's recursion limit.
    """
    graph = _dependency_graph(workbook)
    found: list[CircularReference] = []
    seen: set[CellRef] = set()
    reported: set[frozenset[CellRef]] = set()

    for start in sorted(graph, key=_order):
        if start in seen:
            continue
        _walk(start, graph, seen, found, reported)

    return found


def _order(ref: CellRef) -> tuple[str, int, int]:
    return (ref.sheet, ref.row, ref.col)


def _walk(
    start: CellRef,
    graph: dict[CellRef, set[CellRef]],
    seen: set[CellRef],
    found: list[CircularReference],
    reported: set[frozenset[CellRef]],
) -> None:
    path: list[CellRef] = [start]
    on_path: set[CellRef] = {start}
    stack: list[tuple[CellRef, list[CellRef], int]] = [
        (start, sorted(graph.get(start, ()), key=_order), 0)
    ]

    while stack:
        node, neighbours, index = stack[-1]
        if index < len(neighbours):
            stack[-1] = (node, neighbours, index + 1)
            neighbour = neighbours[index]
            if neighbour in on_path:
                cycle = path[path.index(neighbour) :]
                key = frozenset(cycle)
                if key not in reported:
                    reported.add(key)
                    found.append(CircularReference(cells=[*cycle, neighbour]))
            elif neighbour not in seen:
                path.append(neighbour)
                on_path.add(neighbour)
                stack.append((neighbour, sorted(graph.get(neighbour, ()), key=_order), 0))
        else:
            stack.pop()
            path.pop()
            on_path.discard(node)
            seen.add(node)


def _dependency_graph(workbook: Workbook) -> dict[CellRef, set[CellRef]]:
    """Map each cell to the cells it reads, ignoring self-reads and blanks.

    Cross-sheet references are followed to the sheet they name, so a cycle that
    runs from one sheet into another is still found.
    """
    graph: dict[CellRef, set[CellRef]] = {}
    index = CellIndex(ref for sheet in workbook.sheets for ref in sheet.cells)
    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            if cell.kind is not CellKind.FORMULA:
                continue
            targets: set[CellRef] = set()
            for reference in cell.refs:
                if isinstance(reference, TableRef):
                    resolved: Reference | None = reference.resolve(workbook.tables)
                    if resolved is None:
                        continue
                else:
                    if workbook.sheet_by_name(reference.sheet) is None:
                        continue
                    resolved = reference
                targets.update(
                    ref for ref in index.covered(resolved) if ref != cell.ref
                )
            if targets:
                graph[cell.ref] = targets
    return graph
