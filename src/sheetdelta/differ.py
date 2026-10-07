"""Compare two workbooks and describe what changed.

The diff answers three questions in order: which sheets changed, which cells
changed, and which of those changes reach other cells. The third is the one
that matters -- a formula edit that nothing depends on is cosmetic, the same
edit feeding a summary table is the bug that ships.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from .model import Cell, CellKind, CellRef, RangeRef, Reference, Workbook


class ChangeKind(str, Enum):
    ADDED = "added"
    REMOVED = "removed"
    FORMULA = "formula"
    VALUE = "value"
    STALE = "stale"


class Severity(str, Enum):
    """How much a change should worry someone reading the report."""

    BREAKING = "breaking"  # reaches other cells that now hold a wrong number
    WARNING = "warning"  # a change worth a look, with no downstream reach
    INFO = "info"  # cosmetic: a new sheet, a rename


@dataclass
class CellChange:
    """One cell that differs between the two workbooks."""

    ref: CellRef
    kind: ChangeKind
    severity: Severity
    old: str | None = None
    new: str | None = None
    detail: str = ""
    affected: list[CellRef] = field(default_factory=list)

    @property
    def has_impact(self) -> bool:
        return bool(self.affected)


@dataclass
class SheetChange:
    """A sheet added, removed, renamed or edited between the two workbooks."""

    kind: str  # "added" | "removed" | "renamed" | "changed"
    name: str
    old_name: str | None = None
    cell_changes: list[CellChange] = field(default_factory=list)

    @property
    def breaking_count(self) -> int:
        return sum(1 for c in self.cell_changes if c.severity is Severity.BREAKING)


@dataclass
class DiffResult:
    """The whole comparison: sheet changes plus the counts a caller needs."""

    old_path: str
    new_path: str
    sheet_changes: list[SheetChange] = field(default_factory=list)

    @property
    def cell_changes(self) -> list[CellChange]:
        return [c for s in self.sheet_changes for c in s.cell_changes]

    @property
    def breaking(self) -> list[CellChange]:
        return [c for c in self.cell_changes if c.severity is Severity.BREAKING]

    @property
    def has_changes(self) -> bool:
        return bool(self.sheet_changes)

    def counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for change in self.cell_changes:
            out[change.kind.value] = out.get(change.kind.value, 0) + 1
        return out


def diff_workbooks(old: Workbook, new: Workbook) -> DiffResult:
    """Compare two parsed workbooks."""
    result = DiffResult(old_path=old.path, new_path=new.path)

    renames = _match_renamed_sheets(old, new)
    renamed_from = {old_name for old_name, _ in renames}
    renamed_to = {new_name for _, new_name in renames}

    # The dependency graph spans the whole workbook, not one sheet, so a change
    # in Revenue is followed into the Summary that reads it.
    dependents = _global_dependents(old, new, renames)

    for old_name, new_name in renames:
        old_sheet = old.sheet_by_name(old_name)
        new_sheet = new.sheet_by_name(new_name)
        assert old_sheet is not None and new_sheet is not None
        result.sheet_changes.append(
            SheetChange(
                kind="renamed",
                name=new_name,
                old_name=old_name,
                cell_changes=_diff_cells(
                    _rehome(old_sheet.cells, old_name, new_name),
                    new_sheet.cells,
                    dependents,
                ),
            )
        )

    for name in old.sheet_names:
        if name in renamed_from or new.sheet_by_name(name) is not None:
            continue
        result.sheet_changes.append(
            SheetChange(kind="removed", name=name, old_name=name)
        )

    for name in new.sheet_names:
        if name in renamed_to:
            continue
        old_sheet = old.sheet_by_name(name)
        if old_sheet is None:
            result.sheet_changes.append(SheetChange(kind="added", name=name))
            continue
        new_sheet = new.sheet_by_name(name)
        assert new_sheet is not None
        changes = _diff_cells(old_sheet.cells, new_sheet.cells, dependents)
        if changes:
            result.sheet_changes.append(
                SheetChange(kind="changed", name=name, old_name=name, cell_changes=changes)
            )

    return result


def _global_dependents(
    old: Workbook, new: Workbook, renames: list[tuple[str, str]]
) -> dict[CellRef, set[CellRef]]:
    """Build the reverse dependency map across every sheet.

    Both versions contribute, because a cell deleted in the new file still has
    readers there, and those readers are exactly what the deletion breaks. When
    a sheet was renamed its old cells are re-keyed so the two versions line up.
    """
    old_names = dict(renames)
    combined: dict[CellRef, Cell] = {}

    for workbook, name_map in ((old, old_names), (new, {})):
        for sheet in workbook.sheets:
            target = name_map.get(sheet.name, sheet.name)
            for ref, cell in sheet.cells.items():
                moved = CellRef(target, ref.col, ref.row)
                combined[moved] = Cell(
                    ref=moved,
                    kind=cell.kind,
                    formula=cell.formula,
                    cached_value=cell.cached_value,
                    value_type=cell.value_type,
                    refs=frozenset(
                        _rehome_ref(reference, sheet.name, target) for reference in cell.refs
                    ),
                )

    return _dependents(combined)


def _match_renamed_sheets(old: Workbook, new: Workbook) -> list[tuple[str, str]]:
    """Pair a removed sheet with an added one when their contents match.

    A rename shows up in the file format as a delete plus an insert. Reporting
    that as "everything in this sheet changed" would bury the real diff, so
    sheets with identical contents are reported as renamed instead.
    """
    removed = [n for n in old.sheet_names if new.sheet_by_name(n) is None]
    added = [n for n in new.sheet_names if old.sheet_by_name(n) is None]

    pairs: list[tuple[str, str]] = []
    used: set[str] = set()
    for old_name in removed:
        old_sheet = old.sheet_by_name(old_name)
        if old_sheet is None:
            continue
        for new_name in added:
            if new_name in used:
                continue
            new_sheet = new.sheet_by_name(new_name)
            if new_sheet is None:
                continue
            if _same_sheet(old_sheet.cells, new_sheet.cells, new_name):
                pairs.append((old_name, new_name))
                used.add(new_name)
                break
    return pairs


def _same_sheet(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell], new_name: str
) -> bool:
    """Whether two sheets hold the same cells, ignoring which sheet they sit on."""
    if len(old_cells) != len(new_cells):
        return False
    for ref, cell in old_cells.items():
        other = new_cells.get(CellRef(new_name, ref.col, ref.row))
        if other is None:
            return False
        if cell.formula != other.formula or cell.cached_value != other.cached_value:
            return False
    return True


def _rehome(cells: dict[CellRef, Cell], old_name: str, new_name: str) -> dict[CellRef, Cell]:
    """Rewrite a sheet's cells as if the sheet had always had its new name.

    Without this, a renamed sheet compares cells keyed ``Data!A1`` against
    ``Numbers!A1`` and reports the whole sheet as replaced.
    """
    out: dict[CellRef, Cell] = {}
    for ref, cell in cells.items():
        moved = CellRef(new_name, ref.col, ref.row)
        refs = frozenset(_rehome_ref(reference, old_name, new_name) for reference in cell.refs)
        out[moved] = Cell(
            ref=moved,
            kind=cell.kind,
            formula=cell.formula,
            cached_value=cell.cached_value,
            value_type=cell.value_type,
            refs=refs,
        )
    return out


def _rehome_ref(reference: Reference, old_name: str, new_name: str) -> Reference:
    if reference.sheet != old_name:
        return reference
    if isinstance(reference, CellRef):
        return CellRef(new_name, reference.col, reference.row)
    return RangeRef(
        new_name, reference.min_col, reference.max_col, reference.min_row, reference.max_row
    )


def _diff_cells(
    old_cells: dict[CellRef, Cell],
    new_cells: dict[CellRef, Cell],
    dependents: dict[CellRef, set[CellRef]],
) -> list[CellChange]:
    """Compare the cells of one sheet, in reading order."""
    changes: list[CellChange] = []

    for ref in sorted(set(old_cells) | set(new_cells), key=lambda r: (r.row, r.col)):
        old = old_cells.get(ref)
        new = new_cells.get(ref)

        if old is None and new is not None:
            changes.append(
                CellChange(ref=ref, kind=ChangeKind.ADDED, severity=Severity.INFO, new=new.display)
            )
        elif new is None and old is not None:
            affected = _reachable(ref, dependents)
            changes.append(
                CellChange(
                    ref=ref,
                    kind=ChangeKind.REMOVED,
                    severity=Severity.BREAKING if affected else Severity.WARNING,
                    old=old.display,
                    detail="cell deleted",
                    affected=affected,
                )
            )
        elif old is not None and new is not None:
            changes.extend(_compare_cell(ref, old, new, dependents))

    return changes


def _compare_cell(
    ref: CellRef, old: Cell, new: Cell, dependents: dict[CellRef, set[CellRef]]
) -> list[CellChange]:
    """Describe how one cell changed, if it did."""
    if old.formula != new.formula:
        affected = _reachable(ref, dependents)
        change = CellChange(
            ref=ref,
            kind=ChangeKind.FORMULA,
            severity=Severity.BREAKING if affected else Severity.WARNING,
            old=f"={old.formula}" if old.formula else old.display,
            new=f"={new.formula}" if new.formula else new.display,
            detail="formula changed",
            affected=affected,
        )
        # A formula edit whose cached value did not move means the workbook was
        # saved without recalculating. Excel would show the new formula and the
        # old number side by side, which is exactly how a wrong total ships. A
        # file that never held a cached value is a different case: it is not
        # stale, it was never calculated, so it stays a plain formula change.
        if old.cached_value is not None and old.cached_value == new.cached_value:
            change.kind = ChangeKind.STALE
            change.detail = "formula changed, cached value did not move (not recalculated)"
        return [change]

    if old.cached_value != new.cached_value:
        affected = _reachable(ref, dependents)
        return [
            CellChange(
                ref=ref,
                kind=ChangeKind.VALUE,
                severity=Severity.BREAKING if affected else Severity.WARNING,
                old=old.cached_value,
                new=new.cached_value,
                detail="value changed",
                affected=affected,
            )
        ]

    return []


def _dependents(cells: dict[CellRef, Cell]) -> dict[CellRef, set[CellRef]]:
    """Build the reverse dependency map for a set of cells.

    Maps each referenced cell to the cells that read it, so a change can be
    followed forward to everything it reaches. A formula that reads its own
    cell -- ``D12`` inside ``SUM(D2:D13)``, for instance -- is not a dependency
    of itself and is left out.
    """
    out: dict[CellRef, set[CellRef]] = {}
    for cell in cells.values():
        if cell.kind is not CellKind.FORMULA or not cell.refs:
            continue
        for reference in cell.refs:
            for target in _targets_in(reference, cells):
                if target != cell.ref:
                    out.setdefault(target, set()).add(cell.ref)
    return out


def _targets_in(reference: Reference, cells: dict[CellRef, Cell]) -> list[CellRef]:
    """Return the cells in ``cells`` that a reference points at."""
    return [ref for ref in cells if reference.contains(ref)]


def _reachable(ref: CellRef, dependents: dict[CellRef, set[CellRef]]) -> list[CellRef]:
    """Every cell that transitively reads ``ref``, in reading order."""
    seen: set[CellRef] = set()
    queue = list(dependents.get(ref, ()))
    while queue:
        current = queue.pop()
        if current in seen:
            continue
        seen.add(current)
        queue.extend(dependents.get(current, ()))
    return sorted(seen, key=lambda r: (r.sheet, r.row, r.col))
