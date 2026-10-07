"""Compare two workbooks and describe what changed.

The diff answers three questions in order: which sheets changed, which cells
changed, and which of those changes reach other cells. The third is the one
that matters -- a formula edit that nothing depends on is cosmetic, the same
edit feeding a summary table is the bug that ships.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from enum import Enum

from .model import (
    Cell,
    CellIndex,
    CellKind,
    CellRef,
    RangeRef,
    Reference,
    Table,
    TableRef,
    Workbook,
)


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
class Shift:
    """A block of rows or columns inserted or removed, moving everything after it.

    Excel has no "insert row" in the file format: it rewrites every cell below
    to a new address. Comparing addresses directly therefore reports the whole
    lower half of the sheet as changed. Detecting the shift lets the report say
    "a row was inserted at 5" instead, and lets the cells that merely moved line
    up so the real edits stand out.
    """

    axis: str  # "row" or "column"
    inserted: bool
    count: int
    at: int  # 1-based first row/column of the inserted or removed block

    @property
    def label(self) -> str:
        noun = "row" if self.axis == "row" else "column"
        plural = "" if self.count == 1 else "s"
        verb = "inserted" if self.inserted else "removed"
        return f"{self.count} {noun}{plural} {verb} at {noun} {self.at}"


@dataclass
class SheetChange:
    """A sheet added, removed, renamed or edited between the two workbooks."""

    kind: str  # "added" | "removed" | "renamed" | "changed"
    name: str
    old_name: str | None = None
    cell_changes: list[CellChange] = field(default_factory=list)
    shift: Shift | None = None

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
        cell_changes, shift = _diff_cells(
            _rehome(old_sheet.cells, old_name, new_name),
            new_sheet.cells,
            dependents,
        )
        result.sheet_changes.append(
            SheetChange(
                kind="renamed",
                name=new_name,
                old_name=old_name,
                cell_changes=cell_changes,
                shift=shift,
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
        changes, shift = _diff_cells(old_sheet.cells, new_sheet.cells, dependents)
        if changes or shift is not None:
            result.sheet_changes.append(
                SheetChange(
                    kind="changed",
                    name=name,
                    old_name=name,
                    cell_changes=changes,
                    shift=shift,
                )
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

    # Tables from both versions are merged: a formula in the old file may read a
    # table the new file has dropped, and that reader is what the drop breaks.
    tables = {**old.tables, **new.tables}
    return _dependents(combined, tables)


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
    if isinstance(reference, TableRef):
        return reference  # a structured reference names a table, not a sheet
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
) -> tuple[list[CellChange], Shift | None]:
    """Compare the cells of one sheet, in reading order.

    Returns the changes plus, when the sheet shifted, the shift that was found.
    The old cells are re-keyed onto their new addresses first, so a cell that
    only moved is not reported as changed.
    """
    shift = _detect_shift(old_cells, new_cells)
    if shift is not None:
        old_cells = _apply_shift(old_cells, shift)

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

    return changes, shift


def _detect_shift(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell]
) -> Shift | None:
    """Find a row or column insert/remove that explains most of the difference.

    A shift is only reported when the rows or columns after it agree once moved,
    which is what separates an insert from an edit: two unrelated sheets with
    different heights must not be described as one inserting rows into the other.
    """
    for axis in ("row", "column"):
        shift = _detect_axis(old_cells, new_cells, axis)
        if shift is not None:
            return shift
    return None


def _detect_axis(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell], axis: str
) -> Shift | None:
    """Find a shift on one axis by lining the contents up, not the addresses.

    Addresses alone cannot say where an insert happened: inserting a row pushes
    every later row up, so the set of new addresses looks the same whether the
    row went in at the top or the bottom. What identifies the position is the
    first line whose contents differ -- everything above it stayed put. The
    insert is then confirmed by checking the old line reappears further down
    with everything after it moved by the same amount.
    """
    old_sig = _line_signatures(old_cells, axis)
    new_sig = _line_signatures(new_cells, axis)
    if old_sig == new_sig:
        return None

    lines = sorted(set(old_sig) | set(new_sig))
    first = next(i for i in lines if old_sig.get(i) != new_sig.get(i))

    for inserted in (True, False):
        source = old_sig if inserted else new_sig
        target = new_sig if inserted else old_sig
        if first not in source:
            continue
        anchor = next(
            (line for line in sorted(target) if line > first and target[line] == source[first]),
            None,
        )
        if anchor is None:
            continue
        shift = Shift(axis=axis, inserted=inserted, count=anchor - first, at=first)
        if _shifts_agree(old_cells, new_cells, shift):
            return shift
    return None


def _line_signatures(
    cells: dict[CellRef, Cell], axis: str
) -> dict[int, tuple[tuple[int, str], ...]]:
    """Describe each row (or column) by what it holds, ignoring its address.

    Two lines with the same signature hold the same contents, which is what
    makes them comparable across the shift.
    """
    gathered: dict[int, list[tuple[int, str]]] = {}
    for ref, cell in cells.items():
        line = ref.row if axis == "row" else ref.col
        other = ref.col if axis == "row" else ref.row
        gathered.setdefault(line, []).append((other, cell.display))
    return {line: tuple(sorted(items)) for line, items in gathered.items()}


def _shifts_agree(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell], shift: Shift
) -> bool:
    """Check that the moved cells line up once shifted.

    A shift is a claim about the whole block below it, so most of that block has
    to agree. Requiring every cell to match would reject a real insert that
    came with an edit; requiring only a majority keeps the claim honest while
    letting the edit through as a change in its own right.
    """
    moved = _apply_shift(old_cells, shift)
    matched = mismatched = 0
    for ref, old in moved.items():
        line = ref.row if shift.axis == "row" else ref.col
        if line < shift.at:
            continue
        new = new_cells.get(ref)
        if new is not None and new.display == old.display:
            matched += 1
        else:
            mismatched += 1
    return matched > 0 and mismatched <= matched


def _apply_shift(cells: dict[CellRef, Cell], shift: Shift) -> dict[CellRef, Cell]:
    """Re-key the cells at or after the shift onto their new addresses."""
    out: dict[CellRef, Cell] = {}
    for ref, cell in cells.items():
        coordinate = ref.row if shift.axis == "row" else ref.col
        if coordinate < shift.at:
            out[ref] = cell
            continue
        if not shift.inserted and coordinate < shift.at + shift.count:
            continue  # the removed lines are gone
        delta = shift.count if shift.inserted else -shift.count
        if shift.axis == "row":
            moved = CellRef(ref.sheet, ref.col, ref.row + delta)
        else:
            moved = CellRef(ref.sheet, ref.col + delta, ref.row)
        out[moved] = cell
    return out


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


def _dependents(
    cells: dict[CellRef, Cell],
    tables: Mapping[str, Table] | None = None,
) -> dict[CellRef, set[CellRef]]:
    """Build the reverse dependency map for a set of cells.

    Maps each referenced cell to the cells that read it, so a change can be
    followed forward to everything it reaches. A formula that reads its own
    cell -- ``D12`` inside ``SUM(D2:D13)``, for instance -- is not a dependency
    of itself and is left out.
    """
    out: dict[CellRef, set[CellRef]] = {}
    index = CellIndex(cells)
    for cell in cells.values():
        if cell.kind is not CellKind.FORMULA or not cell.refs:
            continue
        for reference in cell.refs:
            resolved = _resolve(reference, tables)
            if resolved is None:
                continue
            for target in index.covered(resolved):
                if target != cell.ref:
                    out.setdefault(target, set()).add(cell.ref)
    return out


def _resolve(reference: Reference, tables: Mapping[str, Table] | None) -> Reference | None:
    """Turn a structured reference into a grid range, if the tables are known."""
    if isinstance(reference, TableRef):
        if tables is None:
            return None
        return reference.resolve(tables)
    return reference


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
