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
    Sheet,
    Table,
    TableRef,
    VolatileRef,
    Workbook,
)

# How alike two sheets must be to be called a rename rather than a remove plus
# an add. Exact renames score 1.0; a rename committed together with a light
# edit scores just under it. The value is strict enough that two unrelated
# sheets, which share at most the odd coincidental cell, stay well below it.
_RENAME_THRESHOLD = 0.9


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

    A sheet can contain more than one insert, so the axis and the at/count pair
    describe one event and a sheet holds a list of them. Offsets are in the
    original (pre-shift) coordinates, which is the form the report shows and the
    form two shifts on one axis combine need to be expressed in.
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

    @property
    def delta(self) -> int:
        return self.count if self.inserted else -self.count


def shift_cell(ref: CellRef, shifts: list[Shift]) -> CellRef | None:
    """Apply every shift to a cell, or None if one of them deletes it.

    Shifts are stored in the original (pre-shift) coordinates, so every test is
    made against the cell's original address and the offsets are summed. Two
    shifts on one axis never overlap, which makes the sum well defined.
    """
    row_offset = col_offset = 0
    for shift in sorted(shifts, key=lambda s: s.at):
        coordinate = ref.row if shift.axis == "row" else ref.col
        if coordinate < shift.at:
            continue
        if not shift.inserted and coordinate < shift.at + shift.count:
            return None  # a removed line
        if shift.axis == "row":
            row_offset += shift.delta
        else:
            col_offset += shift.delta
    return CellRef(ref.sheet, ref.col + col_offset, ref.row + row_offset)


@dataclass
class RowMove:
    """A whole row that kept its contents but moved to a different row number.

    Excel has no "move row" either: dragging a row rewrites both the vacated and
    the occupied addresses. Comparing addresses directly reports every cell of
    every affected row as changed, which buries the fact that nothing in them
    actually changed -- only their position did.
    """

    old_row: int
    new_row: int

    @property
    def label(self) -> str:
        return f"row {self.old_row} moved to {self.new_row}"


@dataclass
class SheetChange:
    """A sheet added, removed, renamed or edited between the two workbooks."""

    kind: str  # "added" | "removed" | "renamed" | "changed"
    name: str
    old_name: str | None = None
    cell_changes: list[CellChange] = field(default_factory=list)
    shifts: list[Shift] = field(default_factory=list)
    moves: list[RowMove] = field(default_factory=list)

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
        cell_changes, shifts, moves = _diff_cells(
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
                shifts=shifts,
                moves=moves,
            )
        )

    for name in old.sheet_names:
        if name in renamed_from or new.sheet_by_name(name) is not None:
            continue
        result.sheet_changes.append(
            SheetChange(
                kind="removed",
                name=name,
                old_name=name,
                cell_changes=_removal_changes(old.sheet_by_name(name), dependents),
            )
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
        changes, shifts, moves = _diff_cells(old_sheet.cells, new_sheet.cells, dependents)
        if changes or shifts or moves:
            result.sheet_changes.append(
                SheetChange(
                    kind="changed",
                    name=name,
                    old_name=name,
                    cell_changes=changes,
                    shifts=shifts,
                    moves=moves,
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
                    raw_value=cell.raw_value,
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
    sheets with matching contents are reported as renamed instead. The contents
    do not have to match exactly: a sheet renamed in the same commit as a small
    edit is still the same sheet, so a best score above ``_RENAME_THRESHOLD`` is
    accepted. A sheet whose contents are genuinely different stays below it and
    is left as a remove plus an add.
    """
    removed = [n for n in old.sheet_names if new.sheet_by_name(n) is None]
    added = [n for n in new.sheet_names if old.sheet_by_name(n) is None]

    candidates: list[tuple[float, str, str]] = []
    for old_name in removed:
        old_sheet = old.sheet_by_name(old_name)
        if old_sheet is None:
            continue
        for new_name in added:
            new_sheet = new.sheet_by_name(new_name)
            if new_sheet is None:
                continue
            score = _sheet_similarity(old_sheet.cells, new_sheet.cells, new_name)
            if score >= _RENAME_THRESHOLD:
                candidates.append((score, old_name, new_name))

    # Pair the closest matches first so each sheet is claimed only once.
    pairs: list[tuple[str, str]] = []
    taken_old: set[str] = set()
    taken_new: set[str] = set()
    for _, old_name, new_name in sorted(candidates, key=lambda c: (-c[0], c[1], c[2])):
        if old_name in taken_old or new_name in taken_new:
            continue
        pairs.append((old_name, new_name))
        taken_old.add(old_name)
        taken_new.add(new_name)
    return sorted(pairs)


def _sheet_similarity(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell], new_name: str
) -> float:
    """The fraction of the two sheets' cells that sit at the same place and match.

    An exact rename has a similarity of 1. A sheet that was renamed and then
    lightly edited scores just under it, which is what lets the same sheet be
    recognised instead of being reported as a remove plus an add -- while a
    sheet with genuinely different contents stays below the threshold and is
    left alone.
    """
    if not old_cells and not new_cells:
        return 1.0
    total = len(old_cells) + len(new_cells)
    matched = 0
    for ref, cell in old_cells.items():
        other = new_cells.get(CellRef(new_name, ref.col, ref.row))
        if other is not None and cell.signature == other.signature:
            matched += 1
    return 2 * matched / total


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
    if isinstance(reference, VolatileRef):
        return VolatileRef(new_name, reference.kind, reference.scope, reference.literal)
    return RangeRef(
        new_name, reference.min_col, reference.max_col, reference.min_row, reference.max_row
    )


def _diff_cells(
    old_cells: dict[CellRef, Cell],
    new_cells: dict[CellRef, Cell],
    dependents: dict[CellRef, set[CellRef]],
) -> tuple[list[CellChange], list[Shift], list[RowMove]]:
    """Compare the cells of one sheet, in reading order.

    Returns the changes plus every shift and row move found. The old cells are
    re-keyed onto their new addresses first, so a cell that only moved is not
    reported as changed. A sheet can hold more than one insert, so shifts are
    collected until the remaining difference no longer lines up.
    """
    # A reorder is checked before the shifts, because the shift detector also
    # explains a moved row as an insert plus a remove. Only an exact permutation
    # of the rows is claimed as a move, so an insert is never mistaken for one.
    moves = _detect_moves(old_cells, new_cells)
    if moves:
        old_cells = _apply_moves(old_cells, moves)

    raw = _detect_shifts(old_cells, new_cells)
    if raw:
        old_cells = _apply_shifts(old_cells, raw)
    shifts = _display_shifts(raw)

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

    return changes, shifts, moves


def _detect_moves(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell]
) -> list[RowMove]:
    """Find rows that only changed position, and can be re-keyed to line up.

    A reorder is written the way Excel writes one: every cell of the moved row
    is rewritten to its new address, and the rows in between are untouched --
    which is also exactly what one insert plus one remove below looks like. The
    two are told apart by requiring the multiset of row contents to be identical
    on both sides: a reorder permutes rows, it does not add, remove or edit one.
    When that holds, the move is the more faithful reading; when it does not,
    the insert/remove detector is left to explain the difference.

    A row whose contents appear more than once on either side is ambiguous --
    two identical rows cannot be told apart -- so it is skipped rather than
    guessed at, and rows that did not move are never paired.
    """
    old_lines = _line_signatures(old_cells, "row")
    new_lines = _line_signatures(new_cells, "row")
    if old_lines == new_lines:
        return []

    # Only a clean permutation counts. Any added, removed or edited row brings
    # the two multisets out of balance, and the change is left to the cell diff.
    if sorted(old_lines.values()) != sorted(new_lines.values()):
        return []

    # A signature that is not unique on a side is ambiguous: either source row
    # would fit, and choosing one would be a guess. Those rows are left as
    # changes rather than re-keyed.
    old_counts: dict[tuple[tuple[int, str | None], ...], int] = {}
    for sig in old_lines.values():
        old_counts[sig] = old_counts.get(sig, 0) + 1
    new_counts: dict[tuple[tuple[int, str | None], ...], int] = {}
    for sig in new_lines.values():
        new_counts[sig] = new_counts.get(sig, 0) + 1

    by_sig: dict[tuple[tuple[int, str | None], ...], int] = {}
    for row, sig in new_lines.items():
        if old_counts.get(sig, 0) == 1 and new_counts.get(sig, 0) == 1:
            by_sig[sig] = row

    moves: list[RowMove] = []
    for old_row, sig in old_lines.items():
        new_row = by_sig.get(sig)
        if new_row is None or new_row == old_row:
            continue
        moves.append(RowMove(old_row=old_row, new_row=new_row))
    return sorted(moves, key=lambda m: (m.new_row, m.old_row))


def _apply_moves(cells: dict[CellRef, Cell], moves: list[RowMove]) -> dict[CellRef, Cell]:
    """Re-key each moved row's cells to their new row number."""
    by_old = {move.old_row: move.new_row for move in moves}
    out: dict[CellRef, Cell] = {}
    for ref, cell in cells.items():
        new_row = by_old.get(ref.row, ref.row)
        moved = CellRef(ref.sheet, ref.col, new_row)
        out[moved] = cell
    return out


def _detect_shifts(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell]
) -> list[Shift]:
    """Find every row or column insert/remove that explains the difference.

    A sheet can hold more than one insert, so the lines are traced together and
    each time the two sides fall out of step an insert or a removal is recorded;
    when neither side can be resynchronised the line is an edit and both sides
    advance. The set is only kept if the blocks really do line up once moved,
    which is what separates an insert from an edit: two unrelated sheets must
    not be described as one inserting rows into the other.
    """
    for axis in ("row", "column"):
        shifts = _trace_axis(old_cells, new_cells, axis)
        if shifts and _shifts_line_up(old_cells, new_cells, shifts):
            return shifts
    return []


def _trace_axis(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell], axis: str
) -> list[Shift]:
    """Walk the lines of both sheets and record where they fall in and out of step."""
    old_sig = _line_signatures(old_cells, axis)
    new_sig = _line_signatures(new_cells, axis)
    if old_sig == new_sig:
        return []

    old_lines = sorted(old_sig)
    new_lines = sorted(new_sig)
    shifts: list[Shift] = []
    oi = ni = 0

    while oi < len(old_lines) and ni < len(new_lines):
        old_line, new_line = old_lines[oi], new_lines[ni]
        if old_sig[old_line] == new_sig[new_line]:
            oi += 1
            ni += 1
            continue

        # The line we are on has no counterpart here. Look for it ahead on the
        # other side: a jump forward in the new sheet is an insertion, a jump in
        # the old sheet a removal. Whichever resynchronises sooner wins.
        inserted_at = _find_line(new_lines, ni + 1, new_sig, old_sig[old_line])
        removed_at = _find_line(old_lines, oi + 1, old_sig, new_sig[new_line])
        if inserted_at is not None and (removed_at is None or inserted_at - ni <= removed_at - oi):
            shifts.append(Shift(axis=axis, inserted=True, count=inserted_at - ni, at=old_line))
            ni = inserted_at
        elif removed_at is not None:
            shifts.append(Shift(axis=axis, inserted=False, count=removed_at - oi, at=old_line))
            oi = removed_at
        else:
            oi += 1  # a genuine edit; move past it on both sides
            ni += 1

    return shifts


def _find_line(
    lines: list[int],
    start: int,
    signatures: dict[int, tuple[tuple[int, str | None], ...]],
    value: tuple[tuple[int, str | None], ...],
) -> int | None:
    """The first index at or after ``start`` whose line holds ``value``."""
    for index in range(start, len(lines)):
        if signatures[lines[index]] == value:
            return index
    return None


def _line_signatures(
    cells: dict[CellRef, Cell], axis: str
) -> dict[int, tuple[tuple[int, str | None], ...]]:
    """Describe each row (or column) by what it holds, ignoring its address.

    Two lines with the same signature hold the same contents, which is what
    makes them comparable across a shift.
    """
    gathered: dict[int, list[tuple[int, str | None]]] = {}
    for ref, cell in cells.items():
        line = ref.row if axis == "row" else ref.col
        other = ref.col if axis == "row" else ref.row
        gathered.setdefault(line, []).append((other, cell.signature))
    return {line: tuple(sorted(items)) for line, items in gathered.items()}


def _shifts_line_up(
    old_cells: dict[CellRef, Cell], new_cells: dict[CellRef, Cell], shifts: list[Shift]
) -> bool:
    """Check that the blocks line up once every shift is applied.

    A shift is a claim about the whole sheet, so most of the moved cells have to
    agree with their new addresses. Requiring every cell to match would reject a
    real insert that came with an edit; requiring only a majority keeps the
    claim honest while letting the edit through as a change in its own right.
    """
    moved = _apply_shifts(old_cells, shifts)
    matched = mismatched = 0
    for ref, old in moved.items():
        new = new_cells.get(ref)
        if new is not None and new.signature == old.signature:
            matched += 1
        else:
            mismatched += 1
    return matched > 0 and mismatched <= matched


def _display_shifts(shifts: list[Shift]) -> list[Shift]:
    """Re-express shifts in the new sheet's numbering.

    Detection works in the old sheet's coordinates, where applying the shifts in
    order is simple. A person reads the row number off the new sheet, so each
    later shift is moved forward by the inserts (or backward by the removals)
    before it.
    """
    out: list[Shift] = []
    offsets = {"row": 0, "column": 0}
    for shift in sorted(shifts, key=lambda s: (s.axis, s.at)):
        out.append(
            Shift(shift.axis, shift.inserted, shift.count, shift.at + offsets[shift.axis])
        )
        offsets[shift.axis] += shift.delta
    return out


def _apply_shifts(cells: dict[CellRef, Cell], shifts: list[Shift]) -> dict[CellRef, Cell]:
    """Re-key every cell under all the shifts at once."""
    out: dict[CellRef, Cell] = {}
    for ref, cell in cells.items():
        moved = shift_cell(ref, shifts)
        if moved is not None:
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
        if old.value_signature is not None and old.value_signature == new.value_signature:
            change.kind = ChangeKind.STALE
            change.detail = "formula changed, cached value did not move (not recalculated)"
        return [change]

    if old.signature != new.signature:
        affected = _reachable(ref, dependents)
        return [
            CellChange(
                ref=ref,
                kind=ChangeKind.VALUE,
                severity=Severity.BREAKING if affected else Severity.WARNING,
                old=old.display,
                new=new.display,
                detail="value changed",
                affected=affected,
            )
        ]

    return []


def _removal_changes(
    sheet: Sheet | None, dependents: dict[CellRef, set[CellRef]]
) -> list[CellChange]:
    """The cells of a removed sheet that other cells still read.

    A removed sheet is reported as a sheet-level removal with no cell detail,
    which hides the real risk. When another sheet consumed a cell on it, Excel
    rewrites that reader to ``#REF!``; a file written by some other tool keeps
    the stale reference instead, and the diff of the two files shows nothing
    wrong. Reporting the removed cell as breaking, with the readers it strands,
    is what makes ``--fail-on breaking`` catch the silent case.
    """
    if sheet is None:
        return []
    out: list[CellChange] = []
    for ref in sorted(sheet.cells, key=lambda r: (r.sheet, r.row, r.col)):
        affected = _reachable(ref, dependents)
        if not affected:
            continue
        cell = sheet.cells[ref]
        out.append(
            CellChange(
                ref=ref,
                kind=ChangeKind.REMOVED,
                severity=Severity.BREAKING,
                old=cell.display,
                new=None,
                detail=f"cell removed with sheet '{sheet.name}'",
                affected=affected,
            )
        )
    return out


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
