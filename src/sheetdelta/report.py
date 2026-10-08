"""Render a diff as text for a human, or JSON for a machine.

The text report is the product. Anyone can print two lists of changed cells;
the reason to use this instead of reading the files yourself is the line that
says what a change reaches downstream.
"""

from __future__ import annotations

import json
from typing import Any

from .audit import AuditResult
from .differ import CellChange, DiffResult, Severity, SheetChange, Shift
from .model import CellRef

_MARK = {
    Severity.BREAKING: "!!",
    Severity.WARNING: "! ",
    Severity.INFO: "  ",
}


def render_text(result: DiffResult) -> str:
    """Render the diff as a report meant to be read."""
    lines: list[str] = []
    lines.append(f"{result.old_path}  ->  {result.new_path}")

    if not result.sheet_changes:
        lines.append("")
        lines.append("No changes.")
        return "\n".join(lines)

    for sheet in result.sheet_changes:
        lines.append("")
        lines.extend(_render_sheet(sheet))

    lines.append("")
    lines.append(_summary(result))
    return "\n".join(lines)


def _render_sheet(sheet: SheetChange) -> list[str]:
    if sheet.kind == "added":
        return [f"Sheet '{sheet.name}'  added"]
    if sheet.kind == "removed":
        return [f"Sheet '{sheet.name}'  removed"]
    if sheet.kind == "renamed":
        header = f"Sheet '{sheet.old_name}'  renamed -> '{sheet.name}'"
    else:
        header = f"Sheet '{sheet.name}'"

    count = len(sheet.cell_changes)
    breaking = sheet.breaking_count
    if count:
        detail = f"{count} change{'s' if count != 1 else ''}"
        if breaking:
            detail += f", {breaking} breaking"
        header += f"  ({detail})"

    lines = [header]
    if sheet.shift is not None:
        lines.append(f"  {_MARK[Severity.INFO]} {sheet.shift.label}")
    for change in sheet.cell_changes:
        lines.extend(_render_change(change))
    return lines


def _render_change(change: CellChange) -> list[str]:
    mark = _MARK[change.severity]
    head = f"  {mark} {change.ref.a1:<6} {change.detail}"
    lines = [head]

    if change.old is not None and change.new is not None and change.old != change.new:
        lines.append(f"          -  {change.old}")
        lines.append(f"          +  {change.new}")
    elif change.old is not None:
        lines.append(f"          -  {change.old}")
    elif change.new is not None:
        lines.append(f"          +  {change.new}")

    if change.affected:
        lines.append(f"        affects {_describe_affected(change.affected)}")
    return lines


def _describe_affected(affected: list[CellRef]) -> str:
    """Name the downstream cells, shortening a long list."""
    shown = ", ".join(str(ref) for ref in affected[:6])
    if len(affected) > 6:
        shown += f", and {len(affected) - 6} more"
    noun = "cell" if len(affected) == 1 else "cells"
    return f"{len(affected)} {noun} downstream: {shown}"


def render_summary(result: DiffResult) -> str:
    """Render a one-line-per-sheet count, for a CI log.

    The full report is the product, but a CI job often only wants the shape of
    the change: how many sheets moved and how many breaks there are. This is
    that, with no cell detail.
    """
    if not result.sheet_changes:
        return "No changes."

    lines = [f"{result.old_path}  ->  {result.new_path}"]
    for sheet in result.sheet_changes:
        if sheet.kind == "added":
            lines.append(f"  {sheet.name}: sheet added")
            continue
        if sheet.kind == "removed":
            lines.append(f"  {sheet.name}: sheet removed")
            continue
        if sheet.kind == "renamed":
            head = f"  {sheet.old_name} -> {sheet.name}: renamed"
        else:
            head = f"  {sheet.name}"

        bits: list[str] = []
        if sheet.shift is not None:
            bits.append(sheet.shift.label)
        count = len(sheet.cell_changes)
        if count:
            bits.append(f"{count} change{'s' if count != 1 else ''}")
        breaking = sheet.breaking_count
        if breaking:
            bits.append(f"{breaking} breaking")
        if bits:
            lines.append(f"{head}: {', '.join(bits)}")
        else:
            lines.append(head)

    lines.append(_summary(result))
    return "\n".join(lines)


def _summary(result: DiffResult) -> str:
    counts = result.counts()
    parts = [f"{count} {kind}" for kind, count in sorted(counts.items())]
    changed_sheets = sum(1 for s in result.sheet_changes if s.kind == "changed")
    renamed = sum(1 for s in result.sheet_changes if s.kind == "renamed")
    added = sum(1 for s in result.sheet_changes if s.kind == "added")
    removed = sum(1 for s in result.sheet_changes if s.kind == "removed")

    sheet_bits = []
    if changed_sheets:
        sheet_bits.append(f"{changed_sheets} edited")
    if added:
        sheet_bits.append(f"{added} added")
    if removed:
        sheet_bits.append(f"{removed} removed")
    if renamed:
        sheet_bits.append(f"{renamed} renamed")

    summary = "; ".join(parts) if parts else "no cell changes"
    if sheet_bits:
        summary = f"{summary} across {', '.join(sheet_bits)} sheet(s)"

    breaking = len(result.breaking)
    affected = {str(ref) for change in result.breaking for ref in change.affected}
    if breaking:
        return (
            f"{summary}. {breaking} breaking change"
            f"{'s' if breaking != 1 else ''} reaching {len(affected)} downstream cell"
            f"{'s' if len(affected) != 1 else ''}."
        )
    return f"{summary}."


def render_json(result: DiffResult) -> str:
    """Render the diff as JSON, with stable keys and sorted lists."""
    return json.dumps(to_dict(result), indent=2, sort_keys=False)


def to_dict(result: DiffResult) -> dict[str, Any]:
    return {
        "old": result.old_path,
        "new": result.new_path,
        "has_changes": result.has_changes,
        "summary": {
            "cell_changes": result.counts(),
            "breaking": len(result.breaking),
        },
        "sheets": [_sheet_dict(sheet) for sheet in result.sheet_changes],
    }


def _sheet_dict(sheet: SheetChange) -> dict[str, Any]:
    out: dict[str, Any] = {
        "kind": sheet.kind,
        "name": sheet.name,
        "old_name": sheet.old_name,
        "changes": [_change_dict(change) for change in sheet.cell_changes],
    }
    if sheet.shift is not None:
        out["shift"] = _shift_dict(sheet.shift)
    return out


def _shift_dict(shift: Shift) -> dict[str, Any]:
    return {
        "axis": shift.axis,
        "inserted": shift.inserted,
        "count": shift.count,
        "at": shift.at,
        "label": shift.label,
    }


def _change_dict(change: CellChange) -> dict[str, Any]:
    out: dict[str, Any] = {
        "cell": str(change.ref),
        "a1": change.ref.a1,
        "kind": change.kind.value,
        "severity": change.severity.value,
        "detail": change.detail,
    }
    if change.old is not None:
        out["old"] = change.old
    if change.new is not None:
        out["new"] = change.new
    if change.affected:
        out["affected"] = [str(ref) for ref in change.affected]
    return out


def exit_code(result: DiffResult, fail_on: str) -> int:
    """Decide the process exit code from the diff and the --fail-on level.

    ``fail_on`` is one of ``never``, ``any`` or ``breaking``. A CI job uses
    ``breaking`` so a cosmetic edit passes and a formula that corrupts a total
    fails the build.
    """
    if fail_on == "never":
        return 0
    if fail_on == "any":
        return 1 if result.has_changes else 0
    return 1 if result.breaking else 0


def render_audit_text(result: AuditResult) -> str:
    """Render an audit for a human."""
    lines = [f"{result.path}"]
    lines.append(
        f"  {len(result.sheets)} sheet(s), {result.cell_count} cell(s), "
        f"{result.formula_count} formula(s)"
    )
    if result.defined_names:
        lines.append(f"  {len(result.defined_names)} defined name(s)")

    if result.is_sound:
        lines.append("")
        lines.append("No broken references, no circular references.")
        return "\n".join(lines)

    if result.incomplete:
        lines.append("")
        lines.append(f"{len(result.incomplete)} worksheet(s) could not be read:")
        for name in result.incomplete:
            lines.append(f"  !! {name}: part missing or malformed (result is incomplete)")

    if result.broken:
        lines.append("")
        lines.append(f"{len(result.broken)} broken reference(s):")
        for broken in result.broken:
            lines.append(f"  !! {broken.ref.a1:<6} {broken.target}")
            lines.append(f"        {broken.reason}")
            lines.append(f"        ={broken.formula}")

    if result.cycles:
        lines.append("")
        lines.append(f"{len(result.cycles)} circular reference(s):")
        for cycle in result.cycles:
            lines.append(f"  !! {cycle.display}")

    return "\n".join(lines)


def render_audit_json(result: AuditResult) -> str:
    """Render an audit as JSON. The ``incomplete`` list names unread sheets."""
    return json.dumps(
        {
            "path": result.path,
            "sheets": result.sheets,
            "cells": result.cell_count,
            "formulas": result.formula_count,
            "sound": result.is_sound,
            "incomplete": result.incomplete,
            "broken": [
                {
                    "cell": str(b.ref),
                    "target": b.target,
                    "reason": b.reason,
                    "formula": f"={b.formula}",
                }
                for b in result.broken
            ],
            "cycles": [[str(ref) for ref in cycle.cells] for cycle in result.cycles],
        },
        indent=2,
    )
