"""Render a diff as text for a human, or JSON for a machine.

The text report is the product. Anyone can print two lists of changed cells;
the reason to use this instead of reading the files yourself is the line that
says what a change reaches downstream.
"""

from __future__ import annotations

import json
from typing import Any

from .audit import AuditResult
from .batch import DirectoryDiff, WorkbookEntry
from .differ import CellChange, DiffResult, Move, Severity, SheetChange, Shift
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
    for shift in sheet.shifts:
        lines.append(f"  {_MARK[Severity.INFO]} {shift.label}")
    for move in sheet.moves:
        lines.append(f"  {_MARK[Severity.INFO]} {move.label}")
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
        for shift in sheet.shifts:
            bits.append(shift.label)
        for move in sheet.moves:
            bits.append(move.label)
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


def _md(text: str) -> str:
    """Escape a value for a Markdown table cell."""
    return text.replace("|", "\\|").replace("\n", " ")


def render_markdown(result: DiffResult) -> str:
    """Render the diff as Markdown, for a pull-request comment or job summary.

    The shape is the same as the text report -- a table of sheets, then the
    cell detail -- but it is meant to be pasted somewhere that renders Markdown,
    so the summary is a bolded line and each change is a list item.
    """
    lines = ["# Workbook diff", "", f"`{result.old_path}` -> `{result.new_path}`", ""]

    if not result.sheet_changes:
        lines.append("No changes.")
        return "\n".join(lines)

    lines.append("| Sheet | Change | Cells | Breaking |")
    lines.append("| --- | --- | --- | --- |")
    for sheet in result.sheet_changes:
        name = _md(sheet.name)
        if sheet.kind == "renamed":
            name = f"{_md(sheet.old_name or '')} -> {name}"
        lines.append(
            f"| {name} | {sheet.kind} | {len(sheet.cell_changes)} | {sheet.breaking_count} |"
        )

    for sheet in result.sheet_changes:
        if sheet.kind == "added":
            lines.extend(["", f"## Sheet '{_md(sheet.name)}' added"])
            continue
        if sheet.kind == "removed":
            lines.extend(["", f"## Sheet '{_md(sheet.name)}' removed"])
            continue
        if sheet.kind == "renamed":
            header = f"## Sheet '{_md(sheet.old_name or '')}' renamed to '{_md(sheet.name)}'"
        else:
            header = f"## Sheet '{_md(sheet.name)}'"
        lines.extend(["", header, ""])
        for shift in sheet.shifts:
            lines.append(f"- {shift.label}")
        for move in sheet.moves:
            lines.append(f"- {move.label}")
        for change in sheet.cell_changes:
            lines.append(f"- **{change.severity.value}** `{change.ref.a1}` {change.detail}")
            if change.old is not None and change.new is not None and change.old != change.new:
                lines.append(f"  - `{_md(change.old)}` -> `{_md(change.new)}`")
            elif change.old is not None:
                lines.append(f"  - was `{_md(change.old)}`")
            elif change.new is not None:
                lines.append(f"  - now `{_md(change.new)}`")
            if change.affected:
                readers = ", ".join(f"`{ref.a1}`" for ref in change.affected[:6])
                if len(change.affected) > 6:
                    readers += f", and {len(change.affected) - 6} more"
                lines.append(f"  - affects {len(change.affected)} cell(s): {readers}")

    lines.extend(["", f"**{_summary(result)}**"])
    return "\n".join(lines)


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
    return {
        "kind": sheet.kind,
        "name": sheet.name,
        "old_name": sheet.old_name,
        "shifts": [_shift_dict(shift) for shift in sheet.shifts],
        "moves": [_move_dict(move) for move in sheet.moves],
        "changes": [_change_dict(change) for change in sheet.cell_changes],
    }


def _move_dict(move: Move) -> dict[str, Any]:
    return {
        "axis": move.axis,
        "old": move.old,
        "new": move.new,
        "label": move.label,
    }


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


def render_directory_text(result: DirectoryDiff) -> str:
    """Render a directory diff as a report meant to be read."""
    lines = [f"{result.old_dir}  ->  {result.new_dir}"]
    if not result.has_changes:
        lines.append("")
        lines.append("No changes.")
        return "\n".join(lines)

    lines.append("")
    for entry in result.entries:
        if entry.status == "unchanged":
            continue
        lines.append(f"  {entry.name}  {_directory_detail(entry)}")

    lines.append("")
    lines.append(_directory_summary(result))
    return "\n".join(lines)


def _directory_detail(entry: WorkbookEntry) -> str:
    if entry.status == "added":
        return "added"
    if entry.status == "removed":
        return "removed"
    bits = [f"{entry.cell_change_count} change{'s' if entry.cell_change_count != 1 else ''}"]
    if entry.breaking_count:
        bits.append(f"{entry.breaking_count} breaking")
    return f"changed ({', '.join(bits)})"


def _directory_summary(result: DirectoryDiff) -> str:
    changed = sum(1 for e in result.entries if e.status == "changed")
    added = sum(1 for e in result.entries if e.status == "added")
    removed = sum(1 for e in result.entries if e.status == "removed")
    bits = []
    if changed:
        bits.append(f"{changed} changed")
    if added:
        bits.append(f"{added} added")
    if removed:
        bits.append(f"{removed} removed")
    total = len(result.entries)
    summary = f"{', '.join(bits)} of {total} workbook(s)"
    breaking = len(result.breaking)
    if breaking:
        return f"{summary}; {breaking} workbook(s) with breaking changes."
    return f"{summary}."


def render_directory_summary(result: DirectoryDiff) -> str:
    """One line per changed workbook, for a CI log."""
    if not result.has_changes:
        return "No changes."
    lines = [f"{result.old_dir}  ->  {result.new_dir}"]
    for entry in result.entries:
        if entry.status != "unchanged":
            lines.append(f"  {entry.name}: {_directory_detail(entry)}")
    lines.append(_directory_summary(result))
    return "\n".join(lines)


def render_directory_markdown(result: DirectoryDiff) -> str:
    """Render a directory diff as a Markdown table, for a PR comment."""
    lines = ["# Workbook diff", "", f"`{_md(result.old_dir)}` -> `{_md(result.new_dir)}`", ""]
    if not result.has_changes:
        lines.append("No changes.")
        return "\n".join(lines)

    lines.append("| Workbook | Change | Cells | Breaking |")
    lines.append("| --- | --- | --- | --- |")
    for entry in result.entries:
        if entry.status == "unchanged":
            continue
        lines.append(
            f"| {_md(entry.name)} | {entry.status} | "
            f"{entry.cell_change_count} | {entry.breaking_count} |"
        )
    lines.append("")
    lines.append(_directory_summary(result))
    return "\n".join(lines)


def to_directory_dict(result: DirectoryDiff) -> dict[str, Any]:
    return {
        "old": result.old_dir,
        "new": result.new_dir,
        "has_changes": result.has_changes,
        "summary": {
            "changed": sum(1 for e in result.entries if e.status == "changed"),
            "added": sum(1 for e in result.entries if e.status == "added"),
            "removed": sum(1 for e in result.entries if e.status == "removed"),
            "breaking": len(result.breaking),
        },
        "workbooks": [
            {
                "name": entry.name,
                "status": entry.status,
                "cell_changes": entry.cell_change_count,
                "breaking": entry.breaking_count,
                "diff": to_dict(entry.result) if entry.result is not None else None,
            }
            for entry in result.entries
            if entry.status != "unchanged"
        ],
    }


def render_directory_json(result: DirectoryDiff) -> str:
    """Render a directory diff as JSON, with the per-workbook diff nested."""
    return json.dumps(to_directory_dict(result), indent=2, sort_keys=False)


def render_directory_github(result: DirectoryDiff) -> str:
    """Emit one set of annotations per changed workbook.

    The per-workbook report already knows how to annotate its own sheets, so it
    is reused as-is; the paths it carries are the real file paths, which is what
    the annotations need. A run with nothing changed still emits a single notice
    so the step does not look skipped.
    """
    out: list[str] = []
    for entry in result.entries:
        if entry.status == "changed" and entry.result is not None:
            rendered = render_github(entry.result)
            if rendered:
                out.append(rendered)
        elif entry.status == "added":
            out.append(_annotation("notice", f"{entry.name}: added", "workbook added", entry.name))
        elif entry.status == "removed":
            out.append(
                _annotation("notice", f"{entry.name}: removed", "workbook removed", entry.name)
            )
    if not out:
        return _annotation("notice", "sheetdelta", "no changes")
    return "\n".join(out)


def directory_exit_code(result: DirectoryDiff, fail_on: str) -> int:
    """The directory-mode twin of :func:`exit_code`."""
    if fail_on == "never":
        return 0
    if fail_on == "any":
        return 1 if result.has_changes else 0
    return 1 if result.breaking else 0


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

    if result.volatile:
        lines.append("")
        lines.append(f"{len(result.volatile)} volatile reference(s):")
        for volatile in result.volatile:
            where = f"{volatile.kind}({volatile.literal})" if volatile.literal else volatile.kind
            lines.append(f"  ?  {volatile.ref.a1:<6} {where}")
            lines.append(f"        ={volatile.formula}")
        lines.append("")
        lines.append(
            "These formulas compute their target at runtime, so the cells they "
            "read cannot be known without evaluating them. The dependency graph "
            "treats them as reaching the whole sheet; a broken reference or cycle "
            "hidden behind one will not be reported."
        )

    return "\n".join(lines)


_SEVERITY_LEVEL = {
    Severity.BREAKING: "error",
    Severity.WARNING: "warning",
    Severity.INFO: "notice",
}


def _escape_github(message: str) -> str:
    """Escape the characters GitHub treats as command separators in a message."""
    return (
        message.replace("%", "%25")
        .replace("\r", "%0D")
        .replace("\n", "%0A")
        .replace(":", "%3A")
        .replace(",", "%2C")
    )


def _annotation(level: str, title: str, message: str, file: str | None = None) -> str:
    parts = [f"::{level}"]
    attrs: list[str] = []
    if file:
        # A workbook is not a tracked text file, but the annotation still needs
        # a file to attach to; the workbook path is the closest true answer.
        attrs.append(f"file={_escape_github(file)}")
    attrs.append(f"title={_escape_github(title)}")
    parts.append(" " + ",".join(attrs))
    parts.append(f"::{_escape_github(message)}")
    return "".join(parts)


def render_github(result: DiffResult) -> str:
    """Render a diff as GitHub Actions annotations.

    Each change becomes one ``::error``, ``::warning`` or ``::notice`` line so
    it shows against the pull request's diff. The severity carries over: a
    formula that corrupts a downstream total is an error, a cosmetic edit is a
    notice. Everything is on one line, as the workflow command format requires.
    """
    lines: list[str] = []
    for sheet in result.sheet_changes:
        if sheet.kind != "changed":
            lines.append(
                _annotation(
                    "notice",
                    f"{sheet.name}: sheet {sheet.kind}",
                    _sheet_note(sheet),
                    result.new_path,
                )
            )
        for shift in sheet.shifts:
            lines.append(
                _annotation("warning", f"{sheet.name}: {shift.label}", shift.label, result.new_path)
            )
        for move in sheet.moves:
            lines.append(
                _annotation("notice", f"{sheet.name}: {move.label}", move.label, result.new_path)
            )
        for change in sheet.cell_changes:
            level = _SEVERITY_LEVEL[change.severity]
            message = _change_message(sheet.name, change)
            title = f"{sheet.name}!{change.ref.a1}"
            lines.append(_annotation(level, title, message, result.new_path))
    if not lines:
        # A silent log looks like the step never ran; say so explicitly.
        lines.append(_annotation("notice", "sheetdelta", "No changes", result.new_path))
    return "\n".join(lines)


def _sheet_note(sheet: SheetChange) -> str:
    if sheet.kind == "renamed":
        return f"Sheet '{sheet.old_name}' renamed to '{sheet.name}'"
    return f"Sheet '{sheet.name}' {sheet.kind}"


def _change_message(sheet_name: str, change: CellChange) -> str:
    message = f"{sheet_name}!{change.ref.a1}: {change.kind.value}"
    if change.old is not None or change.new is not None:
        message += f" ({change.old or ''} -> {change.new or ''})"
    if change.detail:
        message += f" -- {change.detail}"
    if change.affected:
        readers = ", ".join(ref.a1 for ref in change.affected)
        message += f" [affects {readers}]"
    return message


def render_audit_github(result: AuditResult) -> str:
    """Render an audit as GitHub Actions annotations."""
    lines: list[str] = []
    file = result.path
    for name in result.incomplete:
        lines.append(
            _annotation("warning", f"{name}: unreadable", f"Sheet '{name}' could not be read", file)
        )
    for broken in result.broken:
        lines.append(
            _annotation(
                "error",
                f"{broken.ref.a1}: broken reference",
                f"{broken.ref.a1}: {broken.target} -- {broken.reason} (={broken.formula})",
                file,
            )
        )
    for cycle in result.cycles:
        lines.append(
            _annotation(
                "error",
                "circular reference",
                f"Circular reference: {cycle.display}",
                file,
            )
        )
    for volatile in result.volatile:
        where = f"{volatile.kind}({volatile.literal})" if volatile.literal else volatile.kind
        lines.append(
            _annotation(
                "warning",
                f"{volatile.ref.a1}: volatile reference",
                f"{volatile.ref.a1}: {where} -- target computed at runtime "
                f"(={volatile.formula})",
                file,
            )
        )
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
            "volatile": [
                {
                    "cell": str(v.ref),
                    "kind": v.kind,
                    "literal": v.literal,
                    "formula": f"={v.formula}",
                }
                for v in result.volatile
            ],
        },
        indent=2,
    )
