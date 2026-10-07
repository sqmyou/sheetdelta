"""The ``sheetdelta`` command line.

Two subcommands: ``diff`` compares two workbooks, ``audit`` inspects one. The
exit code is the point of the tool in CI, so ``--fail-on`` is explicit rather
than implied.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence

from . import __version__
from .audit import AuditResult, audit_workbook
from .differ import diff_workbooks
from .errors import SheetDeltaError
from .reader import read_workbook
from .report import exit_code, render_json, render_text

FAIL_ON = ("never", "any", "breaking")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sheetdelta",
        description="Diff Excel workbooks without Excel.",
        epilog="Exit status: 0 no changes worth failing on, 1 changes found, 2 error.",
    )
    parser.add_argument("--version", action="version", version=f"sheetdelta {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    diff = sub.add_parser("diff", help="compare two workbooks")
    diff.add_argument("old", metavar="OLD.xlsx", help="the earlier workbook")
    diff.add_argument("new", metavar="NEW.xlsx", help="the later workbook")
    diff.add_argument("--json", action="store_true", help="emit JSON instead of text")
    diff.add_argument(
        "--fail-on",
        choices=FAIL_ON,
        default="breaking",
        help="when to exit non-zero (default: breaking)",
    )

    audit = sub.add_parser("audit", help="inspect one workbook for broken references")
    audit.add_argument("workbook", metavar="FILE.xlsx", help="the workbook to inspect")
    audit.add_argument("--json", action="store_true", help="emit JSON")
    audit.add_argument(
        "--fail-on",
        choices=FAIL_ON,
        default="any",
        help="when to exit non-zero (default: any)",
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "diff":
            return _cmd_diff(args)
        if args.command == "audit":
            return _cmd_audit(args)
    except SheetDeltaError as exc:
        print(f"sheetdelta: {exc}", file=sys.stderr)
        return 2
    return 2


def _cmd_diff(args: argparse.Namespace) -> int:
    result = diff_workbooks(read_workbook(args.old), read_workbook(args.new))
    print(render_json(result) if args.json else render_text(result))
    return exit_code(result, args.fail_on)


def _cmd_audit(args: argparse.Namespace) -> int:
    result = audit_workbook(read_workbook(args.workbook))
    print(render_audit_json(result) if args.json else render_audit_text(result))
    if args.fail_on == "never":
        return 0
    if args.fail_on == "any":
        return 1 if result.issue_count else 0
    return 0


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
    return json.dumps(
        {
            "path": result.path,
            "sheets": result.sheets,
            "cells": result.cell_count,
            "formulas": result.formula_count,
            "sound": result.is_sound,
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

