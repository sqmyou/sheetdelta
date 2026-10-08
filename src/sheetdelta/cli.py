"""The ``sheetdelta`` command line.

Two subcommands: ``diff`` compares two workbooks, ``audit`` inspects one. The
exit code is the point of the tool in CI, so ``--fail-on`` is explicit rather
than implied.
"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

from . import __version__
from .audit import audit_workbook
from .differ import diff_workbooks
from .errors import SheetDeltaError
from .reader import read_workbook
from .report import (
    exit_code,
    render_audit_github,
    render_audit_json,
    render_audit_text,
    render_github,
    render_json,
    render_summary,
    render_text,
)

FAIL_ON = ("never", "any", "breaking")
DIFF_FORMATS = ("text", "json", "summary", "github")
AUDIT_FORMATS = ("text", "json", "github")


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
    output = diff.add_mutually_exclusive_group()
    output.add_argument("--json", action="store_true", help="alias for --format=json")
    output.add_argument("--summary", action="store_true", help="alias for --format=summary")
    output.add_argument(
        "--format",
        choices=DIFF_FORMATS,
        help="output style (default: text); 'github' emits Actions annotations",
    )
    diff.add_argument(
        "--fail-on",
        choices=FAIL_ON,
        default="breaking",
        help="when to exit non-zero (default: breaking)",
    )

    audit = sub.add_parser("audit", help="inspect one workbook for broken references")
    audit.add_argument("workbook", metavar="FILE.xlsx", help="the workbook to inspect")
    audit.add_argument("--json", action="store_true", help="alias for --format=json")
    audit.add_argument(
        "--format",
        choices=AUDIT_FORMATS,
        help="output style (default: text); 'github' emits Actions annotations",
    )
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
    fmt = args.format or ("json" if args.json else "summary" if args.summary else "text")
    if fmt == "json":
        print(render_json(result))
    elif fmt == "summary":
        print(render_summary(result))
    elif fmt == "github":
        output = render_github(result)
        if output:
            print(output)
    else:
        print(render_text(result))
    return exit_code(result, args.fail_on)


def _cmd_audit(args: argparse.Namespace) -> int:
    result = audit_workbook(read_workbook(args.workbook))
    fmt = args.format or ("json" if args.json else "text")
    if fmt == "json":
        print(render_audit_json(result))
    elif fmt == "github":
        output = render_audit_github(result)
        if output:
            print(output)
    else:
        print(render_audit_text(result))
    # An audit has no "breaking" tier: any issue it finds is a real defect, so
    # "any" and "breaking" both fail, and only "never" always passes. Leaving
    # "breaking" to fall through to 0 would let a CI job pass on an issue.
    if args.fail_on == "never":
        return 0
    return 1 if result.issue_count else 0

