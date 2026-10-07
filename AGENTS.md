# AGENTS.md

Repository notes for anyone (human or agent) working on sheetdelta.

## What this is

A headless `.xlsx` differ. It compares two workbooks cell by cell, including
the formulas and the values Excel cached, then follows each change through the
workbook's dependency graph to report what it reaches downstream. Read-only,
standard library only, no Excel.

## Commands

```console
python -m pip install -e ".[dev]"   # install with dev tools
python -m pytest                     # tests
python -m ruff check src tests       # lint
python -m mypy                       # types (strict on src)
python -m build                      # wheel + sdist
```

Run all three checks before pushing; CI runs them on 3.10 through 3.13.

## Layout

```
src/sheetdelta/
  model.py       dataclasses: CellRef, RangeRef, Cell, Sheet, Workbook
  errors.py      the exception hierarchy, all under SheetDeltaError
  references.py  pulls cell/range references out of a formula
  reader.py      parses .xlsx (zipfile + xml.etree)
  differ.py      compares two workbooks, builds the dependency graph
  batch.py       compares two directory trees, one workbook pair at a time
  audit.py       checks one workbook for broken and circular references
  report.py      text and JSON rendering, exit codes
  cli.py         argparse entry point
```

## Things worth knowing before you change something

- **The reader has no dependencies on purpose.** It is `zipfile` and
  `xml.etree`. Do not add openpyxl as a runtime dependency; it is a dev-only
  tool used to write reference workbooks for `tests/test_interop.py`.
- **Formulas are never evaluated.** Only the value Excel cached is compared.
  There is no calculation engine and adding one would defeat the point.
- **Ranges are never expanded.** `SUM(A:A)` covers a million cells; the graph
  asks `RangeRef.contains` instead of materialising them.
- **The dependency graph spans the whole workbook**, not one sheet, so a change
  in one sheet is followed into the sheet that reads it. A formula that reads
  its own cell (a total inside its own range) is excluded from the graph.
- **Strings are not always in `<v>`.** openpyxl writes text as
  `<is><t>...</t></is>`. Both shapes are handled in `_raw_value`; the interop
  test is what caught this.
- **Shared formulas must be shifted.** Excel stores a repeated formula once and
  refers to it from other cells with relative offsets applied. See
  `shift_formula`.
- **A row insert is a rewrite, not an edit.** Excel has no "insert row" in the
  file format; it re-addresses every cell below. `differ._detect_shifts` walks
  both sheets' lines together, recording an insert or a removal whenever they
  fall out of step (`_trace_axis`), then re-checks the difference after each so
  a sheet with several inserts is handled. The claim is only kept when the
  blocks really line up once moved (`_shifts_line_up`), so two unrelated sheets
  are never aligned by force. Shifts are detected in the old sheet's
  coordinates and re-expressed in the new sheet's numbering by `_display_shifts`
  before they reach the report.
- **A rename need not be exact.** `differ._match_renamed_sheets` scores every
  removed/added pair with `_sheet_similarity` and pairs the best matches first,
  accepting anything at or above `_RENAME_THRESHOLD` (0.9). A sheet renamed in
  the same commit as a light edit is still recognised.
- **A current-row reference needs the formula's row.** `references.extract_
  references` takes a `row` and passes it down to `_scan`; `_parse_structured`
  reports whether `[@Column]` was used, and the row is pinned only when the
  table is on the formula's own sheet. `_validate` drops a reference whose row
  falls outside the table, so it is never pointed at the wrong table.
- **Structured references need the table definition.** `Sales[Amount]` only
  resolves if the table was read. `references.extract_references` takes a
  `tables` map (lowercased name to `Table`) and `TableRef.resolve` turns it into
  a range. A name that is not a known table is not treated as one, which keeps
  a bare `Foo[Bar]` from looking like a reference.
- **A volatile reference is recorded, not dropped.** `INDIRECT` and `OFFSET`
  compute their target at runtime, so `references._scan` cannot resolve them.
  It yields a `VolatileRef` (kind, scope, optional string literal) instead of
  nothing, and `model.CellIndex.covered` treats it as reaching its whole scope
  (sheet by default, workbook under `--volatile-scope workbook`). Without this
  a change behind such a formula looked like it affected nothing.
  `audit._volatile_refs` reports it separately; it does **not** make the
  workbook unsound, so `issue_count` and the exit code are unaffected.
- **A string literal hides a function name.** `references._scan` runs the
  volatile-finder over the raw formula because the literal argument is the
  useful part, but a call inside a literal (`"see INDIRECT(a1)"`) must not be
  one. The stripped text has literals blanked to spaces, so a differing char at
  the match offset means the name is inside a literal and the match is skipped.
  `_first_argument` only returns a literal that is the whole first argument
  (`"A1"` yes, `"A"&B1` no), so a computed target is not shown as a literal one.
- **A reorder is a move only when it is a clean permutation.**
  `differ._detect_moves` compares the multisets of `_line_signatures`; equal
  means the lines were only reordered, unequal means the shift detector is left
  to explain it. This matters because an insert looks exactly like a move for
  every line below it -- requiring the permutation is what stops one insert
  being reported as many moves. Moves are detected **before** shifts, and a
  signature that occurs more than once on either side is skipped as ambiguous.
- **Both axes are checked, rows first.** `_detect_moves` tries the row axis then
  the column axis and returns the first that is a clean permutation. A row
  reorder almost never also looks like a column reorder, but a symmetric sheet
  (a transposed grid, a diagonal) can satisfy both; preferring rows matches the
  reading order the rest of the diff uses and keeps the report stable. `Move`
  carries `axis`; `RowMove` is an alias for it, kept because it was exported.
- **A move is re-keyed, not just reported.** `_apply_moves` rewrites the old
  cells onto their new row and column numbers before the cell comparison, so the
  moved cells do not also appear as changes. `SheetChange.moves` carries the
  labels; every renderer (text, summary, markdown, github) and `to_dict` include
  them.
- **Directory mode pairs by relative path, not file name.** `batch.py` walks
  both trees keyed on the POSIX path relative to each root, so same-named files
  in different folders stay distinct and a moved file reads as remove-plus-add.
  A path that is not a directory raises `DirectoryError` rather than the
  reader's `WorkbookReadError`, because nothing was opened.
- **The GitHub annotation format has its own escaping.** `report._escape_github`
  percent-encodes `%`, CR, LF, `:` and `,`; a workflow command is one line, so
  any newline in a message would break it. `--format github` reuses the same
  severity mapping as the text renderer (`_SEVERITY_LEVEL`) but always prints a
  notice when a diff is empty, so a step never looks like it was skipped.
- **`__all__` is not checked by any tool.** A name listed there but never
  imported is a runtime `AttributeError` only. `tests/test_exports.py` walks the
  list; keep it passing.

## Releasing

Bump `version` in `pyproject.toml` and `__version__` in `src/sheetdelta/__init__.py`,
add a CHANGELOG entry, then `gh release create vX.Y.Z --notes-file ...`. The
Release workflow builds and publishes to PyPI on `release: published`. PyPI's
JSON endpoint can lag the simple index by a minute; confirm with
`pip install --target` rather than the JSON.

Tag the commit with `git tag -a` and push it so the release points at the right
tree; a GitHub release can exist without a local tag and then drift.

Commits must be authored and committed as `sqmyou <sirsamyoudev@gmail.com>`
with no `openhands` co-author trailer.

## Test fixtures

`tests/xlsx_fixtures.py` writes `.xlsx` files with the standard library, so the
suite runs with no dependencies. `tests/test_interop.py` checks the same reader
against workbooks written by openpyxl, and skips if openpyxl is absent.
