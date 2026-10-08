# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.5.0] - 2026-10-08

### Added

- `INDIRECT` and `OFFSET` are now recognised. Their target is computed at
  runtime, so the dependency graph cannot follow it; before, the scanner found
  no reference at all and the formula looked like it depended on nothing. Each
  call site is recorded as a `VolatileRef` and reported by `audit` as a
  *volatile reference*, separately from broken and circular ones. The graph
  treats a volatile reference as reaching its whole sheet, so a change behind
  one is still reported as affecting the formula's cell. `--volatile-scope
  {sheet,workbook}` on `diff` and `audit` controls how wide that reach is
  (default `sheet`).
- `diff --format markdown` renders the report as a Markdown table plus a list,
  ready for a pull-request comment or a job summary. `render_markdown` is
  exported from the package root.
- A clean row reorder is now reported as a move, not as every cell of every
  affected row. When the rows on both sides are the same and unchanged, only in
  a different order, `sheetdelta` re-keys them and reports `row 4 moved to 2`.
  The JSON `moves` list and the text, summary, markdown and github formats all
  carry it. A reorder mixed with an add, remove or edit is not a permutation and
  is still reported as ordinary cell changes, so a move is never guessed.
- `RowMove` and `VolatileRef` are exported from the package root.

### Changed

- The JSON report now has a `moves` list on every sheet, empty when there was
  no reorder. `shifts` is joined by it, so a consumer reads both as lists.

## [0.4.0] - 2026-10-07

### Added

- More than one insert or remove in a single sheet is now reported. The engine
  finds shifts one at a time and re-checks the remaining difference after each,
  so a sheet with two separate blocks inserted reports both instead of only the
  first. Insert and remove can be mixed in one file. Shift rows are now given in
  the new sheet's numbering, so the row a reader sees on screen is the row the
  report names.
- A current-row structured reference, `[@Amount]`, now pins to a single cell --
  the one on the formula's own row -- instead of the whole table column. A
  formula reading its own row no longer looks like it depends on every row, so
  an edit elsewhere in the column is no longer reported as breaking. A reference
  whose row is outside the table is dropped rather than mis-resolved.
- `--format={text,json,summary,github}` for `diff`, and
  `--format={text,json,github}` for `audit`. `--json` and `--summary` remain as
  aliases for their formats. The `github` format emits one GitHub Actions
  workflow command per change, so a PR shows a marker against the workbook:
  breaking changes as `::error`, warnings as `::warning`, the rest as
  `::notice`. A workbook with no changes emits a single notice so the step does
  not look skipped.
- `render_github` and `render_audit_github` are exported from the package root.
- A reusable composite action under `.github/actions/sheetdelta-diff` that
  installs sheetdelta, runs the selected mode in the `github` format, writes a
  summary to the job page, and exits with the diff's code. An example workflow
  shows it diffing a workbook against the base revision.
- The rename matcher now tolerates a light edit. A sheet renamed in the same
  commit as a small change was previously reported as a remove plus an add,
  which buried the real diff; it is now recognised as a rename as long as at
  least 90% of its cells are unchanged.

### Changed

- The JSON `shifts` field is now a list on every sheet, empty when there was no
  shift. It was a single object or absent before.

## [0.3.0] - 2026-10-07

### Fixed

- A shared formula whose range ends on another sheet, such as
  `='Other'!A1:B2`, shifted both endpoints when the fill moved, so the
  reference pointed at the wrong cells. Excel keeps the whole cross-sheet
  reference fixed; it now does too.
- `audit` crashed with `RecursionError` on a workbook with a long dependency
  chain -- a running-balance column filled down a few thousand rows was
  enough. The cycle walk now uses an explicit stack, like the diff's reach
  already did.
- Deleting a sheet that fed another sheet passed a `--fail-on breaking` check.
  Excel rewrites the reader to `#REF!`, but a file written by another tool
  keeps the stale reference and showed no cell change at all. The removed
  cells that other cells still read are now reported as breaking, with the
  readers they strand.
- Two numbers written as `1` and `1.0` were reported as a value change. They
  are compared as normalized decimals now.

### Changed

- A cell's number format is no longer part of its identity. Switching a cell
  from a date format to a plain one showed `2023-10-01` becoming `45200` even
  though the stored serial was identical; that is a display change, not a
  value change. A cell's raw value is kept alongside its rendered text and
  compared for value changes, while reports still show the formatted text.
- A defined name scoped to a single sheet no longer overwrites a workbook-wide
  name of the same word. Sheet-scoped names are skipped rather than resolved
  against the wrong sheet.
- `--json` and `--summary` are now mutually exclusive instead of silently
  preferring JSON.
- The reader parses `xl/workbook.xml` and its rels once per file instead of
  twice, since both are read on the hot path for large workbooks.

## [0.2.2] - 2026-10-07

### Fixed

- `[#Totals]` in a structured reference resolved to the last data row instead of
  the totals row, so `SUM(Table1[#Totals])` pointed one row too high. The
  existing test asserted the wrong row and was corrected with it.
- `audit --fail-on breaking` exited `0` even when it found broken or circular
  references, which could let a CI job pass on a workbook with a `#REF!`. An
  audit has no "breaking" tier -- any issue it finds is a real defect -- so
  `breaking` now fails like `any`, and only `never` always passes.
- A shared formula that referenced another sheet, such as `='Other Sheet'!A1+B1`,
  shifted its cross-sheet reference along with the local one when the block was
  filled down. A cross-sheet reference is anchored in that sheet's grid, so it
  now stays put while the local reference moves.
- A worksheet part that was missing or malformed inside the package was silently
  dropped, so an audit could report a partial workbook as sound. The unreadable
  sheet is now recorded, reported as an incomplete result, and fails the audit.
- A whole-column range such as `A:A` walked every one of the grid's 1,048,576
  rows even when only a handful held cells. The range lookup now binary-searches
  the rows that hold cells, so it costs the cells it touches.

## [0.1.0] - 2026-10-07

First release.

### Added

- `sheetdelta diff OLD NEW`: compares two `.xlsx` workbooks cell by cell,
  including formulas and the values Excel cached, and follows each change
  through the workbook's dependency graph to the cells it reaches.
- `sheetdelta audit FILE`: reports formulas whose references do not resolve --
  a sheet that is not in the workbook, or a circular reference.
- `--fail-on {never,any,breaking}` to set the exit code, so the tool works as a
  CI check without extra scripting.
- `--json` output for both subcommands.
- A standard-library `.xlsx` reader: no runtime dependencies.
- Rename detection, so renaming a sheet does not report every cell as changed.
- Stale-cell detection, for a formula edited without the workbook being
  recalculated.

## [0.2.1] - 2026-10-07

### Fixed

- `Shift` was listed in `__all__` but never imported, so `sheetdelta.Shift`
  raised `AttributeError` even though it was advertised. A test now walks
  `__all__` and fails if any advertised name is missing.

## [0.2.0] - 2026-10-07

### Added

- Row and column insert/remove detection. An insert is reported once, as
  `1 row inserted at row 3`, instead of every cell below it reading as changed.
  The old and new cells are lined up by content before they are compared, so an
  edit that moved with the insert is still reported. The shift is only claimed
  when the block below it really does line up.
- Structured references to Excel tables, such as `SUM(Sales[Amount])` and
  `Sales[[#Totals],[Amount]]`. The table definition is read from the workbook,
  the reference is resolved to the cells it names, and edits to those cells are
  followed through the dependency graph to the formula that reads them.
- `--summary` for `diff`: one line per sheet with counts, for a CI log.
- `Shift`, `Table` and `TableRef` are exported from the package root.

## [0.1.1] - 2026-10-07

### Fixed

- The dependency graph asked every formula reference against every cell in the
  workbook, which is quadratic. A 16,000-cell workbook took about 30 seconds to
  diff; it now takes under a second. A cell lookup goes straight to its address
  and a range is walked row by row over only the rows that hold cells, so
  `SUM(A:A)` costs the cells it touches rather than a million probes.

[0.5.0]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.5.0
[0.4.0]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.4.0
[0.3.0]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.3.0
[0.2.2]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.2.2
[0.2.1]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.2.1
[0.2.0]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.2.0
[0.1.1]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.1.1
[0.1.0]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.1.0
