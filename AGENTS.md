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
  file format; it re-addresses every cell below. `differ._detect_shift` lines
  the contents up before comparing so the insert reads as one event. The claim
  is only made when the block below really lines up (`_shifts_agree`), so two
  unrelated sheets are never aligned by force.
- **Structured references need the table definition.** `Sales[Amount]` only
  resolves if the table was read. `references.extract_references` takes a
  `table_names` map and `TableRef.resolve` turns it into a range. A table name
  that is not in the map is not treated as a table, which keeps a bare
  `Foo[Bar]` from looking like a reference.
- **`__all__` is not checked by any tool.** A name listed there but never
  imported is a runtime `AttributeError` only. `tests/test_exports.py` walks the
  list; keep it passing.

## Releasing

Bump `version` in `pyproject.toml` and `__version__` in `src/sheetdelta/__init__.py`,
add a CHANGELOG entry, then `gh release create vX.Y.Z --notes-file ...`. The
Release workflow builds and publishes to PyPI on `release: published`. PyPI's
JSON endpoint can lag the simple index by a minute; confirm with
`pip install --target` rather than the JSON.

## Test fixtures

`tests/xlsx_fixtures.py` writes `.xlsx` files with the standard library, so the
suite runs with no dependencies. `tests/test_interop.py` checks the same reader
against workbooks written by openpyxl, and skips if openpyxl is absent.
