# sheetdelta

Compare two Excel workbooks and find out what actually changed.

`sheetdelta` reads two `.xlsx` files, shows you the difference cell by cell, and
then tells you which of those changes reach other cells. It runs on Linux, in
CI, with no copy of Excel anywhere near it.

```console
$ sheetdelta diff budget_v1.xlsx budget_v2.xlsx
budget_v1.xlsx  ->  budget_v2.xlsx

Sheet 'Revenue'  (1 change, 1 breaking)
  !! D12    formula changed
          -  =SUM(D2:D11)
          +  =SUM(D2:D13)
        affects 3 cells downstream: Revenue!E12, Revenue!F12, Summary!B4

Sheet 'Summary'  (1 change)
  !  B7     value changed
          -  1200
          +  1450

1 formula; 1 value across 2 edited sheet(s). 1 breaking change reaching 3 downstream cells.
```

## Why

Two things go wrong when a spreadsheet changes and nobody looks closely.

The first is the change itself. `git diff` on an `.xlsx` gives you binary
noise, so the usual answer is to open both files side by side and squint.

The second is worse. A formula edit looks harmless on the sheet you are
looking at, but the cell feeds a summary three sheets away. The number that
matters is now wrong, and nothing tells you.

`sheetdelta` answers both. Every change is followed through the dependency
graph, so a formula edit reports the cells it reaches, and a CI job can fail
the build on exactly that.

## Install

```console
pip install sheetdelta
```

No dependencies. The reader is `zipfile` and `xml.etree` from the standard
library, so it drops into any environment, including a build container that
has nothing but Python.

## Usage

### Compare two workbooks

```console
sheetdelta diff old.xlsx new.xlsx
```

Exit status is `0` when nothing breaking changed, `1` when something did, `2`
on an error. That makes it a CI check without any extra scripting:

```yaml
- run: pip install sheetdelta
- run: sheetdelta diff before.xlsx after.xlsx --fail-on breaking
```

`--fail-on` takes three values:

| Value | Fails when |
|---|---|
| `breaking` (default) | a change reaches another cell |
| `any` | anything changed at all |
| `never` | never; useful for a report-only run |

### Machine-readable output

All subcommands take `--format`:

```console
sheetdelta diff old.xlsx new.xlsx --format json
```

`--json` and `--summary` are kept as aliases for `--format json` and
`--format summary`. They are mutually exclusive with each other, and with
`--format`; pick one.

```json
{
  "old": "budget_v1.xlsx",
  "new": "budget_v2.xlsx",
  "has_changes": true,
  "summary": {
    "cell_changes": {"formula": 1, "value": 1},
    "breaking": 1
  },
  "sheets": [
    {
      "kind": "changed",
      "name": "Revenue",
      "old_name": "Revenue",
      "shifts": [
        {"axis": "row", "inserted": true, "count": 1, "at": 3, "label": "1 row inserted at row 3"}
      ],
      "changes": [
        {
          "cell": "Revenue!D12",
          "a1": "D12",
          "kind": "formula",
          "severity": "breaking",
          "detail": "formula changed",
          "old": "=SUM(D2:D11)",
          "new": "=SUM(D2:D13)",
          "affected": ["Revenue!E12", "Revenue!F12", "Summary!B4"]
        }
      ]
    }
  ]
}
```

`shifts` is always a list, empty when the sheet held no insert or remove, so a
consumer never has to branch on its presence.

### Summary output

```console
sheetdelta diff old.xlsx new.xlsx --summary
sheetdelta diff old.xlsx new.xlsx --format summary
```

The full report is what you want on a laptop. In a CI log you often want the
shape of the change and nothing else: one line per sheet, with the counts.

```console
$ sheetdelta diff budget_v1.xlsx budget_v2.xlsx --summary
budget_v1.xlsx  ->  budget_v2.xlsx
  Revenue: 1 row inserted at row 3, 2 changes, 1 breaking
  Summary: 1 change
1 formula; 2 added across 2 edited sheet(s). 1 breaking change.
```

### Check one workbook

```console
sheetdelta audit workbook.xlsx
```

An audit looks for the two defects that make a workbook unsound:

- a formula pointing at a sheet that does not exist, which Excel turns into
  `#REF!` when it recalculates; and
- a circular reference, a cell that depends on itself through a chain of other
  cells, which Excel refuses to calculate at all.

```console
$ sheetdelta audit model.xlsx
model.xlsx
  3 sheet(s), 412 cell(s), 118 formula(s)

1 broken reference(s):
  !! C4     Removed Sheet!B2
        sheet 'Removed Sheet' does not exist
        ='Removed Sheet'!B2*1.2

1 circular reference(s):
  !! Revenue!D12 -> Revenue!E12 -> Summary!B4 -> Revenue!D12
```

A reference to a cell that is simply empty is not reported. That is normal in
a spreadsheet, and flagging it would make the audit useless on real files.

### GitHub Actions

Both subcommands can emit GitHub Actions annotations:

```console
sheetdelta diff old.xlsx new.xlsx --format github
sheetdelta audit workbook.xlsx --format github
```

Each line is a workflow command, so the change shows as a marker on the pull
request. The severity carries over: a change that reaches another cell is
`::error`, an unread change is `::warning`, and additions or sheet changes are
`::notice`. When nothing changed, `diff` emits a single notice so the step does
not look skipped.

A reusable composite action wraps all of this. A workflow that compares the
workbook on the branch against the one on the base revision:

```yaml
- uses: actions/checkout@v4
  with:
    fetch-depth: 0
- run: git show "origin/${{ github.base_ref }}:models/forecast.xlsx" > /tmp/before.xlsx
- uses: sqmyou/sheetdelta/.github/actions/sheetdelta-diff@main
  with:
    old: /tmp/before.xlsx
    new: models/forecast.xlsx
    fail-on: breaking
```

The action installs sheetdelta, runs the selected `mode` (`diff` or `audit`) in
the `github` format, writes the human report to the job summary, and exits with
the diff's code. See `.github/workflows/workbook-diff.yml` for a full example.

## What it reports

| Change | Severity |
|---|---|
| Formula changed, and other cells read it | breaking |
| Value changed, and other cells read it | breaking |
| Cell deleted, and other cells read it | breaking |
| Sheet removed, and a cell another sheet read | breaking |
| Formula or value changed, nothing reads it | warning |
| Formula changed but the cached value did not move | stale |
| New cell | info |
| Sheet added or renamed | info |
| Sheet removed, and nothing read it | info |
| Row or column inserted or removed | info |

Formatting alone is not a change. Switching a cell between a date format and a
plain number shows a different string but leaves the stored value identical, so
it is not reported. Likewise `1` and `1.0` compare equal. Reports still show the
formatted text.

A **row or column insert** is reported as one event, not as every cell below
it. Excel has no "insert row" in the file format: it rewrites each cell below
to a new address. Compared naively, inserting one row near the top makes the
whole rest of the sheet look changed and buries the real edit. `sheetdelta`
lines the contents up instead, so it can say `1 row inserted at row 3` and
leave only the genuinely new cells as changes. A sheet can hold more than one:
each insert or remove is found in turn and the remaining difference re-checked,
so two separate blocks inserted far apart are both reported, and an insert and
a remove in the same sheet are too. An insert is only claimed when the block
below it really does line up, so two unrelated sheets are never aligned by
force. A real edit that moved with the insert is still reported. Shift rows are
given in the new sheet's numbering, the number a reader sees on screen.

A **structured reference** into an Excel table is understood, not treated as
text. A formula like `SUM(Sales[Amount])` is resolved against the table's
definition, so the dependency graph knows which cells it reads and an edit to
any of them is reported as reaching the formula. The current-row form,
`[@Amount]`, is resolved to the single cell on the formula's own row rather
than the whole column, so a formula that reads its own row does not look like
it depends on every row.

The **stale** case is worth explaining. Excel stores both a formula and the
last value it calculated for it. When a file is edited by something that does
not recalculate, the two disagree. Excel will happily show you the new
formula next to the old number, which is a quiet way to ship a wrong total.
`sheetdelta` flags it.

A **rename** is detected by matching the contents of a removed sheet against
an added one. Without that, renaming a sheet would look like every cell in it
changed and bury the real diff. The contents need not match exactly: a sheet
renamed in the same commit as a light edit is still recognised, as long as at
least 90% of its cells are unchanged.

## What it does not do

Being clear about this saves you time.

- **It does not evaluate formulas.** There is no calculation engine, and
  there will not be one. Excel already stored the value it last computed, and
  that is what gets compared. This is why the tool needs no Excel and no
  dependencies, and why it is fast on large files.
- **It does not write or edit workbooks.** Read-only, by design.
- **It only reads `.xlsx` and `.xlsm`.** The old `.xls` and binary `.xlsb`
  formats are rejected with a clear message rather than half-parsed.
- **It does not read VBA macros** inside `.xlsm` files.
- **It aligns an insert, but not a sort or a reorder.** Inserting or removing
  rows or columns is recognised as one event. Moving existing rows around
  without changing them is not: matching blocks that were shuffled is a
  different problem and is not attempted.

## Using it as a library

```python
from sheetdelta import diff_workbooks, read_workbook

result = diff_workbooks(read_workbook("old.xlsx"), read_workbook("new.xlsx"))

for change in result.breaking:
    print(change.ref, "->", change.affected)
    print("  was", change.old)
    print("  now", change.new)
```

Every type is exported from the package root, and the whole thing is typed.

## How it works

An `.xlsx` file is a zip containing XML. `sheetdelta` reads the sheet list, the
shared string table, the number formats and each sheet's cells, then stops.

Each formula is scanned for the cells and ranges it mentions. That scan is
more careful than it looks: `LOG10(...)` looks exactly like a reference to
column LOG row 10, and `"see A1"` looks like a reference to A1. Both are
handled, because a dependency graph built on false references is worse than
no graph.

References are kept as ranges rather than expanded. `SUM(A:A)` covers a
million cells, and expanding it would cost more than the entire diff; the
graph asks whether a range contains a cell instead.

Excel tables are read from the table parts, and a structured reference such as
`Sales[Amount]` is resolved to the column of cells it names, so the graph sees
through the table syntax to the addresses underneath.

## Requirements

Python 3.10 or newer. No runtime dependencies.

## License

MIT.
