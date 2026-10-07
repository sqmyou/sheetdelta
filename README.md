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

```console
sheetdelta diff old.xlsx new.xlsx --json
```

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

## What it reports

| Change | Severity |
|---|---|
| Formula changed, and other cells read it | breaking |
| Value changed, and other cells read it | breaking |
| Cell deleted, and other cells read it | breaking |
| Formula or value changed, nothing reads it | warning |
| Formula changed but the cached value did not move | stale |
| New cell | info |
| Sheet added, removed or renamed | info |

The **stale** case is worth explaining. Excel stores both a formula and the
last value it calculated for it. When a file is edited by something that does
not recalculate, the two disagree. Excel will happily show you the new
formula next to the old number, which is a quiet way to ship a wrong total.
`sheetdelta` flags it.

A **rename** is detected by matching the contents of a removed sheet against
an added one. Without that, renaming a sheet would look like every cell in it
changed and bury the real diff.

## What it does not do

Being clear about this saves you time.

- **It does not evaluate formulas.** There is no calculation engine, and
  there will not be one. Excel already stored the value it last computed, and
  that is what gets compared. This is why the tool needs no Excel and no
  dependencies, and why it is fast on large files.
- **It does not write or edit workbooks.** Read-only, by design.
- **It only reads `.xlsx` and `.xlsm`.** The old `.xls` and binary `.xlsb`
  formats are rejected with a clear message rather than half-parsed.
- **It does not align rows across an insertion.** If a row is inserted near
  the top, every cell below it reports as changed. Matching rows by content
  is the hard part of spreadsheet diffing and is not implemented yet.
- **It does not read VBA macros** inside `.xlsm` files.

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

## Requirements

Python 3.10 or newer. No runtime dependencies.

## License

MIT.
