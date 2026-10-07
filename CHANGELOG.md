# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

## [0.1.1] - 2026-10-07

### Fixed

- The dependency graph asked every formula reference against every cell in the
  workbook, which is quadratic. A 16,000-cell workbook took about 30 seconds to
  diff; it now takes under a second. A cell lookup goes straight to its address
  and a range is walked row by row over only the rows that hold cells, so
  `SUM(A:A)` costs the cells it touches rather than a million probes.

[0.1.1]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.1.1
[0.1.0]: https://github.com/sqmyou/sheetdelta/releases/tag/v0.1.0
