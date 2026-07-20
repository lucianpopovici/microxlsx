# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `XLSXPackage.resize_table(table_name, *, add_rows=0, add_cols=0)` — grow or
  shrink a table along the row and/or column axis.
  - **Minimal, cascading shove:** growing shoves only the tables that actually
    collide — **down** for row growth, **right** for column growth — each by the
    least amount needed to clear the target, propagating through further tables.
    Tables that don't overlap on the cross axis, or are separated by a gap, are
    left untouched.
  - **Column metadata:** column growth appends `tableColumn` entries with unique
    ids/names and writes the new header cells; shrinking removes trailing
    columns. The `tableColumns` `count` stays in sync with the range width.
- Move-time reference rewriting, so references follow relocated data instead of
  dangling when a table is shoved:
  - **Formulas** anywhere on the sheet that point into a moved block are shifted
    to the new location — preserving `$`-absolute markers and range endpoints,
    while skipping function names, cross-sheet references, and cells outside the
    block.
  - **Merged cells** (`mergeCell`) contained in the moved block shift with it.
  - **Conditional-formatting** and **data-validation** regions (their `sqref`,
    including the multi-range form) and the formulas inside their rules
    (`cfRule`, `formula1`/`formula2`) shift with the moved block.
  - **Named ranges** (workbook `definedName` entries) whose range points into
    the moved block shift, matched to the moving table's sheet and treating a
    sheet-qualified range as one unit.

### Notes

- **Structured table references** (`Table[Col]`) resolve by name through the
  table definition and therefore follow moves and resizes automatically — no
  rewriting is applied or needed.
- Ranges that only *partially* overlap a moved block are left unchanged.

## [0.1.0]

### Added

- Initial release of MicroXLSX: a surgical, zero-dependency XLSX modifier that
  edits only the XML parts it touches and streams everything else (macros,
  images, themes) through byte-for-byte.
- `XLSXPackage` with `update_cell`, `update_table_cell` (auto-expanding a
  table's range), `merge_cells`, and `save`.
- Sheet and table relationship mapping, and cell-reference conversion helpers.

[Unreleased]: https://github.com/lucianpopovici/microxlsx/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/lucianpopovici/microxlsx/releases/tag/v0.1.0
