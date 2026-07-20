# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Inspection helpers: `sheet_names()`, `table_names()`, and
  `table_dimensions(table_name)` (rows × cols, header included).
- `clear_cell(sheet_name, cell_ref)` — remove a cell (the row is kept) and flag
  a recalc.
- `append_table_row(table_name, values)` — append a data row (dict by column
  name or positional list), growing the table and shoving any table directly
  below it out of the way (minimal, cascading), just like `resize_table`.
- `set_column_width(sheet_name, column, width)` and
  `set_row_height(sheet_name, row, height)`.
- `XLSXPackage.add_number_format(format_code)` — register a custom number
  format in `xl/styles.xml` and get back a reusable `style_id` (deduped by
  code). Enables currency / percentage / date formatting from scratch.
- Date support: `update_cell(value=datetime.date | datetime.datetime)` writes
  the value as an Excel serial number and, unless a `style_id` is given,
  auto-applies a default date/datetime number format.
- `XLSXPackage.get_cell(sheet_name, cell_ref)` and
  `get_table_cell(table_name, row_offset, col_name)` — read a cell's value,
  resolving shared strings (including rich-text runs), inline strings, booleans,
  numbers, error values, and a formula cell's cached result. Reading a sheet
  that hasn't been modified leaves it streamed through untouched on save.
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
  - After a move rewrites formulas, the workbook is flagged
    `calcPr fullCalcOnLoad="1"` so Excel recomputes on open and no stale cached
    `<v>` result is trusted.
- Editing or moving a formula now invalidates the cached calculation chain:
  `xl/calcChain.xml` is dropped and its `[Content_Types].xml` override and
  workbook relationship are removed, so Excel rebuilds it cleanly instead of
  warning about recovered content.
- Editing a cell value now also flags the workbook for recalculation
  (`fullCalcOnLoad`), so a formula that reads the changed cell doesn't keep a
  stale cached result. The calc chain is kept for value edits (its structure is
  unchanged).

### Fixed

- `update_cell(value=True/False)` now writes a proper boolean cell (`t="b"`
  with `1`/`0`) instead of an invalid bare `<v>True</v>`.

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
