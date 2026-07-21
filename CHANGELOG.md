# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-07-21

### Fixed

- Package-absolute relationship targets (`Target="/xl/worksheets/sheet1.xml"`,
  as written by Excel and openpyxl) are now resolved correctly. Previously they
  produced a malformed `xl//xl/...` path, so microxlsx could not open workbooks
  authored by those tools at all — only files with relative targets. Table and
  worksheet-rels paths are resolved the same way (handling `..` and the owning
  part's directory).

### Added

- `add_hyperlink(sheet, cell, url, tooltip=)` — attach an external hyperlink
  (worksheet `hyperlink` + external relationship).
- `add_image(sheet, cell, image, width=, height=)` — anchor a PNG/JPEG (path or
  bytes) via a media part + DrawingML anchor + relationships; PNG dimensions
  are auto-detected.
- `add_comment(sheet, cell, text, author=)` — attach a cell note (comments part
  + VML drawing + `legacyDrawing`); multiple comments per sheet and repeated
  authors are handled.
- `set_auto_filter(sheet, ref)` adds worksheet filter dropdowns; new tables
  from `add_table` include an autoFilter by default, kept in sync with the
  table's `ref` through resize / move / insert / delete.
- Page setup: `set_page_setup(sheet, orientation=, fit_to_width=, fit_to_height=)`,
  `set_print_area(sheet, ref)` (an `_xlnm.Print_Area` defined name), and
  `set_header_footer(sheet, header=, footer=)`.
- `protect_sheet(sheet, password=None)` enables worksheet protection using
  Excel's legacy 16-bit password hash.
- `freeze_panes(sheet, cell)` — freeze rows above / columns left of a cell.
- `rename_sheet(old, new)` — rename a sheet and rewrite `OldName!` /
  `'Old Name'!` qualifiers across every sheet's formulas and the workbook's
  defined names (quoting the new name when needed).
- `add_defined_name(name, ref, sheet_name=None)` — create a workbook-global or
  sheet-scoped defined name.
- `add_table` applies a banded built-in table style by default
  (`style_name="TableStyleMedium2"`, `None` to omit).
- Test suite validates output against **openpyxl**, an independent reader:
  baselines authored by openpyxl are modified with microxlsx and reopened with
  openpyxl to confirm the workbook stays valid and each change took effect. CI
  now runs the test suite (previously it only linted).
- `insert_rows` / `delete_rows` / `insert_cols` / `delete_cols` — structural
  edits mid-sheet. Cell data, row heights, column widths, merges,
  conditional-formatting and data-validation regions, formulas, and named
  ranges all shift; ranges straddling an insertion point grow, and a table
  straddling it grows too (a table below/right of it moves). On delete,
  references into the removed band become `#REF!` (single cells) or are
  clamped (ranges); merges/regions swallowed by the band are removed.
  Guards: deleting a table's header row or inserting/deleting through a
  table's column span raises `ValueError` (use
  `resize_table`/`remove_table`).
- `get_cell_style(sheet, ref)` — a cell's raw `style_id` for faithful reuse —
  and `get_style(style_id)` — best-effort decode into `add_style`-compatible
  kwargs (theme colors and mixed borders come back as raw dicts).
- Bulk range operations: `write_range(sheet, start_ref, rows)` writes a 2D
  block in a single tree pass (`None` entries leave cells untouched),
  `get_range(sheet, ref)` reads a rectangle as lists (missing cells → `None`),
  and `iter_table_rows(table)` yields each data row as a
  `{column_name: value}` dict.
- `add_style(...)` — compose a full cell style (bold/italic, font size, name
  and color, solid fill, uniform border, horizontal/vertical alignment, wrap,
  optional number format) into a reusable `style_id`. Identical calls dedupe.
- Structural operations: `add_sheet(name)`, `remove_sheet(name)` (removes its
  tables too; refuses to remove the last sheet), `add_table(sheet, name, ref,
  columns)`, and `remove_table(name)`. These maintain `[Content_Types].xml`
  overrides, workbook/worksheet relationships, `tableParts`, and part files, and
  `save` now writes newly created parts and omits removed ones.
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

### Changed

- `get_cell` / `get_range` now return `datetime.date` / `datetime.datetime`
  for numeric cells whose style carries a date number format (built-in ids and
  custom codes are both detected), so dates written with `update_cell` round-
  trip instead of coming back as raw serial numbers.

### Fixed

- Date serial conversion honours the Mac 1904 date system
  (`workbookPr date1904="1"`) in both directions; previously such workbooks
  would get 1900-based serials (~4 years off).
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

[Unreleased]: https://github.com/lucianpopovici/microxlsx/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/lucianpopovici/microxlsx/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/lucianpopovici/microxlsx/releases/tag/v0.1.0
