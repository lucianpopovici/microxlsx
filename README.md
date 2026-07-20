# MicroXLSX: The Surgical Excel Modifier

MicroXLSX is a lightweight, zero-dependency Python library designed for one thing: modifying existing Excel files without breaking them.
Why use MicroXLSX instead of OpenPyXL?

Most libraries (OpenPyXL, Pandas) work by parsing the entire spreadsheet into an object model and then re-generating the file from scratch. This often strips out:

  *  VBA Macros (.xlsm)

  *  Custom UI Ribbons

  *  Pivot Table Cache Definitions

  *  Complex DrawingML Effects

MicroXLSX uses a "Surgical" approach. It only opens the specific XML files you want to change and streams the rest of the ZIP archive (Macros, Images, Themes) bit-for-bit from source to destination.
Features

   * 🚀 Blazing Fast: No overhead of parsing the whole workbook.

   * 🛡️ Non-Destructive: 100% preservation of unknown XML parts and binary blobs.

   * 📦 Zero Dependencies: Uses only the Python Standard Library.

   * 📎 Macro-Friendly: Perfect for updating data in .xlsm templates.

** 🚀 Usage Example

```python
import datetime
from microxlsx import XLSXPackage

pkg = XLSXPackage("template.xlsm")

# Read existing values (resolves shared strings, typed cells, cached formulas)
current = pkg.get_cell("Sheet1", "B2")
qty = pkg.get_table_cell("SalesTable", 1, "Amount")

# Write values, formulas, booleans, dates, merges
pkg.update_table_cell("SalesTable", 1, "Amount", 500.25)
pkg.update_cell("Sheet1", "D10", formula="SUM(B1:B9)")
pkg.update_cell("Sheet1", "E10", value=True)
pkg.update_cell("Sheet1", "F10", value=datetime.date(2026, 1, 31))  # auto date format

# Add a number format and reuse its style id
money = pkg.add_number_format("$#,##0.00")
pkg.update_cell("Sheet1", "G10", value=1999.9, style_id=money)

# Compose full cell styles: font, fill, border, alignment
header = pkg.add_style(bold=True, fill_color="#DDEBF7", border="thin",
                       align="center")
pkg.update_cell("Sheet1", "A1", value="Report", style_id=header)

# Bulk block writes/reads — one tree pass instead of a call per cell
pkg.write_range("Sheet1", "B2", [
    ["Region", "Amount", "When"],
    ["West",   320.5,    datetime.date(2026, 1, 31)],
    ["East",   210.0,    datetime.date(2026, 2, 28)],
])
block = pkg.get_range("Sheet1", "B2:D4")
for row in pkg.iter_table_rows("SalesTable"):   # dicts keyed by column name
    print(row["Region"], row["Amount"])

# Inspect, append rows, clear cells, size columns/rows
pkg.sheet_names(); pkg.table_names(); pkg.table_dimensions("SalesTable")
pkg.append_table_row("SalesTable", {"Region": "West", "Amount": 320})
pkg.clear_cell("Sheet1", "H10")
pkg.set_column_width("Sheet1", "B", 18)
pkg.set_row_height("Sheet1", 1, 24)

# Add / remove sheets and tables (relationships + content-types handled)
pkg.add_sheet("Summary")
pkg.add_table("Summary", "Totals", "A1:C10", ["Region", "Q1", "Q2"])
pkg.remove_table("OldTable")
pkg.remove_sheet("Scratch")

# Insert / delete rows and columns — references follow, tables adjust
pkg.insert_rows("Sheet1", 4, count=2)   # everything below shifts down
pkg.delete_cols("Sheet1", "F")          # refs into F become #REF!, rest shift

# Reuse or inspect existing styles
tmpl = pkg.get_cell_style("Sheet1", "B2")            # raw id, always faithful
pkg.update_cell("Sheet1", "B10", value=99, style_id=tmpl)
pkg.get_style(tmpl)                                  # decoded add_style kwargs

pkg.merge_cells("Sheet1", "A1:C1")
pkg.save("output.xlsm")
```

`append_table_row` grows the table and shoves any table directly below it out
of the way (minimal, cascading) — the same collision handling as `resize_table`.

`get_cell` / `get_table_cell` return `str` / `int` / `float` / `bool` (or a
formula cell's cached result), or `None` for an empty cell. Reading a sheet you
haven't modified doesn't disturb it — it's still streamed through untouched on
save. `datetime.date` / `datetime.datetime` values are written as Excel serial
numbers and auto-formatted (pass an explicit `style_id` to override) — and read
**back** as `date`/`datetime` when the cell carries a date number format, so
dates round-trip. The Mac 1904 date system (`workbookPr date1904`) is honoured
in both directions. `None` entries in a `write_range` block leave the existing
cells untouched.
`add_number_format` registers a custom format in `xl/styles.xml` and returns a
reusable `style_id`. Editing a formula (or a value a formula depends on)
invalidates the cached results — the workbook is flagged for recalculation, and
a formula edit additionally drops the stale `xl/calcChain.xml` and its
references — so Excel opens cleanly instead of warning about recovered content.

** 📐 Resizing tables with minimal movement

`resize_table` grows (or shrinks) a table by a number of rows and/or columns.
When growing, any tables that would collide are shoved by the **minimal amount**
needed to clear the target — **down** for row growth, **right** for column
growth — cascading through further tables. Tables that don't overlap on the
cross axis, or that have an existing gap, are left untouched.

```python
pkg = XLSXPackage("report.xlsx")
pkg.resize_table("SalesTable", add_rows=5)            # grow rows, push down
pkg.resize_table("SalesTable", add_cols=2)            # grow cols, push right
pkg.resize_table("SalesTable", add_rows=3, add_cols=1)  # both at once
pkg.resize_table("Notes", add_rows=-2)               # shrink (never moves others)
pkg.save("output.xlsx")
```

Column growth also appends `tableColumn` metadata (with unique ids/names) and
writes the new header cells. When a table is moved, its cell block (values,
formulas, styles) is relocated, and references follow the data:

  * **Formulas** anywhere on the sheet that point at a moved cell are rewritten
    to its new location — including `$`-absolute refs and range endpoints.
    Function names, cross-sheet references, and cells outside the moved block
    are left untouched.
  * **Merged-cell ranges** contained in the moved block shift with it.
  * **Conditional-formatting** and **data-validation** regions (their `sqref`)
    and the formulas inside their rules shift with the moved block.
  * **Named ranges** (workbook `definedName` entries) whose range points into
    the moved block are shifted, matched to the moving table's sheet.

**Structured table references** (`Table[Col]`) resolve by name through the
table definition, so they follow moves and resizes automatically — no rewriting
needed. Ranges that only *partially* overlap the moved block are left unchanged.

Because moves rewrite formula text but don't recompute cached results, the
workbook is flagged for a full recalculation on load (`fullCalcOnLoad`) so no
stale cached value is trusted. Shared strings and shared formulas are preserved
as-is and need no special handling.
