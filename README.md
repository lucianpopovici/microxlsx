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
from microxlsx import XLSXPackage

pkg = XLSXPackage("template.xlsm")
pkg.update_table_cell("SalesTable", 1, "Amount", 500.25)
pkg.update_cell("Sheet1", "D10", formula="SUM(B1:B9)")
pkg.merge_cells("Sheet1", "A1:C1")
pkg.save("output.xlsm")
```

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
