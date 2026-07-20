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

`resize_table` grows (or shrinks) a table by a number of rows. When growing,
any tables that would collide below it are shoved **down by the minimal amount**
needed to clear the target, cascading through further tables. Tables in other
columns, or with an existing gap, are left untouched.

```python
pkg = XLSXPackage("report.xlsx")
pkg.resize_table("SalesTable", add_rows=5)   # grow, push colliding tables down
pkg.resize_table("Notes", add_rows=-2)       # shrink (never moves other tables)
pkg.save("output.xlsx")
```

Current scope: row-axis resizing only. When a table is moved, its cell block
(values, formulas, styles) is relocated, but formula *references* into the
moved region, merged-cell ranges, and conditional-formatting ranges are not
yet rewritten. Column-axis resizing is planned.
