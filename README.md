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
