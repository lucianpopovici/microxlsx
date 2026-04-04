"""
Unit tests for microxlsx.core
"""
# pylint: disable=missing-class-docstring,missing-function-docstring
import zipfile

from microxlsx.core import XLSXPackage


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# ---------------------------------------------------------------------------
# Minimal XLSX fixture helpers
# ---------------------------------------------------------------------------

WORKBOOK_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    b"<sheets>"
    b'<sheet name="Sheet1" r:id="rId1"/>'
    b"</sheets>"
    b"</workbook>"
)

WORKBOOK_RELS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1"'
    b' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
    b' Target="worksheets/sheet1.xml"/>'
    b"</Relationships>"
)

SHEET1_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b"<sheetData>"
    b'<row r="1">'
    b'<c r="A1" t="inlineStr"><is><t>Hello</t></is></c>'
    b"</row>"
    b"</sheetData>"
    b"</worksheet>"
)

SHEET1_RELS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1"'
    b' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"'
    b' Target="../tables/table1.xml"/>'
    b"</Relationships>"
)

TABLE1_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    b' displayName="SalesTable" ref="A1:B3">'
    b'<tableColumns count="2">'
    b'<tableColumn name="Name" id="1"/>'
    b'<tableColumn name="Amount" id="2"/>'
    b"</tableColumns>"
    b"</table>"
)


def make_xlsx(tmp_path, *, with_table=False):
    """Create a minimal in-memory XLSX (ZIP) file for testing."""
    path = str(tmp_path / "test.xlsx")
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", SHEET1_XML)
        if with_table:
            zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET1_RELS_XML)
            zf.writestr("xl/tables/table1.xml", TABLE1_XML)
    return path


# ---------------------------------------------------------------------------
# Tests: initialisation and mapping
# ---------------------------------------------------------------------------


class TestInit:
    def test_sheet_map_populated(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        assert "Sheet1" in pkg.sheet_map

    def test_sheet_map_path(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        assert pkg.sheet_map["Sheet1"] == "xl/worksheets/sheet1.xml"

    def test_table_map_empty_without_table(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        assert not pkg.table_map

    def test_table_map_populated(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        assert "SalesTable" in pkg.table_map

    def test_table_columns_mapped(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        cols = pkg.table_map["SalesTable"]["columns"]
        assert cols["Name"] == 0
        assert cols["Amount"] == 1

    def test_table_range(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        assert pkg.table_map["SalesTable"]["range"] == ["A1", "B3"]


# ---------------------------------------------------------------------------
# Tests: update_cell
# ---------------------------------------------------------------------------


class TestUpdateCell:
    def _get_cell(self, pkg, cell_ref):
        sheet_path = pkg.sheet_map["Sheet1"]
        root = pkg.trees[sheet_path].getroot()
        return root.find(f".//{{{NS}}}c[@r='{cell_ref}']")

    def test_numeric_value_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B2", value=42)
        cell = self._get_cell(pkg, "B2")
        assert cell is not None
        assert cell.find(f"{{{NS}}}v").text == "42"

    def test_float_value_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "C3", value=3.14)
        cell = self._get_cell(pkg, "C3")
        assert cell.find(f"{{{NS}}}v").text == "3.14"

    def test_numeric_removes_type_attr(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        # A1 already has t="inlineStr"; overwriting with a number should clear it
        pkg.update_cell("Sheet1", "A1", value=99)
        cell = self._get_cell(pkg, "A1")
        assert cell.get("t") is None

    def test_string_value_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value="World")
        cell = self._get_cell(pkg, "A1")
        assert cell.get("t") == "inlineStr"
        assert cell.find(f".//{{{NS}}}t").text == "World"

    def test_string_creates_inline_str(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "D4", value="NewCell")
        cell = self._get_cell(pkg, "D4")
        assert cell is not None
        assert cell.get("t") == "inlineStr"
        assert cell.find(f".//{{{NS}}}t").text == "NewCell"

    def test_formula_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "C3", formula="=SUM(A1:A2)")
        cell = self._get_cell(pkg, "C3")
        assert cell is not None
        f_node = cell.find(f"{{{NS}}}f")
        assert f_node is not None
        # Leading '=' should be stripped
        assert f_node.text == "SUM(A1:A2)"

    def test_formula_without_leading_equals(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "C3", formula="SUM(A1:A2)")
        cell = self._get_cell(pkg, "C3")
        assert cell.find(f"{{{NS}}}f").text == "SUM(A1:A2)"

    def test_style_id_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=1, style_id=5)
        cell = self._get_cell(pkg, "A1")
        assert cell.get("s") == "5"

    def test_style_id_zero_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=1, style_id=0)
        cell = self._get_cell(pkg, "A1")
        assert cell.get("s") == "0"

    def test_creates_new_row(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A5", value=10)
        sheet_path = pkg.sheet_map["Sheet1"]
        root = pkg.trees[sheet_path].getroot()
        row = root.find(f".//{{{NS}}}row[@r='5']")
        assert row is not None

    def test_sheet_tree_cached_after_update(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=1)
        assert "xl/worksheets/sheet1.xml" in pkg.trees


# ---------------------------------------------------------------------------
# Tests: merge_cells
# ---------------------------------------------------------------------------


class TestMergeCells:
    def _get_merge_cells(self, pkg):
        sheet_path = pkg.sheet_map["Sheet1"]
        root = pkg.trees[sheet_path].getroot()
        return root.find(f"{{{NS}}}mergeCells")

    def test_merge_cells_creates_element(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "A1:C1")
        mc = self._get_merge_cells(pkg)
        assert mc is not None

    def test_merge_cells_ref_set(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "A1:C1")
        mc = self._get_merge_cells(pkg)
        child = mc.find(f"{{{NS}}}mergeCell[@ref='A1:C1']")
        assert child is not None

    def test_merge_cells_count_one(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "A1:C1")
        mc = self._get_merge_cells(pkg)
        assert mc.get("count") == "1"

    def test_merge_cells_multiple(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "A1:C1")
        pkg.merge_cells("Sheet1", "A2:B2")
        mc = self._get_merge_cells(pkg)
        assert mc.get("count") == "2"
        assert len(list(mc)) == 2


# ---------------------------------------------------------------------------
# Tests: save
# ---------------------------------------------------------------------------


class TestSave:
    def test_save_produces_valid_zip(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value="Test")
        out = str(tmp_path / "output.xlsx")
        pkg.save(out)
        assert zipfile.is_zipfile(out)

    def test_save_contains_all_original_entries(self, tmp_path):
        path = make_xlsx(tmp_path)
        pkg = XLSXPackage(path)
        pkg.update_cell("Sheet1", "A1", value="Test")
        out = str(tmp_path / "output.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(out, "r") as dst:
            assert set(src.namelist()).issubset(set(dst.namelist()))

    def test_save_modified_sheet_persisted(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B1", value=99)
        out = str(tmp_path / "output.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            content = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "B1" in content
        assert "99" in content

    def test_save_workbook_xml_content_preserved(self, tmp_path):
        # workbook.xml is always loaded into trees during _map_sheets and re-serialized
        # on save; verify the logical content (sheet name) is still intact.
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value="Test")
        out = str(tmp_path / "output.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            content = zf.read("xl/workbook.xml").decode("utf-8")
        assert "Sheet1" in content

    def test_save_xml_declaration_present(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value="x")
        out = str(tmp_path / "output.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            content = zf.read("xl/worksheets/sheet1.xml")
        assert content.startswith(b"<?xml")


# ---------------------------------------------------------------------------
# Tests: update_table_cell
# ---------------------------------------------------------------------------


class TestUpdateTableCell:
    def _get_cell(self, pkg, cell_ref):
        sheet_path = pkg.sheet_map["Sheet1"]
        root = pkg.trees[sheet_path].getroot()
        return root.find(f".//{{{NS}}}c[@r='{cell_ref}']")

    def test_updates_correct_cell(self, tmp_path):
        # Table starts at A1. row_offset=2, col='Amount'(idx=1) → abs (row=2,col=1) → B3
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 2, "Amount", 500.25)
        cell = self._get_cell(pkg, "B3")
        assert cell is not None
        assert cell.find(f"{{{NS}}}v").text == "500.25"

    def test_updates_first_column(self, tmp_path):
        # row_offset=1, col='Name'(idx=0) → abs (row=1,col=0) → A2
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 1, "Name", "Alice")
        cell = self._get_cell(pkg, "A2")
        assert cell is not None
        assert cell.find(f".//{{{NS}}}t").text == "Alice"

    def test_style_id_forwarded(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 1, "Amount", 10, style_id=3)
        cell = self._get_cell(pkg, "B2")
        assert cell.get("s") == "3"

    def test_no_range_expansion_within_bounds(self, tmp_path):
        # abs_row=2 == curr_end_row=2, no expansion expected
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 2, "Amount", 1)
        assert pkg.table_map["SalesTable"]["range"][1] == "B3"

    def test_range_expansion_beyond_bounds(self, tmp_path):
        # abs_row=5 > curr_end_row=2; end cell should expand to B6
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 5, "Amount", 999)
        assert pkg.table_map["SalesTable"]["range"][1] == "B6"

    def test_range_expansion_updates_xml(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 5, "Amount", 999)
        t_root = pkg.trees[pkg.table_map["SalesTable"]["xml_path"]].getroot()
        assert t_root.get("ref") == "A1:B6"

    def test_range_expansion_correct_cell_written(self, tmp_path):
        # abs_row=5, abs_col=1 → B6
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.update_table_cell("SalesTable", 5, "Amount", 42)
        cell = self._get_cell(pkg, "B6")
        assert cell is not None
        assert cell.find(f"{{{NS}}}v").text == "42"
