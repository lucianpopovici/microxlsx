"""
Unit tests for microxlsx.core
"""
import zipfile
import xml.etree.ElementTree as ET
import pytest

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


def _sheet_rels(*rids):
    parts = [
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        b'<Relationships xmlns="http://schemas.openxmlformats.org/'
        b'package/2006/relationships">',
    ]
    for i, _ in enumerate(rids, start=1):
        parts.append(
            b'<Relationship Id="rId%d"'
            b' Type="http://schemas.openxmlformats.org/officeDocument/'
            b'2006/relationships/table"'
            b' Target="../tables/table%d.xml"/>' % (i, i)
        )
    parts.append(b"</Relationships>")
    return b"".join(parts)


def _table_xml(display_name, ref):
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        b' displayName="%s" ref="%s">'
        b'<tableColumns count="2">'
        b'<tableColumn name="%s_c1" id="1"/>'
        b'<tableColumn name="%s_c2" id="2"/>'
        b"</tableColumns>"
        b"</table>"
        % (display_name, ref, display_name, display_name)
    )


def make_multi_table_xlsx(tmp_path):
    """XLSX with four tables on one sheet for collision/shove testing.

    Layout (rows are 1-based, cols shown):
        Top    A1:B3   (cols A-B)
        Bottom A5:B7   (cols A-B, directly under Top)
        Far    A9:B11  (cols A-B, under Bottom)
        Side   D1:E7   (cols D-E, never collides with A-B growth)
    Tracker cells: A5=B, A9=F, D1=S.
    """
    path = str(tmp_path / "multi.xlsx")
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b"<sheetData>"
        b'<row r="1"><c r="A1" t="inlineStr"><is><t>Top</t></is></c>'
        b'<c r="D1" t="inlineStr"><is><t>S</t></is></c></row>'
        b'<row r="5"><c r="A5" t="inlineStr"><is><t>B</t></is></c></row>'
        b'<row r="9"><c r="A9" t="inlineStr"><is><t>F</t></is></c></row>'
        b"</sheetData>"
        b"</worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            _sheet_rels("rId1", "rId2", "rId3", "rId4"),
        )
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Top", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Bottom", b"A5:B7"))
        zf.writestr("xl/tables/table3.xml", _table_xml(b"Far", b"A9:B11"))
        zf.writestr("xl/tables/table4.xml", _table_xml(b"Side", b"D1:E7"))
    return path


def make_hcollision_xlsx(tmp_path):
    """XLSX with three tables for column-axis collision testing.

    Layout:
        Left  A1:B3  (cols A-B)
        Right D1:E3  (cols D-E, to the right of Left with a one-col gap)
        Below A5:B7  (cols A-B, different rows -> unaffected by col growth)
    Tracker cells: A1=L, D1=R, A5=Btm.
    """
    path = str(tmp_path / "hcol.xlsx")
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b"<sheetData>"
        b'<row r="1"><c r="A1" t="inlineStr"><is><t>L</t></is></c>'
        b'<c r="D1" t="inlineStr"><is><t>R</t></is></c></row>'
        b'<row r="5"><c r="A5" t="inlineStr"><is><t>Btm</t></is></c></row>'
        b"</sheetData>"
        b"</worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            _sheet_rels("rId1", "rId2", "rId3"),
        )
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Left", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Right", b"D1:E3"))
        zf.writestr("xl/tables/table3.xml", _table_xml(b"Below", b"A5:B7"))
    return path


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
        assert pkg.table_map == {}

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


# ---------------------------------------------------------------------------
# Tests: resize_table (collision-aware minimal shoving)
# ---------------------------------------------------------------------------


class TestResizeTable:
    def _get_cell(self, pkg, cell_ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        return root.find(f".//{{{NS}}}c[@r='{cell_ref}']")

    def _table_ref(self, pkg, name):
        return pkg.trees[pkg.table_map[name]["xml_path"]].getroot().get("ref")

    def test_noop_when_zero(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path, with_table=True))
        pkg.resize_table("SalesTable", add_rows=0)
        assert self._table_ref(pkg, "SalesTable") == "A1:B3"

    def test_grows_target_ref(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert pkg.table_map["Top"]["range"] == ["A1", "B6"]
        assert self._table_ref(pkg, "Top") == "A1:B6"

    def test_colliding_table_shifted(self, tmp_path):
        # Top A1:B3 grows +3 → bottom row 6; Bottom (A5:B7) must clear to A7:B9.
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert pkg.table_map["Bottom"]["range"] == ["A7", "B9"]
        assert self._table_ref(pkg, "Bottom") == "A7:B9"

    def test_colliding_table_cell_physically_moved(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        moved = self._get_cell(pkg, "A7")
        assert moved is not None
        assert moved.find(f".//{{{NS}}}t").text == "B"
        assert self._get_cell(pkg, "A5") is None  # vacated

    def test_cascade_shift(self, tmp_path):
        # Bottom shifts down 2 → collides with Far (A9:B11); Far shifts down 1.
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert pkg.table_map["Far"]["range"] == ["A10", "B12"]
        far_cell = self._get_cell(pkg, "A10")
        assert far_cell is not None
        assert far_cell.find(f".//{{{NS}}}t").text == "F"

    def test_non_overlapping_columns_not_moved(self, tmp_path):
        # Side (D1:E7) shares rows but different columns → must stay put.
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert pkg.table_map["Side"]["range"] == ["D1", "E7"]
        assert self._get_cell(pkg, "D1").find(f".//{{{NS}}}t").text == "S"

    def test_gap_preserved_no_shift(self, tmp_path):
        # Growing by only 1 keeps Top's bottom at row 4, below is row 5 → no clash.
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=1)
        assert pkg.table_map["Bottom"]["range"] == ["A5", "B7"]

    def test_shrink_updates_ref_only(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Bottom", add_rows=-1)
        assert pkg.table_map["Bottom"]["range"] == ["A5", "B6"]
        assert pkg.table_map["Far"]["range"] == ["A9", "B11"]  # unaffected

    def test_shrink_below_header_raises(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.resize_table("Top", add_rows=-3)

    def test_save_roundtrip_persists_moves(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            table2 = zf.read("xl/tables/table2.xml").decode("utf-8")
        assert 'r="A7"' in sheet
        assert 'ref="A7:B9"' in table2


class TestResizeTableColumns:
    def _get_cell(self, pkg, cell_ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        return root.find(f".//{{{NS}}}c[@r='{cell_ref}']")

    def _table_root(self, pkg, name):
        return pkg.trees[pkg.table_map[name]["xml_path"]].getroot()

    def test_grows_target_ref(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        assert pkg.table_map["Left"]["range"] == ["A1", "E3"]
        assert self._table_root(pkg, "Left").get("ref") == "A1:E3"

    def test_table_columns_appended(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        cols = self._table_root(pkg, "Left").find(f"{{{NS}}}tableColumns")
        assert cols.get("count") == "5"
        assert len(list(cols)) == 5
        # ids stay unique
        ids = [c.get("id") for c in cols]
        assert len(set(ids)) == 5

    def test_columns_map_updated(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=2)
        cols = pkg.table_map["Left"]["columns"]
        assert cols["Column3"] == 2
        assert cols["Column4"] == 3

    def test_header_cells_written(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        # New columns C1/D1/E1 carry their column names as header text.
        assert self._get_cell(pkg, "C1").find(f".//{{{NS}}}t").text == "Column3"
        assert self._get_cell(pkg, "E1").find(f".//{{{NS}}}t").text == "Column5"

    def test_colliding_table_shifted_right(self, tmp_path):
        # Left A1:B3 +3 cols -> right edge E; Right (D1:E3) must clear to F1:G3.
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        assert pkg.table_map["Right"]["range"] == ["F1", "G3"]
        assert self._table_root(pkg, "Right").get("ref") == "F1:G3"

    def test_colliding_cell_physically_moved(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        moved = self._get_cell(pkg, "F1")
        assert moved is not None
        assert moved.find(f".//{{{NS}}}t").text == "R"

    def test_different_rows_not_moved(self, tmp_path):
        # Below (A5:B7) shares columns but not rows -> stays put on col growth.
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        assert pkg.table_map["Below"]["range"] == ["A5", "B7"]
        assert self._get_cell(pkg, "A5").find(f".//{{{NS}}}t").text == "Btm"

    def test_gap_preserved_no_shift(self, tmp_path):
        # Growing by 1 keeps Left's right edge at C, clear of Right at D.
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=1)
        assert pkg.table_map["Right"]["range"] == ["D1", "E3"]

    def test_shrink_removes_columns(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Right", add_cols=-1)
        assert pkg.table_map["Right"]["range"] == ["D1", "D3"]
        cols = self._table_root(pkg, "Right").find(f"{{{NS}}}tableColumns")
        assert cols.get("count") == "1"
        assert "Right_c2" not in pkg.table_map["Right"]["columns"]

    def test_shrink_all_columns_raises(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.resize_table("Left", add_cols=-3)

    def test_combined_row_and_col_resize(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=1, add_cols=1)
        assert pkg.table_map["Top"]["range"] == ["A1", "C4"]
        cols = self._table_root(pkg, "Top").find(f"{{{NS}}}tableColumns")
        assert cols.get("count") == "3"
        assert self._get_cell(pkg, "C1") is not None

    def test_save_roundtrip_persists_col_growth(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            table1 = zf.read("xl/tables/table1.xml").decode("utf-8")
            table2 = zf.read("xl/tables/table2.xml").decode("utf-8")
        assert 'ref="A1:E3"' in table1
        assert 'count="5"' in table1
        assert 'ref="F1:G3"' in table2
        assert 'r="F1"' in sheet
