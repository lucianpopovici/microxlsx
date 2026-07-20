"""
Unit tests for microxlsx.core
"""
import zipfile
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


def make_formula_xlsx(tmp_path):
    """XLSX with formulas + merged cells for row-move rewrite testing.

    Tables Top (A1:B3) and Bottom (A5:B7) stacked in cols A-B. Formulas live
    both inside Bottom (B5) and loose in col D/E; merges cover Top and Bottom.
    """
    path = str(tmp_path / "formula.xlsx")
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b"<sheetData>"
        b'<row r="1">'
        b'<c r="A1" t="inlineStr"><is><t>Top</t></is></c>'
        b'<c r="D1"><f>SUM(A5:A7)</f></c>'
        b'<c r="E1"><f>A1+1</f></c>'
        b"</row>"
        b'<row r="3">'
        b'<c r="D3"><f>SUM(A5:A7)+LOG10(A5)</f></c>'
        b'<c r="D4"><f>$A$5</f></c>'
        b'<c r="D5"><f>Sheet2!A5</f></c>'
        b"</row>"
        b'<row r="5">'
        b'<c r="A5" t="inlineStr"><is><t>x</t></is></c>'
        b'<c r="B5"><f>A5*2</f></c>'
        b"</row>"
        b"</sheetData>"
        b'<mergeCells count="2">'
        b'<mergeCell ref="A1:B1"/>'
        b'<mergeCell ref="A5:B5"/>'
        b"</mergeCells>"
        b"</worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels", _sheet_rels("rId1", "rId2")
        )
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Top", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Bottom", b"A5:B7"))
    return path


def make_hformula_xlsx(tmp_path):
    """XLSX with formulas referencing a table that gets shoved right."""
    path = str(tmp_path / "hformula.xlsx")
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b"<sheetData>"
        b'<row r="1">'
        b'<c r="A1" t="inlineStr"><is><t>Left</t></is></c>'
        b'<c r="D1" t="inlineStr"><is><t>R</t></is></c>'
        b'<c r="E1"><f>D1</f></c>'
        b"</row>"
        b'<row r="10"><c r="A10"><f>D1*2</f></c></row>'
        b"</sheetData>"
        b"</worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels", _sheet_rels("rId1", "rId2")
        )
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Left", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Right", b"D1:E3"))
    return path


def make_cfdv_xlsx(tmp_path):
    """XLSX with conditional formatting + data validation for move testing.

    Tables Top (A1:B3) and Bottom (A5:B7) in cols A-B. A CF region and a data
    validation cover Bottom; a second CF covers Top (which never moves).
    """
    path = str(tmp_path / "cfdv.xlsx")
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b"<sheetData>"
        b'<row r="5"><c r="A5" t="inlineStr"><is><t>x</t></is></c></row>'
        b"</sheetData>"
        b'<conditionalFormatting sqref="A5:B7">'
        b'<cfRule type="expression" priority="1"><formula>A5&gt;0</formula></cfRule>'
        b"</conditionalFormatting>"
        b'<conditionalFormatting sqref="A1:B3">'
        b'<cfRule type="expression" priority="2"><formula>A1&gt;0</formula></cfRule>'
        b"</conditionalFormatting>"
        b'<conditionalFormatting sqref="A6:B6 D1:D2">'
        b'<cfRule type="expression" priority="3"><formula>1</formula></cfRule>'
        b"</conditionalFormatting>"
        b'<dataValidations count="1">'
        b'<dataValidation type="whole" operator="greaterThan" sqref="A5:A7">'
        b"<formula1>A5</formula1>"
        b"</dataValidation>"
        b"</dataValidations>"
        b"</worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels", _sheet_rels("rId1", "rId2")
        )
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Top", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Bottom", b"A5:B7"))
    return path


def make_named_range_xlsx(tmp_path):
    """XLSX whose workbook.xml carries defined names for move testing.

    Tables Top (A1:B3) and Bottom (A5:B7) in cols A-B. Named ranges cover
    Bottom, Top, another sheet, and a single cell.
    """
    path = str(tmp_path / "named.xlsx")
    workbook_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships">'
        b'<sheets><sheet name="Sheet1" r:id="rId1"/></sheets>'
        b"<definedNames>"
        b'<definedName name="BottomData">Sheet1!$A$5:$A$7</definedName>'
        b'<definedName name="TopData">Sheet1!$A$1:$A$3</definedName>'
        b'<definedName name="OtherSheet">Sheet2!$A$5:$A$7</definedName>'
        b'<definedName name="SingleCell">Sheet1!$A$5</definedName>'
        b'<definedName name="Summed">SUM(Sheet1!$A$5:$A$7)</definedName>'
        b"</definedNames>"
        b"</workbook>"
    )
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b'<sheetData><row r="5">'
        b'<c r="A5" t="inlineStr"><is><t>x</t></is></c>'
        b"</row></sheetData>"
        b"</worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels", _sheet_rels("rId1", "rId2")
        )
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Top", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Bottom", b"A5:B7"))
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


class TestFormulaMergeRewrite:
    def _f(self, pkg, cell_ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        cell = root.find(f".//{{{NS}}}c[@r='{cell_ref}']")
        return None if cell is None else cell.find(f"{{{NS}}}f").text

    def _merges(self, pkg):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        mc = root.find(f"{{{NS}}}mergeCells")
        return {m.get("ref") for m in mc.findall(f"{{{NS}}}mergeCell")}

    def test_range_ref_to_moved_cells_shifts(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)  # Bottom shifts down 2
        assert self._f(pkg, "D1") == "SUM(A7:A9)"

    def test_ref_to_unmoved_cell_unchanged(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._f(pkg, "E1") == "A1+1"

    def test_function_name_not_mangled(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._f(pkg, "D3") == "SUM(A7:A9)+LOG10(A7)"

    def test_absolute_ref_shifts_and_keeps_markers(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._f(pkg, "D4") == "$A$7"

    def test_cross_sheet_ref_untouched(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._f(pkg, "D5") == "Sheet2!A5"

    def test_formula_inside_moved_table_follows(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._f(pkg, "B7") == "A7*2"  # B5 -> B7, A5 -> A7

    def test_merged_cell_in_moved_block_shifts(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        merges = self._merges(pkg)
        assert "A7:B7" in merges  # A5:B5 followed Bottom down
        assert "A1:B1" in merges  # Top's merge stayed put

    def test_column_move_shifts_refs(self, tmp_path):
        pkg = XLSXPackage(make_hformula_xlsx(tmp_path))
        pkg.resize_table("Left", add_cols=3)  # Right shifts right 2
        assert pkg.table_map["Right"]["range"] == ["F1", "G3"]
        assert self._f(pkg, "A10") == "F1*2"  # loose ref D1 -> F1
        assert self._f(pkg, "G1") == "F1"     # E1 formula -> G1, D1 -> F1

    def test_save_roundtrip_persists_rewrites(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "SUM(A7:A9)" in sheet
        assert "A7:B7" in sheet


class TestConditionalFormattingDataValidation:
    def _root(self, pkg):
        return pkg.trees[pkg.sheet_map["Sheet1"]].getroot()

    def _cf_map(self, pkg):
        out = {}
        for cf in self._root(pkg).findall(f"{{{NS}}}conditionalFormatting"):
            formula = cf.find(f".//{{{NS}}}formula")
            out[cf.get("sqref")] = None if formula is None else formula.text
        return out

    def _dv(self, pkg):
        return self._root(pkg).find(f".//{{{NS}}}dataValidation")

    def test_cf_region_over_moved_table_shifts(self, tmp_path):
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)  # Bottom shifts down 2
        cf = self._cf_map(pkg)
        assert "A7:B9" in cf
        assert cf["A7:B9"] == "A7>0"  # rule formula followed too

    def test_cf_region_over_static_table_untouched(self, tmp_path):
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        cf = self._cf_map(pkg)
        assert cf.get("A1:B3") == "A1>0"

    def test_data_validation_region_shifts(self, tmp_path):
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        dv = self._dv(pkg)
        assert dv.get("sqref") == "A7:A9"

    def test_data_validation_formula_shifts(self, tmp_path):
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        dv = self._dv(pkg)
        assert dv.find(f"{{{NS}}}formula1").text == "A7"

    def test_multi_range_sqref_partial_shift(self, tmp_path):
        # A CF whose sqref lists one range inside Bottom (A6:B6) and one
        # outside (D1:D2): only the contained range shifts.
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert "A8:B8 D1:D2" in self._cf_map(pkg)

    def test_save_roundtrip_persists(self, tmp_path):
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert 'sqref="A7:B9"' in sheet
        assert 'sqref="A7:A9"' in sheet


class TestDefinedNames:
    def _names(self, pkg):
        root = pkg.trees["xl/workbook.xml"].getroot()
        dn = root.find(f"{{{NS}}}definedNames")
        return {n.get("name"): n.text for n in dn.findall(f"{{{NS}}}definedName")}

    def test_named_range_over_moved_table_shifts(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)  # Bottom shifts down 2
        assert self._names(pkg)["BottomData"] == "Sheet1!$A$7:$A$9"

    def test_named_range_over_static_table_untouched(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._names(pkg)["TopData"] == "Sheet1!$A$1:$A$3"

    def test_named_range_other_sheet_untouched(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._names(pkg)["OtherSheet"] == "Sheet2!$A$5:$A$7"

    def test_single_cell_named_range_shifts(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._names(pkg)["SingleCell"] == "Sheet1!$A$7"

    def test_formula_valued_name_shifts(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        assert self._names(pkg)["Summed"] == "SUM(Sheet1!$A$7:$A$9)"

    def test_column_move_shifts_named_range(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_cols=3)  # no vertical collision; Bottom stays
        # Bottom didn't move (different rows), so BottomData is unchanged.
        assert self._names(pkg)["BottomData"] == "Sheet1!$A$5:$A$7"

    def test_save_roundtrip_persists(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            wb = zf.read("xl/workbook.xml").decode("utf-8")
        assert "Sheet1!$A$7:$A$9" in wb
        assert "Sheet2!$A$5:$A$7" in wb
