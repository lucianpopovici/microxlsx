"""
Unit tests for microxlsx.core
"""
import datetime
import struct
import zipfile
import zlib
import xml.etree.ElementTree as ET
import pytest

from microxlsx.core import XLSXPackage


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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


def make_calcpr_xlsx(tmp_path):
    """XLSX whose workbook.xml already has a calcPr (with a calcId to preserve)."""
    path = str(tmp_path / "calcpr.xlsx")
    workbook_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships">'
        b'<sheets><sheet name="Sheet1" r:id="rId1"/></sheets>'
        b'<calcPr calcId="191029"/>'
        b"</workbook>"
    )
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b'<sheetData><row r="5"><c r="B5"><f>A5*2</f><v>0</v></c></row></sheetData>'
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


def make_read_xlsx(tmp_path):
    """XLSX exercising every readable cell type + a shared-strings table."""
    path = str(tmp_path / "read.xlsx")
    shared = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        b' count="2" uniqueCount="2">'
        b"<si><t>Plain</t></si>"
        b"<si><r><t>Rich</t></r><r><t>Text</t></r></si>"
        b"</sst>"
    )
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"><sheetData><row r="1">'
        b'<c r="A1" t="s"><v>0</v></c>'          # shared string "Plain"
        b'<c r="B1" t="s"><v>1</v></c>'          # shared string (rich) "RichText"
        b'<c r="C1" t="inlineStr"><is><t>Inline</t></is></c>'
        b'<c r="D1"><v>42</v></c>'               # int
        b'<c r="E1"><v>3.5</v></c>'              # float
        b'<c r="F1" t="b"><v>1</v></c>'          # bool
        b'<c r="G1"><f>A1</f><v>7</v></c>'       # formula -> cached 7
        b'<c r="H1" t="e"><v>#DIV/0!</v></c>'    # error
        b"</row></sheetData></worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/sharedStrings.xml", shared)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", _sheet_rels("rId1"))
        zf.writestr("xl/tables/table1.xml", _table_xml(b"SalesTable", b"A1:B3"))
    return path


def make_calcchain_xlsx(tmp_path):
    """XLSX with a calcChain part registered in content-types + workbook rels."""
    path = str(tmp_path / "calc.xlsx")
    workbook_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships">'
        b'<sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>'
    )
    wb_rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/'
        b'package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/worksheet"'
        b' Target="worksheets/sheet1.xml"/>'
        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>'
        b"</Relationships>"
    )
    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/xl/workbook.xml" ContentType="wb"/>'
        b'<Override PartName="/xl/calcChain.xml"'
        b' ContentType="application/vnd.openxmlformats-officedocument.'
        b'spreadsheetml.calcChain+xml"/>'
        b"</Types>"
    )
    calc_chain = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<calcChain xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"><c r="B5" i="1"/></calcChain>'
    )
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"><sheetData>'
        b'<row r="5"><c r="B5"><f>A5*2</f><v>0</v></c></row>'
        b"</sheetData></worksheet>"
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/calcChain.xml", calc_chain)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", _sheet_rels("rId1", "rId2"))
        zf.writestr("xl/tables/table1.xml", _table_xml(b"Top", b"A1:B3"))
        zf.writestr("xl/tables/table2.xml", _table_xml(b"Bottom", b"A5:B7"))
    return path


def make_styles_xlsx(tmp_path):
    """XLSX with a minimal but valid styles part (one base cellXfs entry)."""
    path = str(tmp_path / "styles.xlsx")
    styles = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<styleSheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main">'
        b'<fonts count="1"><font><sz val="11"/></font></fonts>'
        b'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        b'<borders count="1"><border/></borders>'
        b'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0"'
        b' borderId="0"/></cellStyleXfs>'
        b'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0"'
        b' borderId="0" xfId="0"/></cellXfs>'
        b"</styleSheet>"
    )
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/'
        b'spreadsheetml/2006/main"><sheetData/></worksheet>'
    )
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/styles.xml", styles)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return path


def make_opc_xlsx(tmp_path):
    """A conformant package (content-types + rels) with one sheet and one table."""
    path = str(tmp_path / "opc.xlsx")
    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
        b'package.relationships+xml"/>'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b'<Override PartName="/xl/workbook.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        b'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        b'<Override PartName="/xl/tables/table1.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
        b"</Types>"
    )
    workbook = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships">'
        b'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    wb_rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/'
        b'package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        b"</Relationships>"
    )
    sheet = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships"><sheetData/>'
        b'<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
    )
    sheet_rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/'
        b'package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
        b"</Relationships>"
    )
    table = _table_xml(b"T1", b"A1:B2")
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet)
        zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels)
        zf.writestr("xl/tables/table1.xml", table)
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


class TestFullCalcOnLoad:
    def _calc_pr(self, pkg):
        root = pkg.trees["xl/workbook.xml"].getroot()
        return root.find(f"{{{NS}}}calcPr")

    def test_calc_pr_created_on_move(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        assert self._calc_pr(pkg) is None  # workbook has none to start
        pkg.resize_table("Top", add_rows=3)  # moves Bottom
        assert self._calc_pr(pkg).get("fullCalcOnLoad") == "1"

    def test_existing_calc_pr_preserved(self, tmp_path):
        pkg = XLSXPackage(make_calcpr_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        calc_pr = self._calc_pr(pkg)
        assert calc_pr.get("fullCalcOnLoad") == "1"
        assert calc_pr.get("calcId") == "191029"  # existing attrs kept

    def test_calc_pr_placed_after_defined_names(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        root = pkg.trees["xl/workbook.xml"].getroot()
        tags = [child.tag.split("}")[-1] for child in root]
        assert tags.index("calcPr") == tags.index("definedNames") + 1

    def test_no_move_leaves_workbook_untouched(self, tmp_path):
        # A pure shrink moves nothing, so no recalc flag is added.
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Bottom", add_rows=-1)
        assert self._calc_pr(pkg) is None

    def test_save_roundtrip_persists(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            wb = zf.read("xl/workbook.xml").decode("utf-8")
        assert 'fullCalcOnLoad="1"' in wb


class TestGetCell:
    def test_shared_string(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "A1") == "Plain"

    def test_shared_string_rich_text_concatenated(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "B1") == "RichText"

    def test_inline_string(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "C1") == "Inline"

    def test_int(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        val = pkg.get_cell("Sheet1", "D1")
        assert val == 42 and isinstance(val, int)

    def test_float(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "E1") == 3.5

    def test_bool(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "F1") is True

    def test_formula_returns_cached_value(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "G1") == 7

    def test_error_value(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "H1") == "#DIV/0!"

    def test_missing_cell_is_none(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        assert pkg.get_cell("Sheet1", "Z9") is None

    def test_reflects_pending_write(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "D1", value=99)
        assert pkg.get_cell("Sheet1", "D1") == 99

    def test_read_only_does_not_pollute_trees(self, tmp_path):
        # Reading an untouched sheet must not pull it into the re-serialized set.
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        pkg.get_cell("Sheet1", "A1")
        assert pkg.sheet_map["Sheet1"] not in pkg.trees

    def test_get_table_cell(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        # SalesTable at A1:B3; row_offset 0, col "SalesTable_c1" -> A1 -> "Plain"
        assert pkg.get_table_cell("SalesTable", 0, "SalesTable_c1") == "Plain"


class TestBooleanWrite:
    def _cell(self, pkg, ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        return root.find(f".//{{{NS}}}c[@r='{ref}']")

    def test_true_writes_boolean_type(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B2", value=True)
        cell = self._cell(pkg, "B2")
        assert cell.get("t") == "b"
        assert cell.find(f"{{{NS}}}v").text == "1"

    def test_false_writes_zero(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B2", value=False)
        assert self._cell(pkg, "B2").find(f"{{{NS}}}v").text == "0"

    def test_bool_overwrites_inline_string(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))  # A1 starts as inlineStr "Hello"
        pkg.update_cell("Sheet1", "A1", value=True)
        cell = self._cell(pkg, "A1")
        assert cell.get("t") == "b"
        assert cell.find(f"{{{NS}}}is") is None

    def test_bool_round_trips_through_read(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "C3", value=False)
        assert pkg.get_cell("Sheet1", "C3") is False


class TestCalcChainInvalidation:
    def _saved(self, pkg, tmp_path):
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        return zipfile.ZipFile(out, "r")

    def test_formula_edit_drops_calc_chain(self, tmp_path):
        pkg = XLSXPackage(make_calcchain_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B5", formula="=A5*3")
        with self._saved(pkg, tmp_path) as zf:
            assert "xl/calcChain.xml" not in zf.namelist()
            assert b"calcChain" not in zf.read("[Content_Types].xml")
            assert b"calcChain" not in zf.read("xl/_rels/workbook.xml.rels")

    def test_other_parts_preserved(self, tmp_path):
        pkg = XLSXPackage(make_calcchain_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B5", formula="=A5*3")
        with self._saved(pkg, tmp_path) as zf:
            ct = zf.read("[Content_Types].xml")
            rels = zf.read("xl/_rels/workbook.xml.rels")
        assert b"/xl/workbook.xml" in ct            # unrelated override kept
        assert b"worksheets/sheet1.xml" in rels     # worksheet rel kept

    def test_move_drops_calc_chain(self, tmp_path):
        pkg = XLSXPackage(make_calcchain_xlsx(tmp_path))
        pkg.resize_table("Top", add_rows=3)  # moves Bottom (has a formula)
        with self._saved(pkg, tmp_path) as zf:
            assert "xl/calcChain.xml" not in zf.namelist()

    def test_value_only_edit_keeps_calc_chain(self, tmp_path):
        pkg = XLSXPackage(make_calcchain_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B5", value=10)
        with self._saved(pkg, tmp_path) as zf:
            assert "xl/calcChain.xml" in zf.namelist()
            assert b"calcChain" in zf.read("[Content_Types].xml")


class TestNumberFormats:
    def _styles(self, pkg):
        return pkg.trees["xl/styles.xml"].getroot()

    def test_returns_style_id(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        assert pkg.add_number_format("#,##0.00") == 1  # base xf is 0

    def test_creates_numfmt_and_xf(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.add_number_format("#,##0.00")
        root = self._styles(pkg)
        num_fmts = root.find(f"{{{NS}}}numFmts")
        assert num_fmts.get("count") == "1"
        entry = num_fmts.find(f"{{{NS}}}numFmt")
        assert entry.get("numFmtId") == "164"
        assert entry.get("formatCode") == "#,##0.00"
        cell_xfs = root.find(f"{{{NS}}}cellXfs")
        assert cell_xfs.get("count") == "2"
        assert cell_xfs.findall(f"{{{NS}}}xf")[1].get("numFmtId") == "164"

    def test_numfmts_is_first_child(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.add_number_format("0%")
        tags = [c.tag.split("}")[-1] for c in self._styles(pkg)]
        assert tags[0] == "numFmts"

    def test_dedup_same_code(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        first = pkg.add_number_format("#,##0.00")
        second = pkg.add_number_format("#,##0.00")
        assert first == second
        assert self._styles(pkg).find(f"{{{NS}}}cellXfs").get("count") == "2"

    def test_distinct_codes_get_distinct_ids(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        a = pkg.add_number_format("#,##0.00")
        b = pkg.add_number_format("0%")
        assert a != b
        ids = [e.get("numFmtId")
               for e in self._styles(pkg).find(f"{{{NS}}}numFmts")]
        assert ids == ["164", "165"]

    def test_style_applies_to_cell(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        fmt = pkg.add_number_format("#,##0.00")
        pkg.update_cell("Sheet1", "A1", value=1234.5, style_id=fmt)
        cell = pkg.trees[pkg.sheet_map["Sheet1"]].getroot().find(f".//{{{NS}}}c[@r='A1']")
        assert cell.get("s") == str(fmt)

    def test_raises_without_styles_part(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))  # no styles.xml
        with pytest.raises(FileNotFoundError):
            pkg.add_number_format("0.00")


class TestDateValues:
    def _cell(self, pkg, ref):
        return pkg.trees[pkg.sheet_map["Sheet1"]].getroot().find(f".//{{{NS}}}c[@r='{ref}']")

    def test_date_serial(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        assert self._cell(pkg, "A1").find(f"{{{NS}}}v").text == "45306"  # serial
        # get_cell sees the date number format and converts back.
        assert pkg.get_cell("Sheet1", "A1") == datetime.date(2024, 1, 15)

    def test_datetime_serial_has_time_fraction(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.datetime(2024, 1, 15, 6))
        assert self._cell(pkg, "A1").find(f"{{{NS}}}v").text == "45306.25"
        assert pkg.get_cell("Sheet1", "A1") == datetime.datetime(2024, 1, 15, 6)

    def test_date_cell_is_numeric_not_text(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        assert self._cell(pkg, "A1").get("t") is None

    def test_date_auto_applies_date_format(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        style_id = int(self._cell(pkg, "A1").get("s"))
        xf = pkg.trees["xl/styles.xml"].getroot().find(f"{{{NS}}}cellXfs")[style_id]
        num_fmt_id = xf.get("numFmtId")
        codes = {e.get("numFmtId"): e.get("formatCode")
                 for e in pkg.trees["xl/styles.xml"].getroot().find(f"{{{NS}}}numFmts")}
        assert "yyyy-mm-dd" in codes[num_fmt_id]

    def test_explicit_style_overrides_auto_date_format(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        custom = pkg.add_number_format("dd/mm/yy")
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15), style_id=custom)
        assert self._cell(pkg, "A1").get("s") == str(custom)

    def test_date_without_styles_part_still_writes_serial(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))  # no styles.xml
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        assert pkg.get_cell("Sheet1", "A1") == 45306
        assert self._cell(pkg, "A1").get("s") is None  # unformatted


class TestRecalcOnValueEdit:
    def _calc_pr(self, pkg):
        return pkg.trees["xl/workbook.xml"].getroot().find(f"{{{NS}}}calcPr")

    def test_value_edit_sets_full_calc_on_load(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "D1", value=99)
        assert self._calc_pr(pkg).get("fullCalcOnLoad") == "1"

    def test_value_edit_keeps_calc_chain(self, tmp_path):
        # Changing an input value keeps calcChain (structure unchanged).
        pkg = XLSXPackage(make_calcchain_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B5", value=10)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            assert "xl/calcChain.xml" in zf.namelist()
            assert zf.read("xl/workbook.xml").decode().count("fullCalcOnLoad") == 1


class TestInspection:
    def test_sheet_names(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        assert pkg.sheet_names() == ["Sheet1"]

    def test_table_names(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        assert sorted(pkg.table_names()) == ["Bottom", "Far", "Side", "Top"]

    def test_table_dimensions(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        assert pkg.table_dimensions("Top") == (3, 2)   # A1:B3 incl. header
        assert pkg.table_dimensions("Side") == (7, 2)  # D1:E7


class TestClearCell:
    def _cell(self, pkg, ref):
        return pkg.trees[pkg.sheet_map["Sheet1"]].getroot().find(f".//{{{NS}}}c[@r='{ref}']")

    def test_clears_existing_cell(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))  # A1 = "Hello"
        pkg.clear_cell("Sheet1", "A1")
        assert self._cell(pkg, "A1") is None
        assert pkg.get_cell("Sheet1", "A1") is None

    def test_missing_cell_is_noop(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.clear_cell("Sheet1", "Z9")  # should not raise

    def test_row_survives(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.clear_cell("Sheet1", "A1")
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        assert root.find(f".//{{{NS}}}row[@r='1']") is not None

    def test_clear_flags_recalc(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.clear_cell("Sheet1", "A1")
        assert pkg.trees["xl/workbook.xml"].getroot().find(f"{{{NS}}}calcPr") is not None


class TestAppendTableRow:
    def test_dict_values(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        offset = pkg.append_table_row("Top", {"Top_c1": "x", "Top_c2": 5})
        assert offset == 3
        assert pkg.table_map["Top"]["range"] == ["A1", "B4"]
        assert pkg.get_cell("Sheet1", "A4") == "x"
        assert pkg.get_cell("Sheet1", "B4") == 5

    def test_positional_values(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.append_table_row("Top", ["a", "b"])
        assert pkg.get_cell("Sheet1", "A4") == "a"
        assert pkg.get_cell("Sheet1", "B4") == "b"

    def test_append_shoves_colliding_table(self, tmp_path):
        # Top A1:B3, Bottom A5:B7 (gap at row 4). Two appends push into row 5.
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.append_table_row("Top", {"Top_c1": "1"})
        pkg.append_table_row("Top", {"Top_c1": "2"})
        assert pkg.table_map["Top"]["range"] == ["A1", "B5"]
        assert pkg.table_map["Bottom"]["range"] == ["A6", "B8"]  # shoved down 1


class TestColumnRowSizing:
    def _root(self, pkg):
        return pkg.trees[pkg.sheet_map["Sheet1"]].getroot()

    def test_set_column_width_by_letter(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_column_width("Sheet1", "C", 25)
        col = self._root(pkg).find(f"{{{NS}}}cols/{{{NS}}}col")
        assert col.get("min") == "3" and col.get("max") == "3"
        assert col.get("width") == "25" and col.get("customWidth") == "1"

    def test_set_column_width_by_index(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_column_width("Sheet1", 0, 12)  # 0-based -> column A
        col = self._root(pkg).find(f"{{{NS}}}cols/{{{NS}}}col")
        assert col.get("min") == "1"

    def test_cols_precedes_sheetdata(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_column_width("Sheet1", "A", 10)
        root = self._root(pkg)
        children = list(root)
        assert children.index(root.find(f"{{{NS}}}cols")) < \
            children.index(root.find(f"{{{NS}}}sheetData"))

    def test_column_width_updates_existing(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_column_width("Sheet1", "A", 10)
        pkg.set_column_width("Sheet1", "A", 20)
        cols = self._root(pkg).findall(f"{{{NS}}}cols/{{{NS}}}col")
        assert len(cols) == 1 and cols[0].get("width") == "20"

    def test_set_row_height(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_row_height("Sheet1", 2, 30)
        row = self._root(pkg).find(f".//{{{NS}}}row[@r='2']")
        assert row.get("ht") == "30" and row.get("customHeight") == "1"

    def test_set_row_height_existing_row(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))  # row 1 already exists
        pkg.set_row_height("Sheet1", 1, 42)
        rows = self._root(pkg).findall(f".//{{{NS}}}row[@r='1']")
        assert len(rows) == 1 and rows[0].get("ht") == "42"


NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _reopen(tmp_path, pkg, name="out.xlsx"):
    out = str(tmp_path / name)
    pkg.save(out)
    return out, XLSXPackage(out)


def _read(out, part):
    with zipfile.ZipFile(out, "r") as zf:
        return zf.read(part)


def _names(out):
    with zipfile.ZipFile(out, "r") as zf:
        return set(zf.namelist())


class TestAddSheet:
    def test_appears_in_sheet_names(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        _, reopened = _reopen(tmp_path, pkg)
        assert reopened.sheet_names() == ["Sheet1", "Data"]

    def test_new_part_and_content_type(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        out, _ = _reopen(tmp_path, pkg)
        assert "xl/worksheets/sheet2.xml" in _names(out)
        assert b"/xl/worksheets/sheet2.xml" in _read(out, "[Content_Types].xml")

    def test_workbook_relationship_added(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        out, _ = _reopen(tmp_path, pkg)
        rels = _read(out, "xl/_rels/workbook.xml.rels")
        assert b"worksheets/sheet2.xml" in rels

    def test_duplicate_name_raises(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.add_sheet("Sheet1")

    def test_can_write_to_new_sheet(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        pkg.update_cell("Data", "A1", value="hello")
        _, reopened = _reopen(tmp_path, pkg)
        assert reopened.get_cell("Data", "A1") == "hello"


class TestRemoveSheet:
    def test_removed_from_names(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        _, reopened = _reopen(tmp_path, pkg)
        reopened.remove_sheet("Data")
        _, again = _reopen(tmp_path, reopened, "again.xlsx")
        assert again.sheet_names() == ["Sheet1"]

    def test_parts_and_content_type_dropped(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        _, mid = _reopen(tmp_path, pkg)
        mid.remove_sheet("Data")
        out, _ = _reopen(tmp_path, mid, "final.xlsx")
        assert "xl/worksheets/sheet2.xml" not in _names(out)
        assert b"/xl/worksheets/sheet2.xml" not in _read(out, "[Content_Types].xml")

    def test_last_sheet_guard(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.remove_sheet("Sheet1")

    def test_removing_sheet_drops_its_tables(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        pkg.add_table("Data", "T2", "A1:B3", ["a", "b"])
        out, _ = _reopen(tmp_path, pkg)
        pkg2 = XLSXPackage(out)
        pkg2.remove_sheet("Data")
        out2, reopened = _reopen(tmp_path, pkg2, "f.xlsx")
        assert "T2" not in reopened.table_names()
        assert "xl/tables/table2.xml" not in _names(out2)


class TestAddTable:
    def test_appears_in_table_names(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        pkg.add_table("Data", "T2", "A1:C4", ["x", "y", "z"])
        _, reopened = _reopen(tmp_path, pkg)
        assert "T2" in reopened.table_names()
        assert reopened.table_dimensions("T2") == (4, 3)

    def test_part_rels_and_content_type(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_sheet("Data")
        pkg.add_table("Data", "T2", "A1:C4", ["x", "y", "z"])
        out, _ = _reopen(tmp_path, pkg)
        assert "xl/tables/table2.xml" in _names(out)
        assert "xl/worksheets/_rels/sheet2.xml.rels" in _names(out)
        assert b"/xl/tables/table2.xml" in _read(out, "[Content_Types].xml")

    def test_unique_table_id(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "T2", "D1:E2", ["p", "q"])
        out, _ = _reopen(tmp_path, pkg)
        ids = set()
        with zipfile.ZipFile(out) as zf:
            for n in zf.namelist():
                if n.startswith("xl/tables/"):
                    ids.add(ET.fromstring(zf.read(n)).get("id"))
        assert len(ids) == 2  # distinct ids for table1 and table2

    def test_can_write_via_new_table(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "T2", "D1:E2", ["p", "q"])
        pkg.update_table_cell("T2", 1, "p", 7)
        _, reopened = _reopen(tmp_path, pkg)
        assert reopened.get_table_cell("T2", 1, "p") == 7

    def test_duplicate_name_raises(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.add_table("Sheet1", "T1", "D1:E2", ["p", "q"])


class TestRemoveTable:
    def test_removed_from_names(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.remove_table("T1")
        _, reopened = _reopen(tmp_path, pkg)
        assert not reopened.table_names()

    def test_part_content_type_and_tablepart_dropped(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.remove_table("T1")
        out, _ = _reopen(tmp_path, pkg)
        assert "xl/tables/table1.xml" not in _names(out)
        assert b"/xl/tables/table1.xml" not in _read(out, "[Content_Types].xml")
        assert b"tablePart" not in _read(out, "xl/worksheets/sheet1.xml")

    def test_relationship_dropped(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.remove_table("T1")
        out, _ = _reopen(tmp_path, pkg)
        assert b"table1.xml" not in _read(out, "xl/worksheets/_rels/sheet1.xml.rels")


def make_date1904_xlsx(tmp_path):
    """The styles fixture with workbookPr date1904 switched on."""
    src = make_styles_xlsx(tmp_path)
    path = str(tmp_path / "d1904.xlsx")
    workbook = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
        b'2006/relationships"><workbookPr date1904="1"/>'
        b'<sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>'
    )
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(path, "w") as zout:
        for name in zin.namelist():
            zout.writestr(name, workbook if name == "xl/workbook.xml" else zin.read(name))
    return path


class TestWriteRange:
    def test_block_written_and_read_back(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.write_range("Sheet1", "B2", [["a", 1], ["b", 2.5]])
        assert pkg.get_range("Sheet1", "B2:C3") == [["a", 1], ["b", 2.5]]

    def test_none_skips_cell(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "C2", value="keep")
        pkg.write_range("Sheet1", "B2", [["new", None]])
        assert pkg.get_cell("Sheet1", "C2") == "keep"

    def test_merges_with_existing_cells_sorted(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "D1", value="pre")
        pkg.write_range("Sheet1", "A1", [[1, 2]])
        row = pkg.trees[pkg.sheet_map["Sheet1"]].getroot().find(f".//{{{NS}}}row[@r='1']")
        refs = [c.get("r") for c in row]
        assert refs == ["A1", "B1", "D1"]  # sorted, D1 preserved

    def test_style_id_applied_to_block(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        fmt = pkg.add_number_format("0.00")
        pkg.write_range("Sheet1", "A1", [[1, 2]], style_id=fmt)
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        assert root.find(f".//{{{NS}}}c[@r='B1']").get("s") == str(fmt)

    def test_dates_in_block_auto_formatted(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.write_range("Sheet1", "A1", [[datetime.date(2024, 1, 15)]])
        assert pkg.get_cell("Sheet1", "A1") == datetime.date(2024, 1, 15)

    def test_flags_recalc(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.write_range("Sheet1", "A1", [[1]])
        assert pkg.trees["xl/workbook.xml"].getroot().find(f"{{{NS}}}calcPr") is not None

    def test_save_roundtrip(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.write_range("Sheet1", "B2", [["x", 1], ["y", 2]])
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        assert XLSXPackage(out).get_range("Sheet1", "B2:C3") == [["x", 1], ["y", 2]]


class TestGetRange:
    def test_missing_cells_are_none(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        block = pkg.get_range("Sheet1", "C1:D2")
        assert block == [["Inline", 42], [None, None]]

    def test_read_only_does_not_pollute_trees(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        pkg.get_range("Sheet1", "A1:B1")
        assert pkg.sheet_map["Sheet1"] not in pkg.trees


class TestIterTableRows:
    def test_yields_dicts_in_column_order(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.append_table_row("Top", {"Top_c1": "a", "Top_c2": 1})
        rows = list(pkg.iter_table_rows("Top"))
        assert rows[-1] == {"Top_c1": "a", "Top_c2": 1}
        assert len(rows) == 3  # A1:B4 after append -> 3 data rows

    def test_header_only_table_yields_nothing(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "H", "D1:E1", ["p", "q"])
        assert not list(pkg.iter_table_rows("H"))


class TestAddStyle:
    def _styles(self, pkg):
        return pkg.trees["xl/styles.xml"].getroot()

    def test_returns_distinct_ids_and_dedups(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        a = pkg.add_style(bold=True)
        b = pkg.add_style(bold=True)
        c = pkg.add_style(italic=True)
        assert a == b and a != c

    def test_font_attributes(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        sid = pkg.add_style(bold=True, font_size=14, font_name="Arial",
                            font_color="#FF0000")
        xf = self._styles(pkg).find(f"{{{NS}}}cellXfs")[sid]
        font = self._styles(pkg).find(f"{{{NS}}}fonts")[int(xf.get("fontId"))]
        assert font.find(f"{{{NS}}}b") is not None
        assert font.find(f"{{{NS}}}sz").get("val") == "14"
        assert font.find(f"{{{NS}}}name").get("val") == "Arial"
        assert font.find(f"{{{NS}}}color").get("rgb") == "FFFF0000"

    def test_fill_and_border(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        sid = pkg.add_style(fill_color="DDEBF7", border="thin")
        xf = self._styles(pkg).find(f"{{{NS}}}cellXfs")[sid]
        fill = self._styles(pkg).find(f"{{{NS}}}fills")[int(xf.get("fillId"))]
        assert fill.find(f".//{{{NS}}}fgColor").get("rgb") == "FFDDEBF7"
        border = self._styles(pkg).find(f"{{{NS}}}borders")[int(xf.get("borderId"))]
        assert border.find(f"{{{NS}}}left").get("style") == "thin"

    def test_alignment_and_number_format(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        sid = pkg.add_style(number_format="0%", align="center", valign="top",
                            wrap=True)
        xf = self._styles(pkg).find(f"{{{NS}}}cellXfs")[sid]
        assert xf.get("applyNumberFormat") == "1"
        alignment = xf.find(f"{{{NS}}}alignment")
        assert alignment.get("horizontal") == "center"
        assert alignment.get("vertical") == "top"
        assert alignment.get("wrapText") == "1"

    def test_counts_stay_in_sync(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.add_style(bold=True, fill_color="EEEEEE", border="thin")
        root = self._styles(pkg)
        for tag in ("fonts", "fills", "borders", "cellXfs"):
            el = root.find(f"{{{NS}}}{tag}")
            assert el.get("count") == str(len(el))

    def test_raises_without_styles_part(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        with pytest.raises(FileNotFoundError):
            pkg.add_style(bold=True)


class TestDateRoundTrip:
    def test_reopened_file_reads_date(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        assert XLSXPackage(out).get_cell("Sheet1", "A1") == datetime.date(2024, 1, 15)

    def test_builtin_date_format_detected(self, tmp_path):
        # A style using builtin numFmtId 14 (m/d/yyyy) with no custom numFmts.
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        # Registering any format loads styles.xml; then append a builtin-id xf.
        pkg.add_number_format("0.00")
        cell_xfs = pkg.trees["xl/styles.xml"].getroot().find(f"{{{NS}}}cellXfs")
        xf = ET.SubElement(cell_xfs, f"{{{NS}}}xf")
        for attr, val in (("numFmtId", "14"), ("fontId", "0"), ("fillId", "0"),
                          ("borderId", "0"), ("xfId", "0"),
                          ("applyNumberFormat", "1")):
            xf.set(attr, val)
        cell_xfs.set("count", str(len(cell_xfs)))
        sid = len(cell_xfs) - 1
        pkg.update_cell("Sheet1", "A1", value=45306, style_id=sid)
        assert pkg.get_cell("Sheet1", "A1") == datetime.date(2024, 1, 15)

    def test_plain_number_not_converted(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        fmt = pkg.add_number_format("#,##0.00")
        pkg.update_cell("Sheet1", "A1", value=45306, style_id=fmt)
        assert pkg.get_cell("Sheet1", "A1") == 45306  # money format, not a date

    def test_1904_serial_written(self, tmp_path):
        pkg = XLSXPackage(make_date1904_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        v = root.find(f".//{{{NS}}}c[@r='A1']/{{{NS}}}v")
        assert v.text == "43844"  # days since 1904-01-01, not 45306

    def test_1904_round_trip(self, tmp_path):
        pkg = XLSXPackage(make_date1904_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=datetime.date(2024, 1, 15))
        assert pkg.get_cell("Sheet1", "A1") == datetime.date(2024, 1, 15)


class TestStyleGetters:
    def test_get_cell_style_returns_id(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        sid = pkg.add_style(bold=True)
        pkg.update_cell("Sheet1", "A1", value=1, style_id=sid)
        assert pkg.get_cell_style("Sheet1", "A1") == sid

    def test_get_cell_style_unstyled_and_missing(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "A1", value=1)
        assert pkg.get_cell_style("Sheet1", "A1") is None
        assert pkg.get_cell_style("Sheet1", "Z9") is None

    def test_get_cell_style_read_only_does_not_pollute(self, tmp_path):
        pkg = XLSXPackage(make_read_xlsx(tmp_path))
        pkg.get_cell_style("Sheet1", "A1")
        assert pkg.sheet_map["Sheet1"] not in pkg.trees

    def test_get_style_decodes_composed_style(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        sid = pkg.add_style(bold=True, font_size=12, font_name="Arial",
                            font_color="#FF0000", fill_color="DDEBF7",
                            border="thin", align="center", valign="top",
                            wrap=True, number_format="$#,##0.00")
        decoded = pkg.get_style(sid)
        assert decoded == {
            "number_format": "$#,##0.00", "bold": True, "italic": False,
            "font_size": 12.0, "font_name": "Arial", "font_color": "FF0000",
            "fill_color": "DDEBF7", "border": "thin", "align": "center",
            "valign": "top", "wrap": True,
        }

    def test_get_style_round_trips_through_add_style(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        sid = pkg.add_style(italic=True, fill_color="EEEEEE", align="right")
        clone = pkg.add_style(**pkg.get_style(sid))
        assert pkg.get_style(clone) == pkg.get_style(sid)

    def test_get_style_plain_base_style(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        decoded = pkg.get_style(0)  # fixture's base xf
        assert decoded["number_format"] is None
        assert decoded["bold"] is False
        assert decoded["fill_color"] is None

    def test_get_style_builtin_numfmt_decoded(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.add_number_format("0.00")  # loads styles into trees
        cell_xfs = pkg.trees["xl/styles.xml"].getroot().find(f"{{{NS}}}cellXfs")
        xf = ET.SubElement(cell_xfs, f"{{{NS}}}xf")
        for attr, val in (("numFmtId", "9"), ("fontId", "0"), ("fillId", "0"),
                          ("borderId", "0"), ("xfId", "0")):
            xf.set(attr, val)
        cell_xfs.set("count", str(len(cell_xfs)))
        assert pkg.get_style(len(cell_xfs) - 1)["number_format"] == "0%"

    def test_get_style_raises_without_styles_part(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        with pytest.raises(FileNotFoundError):
            pkg.get_style(0)


class TestInsertRows:
    def _f(self, pkg, ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        cell = root.find(f".//{{{NS}}}c[@r='{ref}']")
        return None if cell is None else cell.find(f"{{{NS}}}f").text

    def test_cells_and_table_below_shift(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 4, 2)
        assert pkg.table_map["Bottom"]["range"] == ["A7", "B9"]
        assert pkg.get_cell("Sheet1", "A7") == "x"  # was A5

    def test_formulas_rewritten(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 4, 2)
        assert self._f(pkg, "D1") == "SUM(A7:A9)"
        assert self._f(pkg, "E1") == "A1+1"          # above insert point
        assert self._f(pkg, "D4") == "$A$7"          # $ markers kept
        assert self._f(pkg, "D5") == "Sheet2!A5"     # other sheet untouched

    def test_straddling_range_grows(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "F1", formula="SUM(A1:A7)")
        pkg.insert_rows("Sheet1", 4, 2)
        assert self._f(pkg, "F1") == "SUM(A1:A9)"

    def test_straddled_table_grows(self, tmp_path):
        # Insert inside Bottom's data rows (A5:B7): row 6 is between 5 and 7.
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 6, 1)
        assert pkg.table_map["Bottom"]["range"] == ["A5", "B8"]

    def test_merges_shift_and_straddling_merge_grows(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "D1:D5")
        pkg.insert_rows("Sheet1", 3, 2)
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        refs = {m.get("ref") for m in root.find(f"{{{NS}}}mergeCells")}
        assert "A7:B7" in refs   # A5:B5 shifted
        assert "D1:D7" in refs   # straddling merge grew

    def test_row_height_rides_along(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.set_row_height("Sheet1", 5, 33)
        pkg.insert_rows("Sheet1", 2, 3)
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        assert root.find(f".//{{{NS}}}row[@r='8']").get("ht") == "33"

    def test_named_ranges_shift(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 4, 2)
        names = {n.get("name"): n.text
                 for n in pkg.trees["xl/workbook.xml"].getroot()
                 .find(f"{{{NS}}}definedNames")}
        assert names["BottomData"] == "Sheet1!$A$7:$A$9"
        assert names["OtherSheet"] == "Sheet2!$A$5:$A$7"

    def test_cfdv_regions_shift(self, tmp_path):
        pkg = XLSXPackage(make_cfdv_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 4, 2)
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        sqrefs = {cf.get("sqref")
                  for cf in root.findall(f"{{{NS}}}conditionalFormatting")}
        assert "A7:B9" in sqrefs
        dv = root.find(f".//{{{NS}}}dataValidation")
        assert dv.get("sqref") == "A7:A9"
        assert dv.find(f"{{{NS}}}formula1").text == "A7"

    def test_save_roundtrip(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 4, 2)
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        reopened = XLSXPackage(out)
        assert reopened.table_map["Bottom"]["range"] == ["A7", "B9"]
        assert reopened.get_cell("Sheet1", "A7") == "x"


class TestDeleteRows:
    def _f(self, pkg, ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        cell = root.find(f".//{{{NS}}}c[@r='{ref}']")
        return None if cell is None else cell.find(f"{{{NS}}}f").text

    def test_undoes_insert(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.insert_rows("Sheet1", 4, 2)
        pkg.delete_rows("Sheet1", 4, 2)
        assert pkg.table_map["Bottom"]["range"] == ["A5", "B7"]
        assert self._f(pkg, "D1") == "SUM(A5:A7)"
        assert self._f(pkg, "B5") == "A5*2"

    def test_range_clamped_and_table_shrunk(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.delete_rows("Sheet1", 6, 2)  # Bottom data rows
        assert pkg.table_map["Bottom"]["range"] == ["A5", "B5"]
        assert self._f(pkg, "D1") == "SUM(A5:A5)"

    def test_single_ref_becomes_ref_error(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "F1", formula="A4*3")
        pkg.delete_rows("Sheet1", 4, 1)
        assert self._f(pkg, "F1") == "#REF!*3"

    def test_header_guard(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.delete_rows("Sheet1", 5, 3)  # covers Bottom's header

    def test_merge_inside_band_removed_others_kept(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "D6:E7")
        pkg.delete_rows("Sheet1", 6, 2)  # band swallows D6:E7 entirely
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        refs = {m.get("ref") for m in root.find(f"{{{NS}}}mergeCells")}
        assert refs == {"A1:B1", "A5:B5"}  # swallowed merge gone, others kept

    def test_table_below_band_shifts_up(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.delete_rows("Sheet1", 4, 1)  # gap row between Top and Bottom
        assert pkg.table_map["Bottom"]["range"] == ["A4", "B6"]
        assert pkg.table_map["Far"]["range"] == ["A8", "B10"]
        assert pkg.get_cell("Sheet1", "A4") == "B"


class TestInsertDeleteCols:
    def _f(self, pkg, ref):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        cell = root.find(f".//{{{NS}}}c[@r='{ref}']")
        return None if cell is None else cell.find(f"{{{NS}}}f").text

    def test_insert_shifts_table_and_formulas(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "G1", formula="D1*2")
        pkg.insert_cols("Sheet1", "C", 2)
        assert pkg.table_map["Right"]["range"] == ["F1", "G3"]
        assert pkg.table_map["Left"]["range"] == ["A1", "B3"]
        assert self._f(pkg, "I1") == "F1*2"
        assert pkg.get_cell("Sheet1", "F1") == "R"

    def test_insert_shifts_width_entries(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.set_column_width("Sheet1", "D", 22)
        pkg.insert_cols("Sheet1", "C", 2)
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        col = root.find(f"{{{NS}}}cols/{{{NS}}}col")
        assert (col.get("min"), col.get("max")) == ("6", "6")  # D -> F

    def test_delete_undoes_insert(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "G1", formula="D1*2")
        pkg.insert_cols("Sheet1", "C", 2)
        pkg.delete_cols("Sheet1", "C", 2)
        assert pkg.table_map["Right"]["range"] == ["D1", "E3"]
        assert self._f(pkg, "G1") == "D1*2"

    def test_deleted_ref_becomes_ref_error(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "C1", formula="A1+1")
        pkg.delete_cols("Sheet1", "A", 1)
        assert self._f(pkg, "B1") == "#REF!+1"

    def test_mid_table_insert_guard(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.insert_cols("Sheet1", "B", 1)  # inside Left A1:B3

    def test_table_intersect_delete_guard(self, tmp_path):
        pkg = XLSXPackage(make_hcollision_xlsx(tmp_path))
        with pytest.raises(ValueError):
            pkg.delete_cols("Sheet1", "D", 1)  # Right's first column

    def test_column_index_accepts_int(self, tmp_path):
        pkg = XLSXPackage(make_styles_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "B1", value=7)
        pkg.insert_cols("Sheet1", 0, 1)  # 0-based -> before column A
        assert pkg.get_cell("Sheet1", "C1") == 7


class TestFreezePanes:
    def _pane(self, pkg):
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        return root.find(f".//{{{NS}}}sheetViews/{{{NS}}}sheetView/{{{NS}}}pane")

    def test_freeze_both(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.freeze_panes("Sheet1", "B2")
        pane = self._pane(pkg)
        assert pane.get("xSplit") == "1" and pane.get("ySplit") == "1"
        assert pane.get("topLeftCell") == "B2" and pane.get("state") == "frozen"
        assert pane.get("activePane") == "bottomRight"

    def test_freeze_rows_only(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.freeze_panes("Sheet1", "A2")
        pane = self._pane(pkg)
        assert pane.get("xSplit") is None and pane.get("ySplit") == "1"
        assert pane.get("activePane") == "bottomLeft"

    def test_sheetviews_precedes_sheetdata(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_column_width("Sheet1", "A", 10)  # inserts <cols>
        pkg.freeze_panes("Sheet1", "B2")
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        tags = [c.tag.split("}")[-1] for c in root]
        assert tags.index("sheetViews") < tags.index("cols") < tags.index("sheetData")

    def test_freeze_at_a1_is_noop(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.freeze_panes("Sheet1", "A1")
        assert self._pane(pkg) is None


class TestRenameSheet:
    def test_workbook_and_map_updated(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.rename_sheet("Sheet1", "Report")
        assert pkg.sheet_names() == ["Report"]
        sheets = pkg.trees["xl/workbook.xml"].getroot().find(f"{{{NS}}}sheets")
        assert sheets[0].get("name") == "Report"

    def test_table_sheet_refs_updated(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.rename_sheet("Sheet1", "Report")
        assert pkg.table_map["Top"]["sheet"] == "Report"

    def test_formula_qualifiers_rewritten(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "Z1", formula="=Sheet1!A1+Sheet2!A1")
        pkg.rename_sheet("Sheet1", "Data")
        root = pkg.trees[pkg.sheet_map["Data"]].getroot()
        cell = root.find(f".//{{{NS}}}c[@r='Z1']")
        assert cell.find(f"{{{NS}}}f").text == "Data!A1+Sheet2!A1"

    def test_defined_name_qualifier_rewritten(self, tmp_path):
        pkg = XLSXPackage(make_named_range_xlsx(tmp_path))
        pkg.rename_sheet("Sheet1", "Data")
        names = {n.get("name"): n.text
                 for n in pkg.trees["xl/workbook.xml"].getroot()
                 .find(f"{{{NS}}}definedNames")}
        assert names["BottomData"] == "Data!$A$5:$A$7"
        assert names["OtherSheet"] == "Sheet2!$A$5:$A$7"

    def test_quotes_name_with_space(self, tmp_path):
        pkg = XLSXPackage(make_formula_xlsx(tmp_path))
        pkg.update_cell("Sheet1", "Z1", formula="=Sheet1!A1")
        pkg.rename_sheet("Sheet1", "My Data")
        root = pkg.trees[pkg.sheet_map["My Data"]].getroot()
        assert root.find(f".//{{{NS}}}c[@r='Z1']/{{{NS}}}f").text == "'My Data'!A1"

    def test_missing_and_duplicate_guards(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        with pytest.raises(KeyError):
            pkg.rename_sheet("Nope", "X")


class TestAddDefinedName:
    def _names(self, pkg):
        dn = pkg.trees["xl/workbook.xml"].getroot().find(f"{{{NS}}}definedNames")
        return {n.get("name"): n for n in dn.findall(f"{{{NS}}}definedName")}

    def test_global_name_added(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.add_defined_name("TaxRate", "Sheet1!$B$2")
        entry = self._names(pkg)["TaxRate"]
        assert entry.text == "Sheet1!$B$2"
        assert entry.get("localSheetId") is None

    def test_scoped_name_gets_local_sheet_id(self, tmp_path):
        pkg = XLSXPackage(make_multi_table_xlsx(tmp_path))
        pkg.add_defined_name("Local", "Sheet1!$A$1", sheet_name="Sheet1")
        assert self._names(pkg)["Local"].get("localSheetId") == "0"

    def test_definednames_after_sheets(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.add_defined_name("N", "Sheet1!$A$1")
        root = pkg.trees["xl/workbook.xml"].getroot()
        tags = [c.tag.split("}")[-1] for c in root]
        assert tags.index("definedNames") == tags.index("sheets") + 1


class TestAddTableStyle:
    def test_default_style_applied(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "T", "D1:E2", ["x", "y"])
        root = pkg.trees["xl/tables/table2.xml"].getroot()
        style = root.find(f"{{{NS}}}tableStyleInfo")
        assert style.get("name") == "TableStyleMedium2"
        assert style.get("showRowStripes") == "1"

    def test_style_none_omits(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "T", "D1:E2", ["x", "y"], style_name=None)
        root = pkg.trees["xl/tables/table2.xml"].getroot()
        assert root.find(f"{{{NS}}}tableStyleInfo") is None


class TestAutoFilter:
    def test_worksheet_auto_filter(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_auto_filter("Sheet1", "A1:C1")
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        assert root.find(f"{{{NS}}}autoFilter").get("ref") == "A1:C1"

    def test_autofilter_ordered_before_mergecells(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.merge_cells("Sheet1", "A1:B1")
        pkg.set_auto_filter("Sheet1", "A1:C1")
        tags = [c.tag.split("}")[-1]
                for c in pkg.trees[pkg.sheet_map["Sheet1"]].getroot()]
        assert tags.index("autoFilter") < tags.index("mergeCells")

    def test_new_table_has_autofilter(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "T", "D1:E3", ["x", "y"])
        root = pkg.trees["xl/tables/table2.xml"].getroot()
        assert root.find(f"{{{NS}}}autoFilter").get("ref") == "D1:E3"

    def test_table_autofilter_follows_resize(self, tmp_path):
        # A table created via add_table has an autoFilter; resizing keeps it synced.
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_table("Sheet1", "T", "D1:E3", ["x", "y"])
        pkg.resize_table("T", add_rows=2)
        root = pkg.trees[pkg.table_map["T"]["xml_path"]].getroot()
        assert root.get("ref") == "D1:E5"
        assert root.find(f"{{{NS}}}autoFilter").get("ref") == "D1:E5"


class TestPageSetup:
    def _root(self, pkg):
        return pkg.trees[pkg.sheet_map["Sheet1"]].getroot()

    def test_orientation(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_page_setup("Sheet1", orientation="landscape")
        assert self._root(pkg).find(f"{{{NS}}}pageSetup").get("orientation") == "landscape"

    def test_fit_to_page(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_page_setup("Sheet1", fit_to_width=1, fit_to_height=0)
        root = self._root(pkg)
        assert root.find(f"{{{NS}}}pageSetup").get("fitToWidth") == "1"
        page_pr = root.find(f"{{{NS}}}sheetPr/{{{NS}}}pageSetUpPr")
        assert page_pr.get("fitToPage") == "1"

    def test_print_area_defined_name(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_print_area("Sheet1", "$A$1:$C$10")
        names = {n.get("name"): n
                 for n in pkg.trees["xl/workbook.xml"].getroot()
                 .find(f"{{{NS}}}definedNames")}
        entry = names["_xlnm.Print_Area"]
        assert entry.get("localSheetId") == "0"
        assert entry.text == "Sheet1!$A$1:$C$10"

    def test_print_area_replaces_existing(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_print_area("Sheet1", "A1:B2")
        pkg.set_print_area("Sheet1", "A1:C3")
        entries = [n for n in pkg.trees["xl/workbook.xml"].getroot()
                   .find(f"{{{NS}}}definedNames")
                   if n.get("name") == "_xlnm.Print_Area"]
        assert len(entries) == 1 and entries[0].text.endswith("A1:C3")

    def test_header_footer(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.set_header_footer("Sheet1", header="&CTitle", footer="&CPage &P")
        hf = self._root(pkg).find(f"{{{NS}}}headerFooter")
        assert hf.find(f"{{{NS}}}oddHeader").text == "&CTitle"
        assert hf.find(f"{{{NS}}}oddFooter").text == "&CPage &P"


class TestProtectSheet:
    def test_protection_enabled_with_password(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.protect_sheet("Sheet1", password="secret")
        prot = pkg.trees[pkg.sheet_map["Sheet1"]].getroot().find(f"{{{NS}}}sheetProtection")
        assert prot.get("sheet") == "1"
        assert prot.get("password") == "DAA7"  # Excel legacy hash of "secret"

    def test_protection_without_password(self, tmp_path):
        pkg = XLSXPackage(make_xlsx(tmp_path))
        pkg.protect_sheet("Sheet1")
        prot = pkg.trees[pkg.sheet_map["Sheet1"]].getroot().find(f"{{{NS}}}sheetProtection")
        assert prot.get("sheet") == "1" and prot.get("password") is None


def _tiny_png():
    width, height = 2, 3
    raw = b''.join(b'\x00' + b'\xff\x00\x00' * width for _ in range(height))

    def chunk(tag, data):
        return (struct.pack('>I', len(data)) + tag + data
                + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff))
    ihdr = struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr)
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


class TestHyperlink:
    def test_link_and_relationship(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_hyperlink("Sheet1", "A1", "https://example.com", tooltip="Go")
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        link = root.find(f"{{{NS}}}hyperlinks/{{{NS}}}hyperlink")
        assert link.get("ref") == "A1" and link.get("tooltip") == "Go"
        rid = link.get(f"{{{NS_R}}}id")
        rels = pkg.trees["xl/worksheets/_rels/sheet1.xml.rels"].getroot()
        rel = next(r for r in rels if r.get("Id") == rid)
        assert rel.get("Target") == "https://example.com"
        assert rel.get("TargetMode") == "External"


class TestImage:
    def test_parts_created(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_image("Sheet1", "B2", _tiny_png())
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        with zipfile.ZipFile(out, "r") as zf:
            names = set(zf.namelist())
            assert "xl/media/image1.png" in names
            assert "xl/drawings/drawing1.xml" in names
            assert b"image/png" in zf.read("[Content_Types].xml")
            assert zf.read("xl/media/image1.png").startswith(b"\x89PNG")

    def test_anchor_uses_cell_position(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_image("Sheet1", "B2", _tiny_png())
        xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        root = pkg.trees["xl/drawings/drawing1.xml"].getroot()
        frm = root.find(f".//{{{xdr}}}from")
        assert frm.find(f"{{{xdr}}}col").text == "1"  # B
        assert frm.find(f"{{{xdr}}}row").text == "1"  # 2

    def test_png_size_detected(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        assert pkg._png_size(_tiny_png()) == (2, 3)  # pylint: disable=protected-access


class TestComment:
    def test_comment_parts_and_content(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_comment("Sheet1", "A1", "Note text", author="Bob")
        comments = pkg.trees["xl/comments1.xml"].getroot()
        assert comments.find(f".//{{{NS}}}author").text == "Bob"
        comment = comments.find(f".//{{{NS}}}comment")
        assert comment.get("ref") == "A1"
        assert comment.find(f".//{{{NS}}}t").text == "Note text"

    def test_vml_and_legacy_drawing_wired(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_comment("Sheet1", "A1", "x", author="A")
        root = pkg.trees[pkg.sheet_map["Sheet1"]].getroot()
        assert root.find(f"{{{NS}}}legacyDrawing") is not None
        assert "xl/drawings/vmlDrawing1.vml" in pkg._added_bytes  # pylint: disable=protected-access

    def test_two_authors_deduped(self, tmp_path):
        pkg = XLSXPackage(make_opc_xlsx(tmp_path))
        pkg.add_comment("Sheet1", "A1", "one", author="A")
        pkg.add_comment("Sheet1", "B2", "two", author="A")
        authors = pkg.trees["xl/comments1.xml"].getroot().find(f"{{{NS}}}authors")
        assert len(authors) == 1  # same author reused
