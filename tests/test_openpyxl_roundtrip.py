"""Validate microxlsx output against openpyxl, an independent reader.

Baselines are authored *by openpyxl* (so they are known-valid, complete
packages with absolute relationship targets, styles, theme, docProps, etc.),
modified with microxlsx, then reopened with openpyxl to confirm the result is
still a valid workbook and the change took effect.
"""
import datetime

import openpyxl
from openpyxl.worksheet.table import Table

from microxlsx.core import XLSXPackage


def _baseline(path, *, with_table=True, with_second_sheet=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row, (name, amount) in enumerate(
            [("Name", "Amount"), ("ada", 10), ("bob", 20)], start=1):
        ws.cell(row, 1, name)
        ws.cell(row, 2, amount)
    ws["D1"] = "=SUM(B2:B3)"
    if with_table:
        ws.add_table(Table(displayName="Sales", ref="A1:B3"))
    if with_second_sheet:
        wb.create_sheet("Data")["A1"] = 99
    wb.save(path)
    return path


def _apply(tmp_path, operation, **baseline_kwargs):
    """Build a baseline, apply ``operation``, save, and reopen with openpyxl."""
    src = _baseline(str(tmp_path / "in.xlsx"), **baseline_kwargs)
    pkg = XLSXPackage(src)
    operation(pkg)
    out = str(tmp_path / "out.xlsx")
    pkg.save(out)
    workbook = openpyxl.load_workbook(out)
    # Force a full parse of every worksheet and its tables.
    for worksheet in workbook.worksheets:
        list(worksheet.iter_rows())
        dict(worksheet.tables)
    return workbook


class TestOpensRealWorkbooks:
    def test_absolute_rel_targets_resolve(self, tmp_path):
        # openpyxl writes Target="/xl/worksheets/sheetN.xml"; the maps must
        # resolve them rather than producing "xl//xl/...".
        pkg = XLSXPackage(_baseline(str(tmp_path / "b.xlsx")))
        assert pkg.sheet_names() == ["Sheet1", "Data"]
        assert pkg.sheet_map["Sheet1"] == "xl/worksheets/sheet1.xml"
        assert pkg.table_names() == ["Sales"]

    def test_read_values_from_openpyxl_file(self, tmp_path):
        pkg = XLSXPackage(_baseline(str(tmp_path / "b.xlsx")))
        assert pkg.get_cell("Sheet1", "A2") == "ada"
        assert pkg.get_cell("Sheet1", "B3") == 20


class TestValueRoundTrips:
    def test_update_cell(self, tmp_path):
        wb = _apply(tmp_path, lambda p: (
            p.update_cell("Sheet1", "B2", value=42),
            p.update_cell("Sheet1", "C1", formula="=B2*2")))
        assert wb["Sheet1"]["B2"].value == 42
        assert wb["Sheet1"]["C1"].value == "=B2*2"

    def test_bool(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.update_cell("Sheet1", "C2", value=True))
        assert wb["Sheet1"]["C2"].value is True

    def test_date(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.update_cell(
            "Sheet1", "E1", value=datetime.date(2024, 1, 15)))
        assert wb["Sheet1"]["E1"].value == datetime.datetime(2024, 1, 15)
        assert wb["Sheet1"]["E1"].is_date

    def test_write_range(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.write_range(
            "Sheet1", "F1", [["x", 1], ["y", 2]]))
        assert [c.value for c in wb["Sheet1"]["F1:G2"][0]] == ["x", 1]
        assert [c.value for c in wb["Sheet1"]["F1:G2"][1]] == ["y", 2]


class TestStyleRoundTrips:
    def test_composed_style(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.update_cell(
            "Sheet1", "A1", value="H",
            style_id=p.add_style(bold=True, fill_color="DDEBF7", border="thin")))
        cell = wb["Sheet1"]["A1"]
        assert cell.font.bold is True
        assert cell.fill.fgColor.rgb == "FFDDEBF7"
        assert cell.border.left.style == "thin"

    def test_number_format(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.update_cell(
            "Sheet1", "B2", value=1234.5,
            style_id=p.add_number_format("$#,##0.00")))
        assert wb["Sheet1"]["B2"].number_format == "$#,##0.00"


class TestTableRoundTrips:
    def test_resize_table(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.resize_table("Sales", add_rows=2))
        assert dict(wb["Sheet1"].tables)["Sales"].ref == "A1:B5"

    def test_append_table_row(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.append_table_row(
            "Sales", {"Name": "cara", "Amount": 30}))
        assert dict(wb["Sheet1"].tables)["Sales"].ref == "A1:B4"
        assert wb["Sheet1"]["A4"].value == "cara"

    def test_add_table(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.add_table(
            "Data", "Extra", "A1:B3", ["x", "y"]))
        assert "Extra" in dict(wb["Data"].tables)

    def test_remove_table(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.remove_table("Sales"))
        assert not dict(wb["Sheet1"].tables)


class TestStructuralRoundTrips:
    def test_add_sheet(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.add_sheet("Extra"))
        assert "Extra" in wb.sheetnames

    def test_add_sheet_then_write(self, tmp_path):
        def op(pkg):
            pkg.add_sheet("Extra")
            pkg.update_cell("Extra", "A1", value="hi")
        wb = _apply(tmp_path, op)
        assert wb["Extra"]["A1"].value == "hi"

    def test_remove_sheet(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.remove_sheet("Data"))
        assert wb.sheetnames == ["Sheet1"]

    def test_remove_sheet_with_table(self, tmp_path):
        def op(pkg):
            pkg.add_table("Data", "Extra", "A1:B2", ["x", "y"])
            pkg.remove_sheet("Data")
        wb = _apply(tmp_path, op)
        assert wb.sheetnames == ["Sheet1"]


class TestInsertDeleteRoundTrips:
    def test_insert_rows(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.insert_rows("Sheet1", 2, 2))
        assert wb["Sheet1"]["A4"].value == "ada"  # was A2
        assert dict(wb["Sheet1"].tables)["Sales"].ref == "A1:B5"

    def test_delete_rows(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.delete_rows("Sheet1", 3, 1))
        assert wb["Sheet1"]["A2"].value == "ada"

    def test_insert_cols(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.insert_cols("Sheet1", "C", 1))
        assert wb["Sheet1"]["A1"].value == "Name"

    def test_delete_cols(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.delete_cols("Sheet1", "F", 1))
        assert wb["Sheet1"]["B2"].value == 10


class TestLayoutRoundTrips:
    def test_column_width_and_row_height(self, tmp_path):
        def op(pkg):
            pkg.set_column_width("Sheet1", "B", 25)
            pkg.set_row_height("Sheet1", 1, 30)
        wb = _apply(tmp_path, op)
        assert wb["Sheet1"].column_dimensions["B"].width == 25
        assert wb["Sheet1"].row_dimensions[1].height == 30

    def test_merge_cells(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.merge_cells("Sheet1", "F1:H1"))
        assert "F1:H1" in [str(m) for m in wb["Sheet1"].merged_cells.ranges]


class TestChainedOperations:
    def test_full_report_build(self, tmp_path):
        def op(pkg):
            header = pkg.add_style(bold=True, fill_color="DDEBF7", align="center")
            pkg.add_sheet("Report")
            pkg.write_range("Report", "A1", [["Region", "Q1", "Q2"]], style_id=header)
            pkg.write_range("Report", "A2", [
                ["West", 100, datetime.date(2026, 1, 31)],
                ["East", 200, datetime.date(2026, 2, 28)],
            ])
            pkg.add_table("Report", "Totals", "A1:C3", ["Region", "Q1", "Q2"])
            pkg.set_column_width("Report", "A", 18)
        wb = _apply(tmp_path, op)
        report = wb["Report"]
        assert report["A1"].value == "Region"
        assert report["A1"].font.bold is True
        assert report["C2"].value == datetime.datetime(2026, 1, 31)
        assert "Totals" in dict(report.tables)


class TestFinishingTouchesRoundTrips:
    def test_freeze_panes(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.freeze_panes("Sheet1", "B2"))
        assert wb["Sheet1"].freeze_panes == "B2"

    def test_rename_sheet_updates_cross_sheet_formula(self, tmp_path):
        # Data!A1 references Sheet1; renaming Sheet1 must fix that qualifier.
        def build(path):
            wb = openpyxl.Workbook()
            wb.active.title = "Sheet1"
            wb.active["B2"] = 10
            wb.create_sheet("Data")["A1"] = "=Sheet1!B2*2"
            wb.save(path)
            return path
        src = build(str(tmp_path / "in.xlsx"))
        pkg = XLSXPackage(src)
        pkg.rename_sheet("Sheet1", "Sales")
        out = str(tmp_path / "out.xlsx")
        pkg.save(out)
        wb = openpyxl.load_workbook(out)
        assert wb.sheetnames == ["Sales", "Data"]
        assert wb["Data"]["A1"].value == "=Sales!B2*2"

    def test_add_defined_name(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.add_defined_name(
            "TaxRate", "Sheet1!$B$2"))
        assert "TaxRate" in wb.defined_names
        assert wb.defined_names["TaxRate"].value == "Sheet1!$B$2"

    def test_add_scoped_defined_name(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.add_defined_name(
            "Local", "Sheet1!$A$1", sheet_name="Sheet1"))
        assert "Local" in wb["Sheet1"].defined_names

    def test_table_style_reads_back(self, tmp_path):
        wb = _apply(tmp_path, lambda p: p.add_table(
            "Data", "Extra", "A1:B3", ["x", "y"]))
        table = dict(wb["Data"].tables)["Extra"]
        assert table.tableStyleInfo.name == "TableStyleMedium2"
        assert table.tableStyleInfo.showRowStripes is True
