"""
Core functionality for MicroXLSX.
"""
# pylint: disable=too-many-lines
import re
import zipfile
import io
import datetime
import xml.etree.ElementTree as ET
from .utils import cell_to_indices, indices_to_cell

# Excel's day-zero (accounts for the fictional 1900 leap day for later dates).
_EXCEL_EPOCH = datetime.datetime(1899, 12, 30)
# Mac-origin workbooks (workbookPr date1904="1") count days from 1904 instead.
_EXCEL_EPOCH_1904 = datetime.datetime(1904, 1, 1)

# Built-in number-format ids that render as dates/times (ECMA-376 §18.8.30).
_BUILTIN_DATE_FMTS = frozenset(
    list(range(14, 23)) + list(range(27, 37)) + list(range(45, 48))
    + list(range(50, 59))
)

# Format codes for the common built-in number-format ids (ECMA-376 §18.8.30).
_BUILTIN_FMT_CODES = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yyyy h:mm', 37: '#,##0;(#,##0)', 38: '#,##0;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
}

# A1-style reference: optional 'Sheet'! qualifier, optional $ absolute markers.
_CELL_REF_RE = re.compile(
    r"(?P<sheet>(?:'[^']+'|[A-Za-z0-9_.]+)!)?"
    r"(?P<c_abs>\$?)(?P<col>[A-Za-z]{1,3})(?P<r_abs>\$?)(?P<row>[0-9]+)"
)

# A cell or range carrying a single leading sheet qualifier, used for defined
# names where a cross-sheet range must be treated as one unit.
_ENDPOINT = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
_RANGE_RE = re.compile(
    r"(?P<sheet>(?:'[^']+'|[A-Za-z0-9_.]+)!)?"
    r"(?P<a>" + _ENDPOINT + r")(?::(?P<b>" + _ENDPOINT + r"))?"
)

class XLSXPackage:
    """
    Represents an Excel (XLSX) package.
    """
    # pylint: disable=too-many-instance-attributes,too-many-public-methods
    NS = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/'
    }
    # Custom number-format ids must start at 164 (0-163 are reserved built-ins).
    CUSTOM_FMT_BASE = 164
    REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    CT_WORKSHEET = ('application/vnd.openxmlformats-officedocument.'
                    'spreadsheetml.worksheet+xml')
    CT_TABLE = ('application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.table+xml')
    REL_WORKSHEET = ('http://schemas.openxmlformats.org/officeDocument/'
                     '2006/relationships/worksheet')
    REL_TABLE = ('http://schemas.openxmlformats.org/officeDocument/'
                 '2006/relationships/table')

    def __init__(self, filename):
        self.filename = filename
        self.trees = {}
        self.sheet_map = {}
        self.table_map = {}
        self._shared_strings = None
        self._drop_calc_chain = False
        self._added_formats = {}
        self._added_styles = {}
        self._date_style_cache = None
        self._source_parts = set()
        self._dropped_parts = set()
        self._ct_add = {}
        self._ct_remove = set()
        for prefix, uri in self.NS.items():
            ET.register_namespace(prefix if prefix != 'main' else '', uri)
        self._build_maps()

    def _build_maps(self):
        """Builds relationship maps for Sheets and Tables."""
        with zipfile.ZipFile(self.filename, 'r') as zin:
            self._source_parts = set(zin.namelist())
            self._map_sheets(zin)
            self._map_tables(zin)

    def _get_tree(self, zin, path):
        if path not in self.trees:
            with zin.open(path) as f:
                self.trees[path] = ET.parse(f)
        return self.trees[path]

    def _map_sheets(self, zin):
        """Map Sheets to paths."""
        wb_tree = self._get_tree(zin, 'xl/workbook.xml')
        sheets = wb_tree.getroot().find(f"{{{self.NS['main']}}}sheets")
        id_to_name = {s.get(f"{{{self.NS['r']}}}id"): s.get('name') for s in sheets}
        rel_tree = self._get_tree(zin, 'xl/_rels/workbook.xml.rels')
        for rel in rel_tree.getroot():
            rid, target = rel.get('Id'), rel.get('Target')
            if rid in id_to_name:
                path = f"xl/{target}" if not target.startswith('xl/') else target
                self.sheet_map[id_to_name[rid]] = path

    def _map_tables(self, zin):
        """Map Tables to metadata."""
        for sheet_name, sheet_path in self.sheet_map.items():
            rel_path = f"xl/worksheets/_rels/{sheet_path.rsplit('/', maxsplit=1)[-1]}.rels"
            try:
                with zin.open(rel_path) as f:
                    t_rel_tree = ET.parse(f)
                    for rel in t_rel_tree.getroot():
                        if "table" in rel.get('Type'):
                            self._parse_table_rel(zin, sheet_name, rel)
            except (KeyError, FileNotFoundError):
                continue

    def _parse_table_rel(self, zin, sheet_name, rel):
        """Helper to parse table relationship."""
        t_path = f"xl/tables/{rel.get('Target').rsplit('/', maxsplit=1)[-1]}"
        t_tree = self._get_tree(zin, t_path)
        t_root = t_tree.getroot()
        t_name, t_ref = t_root.get('displayName'), t_root.get('ref')
        cols = t_root.find(f"{{{self.NS['main']}}}tableColumns")
        col_map = {c.get('name'): i for i, c in enumerate(cols)}
        start_cell, end_cell = t_ref.split(':')
        self.table_map[t_name] = {
            'xml_path': t_path, 'sheet': sheet_name,
            'range': [start_cell, end_cell],
            'start_indices': cell_to_indices(start_cell),
            'columns': col_map
        }

    def sheet_names(self):
        """Return the workbook's sheet names, in relationship order."""
        return list(self.sheet_map)

    def table_names(self):
        """Return the names of all tables across the workbook."""
        return list(self.table_map)

    def table_dimensions(self, table_name):
        """Return ``(rows, cols)`` of a table, counting its header row."""
        top, left = self.table_map[table_name]['start_indices']
        bottom, right = cell_to_indices(self.table_map[table_name]['range'][1])
        return (bottom - top + 1, right - left + 1)

    def _sheet_root(self, sheet_name):
        """Return a worksheet's root element, loading it for editing if needed."""
        path = self.sheet_map.get(sheet_name, sheet_name)
        if path not in self.trees:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                self._get_tree(zin, path)
        return self.trees[path].getroot()

    def clear_cell(self, sheet_name, cell_ref):
        """Remove a cell entirely (leaves the row); flags a recalc."""
        ns = self.NS['main']
        sheet_data = self._sheet_root(sheet_name).find(f".//{{{ns}}}sheetData")
        row = sheet_data.find(f"{{{ns}}}row[@r='{cell_to_indices(cell_ref)[0] + 1}']")
        if row is None:
            return
        cell = row.find(f"{{{ns}}}c[@r='{cell_ref}']")
        if cell is not None:
            row.remove(cell)
            self._set_full_calc_on_load()

    def append_table_row(self, table_name, values):
        """Append a data row to a table, expanding its range.

        ``values`` may be a ``{column_name: value}`` mapping or a positional
        list/tuple in column order. Returns the new row's 0-based offset.
        """
        table = self.table_map[table_name]
        start_row = table['start_indices'][0]
        end_row = cell_to_indices(table['range'][1])[0]
        offset = (end_row - start_row) + 1
        # Grow via resize_table so any table directly below is shoved aside
        # (minimally, cascading) instead of being overwritten.
        self.resize_table(table_name, add_rows=1)
        if isinstance(values, dict):
            items = list(values.items())
        else:
            names = {idx: name for name, idx in table['columns'].items()}
            items = [(names[i], val) for i, val in enumerate(values)]
        for col_name, val in items:
            self.update_table_cell(table_name, offset, col_name, val)
        return offset

    def set_column_width(self, sheet_name, column, width):
        """Set a column's width. ``column`` is a letter (``"A"``) or 0-based int."""
        ns = self.NS['main']
        root = self._sheet_root(sheet_name)
        idx = column if isinstance(column, int) else cell_to_indices(f"{column}1")[1]
        num = str(idx + 1)  # <col> uses 1-based min/max
        cols = root.find(f"{{{ns}}}cols")
        if cols is None:
            cols = ET.Element(f"{{{ns}}}cols")
            sheet_data = root.find(f"{{{ns}}}sheetData")
            root.insert(list(root).index(sheet_data), cols)  # cols precedes sheetData
        col = next(
            (c for c in cols.findall(f"{{{ns}}}col")
             if c.get('min') == num and c.get('max') == num),
            None,
        )
        if col is None:
            col = ET.SubElement(cols, f"{{{ns}}}col")
            col.set('min', num)
            col.set('max', num)
        col.set('width', str(width))
        col.set('customWidth', '1')

    def set_row_height(self, sheet_name, row, height):
        """Set a row's height. ``row`` is the 1-based row number."""
        ns = self.NS['main']
        sheet_data = self._sheet_root(sheet_name).find(f"{{{ns}}}sheetData")
        row_el = self._row_get_or_create(sheet_data, int(row))
        row_el.set('ht', str(height))
        row_el.set('customHeight', '1')

    def add_sheet(self, name):
        """Create a new empty worksheet and return its name."""
        ns = self.NS['main']
        if name in self.sheet_map:
            raise ValueError(f"sheet '{name}' already exists")
        part = f"xl/worksheets/sheet{self._next_part_number('xl/worksheets/sheet')}.xml"
        self.trees[part] = ET.ElementTree(
            ET.fromstring(f'<worksheet xmlns="{ns}"><sheetData/></worksheet>'))
        rid = self._add_workbook_rel(
            self.REL_WORKSHEET, f"worksheets/{part.rsplit('/', maxsplit=1)[-1]}")
        sheets = self.trees['xl/workbook.xml'].getroot().find(f"{{{ns}}}sheets")
        sheet_ids = [int(s.get('sheetId')) for s in sheets if s.get('sheetId')]
        sheet_el = ET.SubElement(sheets, f"{{{ns}}}sheet")
        sheet_el.set('name', name)
        sheet_el.set('sheetId', str(max(sheet_ids, default=0) + 1))
        sheet_el.set(f"{{{self.NS['r']}}}id", rid)
        self._ct_add[f"/{part}"] = self.CT_WORKSHEET
        self.sheet_map[name] = part
        return name

    def remove_sheet(self, name):
        """Remove a worksheet, its relationship, and any tables it holds."""
        ns = self.NS['main']
        if name not in self.sheet_map:
            raise KeyError(name)
        if len(self.sheet_map) == 1:
            raise ValueError("cannot remove the only sheet")
        path = self.sheet_map[name]
        sheets = self.trees['xl/workbook.xml'].getroot().find(f"{{{ns}}}sheets")
        sheet_el = next(s for s in sheets if s.get('name') == name)
        rid = sheet_el.get(f"{{{self.NS['r']}}}id")
        sheets.remove(sheet_el)
        self._remove_workbook_rel(rid)
        self._drop_part(path)
        self._ct_remove.add(f"/{path}")
        self._drop_part(self._sheet_rels_path(path))
        for table in [t for t, m in self.table_map.items() if m['sheet'] == name]:
            self._drop_table_part(table)
        del self.sheet_map[name]

    # pylint: disable=too-many-locals
    def add_table(self, sheet_name, name, ref, columns):
        """Create a table over ``ref`` on a sheet with the given column names."""
        ns = self.NS['main']
        if name in self.table_map:
            raise ValueError(f"table '{name}' already exists")
        part = f"xl/tables/table{self._next_part_number('xl/tables/table')}.xml"
        cols = ''.join(f'<tableColumn id="{i + 1}" name="{c}"/>'
                       for i, c in enumerate(columns))
        self.trees[part] = ET.ElementTree(ET.fromstring(
            f'<table xmlns="{ns}" id="{self._next_table_id()}" name="{name}"'
            f' displayName="{name}" ref="{ref}">'
            f'<tableColumns count="{len(columns)}">{cols}</tableColumns></table>'))
        rels = self._get_or_create_rels(self._sheet_rels_path(self.sheet_map[sheet_name]))
        rid = self._next_rid(rels.getroot())
        rel = ET.SubElement(rels.getroot(), f"{{{self.REL_NS}}}Relationship")
        rel.set('Id', rid)
        rel.set('Type', self.REL_TABLE)
        rel.set('Target', f"../tables/{part.rsplit('/', maxsplit=1)[-1]}")
        sheet_root = self._sheet_root(sheet_name)
        table_parts = sheet_root.find(f"{{{ns}}}tableParts")
        if table_parts is None:
            table_parts = ET.SubElement(sheet_root, f"{{{ns}}}tableParts")
        ET.SubElement(table_parts, f"{{{ns}}}tablePart").set(f"{{{self.NS['r']}}}id", rid)
        table_parts.set('count', str(len(table_parts)))
        self._ct_add[f"/{part}"] = self.CT_TABLE
        start_cell, end_cell = ref.split(':')
        self.table_map[name] = {
            'xml_path': part, 'sheet': sheet_name,
            'range': [start_cell, end_cell],
            'start_indices': cell_to_indices(start_cell),
            'columns': {c: i for i, c in enumerate(columns)},
        }
        return name

    def remove_table(self, name):
        """Remove a table part and its worksheet/relationship references."""
        ns = self.NS['main']
        meta = self.table_map[name]
        table_file = meta['xml_path'].rsplit('/', maxsplit=1)[-1]
        rels = self._get_or_create_rels(self._sheet_rels_path(self.sheet_map[meta['sheet']]))
        rid = None
        for rel in rels.getroot().findall(f"{{{self.REL_NS}}}Relationship"):
            if rel.get('Target', '').endswith(table_file):
                rid = rel.get('Id')
                rels.getroot().remove(rel)
                break
        sheet_root = self._sheet_root(meta['sheet'])
        table_parts = sheet_root.find(f"{{{ns}}}tableParts")
        if table_parts is not None:
            for part in table_parts.findall(f"{{{ns}}}tablePart"):
                if part.get(f"{{{self.NS['r']}}}id") == rid:
                    table_parts.remove(part)
            table_parts.set('count', str(len(table_parts)))
            if len(table_parts) == 0:
                sheet_root.remove(table_parts)
        self._drop_table_part(name)

    def _drop_table_part(self, name):
        """Drop a table's part + content type and forget it."""
        meta = self.table_map.pop(name)
        self._drop_part(meta['xml_path'])
        self._ct_remove.add(f"/{meta['xml_path']}")

    def _drop_part(self, path):
        """Mark a part for removal on save and evict any cached tree."""
        self._dropped_parts.add(path)
        self.trees.pop(path, None)

    @staticmethod
    def _sheet_rels_path(sheet_path):
        """Return the .rels path for a worksheet part."""
        return f"xl/worksheets/_rels/{sheet_path.rsplit('/', maxsplit=1)[-1]}.rels"

    def _next_part_number(self, prefix):
        """Return the next free integer for ``{prefix}N.xml`` part names."""
        pattern = re.compile(re.escape(prefix) + r'(\d+)\.xml$')
        candidates = (self._source_parts | set(self.trees)) - self._dropped_parts
        nums = [int(m.group(1)) for m in map(pattern.search, candidates) if m]
        return max(nums, default=0) + 1

    def _next_table_id(self):
        """Return a workbook-unique table id."""
        ids = [int(tree.getroot().get('id'))
               for path, tree in self.trees.items()
               if path.startswith('xl/tables/') and tree.getroot().get('id')]
        return max(ids, default=0) + 1

    def _next_rid(self, rels_root):
        """Return the next free ``rIdN`` for a relationships tree."""
        nums = [int(r.get('Id')[3:]) for r in rels_root.findall(f"{{{self.REL_NS}}}Relationship")
                if (r.get('Id') or '').startswith('rId')]
        return f"rId{max(nums, default=0) + 1}"

    def _add_workbook_rel(self, rel_type, target):
        """Append a relationship to the workbook rels tree; return its id."""
        root = self.trees['xl/_rels/workbook.xml.rels'].getroot()
        rid = self._next_rid(root)
        rel = ET.SubElement(root, f"{{{self.REL_NS}}}Relationship")
        rel.set('Id', rid)
        rel.set('Type', rel_type)
        rel.set('Target', target)
        return rid

    def _remove_workbook_rel(self, rid):
        """Remove a workbook relationship by id."""
        root = self.trees['xl/_rels/workbook.xml.rels'].getroot()
        for rel in root.findall(f"{{{self.REL_NS}}}Relationship"):
            if rel.get('Id') == rid:
                root.remove(rel)

    def _get_or_create_rels(self, rels_path):
        """Return a worksheet rels tree, creating an empty one if absent."""
        if rels_path in self.trees:
            return self.trees[rels_path]
        try:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                return self._get_tree(zin, rels_path)
        except KeyError:
            tree = ET.ElementTree(ET.fromstring(
                f'<Relationships xmlns="{self.REL_NS}"/>'))
            self.trees[rels_path] = tree
            return tree

    def merge_cells(self, sheet_name, cell_range):
        """Adds a merge rule to the worksheet."""
        path = self.sheet_map.get(sheet_name, sheet_name)
        with zipfile.ZipFile(self.filename, 'r') as zin:
            tree = self._get_tree(zin, path)
            root = tree.getroot()
            merge_cells = root.find(f"{{{self.NS['main']}}}mergeCells")
            if merge_cells is None:
                merge_cells = ET.SubElement(root, f"{{{self.NS['main']}}}mergeCells")
            ET.SubElement(merge_cells, f"{{{self.NS['main']}}}mergeCell", ref=cell_range)
            merge_cells.set('count', str(len(merge_cells)))

    # pylint: disable=too-many-arguments
    def update_cell(self, sheet_name, cell_ref, *, value=None, formula=None, style_id=None):
        """Updates or creates a cell with values/formulas."""
        if style_id is None and isinstance(value, datetime.date):
            try:
                style_id = self._default_date_style(value)
            except FileNotFoundError:
                pass  # no styles part: still write the serial, just unformatted
        path = self.sheet_map.get(sheet_name, sheet_name)
        with zipfile.ZipFile(self.filename, 'r') as zin:
            tree = self._get_tree(zin, path)
            root = tree.getroot()
            sheet_data = root.find(f".//{{{self.NS['main']}}}sheetData")
            row_idx, _ = cell_to_indices(cell_ref)
            row_num = str(row_idx + 1)
            row = sheet_data.find(f"{{{self.NS['main']}}}row[@r='{row_num}']") or \
                  ET.SubElement(sheet_data, f"{{{self.NS['main']}}}row", r=row_num)
            cell = row.find(f"{{{self.NS['main']}}}c[@r='{cell_ref}']") or \
                   ET.SubElement(row, f"{{{self.NS['main']}}}c", r=cell_ref)
            if style_id is not None:
                cell.set('s', str(style_id))
            if formula:
                self._set_cell_formula(cell, formula)
            if value is not None:
                self._set_cell_value(cell, value)

    def _set_cell_formula(self, cell, formula):
        """Helper to set cell formula."""
        f_node = cell.find(f"{{{self.NS['main']}}}f")
        if f_node is None:
            f_node = ET.SubElement(cell, f"{{{self.NS['main']}}}f")
        f_node.text = formula.lstrip('=')
        self._invalidate_calc_cache()

    def _set_cell_value(self, cell, value):
        """Helper to set cell value."""
        if isinstance(value, bool):
            cell.set('t', 'b')
            self._remove_inline_string(cell)
            self._cell_v(cell).text = '1' if value else '0'
        elif isinstance(value, datetime.date):
            cell.attrib.pop('t', None)
            self._remove_inline_string(cell)
            self._cell_v(cell).text = str(self._to_excel_serial(value))
        elif isinstance(value, (int, float)):
            self._cell_v(cell).text = str(value)
            cell.attrib.pop('t', None)
        else:
            cell.set('t', 'inlineStr')
            is_node = cell.find(f"{{{self.NS['main']}}}is")
            if is_node is None:
                is_node = ET.SubElement(cell, f"{{{self.NS['main']}}}is")
            t_node = is_node.find(f"{{{self.NS['main']}}}t")
            if t_node is None:
                t_node = ET.SubElement(is_node, f"{{{self.NS['main']}}}t")
            t_node.text = str(value)
        # A changed input value leaves dependent formula caches stale.
        self._set_full_calc_on_load()

    def _cell_v(self, cell):
        """Return the cell's ``<v>`` child, creating it if absent."""
        v_node = cell.find(f"{{{self.NS['main']}}}v")
        if v_node is None:
            v_node = ET.SubElement(cell, f"{{{self.NS['main']}}}v")
        return v_node

    def _remove_inline_string(self, cell):
        """Drop an ``<is>`` payload left over from a prior string write."""
        is_node = cell.find(f"{{{self.NS['main']}}}is")
        if is_node is not None:
            cell.remove(is_node)

    def _excel_epoch(self):
        """Return the workbook's day-zero, honouring the 1904 date system."""
        ns = self.NS['main']
        props = self.trees['xl/workbook.xml'].getroot().find(f"{{{ns}}}workbookPr")
        if props is not None and props.get('date1904') in ('1', 'true'):
            return _EXCEL_EPOCH_1904
        return _EXCEL_EPOCH

    def _to_excel_serial(self, value):
        """Convert a date/datetime to an Excel serial number."""
        epoch = self._excel_epoch()
        if isinstance(value, datetime.datetime):
            delta = value - epoch
            return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400
        return (value - epoch.date()).days

    def _from_excel_serial(self, serial):
        """Convert a serial back to ``date`` (whole days) or ``datetime``."""
        epoch = self._excel_epoch()
        if serial == int(serial):
            return epoch.date() + datetime.timedelta(days=int(serial))
        # Round to whole seconds so float noise doesn't yield 05:59:59.999999.
        return epoch + datetime.timedelta(seconds=round(serial * 86400))

    def add_number_format(self, format_code):
        """Register a custom number format, returning a ``style_id`` for it.

        The id indexes ``cellXfs`` and can be passed as ``style_id`` to
        ``update_cell`` / ``update_table_cell``. Repeated calls with the same
        code reuse the same style. Requires an existing ``xl/styles.xml``.
        """
        if format_code in self._added_formats:
            return self._added_formats[format_code]
        ns = self.NS['main']
        root = self._styles_tree().getroot()
        num_fmts = root.find(f"{{{ns}}}numFmts")
        if num_fmts is None:
            num_fmts = ET.Element(f"{{{ns}}}numFmts")
            root.insert(0, num_fmts)  # numFmts is the first child of styleSheet
        fmt_id = self._numfmt_id_for(num_fmts, format_code)
        cell_xfs = root.find(f"{{{ns}}}cellXfs")
        if cell_xfs is None:
            raise ValueError("xl/styles.xml has no cellXfs; cannot register a style")
        xf_node = ET.SubElement(cell_xfs, f"{{{ns}}}xf")
        xf_node.set('numFmtId', str(fmt_id))
        for attr in ('fontId', 'fillId', 'borderId', 'xfId'):
            xf_node.set(attr, '0')
        xf_node.set('applyNumberFormat', '1')
        cell_xfs.set('count', str(len(cell_xfs)))
        style_id = len(cell_xfs) - 1
        self._added_formats[format_code] = style_id
        self._date_style_cache = None
        return style_id

    # pylint: disable=too-many-locals
    def add_style(self, *, number_format=None, bold=False, italic=False,
                  font_size=None, font_name=None, font_color=None,
                  fill_color=None, border=None, align=None, valign=None,
                  wrap=False):
        """Register a composed cell style and return its ``style_id``.

        Colors are hex ``"RRGGBB"`` (with or without ``#``). ``border`` is a
        border style name (e.g. ``"thin"``) applied to all four edges.
        ``align``/``valign`` are OOXML alignment keywords. Identical calls
        reuse the same style. Requires an existing ``xl/styles.xml``.
        """
        key = (number_format, bold, italic, font_size, font_name, font_color,
               fill_color, border, align, valign, wrap)
        if key in self._added_styles:
            return self._added_styles[key]
        ns = self.NS['main']
        root = self._styles_tree().getroot()
        cell_xfs = root.find(f"{{{ns}}}cellXfs")
        if cell_xfs is None:
            raise ValueError("xl/styles.xml has no cellXfs; cannot register a style")
        xf_node = ET.SubElement(cell_xfs, f"{{{ns}}}xf")
        for attr in ('numFmtId', 'fontId', 'fillId', 'borderId', 'xfId'):
            xf_node.set(attr, '0')
        if number_format is not None:
            num_fmts = root.find(f"{{{ns}}}numFmts")
            if num_fmts is None:
                num_fmts = ET.Element(f"{{{ns}}}numFmts")
                root.insert(0, num_fmts)
            xf_node.set('numFmtId', str(self._numfmt_id_for(num_fmts, number_format)))
            xf_node.set('applyNumberFormat', '1')
        if bold or italic or font_size or font_name or font_color:
            xf_node.set('fontId', str(self._add_font(
                bold=bold, italic=italic, size=font_size,
                name=font_name, color=font_color)))
            xf_node.set('applyFont', '1')
        if fill_color is not None:
            xf_node.set('fillId', str(self._add_fill(fill_color)))
            xf_node.set('applyFill', '1')
        if border is not None:
            xf_node.set('borderId', str(self._add_border(border)))
            xf_node.set('applyBorder', '1')
        if align or valign or wrap:
            alignment = ET.SubElement(xf_node, f"{{{ns}}}alignment")
            if align:
                alignment.set('horizontal', align)
            if valign:
                alignment.set('vertical', valign)
            if wrap:
                alignment.set('wrapText', '1')
            xf_node.set('applyAlignment', '1')
        cell_xfs.set('count', str(len(cell_xfs)))
        style_id = len(cell_xfs) - 1
        self._added_styles[key] = style_id
        self._date_style_cache = None
        return style_id

    def _add_font(self, *, bold, italic, size, name, color):
        """Append a ``<font>`` to the styles part; return its index."""
        ns = self.NS['main']
        fonts = self.trees['xl/styles.xml'].getroot().find(f"{{{ns}}}fonts")
        font = ET.SubElement(fonts, f"{{{ns}}}font")
        if bold:
            ET.SubElement(font, f"{{{ns}}}b")
        if italic:
            ET.SubElement(font, f"{{{ns}}}i")
        if size is not None:
            ET.SubElement(font, f"{{{ns}}}sz").set('val', str(size))
        if color is not None:
            ET.SubElement(font, f"{{{ns}}}color").set('rgb', self._argb(color))
        if name is not None:
            ET.SubElement(font, f"{{{ns}}}name").set('val', name)
        fonts.set('count', str(len(fonts)))
        return len(fonts) - 1

    def _add_fill(self, color):
        """Append a solid ``<fill>`` to the styles part; return its index."""
        ns = self.NS['main']
        fills = self.trees['xl/styles.xml'].getroot().find(f"{{{ns}}}fills")
        fill = ET.SubElement(fills, f"{{{ns}}}fill")
        pattern = ET.SubElement(fill, f"{{{ns}}}patternFill")
        pattern.set('patternType', 'solid')
        ET.SubElement(pattern, f"{{{ns}}}fgColor").set('rgb', self._argb(color))
        fills.set('count', str(len(fills)))
        return len(fills) - 1

    def _add_border(self, style):
        """Append a uniform ``<border>`` to the styles part; return its index."""
        ns = self.NS['main']
        borders = self.trees['xl/styles.xml'].getroot().find(f"{{{ns}}}borders")
        border = ET.SubElement(borders, f"{{{ns}}}border")
        for side in ('left', 'right', 'top', 'bottom'):
            ET.SubElement(border, f"{{{ns}}}{side}").set('style', style)
        ET.SubElement(border, f"{{{ns}}}diagonal")
        borders.set('count', str(len(borders)))
        return len(borders) - 1

    @staticmethod
    def _argb(color):
        """Normalize ``"RRGGBB"``/``"#RRGGBB"``/8-digit ARGB to ARGB hex."""
        color = color.lstrip('#').upper()
        return color if len(color) == 8 else f"FF{color}"

    def get_cell_style(self, sheet_name, cell_ref):
        """Return a cell's ``style_id`` (reusable as-is), or None if unstyled.

        Unlike ``get_style``, this never decodes anything, so reuse via
        ``update_cell(style_id=...)`` is always faithful to the original.
        """
        ns = self.NS['main']
        path = self.sheet_map.get(sheet_name, sheet_name)
        if path in self.trees:
            root = self.trees[path].getroot()
        else:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                with zin.open(path) as handle:
                    root = ET.parse(handle).getroot()
        cell = root.find(f".//{{{ns}}}c[@r='{cell_ref}']")
        if cell is None or cell.get('s') is None:
            return None
        return int(cell.get('s'))

    def get_style(self, style_id):
        """Decode a style into ``add_style``-compatible kwargs (best effort).

        Returns a dict with keys ``number_format``, ``bold``, ``italic``,
        ``font_size``, ``font_name``, ``font_color``, ``fill_color``,
        ``border``, ``align``, ``valign``, ``wrap``. Theme colors decode to
        ``{'theme': n, 'tint': t}`` dicts and a mixed border to a per-side
        dict — neither is re-feedable to ``add_style``; unknown built-in
        number formats decode to None. For faithful reuse of any style, pass
        the raw id from ``get_cell_style`` instead.
        """
        root = self._styles_root_readonly()
        if root is None:
            raise FileNotFoundError("xl/styles.xml not found")
        ns = self.NS['main']
        xf_node = root.find(f"{{{ns}}}cellXfs")[style_id]
        out = {'number_format': self._decode_numfmt(root, xf_node),
               'bold': False, 'italic': False, 'font_size': None,
               'font_name': None, 'font_color': None, 'fill_color': None,
               'border': None, 'align': None, 'valign': None, 'wrap': False}
        self._decode_font(root, xf_node, out)
        self._decode_fill_border(root, xf_node, out)
        alignment = xf_node.find(f"{{{ns}}}alignment")
        if alignment is not None:
            out['align'] = alignment.get('horizontal')
            out['valign'] = alignment.get('vertical')
            out['wrap'] = alignment.get('wrapText') == '1'
        return out

    def _decode_numfmt(self, root, xf_node):
        """Resolve an xf's numFmtId to its format code where known."""
        ns = self.NS['main']
        fmt_id = int(xf_node.get('numFmtId', '0'))
        if fmt_id == 0:
            return None
        for fmts in root.findall(f"{{{ns}}}numFmts"):
            for entry in fmts.findall(f"{{{ns}}}numFmt"):
                if int(entry.get('numFmtId')) == fmt_id:
                    return entry.get('formatCode')
        return _BUILTIN_FMT_CODES.get(fmt_id)

    def _decode_font(self, root, xf_node, out):
        """Fill font-related keys of a decoded style dict."""
        ns = self.NS['main']
        fonts = root.find(f"{{{ns}}}fonts")
        idx = int(xf_node.get('fontId', '0'))
        if fonts is None or idx >= len(fonts):
            return
        font = fonts[idx]
        out['bold'] = font.find(f"{{{ns}}}b") is not None
        out['italic'] = font.find(f"{{{ns}}}i") is not None
        size = font.find(f"{{{ns}}}sz")
        if size is not None:
            out['font_size'] = float(size.get('val'))
        name = font.find(f"{{{ns}}}name")
        if name is not None:
            out['font_name'] = name.get('val')
        out['font_color'] = self._decode_color(font.find(f"{{{ns}}}color"))

    def _decode_fill_border(self, root, xf_node, out):
        """Fill fill/border keys of a decoded style dict."""
        ns = self.NS['main']
        fills = root.find(f"{{{ns}}}fills")
        fill_idx = int(xf_node.get('fillId', '0'))
        if fills is not None and fill_idx < len(fills):
            pattern = fills[fill_idx].find(f"{{{ns}}}patternFill")
            if pattern is not None and pattern.get('patternType') == 'solid':
                out['fill_color'] = self._decode_color(
                    pattern.find(f"{{{ns}}}fgColor"))
        borders = root.find(f"{{{ns}}}borders")
        border_idx = int(xf_node.get('borderId', '0'))
        if borders is not None and border_idx < len(borders):
            sides = {side: el.get('style')
                     for side in ('left', 'right', 'top', 'bottom')
                     for el in [borders[border_idx].find(f"{{{ns}}}{side}")]
                     if el is not None and el.get('style')}
            if sides:
                styles = set(sides.values())
                out['border'] = (styles.pop()
                                 if len(styles) == 1 and len(sides) == 4
                                 else sides)

    @staticmethod
    def _decode_color(color_el):
        """Decode a color element: rgb → hex string, theme → raw dict."""
        if color_el is None:
            return None
        rgb = color_el.get('rgb')
        if rgb is not None:
            return rgb[2:] if len(rgb) == 8 and rgb.startswith('FF') else rgb
        if color_el.get('theme') is not None:
            out = {'theme': int(color_el.get('theme'))}
            if color_el.get('tint') is not None:
                out['tint'] = float(color_el.get('tint'))
            return out
        return None

    def _numfmt_id_for(self, num_fmts, format_code):
        """Return the numFmtId for ``format_code``, creating a numFmt if needed."""
        ns = self.NS['main']
        entries = num_fmts.findall(f"{{{ns}}}numFmt")
        for entry in entries:
            if entry.get('formatCode') == format_code:
                return int(entry.get('numFmtId'))
        fmt_id = max([int(e.get('numFmtId')) for e in entries]
                     + [self.CUSTOM_FMT_BASE - 1]) + 1
        entry = ET.SubElement(num_fmts, f"{{{ns}}}numFmt")
        entry.set('numFmtId', str(fmt_id))
        entry.set('formatCode', format_code)
        num_fmts.set('count', str(len(entries) + 1))
        return fmt_id

    def _default_date_style(self, value):
        """Style id for a default date/datetime number format."""
        code = ('yyyy-mm-dd hh:mm:ss' if isinstance(value, datetime.datetime)
                else 'yyyy-mm-dd')
        return self.add_number_format(code)

    def _styles_tree(self):
        """Return the parsed ``xl/styles.xml`` tree, raising if it is absent."""
        if 'xl/styles.xml' not in self.trees:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                try:
                    self._get_tree(zin, 'xl/styles.xml')
                except KeyError:
                    raise FileNotFoundError(
                        "xl/styles.xml not found; a styles part is required"
                    ) from None
        return self.trees['xl/styles.xml']

    # pylint: disable=too-many-arguments
    def update_table_cell(self, table_name, row_offset, col_name, value, *, style_id=None):
        """Updates table cell by column name and expands table range automatically."""
        table = self.table_map[table_name]
        col_idx = table['columns'][col_name]
        abs_row = table['start_indices'][0] + row_offset
        abs_col = table['start_indices'][1] + col_idx

        # Range Expander
        curr_end_row = cell_to_indices(table['range'][1])[0]
        if abs_row > curr_end_row:
            table['range'][1] = indices_to_cell(
                abs_row, table['start_indices'][1] + len(table['columns']) - 1
            )
            root = self.trees[table['xml_path']].getroot()
            root.set('ref', f"{table['range'][0]}:{table['range'][1]}")

        self.update_cell(
            table['sheet'], indices_to_cell(abs_row, abs_col), value=value, style_id=style_id
        )

    def get_cell(self, sheet_name, cell_ref):
        """Read a cell's value, resolving shared strings and typed cells.

        Returns ``str``/``int``/``float``/``bool`` (or the cached result of a
        formula cell), or ``None`` for a missing or empty cell.
        """
        ns = self.NS['main']
        path = self.sheet_map.get(sheet_name, sheet_name)
        if path in self.trees:
            root = self.trees[path].getroot()
        else:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                with zin.open(path) as handle:
                    root = ET.parse(handle).getroot()
        cell = root.find(f".//{{{ns}}}c[@r='{cell_ref}']")
        return None if cell is None else self._cell_value(cell)

    def get_table_cell(self, table_name, row_offset, col_name):
        """Read a table cell by column name (mirrors ``update_table_cell``)."""
        table = self.table_map[table_name]
        abs_row = table['start_indices'][0] + row_offset
        abs_col = table['start_indices'][1] + table['columns'][col_name]
        return self.get_cell(table['sheet'], indices_to_cell(abs_row, abs_col))

    def write_range(self, sheet_name, start_ref, rows, *, style_id=None):
        """Write a 2D block of values starting at ``start_ref``, one pass.

        ``rows`` is an iterable of row iterables. ``None`` entries leave the
        existing cell untouched. Much faster than per-cell ``update_cell``
        for large blocks: the sheet tree is resolved once and each row's
        cells are merged and re-sorted in a single operation.
        """
        ns = self.NS['main']
        sheet_data = self._sheet_root(sheet_name).find(f".//{{{ns}}}sheetData")
        top, left = cell_to_indices(start_ref)
        wrote = False
        for r_off, row_values in enumerate(rows):
            row_values = list(row_values)
            if all(v is None for v in row_values):
                continue
            row = self._row_get_or_create(sheet_data, top + r_off + 1)
            existing = {c.get('r'): c for c in row}
            for c_off, value in enumerate(row_values):
                if value is None:
                    continue
                ref = indices_to_cell(top + r_off, left + c_off)
                cell = existing.get(ref)
                if cell is None:
                    cell = ET.SubElement(row, f"{{{ns}}}c")
                    cell.set('r', ref)
                    existing[ref] = cell
                self._apply_bulk_value(cell, value, style_id)
                wrote = True
            row[:] = sorted(row, key=lambda c: cell_to_indices(c.get('r'))[1])
        if wrote:
            self._set_full_calc_on_load()

    def _apply_bulk_value(self, cell, value, style_id):
        """Set one bulk-write cell: style (or auto date style) + value."""
        if style_id is None and isinstance(value, datetime.date):
            try:
                style_id = self._default_date_style(value)
            except FileNotFoundError:
                pass  # no styles part: write the serial unformatted
        if style_id is not None:
            cell.set('s', str(style_id))
        self._set_cell_value(cell, value)

    def get_range(self, sheet_name, ref):
        """Read a rectangular range as a list of rows (missing cells → None)."""
        ns = self.NS['main']
        path = self.sheet_map.get(sheet_name, sheet_name)
        if path in self.trees:
            root = self.trees[path].getroot()
        else:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                with zin.open(path) as handle:
                    root = ET.parse(handle).getroot()
        start, end = ref.split(':')
        top, left = cell_to_indices(start)
        bottom, right = cell_to_indices(end)
        sheet_data = root.find(f".//{{{ns}}}sheetData")
        by_row = {int(r.get('r')): r for r in sheet_data}
        result = []
        for row_idx in range(top, bottom + 1):
            row = by_row.get(row_idx + 1)
            cells = {} if row is None else {c.get('r'): c for c in row}
            result.append([
                None if cells.get(indices_to_cell(row_idx, col)) is None
                else self._cell_value(cells[indices_to_cell(row_idx, col)])
                for col in range(left, right + 1)
            ])
        return result

    def iter_table_rows(self, table_name):
        """Yield each table data row as a ``{column_name: value}`` dict."""
        table = self.table_map[table_name]
        top, left = table['start_indices']
        bottom, right = cell_to_indices(table['range'][1])
        if bottom == top:
            return  # header-only table
        names = sorted(table['columns'], key=table['columns'].get)
        data_ref = (f"{indices_to_cell(top + 1, left)}:"
                    f"{indices_to_cell(bottom, right)}")
        for row_values in self.get_range(table['sheet'], data_ref):
            yield dict(zip(names, row_values))

    # pylint: disable=too-many-return-statements
    def _cell_value(self, cell):
        """Convert a ``<c>`` element into a Python value."""
        ns = self.NS['main']
        cell_type = cell.get('t')
        if cell_type == 'inlineStr':
            node = cell.find(f".//{{{ns}}}t")
            return node.text if node is not None else None
        value = cell.find(f"{{{ns}}}v")
        if cell_type == 's':
            if value is None or value.text is None:
                return None
            strings = self._shared_strings_list()
            index = int(value.text)
            return strings[index] if 0 <= index < len(strings) else None
        if value is None or value.text is None:
            return None
        if cell_type == 'b':
            return value.text == '1'
        if cell_type in ('str', 'e'):
            return value.text
        number = self._parse_number(value.text)
        style = cell.get('s')
        if style is not None and int(style) in self._date_style_ids():
            return self._from_excel_serial(number)
        return number

    @staticmethod
    def _parse_number(text):
        """Parse a numeric cell payload, preserving int where exact."""
        try:
            return int(text)
        except ValueError:
            return float(text)

    def _date_style_ids(self):
        """Return the set of ``cellXfs`` indexes whose number format is a date."""
        if self._date_style_cache is None:
            ns = self.NS['main']
            self._date_style_cache = set()
            root = self._styles_root_readonly()
            if root is not None:
                custom = {
                    int(e.get('numFmtId')): e.get('formatCode') or ''
                    for fmts in root.findall(f"{{{ns}}}numFmts")
                    for e in fmts.findall(f"{{{ns}}}numFmt")
                }
                cell_xfs = root.find(f"{{{ns}}}cellXfs")
                for idx, xf_node in enumerate(cell_xfs if cell_xfs is not None else []):
                    fmt_id = int(xf_node.get('numFmtId', '0'))
                    if fmt_id in _BUILTIN_DATE_FMTS or \
                            self._format_is_datish(custom.get(fmt_id, '')):
                        self._date_style_cache.add(idx)
        return self._date_style_cache

    def _styles_root_readonly(self):
        """Return the styles root without pulling it into the modified set."""
        if 'xl/styles.xml' in self.trees:
            return self.trees['xl/styles.xml'].getroot()
        with zipfile.ZipFile(self.filename, 'r') as zin:
            try:
                with zin.open('xl/styles.xml') as handle:
                    return ET.parse(handle).getroot()
            except KeyError:
                return None

    @staticmethod
    def _format_is_datish(code):
        """True if a custom format code renders dates/times (y/m/d/h/s tokens)."""
        stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', '', code)
        return any(ch in 'ymdhs' for ch in stripped.lower())

    def _shared_strings_list(self):
        """Lazily load ``xl/sharedStrings.xml`` as a list of plain strings."""
        if self._shared_strings is None:
            ns = self.NS['main']
            self._shared_strings = []
            with zipfile.ZipFile(self.filename, 'r') as zin:
                try:
                    handle = zin.open('xl/sharedStrings.xml')
                except KeyError:
                    return self._shared_strings
                with handle:
                    root = ET.parse(handle).getroot()
            for si_node in root.findall(f"{{{ns}}}si"):
                self._shared_strings.append(
                    ''.join(t.text or '' for t in si_node.iter(f"{{{ns}}}t"))
                )
        return self._shared_strings

    def resize_table(self, table_name, *, add_rows=0, add_cols=0):
        """Grow or shrink a table along the row and/or column axis.

        Growing shoves any tables that would collide below (row growth) or to
        the right (column growth) of the target by the *minimal* amount needed
        to clear it, cascading through further tables. Shrinking (negative
        deltas) only narrows the range and never moves other tables. Column
        growth also appends ``tableColumn`` metadata + header cells. When a
        table is moved, formulas referencing its cells and merged ranges inside
        it are rewritten to follow the move.
        """
        if add_rows == 0 and add_cols == 0:
            return
        table = self.table_map[table_name]
        with zipfile.ZipFile(self.filename, 'r') as zin:
            self._get_tree(zin, self.sheet_map[table['sheet']])
        if add_rows:
            self._resize_rows(table_name, table, add_rows)
        if add_cols:
            self._resize_cols(table_name, table, add_cols)

    def _resize_rows(self, target_name, table, add_rows):
        """Apply a row-axis resize (axis 0), shoving colliding tables down."""
        top, _ = table['start_indices']
        end_r, end_c = cell_to_indices(table['range'][1])
        new_end_r = end_r + add_rows
        if new_end_r < top:
            raise ValueError(
                f"resize_table: add_rows={add_rows} would shrink "
                f"'{target_name}' above its header row"
            )
        table['range'][1] = indices_to_cell(new_end_r, end_c)
        self.trees[table['xml_path']].getroot().set(
            'ref', f"{table['range'][0]}:{table['range'][1]}"
        )
        if add_rows > 0:
            self._resolve_collisions(target_name, table, axis=0)

    def _resize_cols(self, target_name, table, add_cols):
        """Apply a column-axis resize (axis 1), shoving colliding tables right."""
        top, left = table['start_indices']
        end_r, end_c = cell_to_indices(table['range'][1])
        new_end_c = end_c + add_cols
        if new_end_c < left:
            raise ValueError(
                f"resize_table: add_cols={add_cols} would remove all "
                f"columns of '{target_name}'"
            )
        table['range'][1] = indices_to_cell(end_r, new_end_c)
        self.trees[table['xml_path']].getroot().set(
            'ref', f"{table['range'][0]}:{table['range'][1]}"
        )
        added = self._adjust_table_columns(table, add_cols, end_c)
        if add_cols > 0:
            self._resolve_collisions(target_name, table, axis=1)
            # Header cells go in last, after colliding tables have vacated.
            for abs_col, name in added:
                self.update_cell(
                    table['sheet'], indices_to_cell(top, abs_col), value=name
                )

    def _adjust_table_columns(self, table, add_cols, old_end_c):
        """Add/remove ``tableColumn`` entries; return new (abs_col, name) pairs."""
        ns = self.NS['main']
        cols_el = self.trees[table['xml_path']].getroot().find(
            f"{{{ns}}}tableColumns"
        )
        columns = list(cols_el)
        added = []
        if add_cols > 0:
            names = {c.get('name') for c in columns}
            next_id = max((int(c.get('id')) for c in columns), default=0) + 1
            for k in range(add_cols):
                name = self._unique_col_name(names, len(columns) + k + 1)
                names.add(name)
                col_el = ET.SubElement(cols_el, f"{{{ns}}}tableColumn")
                col_el.set('id', str(next_id))
                col_el.set('name', name)
                next_id += 1
                table['columns'][name] = len(columns) + k
                added.append((old_end_c + 1 + k, name))
        else:
            for col_el in columns[len(columns) + add_cols:]:
                table['columns'].pop(col_el.get('name'), None)
                cols_el.remove(col_el)
        cols_el.set('count', str(len(list(cols_el))))
        return added

    @staticmethod
    def _unique_col_name(existing, start):
        """Return the first ``Column{n}`` name not already in ``existing``."""
        i = start
        while f"Column{i}" in existing:
            i += 1
        return f"Column{i}"

    @staticmethod
    def _table_box(table):
        """Return (top, left, bottom, right) 0-based indices of a table."""
        top, left = table['start_indices']
        bottom, right = cell_to_indices(table['range'][1])
        return top, left, bottom, right

    # pylint: disable=too-many-locals
    def _resolve_collisions(self, target_name, target, axis):
        """Compute minimal shifts for tables the target now overlaps.

        ``axis=0`` shoves colliding tables down, ``axis=1`` shoves them right.
        Starting from a valid (non-overlapping) layout, only the target's new
        trailing edge can trigger shifts, so movement stays minimal.
        """
        lead = 0 if axis == 0 else 1          # top / left
        trail = 2 if axis == 0 else 3         # bottom / right
        cross_lo, cross_hi = (1, 3) if axis == 0 else (0, 2)
        tables = {
            name: t for name, t in self.table_map.items()
            if t['sheet'] == target['sheet']
        }
        orig_lead = {name: self._table_box(t)[lead] for name, t in tables.items()}
        shift = {name: 0 for name in tables}

        changed = True
        while changed:
            changed = False
            for a_name, a_tbl in tables.items():
                a_box = self._table_box(a_tbl)
                a_trail = a_box[trail] + shift[a_name]
                for b_name, b_tbl in tables.items():
                    if b_name in (a_name, target_name):
                        continue  # never move the target itself
                    if orig_lead[a_name] >= orig_lead[b_name]:
                        continue  # only push tables that started after A
                    b_box = self._table_box(b_tbl)
                    if a_box[cross_hi] < b_box[cross_lo] or \
                            b_box[cross_hi] < a_box[cross_lo]:
                        continue  # no overlap on the cross axis
                    needed = (a_trail + 1) - (b_box[lead] + shift[b_name])
                    if needed > 0:
                        shift[b_name] += needed
                        changed = True

        # Apply moves furthest-first so each destination band is already clear.
        movers = [n for n in tables if n != target_name and shift[n] > 0]
        movers.sort(key=lambda n: orig_lead[n], reverse=True)
        for name in movers:
            self._move_table(tables[name], shift[name], axis)

    def _move_table(self, table, delta, axis):
        """Relocate a table's cell block by ``delta`` along ``axis`` and update refs."""
        if delta <= 0:
            return
        ns = self.NS['main']
        root = self.trees[self.sheet_map[table['sheet']]].getroot()
        sheet_data = root.find(f".//{{{ns}}}sheetData")
        box = self._table_box(table)
        top, left, bottom, right = box
        # Iterate away from the move direction so a cell is never overwritten
        # before it is relocated.
        rows = range(bottom, top - 1, -1) if axis == 0 else range(top, bottom + 1)
        cols = range(left, right + 1) if axis == 0 else range(right, left - 1, -1)
        for row_idx in rows:
            row = sheet_data.find(f"{{{ns}}}row[@r='{row_idx + 1}']")
            if row is None:
                continue
            for col in cols:
                new_row_idx = row_idx + delta if axis == 0 else row_idx
                new_col = col if axis == 0 else col + delta
                self._relocate_cell(
                    sheet_data, row, indices_to_cell(row_idx, col),
                    new_row_idx=new_row_idx, new_col=new_col
                )

        # References that pointed into the moved block follow it to its new home.
        self._rewrite_formulas(sheet_data, table['sheet'], box, delta=delta, axis=axis)
        self._shift_merged_cells(root, box, delta, axis)
        self._rewrite_range_features(root, table['sheet'], box, delta=delta, axis=axis)
        self._rewrite_defined_names(table['sheet'], box, delta=delta, axis=axis)
        # Rewritten formulas leave cached <v> results stale; force a recalc.
        self._invalidate_calc_cache()

        new_top = top + (delta if axis == 0 else 0)
        new_left = left + (delta if axis == 1 else 0)
        new_bottom = bottom + (delta if axis == 0 else 0)
        new_right = right + (delta if axis == 1 else 0)
        table['start_indices'] = (new_top, new_left)
        table['range'] = [
            indices_to_cell(new_top, new_left),
            indices_to_cell(new_bottom, new_right),
        ]
        self.trees[table['xml_path']].getroot().set(
            'ref', f"{table['range'][0]}:{table['range'][1]}"
        )

    def _rewrite_formulas(self, sheet_data, sheet_name, box, *, delta, axis):
        """Shift every reference into ``box`` by ``delta`` across all formulas."""
        ns = self.NS['main']
        for f_node in sheet_data.iter(f"{{{ns}}}f"):
            if f_node.text:
                f_node.text = self._shift_formula_refs(
                    f_node.text, sheet_name, box, delta=delta, axis=axis
                )
            shared_ref = f_node.get('ref')
            if shared_ref:
                moved = self._shift_range_ref(shared_ref, box, delta, axis)
                if moved is not None:
                    f_node.set('ref', moved)

    def _shift_formula_refs(self, text, sheet_name, box, *, delta, axis):
        """Return ``text`` with each A1 reference inside ``box`` shifted."""
        top, left, bottom, right = box

        def repl(match):
            sheet = match.group('sheet')
            if sheet and sheet[:-1].strip("'") != sheet_name:
                return match.group(0)  # a different sheet
            start = match.start()
            if start > 0 and (text[start - 1].isalnum() or text[start - 1] == '_'):
                return match.group(0)  # part of a longer name
            if match.end() < len(text) and text[match.end()] == '(':
                return match.group(0)  # a function call, not a reference
            row, col = cell_to_indices(match.group('col') + match.group('row'))
            if not (top <= row <= bottom and left <= col <= right):
                return match.group(0)
            if axis == 0:
                row += delta
            else:
                col += delta
            letters = indices_to_cell(0, col)[:-1]
            return (f"{sheet or ''}{match.group('c_abs')}{letters}"
                    f"{match.group('r_abs')}{row + 1}")

        return _CELL_REF_RE.sub(repl, text)

    def _shift_merged_cells(self, root, box, delta, axis):
        """Shift ``mergeCell`` ranges fully contained in ``box`` by ``delta``."""
        ns = self.NS['main']
        merge_cells = root.find(f"{{{ns}}}mergeCells")
        if merge_cells is None:
            return
        for merge in merge_cells.findall(f"{{{ns}}}mergeCell"):
            moved = self._shift_range_ref(merge.get('ref'), box, delta, axis)
            if moved is not None:
                merge.set('ref', moved)

    def _rewrite_range_features(self, root, sheet_name, box, *, delta, axis):
        """Shift conditional-formatting / data-validation regions + formulas."""
        ns = self.NS['main']
        for cf_node in root.findall(f"{{{ns}}}conditionalFormatting"):
            self._shift_sqref(cf_node, box, delta, axis)
            for formula in cf_node.iter(f"{{{ns}}}formula"):
                if formula.text:
                    formula.text = self._shift_formula_refs(
                        formula.text, sheet_name, box, delta=delta, axis=axis
                    )
        validations = root.find(f"{{{ns}}}dataValidations")
        if validations is not None:
            for dv_node in validations.findall(f"{{{ns}}}dataValidation"):
                self._shift_sqref(dv_node, box, delta, axis)
                for tag in ('formula1', 'formula2'):
                    node = dv_node.find(f"{{{ns}}}{tag}")
                    if node is not None and node.text:
                        node.text = self._shift_formula_refs(
                            node.text, sheet_name, box, delta=delta, axis=axis
                        )

    def _shift_sqref(self, elem, box, delta, axis):
        """Shift each fully-contained range in a space-separated ``sqref``."""
        sqref = elem.get('sqref')
        if not sqref:
            return
        parts = [
            self._shift_range_ref(part, box, delta, axis) or part
            for part in sqref.split()
        ]
        elem.set('sqref', ' '.join(parts))

    def _rewrite_defined_names(self, sheet_name, box, *, delta, axis):
        """Shift workbook defined names whose ranges point into the moved block."""
        ns = self.NS['main']
        workbook = self.trees.get('xl/workbook.xml')
        if workbook is None:
            return
        names = workbook.getroot().find(f"{{{ns}}}definedNames")
        if names is None:
            return
        for name in names.findall(f"{{{ns}}}definedName"):
            if name.text:
                name.text = self._shift_name_refs(
                    name.text, sheet_name, box, delta=delta, axis=axis
                )

    def _shift_name_refs(self, text, sheet_name, box, *, delta, axis):
        """Shift each sheet-qualified range in a defined-name expression."""

        def repl(match):
            start = match.start()
            if start > 0 and (text[start - 1].isalnum() or text[start - 1] == '_'):
                return match.group(0)  # part of a longer name
            sheet = match.group('sheet')
            if sheet and sheet[:-1].strip("'") != sheet_name:
                return match.group(0)  # a different sheet
            end_a, end_b = match.group('a'), match.group('b')
            if end_b is None:
                if match.end() < len(text) and text[match.end()] == '(':
                    return match.group(0)  # a function call
                shifted = self._shift_endpoint(end_a, box, delta, axis)
                return match.group(0) if shifted is None else f"{sheet or ''}{shifted}"
            new_a = self._shift_endpoint(end_a, box, delta, axis)
            new_b = self._shift_endpoint(end_b, box, delta, axis)
            if new_a is None or new_b is None:
                return match.group(0)  # not fully contained
            return f"{sheet or ''}{new_a}:{new_b}"

        return _RANGE_RE.sub(repl, text)

    @staticmethod
    def _shift_endpoint(endpoint, box, delta, axis):
        """Shift a single ``$A$5``-style endpoint if inside ``box``, else None."""
        top, left, bottom, right = box
        match = re.match(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)", endpoint)
        c_abs, col_str, r_abs, row_str = match.groups()
        row, col = cell_to_indices(col_str + row_str)
        if not (top <= row <= bottom and left <= col <= right):
            return None
        if axis == 0:
            row += delta
        else:
            col += delta
        return f"{c_abs}{indices_to_cell(0, col)[:-1]}{r_abs}{row + 1}"

    def _invalidate_calc_cache(self):
        """Mark cached results stale: recalc on load + drop the calc chain."""
        self._drop_calc_chain = True
        self._set_full_calc_on_load()

    def _set_full_calc_on_load(self):
        """Flag the workbook so Excel recalculates on open (clears stale caches)."""
        ns = self.NS['main']
        workbook = self.trees.get('xl/workbook.xml')
        if workbook is None:
            return
        root = workbook.getroot()
        calc_pr = root.find(f"{{{ns}}}calcPr")
        if calc_pr is None:
            calc_pr = ET.Element(f"{{{ns}}}calcPr")
            # calcPr follows definedNames (or sheets) in the schema order.
            anchor = root.find(f"{{{ns}}}definedNames")
            if anchor is None:
                anchor = root.find(f"{{{ns}}}sheets")
            children = list(root)
            idx = children.index(anchor) + 1 if anchor is not None else len(children)
            root.insert(idx, calc_pr)
        calc_pr.set('fullCalcOnLoad', '1')

    @staticmethod
    def _shift_range_ref(ref, box, delta, axis):
        """Shift a cell/range ref by ``delta`` if fully inside ``box``, else None."""
        top, left, bottom, right = box
        coords = [cell_to_indices(part) for part in ref.split(':')]
        if any(not (top <= r <= bottom and left <= c <= right) for r, c in coords):
            return None
        shifted = []
        for row, col in coords:
            if axis == 0:
                row += delta
            else:
                col += delta
            shifted.append(indices_to_cell(row, col))
        return ':'.join(shifted)

    def _relocate_cell(self, sheet_data, src_row, old_ref, *, new_row_idx, new_col):
        """Move a single cell from ``old_ref`` to (new_row_idx, new_col)."""
        ns = self.NS['main']
        cell = src_row.find(f"{{{ns}}}c[@r='{old_ref}']")
        if cell is None:
            return
        src_row.remove(cell)
        new_ref = indices_to_cell(new_row_idx, new_col)
        cell.set('r', new_ref)
        dest_row = self._row_get_or_create(sheet_data, new_row_idx + 1)
        existing = dest_row.find(f"{{{ns}}}c[@r='{new_ref}']")
        if existing is not None:
            dest_row.remove(existing)
        self._cell_insert_sorted(dest_row, cell)

    def _row_get_or_create(self, sheet_data, row_num):
        """Return the row element for ``row_num``, inserting it in sorted order."""
        ns = self.NS['main']
        existing = sheet_data.find(f"{{{ns}}}row[@r='{row_num}']")
        if existing is not None:
            return existing
        new_row = ET.Element(f"{{{ns}}}row")
        new_row.set('r', str(row_num))
        idx = len(sheet_data)
        for i, row in enumerate(sheet_data):
            if int(row.get('r')) > row_num:
                idx = i
                break
        sheet_data.insert(idx, new_row)
        return new_row

    @staticmethod
    def _cell_insert_sorted(row, cell):
        """Insert ``cell`` into ``row`` keeping cells ordered by column index."""
        col = cell_to_indices(cell.get('r'))[1]
        idx = len(row)
        for i, existing in enumerate(row):
            if cell_to_indices(existing.get('r'))[1] > col:
                idx = i
                break
        row.insert(idx, cell)

    def save(self, output_path):
        """Preservation Loop: re-serialize edited parts, add new, drop removed."""
        if self._drop_calc_chain:
            # workbook rels live in self.trees (loaded for sheet mapping), so
            # drop the calcChain relationship there; [Content_Types].xml is
            # streamed, so it is patched on the raw bytes below.
            self._strip_calc_chain_relationship()
        dropped = set(self._dropped_parts)
        if self._drop_calc_chain:
            dropped.add('xl/calcChain.xml')
        written = set()
        with zipfile.ZipFile(self.filename, 'r') as zin:
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    name = item.filename
                    if name in dropped:
                        continue
                    written.add(name)
                    if name in self.trees:
                        zout.writestr(name, self._serialize_tree(name))
                    elif name == '[Content_Types].xml':
                        zout.writestr(item, self._patch_content_types(zin.read(name)))
                    else:
                        zout.writestr(item, zin.read(name))
                # Parts created this session (in trees but not in the source).
                for name in self.trees:
                    if name not in written and name not in dropped:
                        zout.writestr(name, self._serialize_tree(name))

    def _serialize_tree(self, name):
        """Serialize a cached tree with the standard XML declaration."""
        buf = io.BytesIO()
        buf.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        self.trees[name].write(buf, encoding='utf-8', xml_declaration=False)
        return buf.getvalue()

    def _patch_content_types(self, data):
        """Apply pending Override add/remove edits to [Content_Types].xml bytes."""
        removals = set(self._ct_remove)
        if self._drop_calc_chain:
            removals.add('/xl/calcChain.xml')
        for part in removals:
            data = re.sub(
                rb'<Override\b[^>]*?PartName="' + re.escape(part).encode()
                + rb'"[^>]*?/>',
                b'', data)
        additions = b''.join(
            b'<Override PartName="%s" ContentType="%s"/>'
            % (part.encode(), content_type.encode())
            for part, content_type in self._ct_add.items())
        if additions:
            data = data.replace(b'</Types>', additions + b'</Types>')
        return data

    def _strip_calc_chain_relationship(self):
        """Remove the calcChain relationship from the workbook rels tree."""
        rels = self.trees.get('xl/_rels/workbook.xml.rels')
        if rels is None:
            return
        rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        root = rels.getroot()
        for rel in root.findall(f"{{{rel_ns}}}Relationship"):
            if (rel.get('Target') or '').endswith('calcChain.xml'):
                root.remove(rel)
