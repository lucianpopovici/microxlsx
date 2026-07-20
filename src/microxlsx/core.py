"""
Core functionality for MicroXLSX.
"""
import zipfile
import io
import xml.etree.ElementTree as ET
from .utils import cell_to_indices, indices_to_cell

class XLSXPackage:
    """
    Represents an Excel (XLSX) package.
    """
    NS = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/'
    }

    def __init__(self, filename):
        self.filename = filename
        self.trees = {}
        self.sheet_map = {}
        self.table_map = {}
        for prefix, uri in self.NS.items():
            ET.register_namespace(prefix if prefix != 'main' else '', uri)
        self._build_maps()

    def _get_tree(self, zin, path):
        if path not in self.trees:
            with zin.open(path) as f:
                self.trees[path] = ET.parse(f)
        return self.trees[path]

    def _build_maps(self):
        """Builds relationship maps for Sheets and Tables."""
        with zipfile.ZipFile(self.filename, 'r') as zin:
            self._map_sheets(zin)
            self._map_tables(zin)

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
            rel_path = f"xl/worksheets/_rels/{sheet_path.split('/')[-1]}.rels"
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
        t_path = f"xl/tables/{rel.get('Target').split('/')[-1]}"
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

    def _set_cell_value(self, cell, value):
        """Helper to set cell value."""
        if isinstance(value, (int, float)):
            v_node = cell.find(f"{{{self.NS['main']}}}v")
            if v_node is None:
                v_node = ET.SubElement(cell, f"{{{self.NS['main']}}}v")
            v_node.text = str(value)
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

    def resize_table(self, table_name, *, add_rows=0):
        """Grow or shrink a table by ``add_rows`` rows along the row axis.

        Growing shoves any tables that would collide below the target down by
        the *minimal* amount needed to clear it (cascading through further
        tables). Shrinking (negative ``add_rows``) only narrows the range and
        never moves other tables. Column-axis resizing and formula/merge-range
        rewriting are intentionally out of scope for now.
        """
        if add_rows == 0:
            return
        table = self.table_map[table_name]
        r1, _ = table['start_indices']
        end_r, end_c = cell_to_indices(table['range'][1])
        new_end_r = end_r + add_rows
        if new_end_r < r1:
            raise ValueError(
                f"resize_table: add_rows={add_rows} would shrink "
                f"'{table_name}' above its header row"
            )
        # Update the target table's own range + ref (header/start is unchanged).
        table['range'][1] = indices_to_cell(new_end_r, end_c)
        t_root = self.trees[table['xml_path']].getroot()
        t_root.set('ref', f"{table['range'][0]}:{table['range'][1]}")

        if add_rows > 0:
            with zipfile.ZipFile(self.filename, 'r') as zin:
                self._get_tree(zin, self.sheet_map[table['sheet']])
                self._resolve_collisions_down(table_name, table)

    @staticmethod
    def _table_box(table):
        """Return (top, left, bottom, right) 0-based indices of a table."""
        top, left = table['start_indices']
        bottom, right = cell_to_indices(table['range'][1])
        return top, left, bottom, right

    # pylint: disable=too-many-locals
    def _resolve_collisions_down(self, target_name, target):
        """Compute minimal downward shifts for tables the target now overlaps."""
        tables = {
            name: t for name, t in self.table_map.items()
            if t['sheet'] == target['sheet']
        }
        orig_top = {name: self._table_box(t)[0] for name, t in tables.items()}
        shift = {name: 0 for name in tables}

        # Relaxation: repeatedly push a lower table down until nothing overlaps.
        # Starting from a valid (non-overlapping) layout, only the target's new
        # bottom edge can trigger shifts, so movement stays minimal.
        changed = True
        while changed:
            changed = False
            for a_name, a_tbl in tables.items():
                _, a_left, a_bottom0, a_right = self._table_box(a_tbl)
                a_bottom = a_bottom0 + shift[a_name]
                for b_name, b_tbl in tables.items():
                    if b_name in (a_name, target_name):
                        continue  # never move the target itself
                    if orig_top[a_name] >= orig_top[b_name]:
                        continue  # only push tables that started below A
                    b_top0, b_left, _, b_right = self._table_box(b_tbl)
                    if a_right < b_left or b_right < a_left:
                        continue  # no horizontal overlap
                    needed = (a_bottom + 1) - (b_top0 + shift[b_name])
                    if needed > 0:
                        shift[b_name] += needed
                        changed = True

        # Apply moves lowest-first so each destination band is already clear.
        movers = [n for n in tables if n != target_name and shift[n] > 0]
        movers.sort(key=lambda n: orig_top[n], reverse=True)
        for name in movers:
            self._move_table_down(tables[name], shift[name])

    # pylint: disable=too-many-locals
    def _move_table_down(self, table, delta):
        """Relocate a table's cell block down by ``delta`` rows and update refs."""
        if delta <= 0:
            return
        ns = self.NS['main']
        sheet_data = self.trees[self.sheet_map[table['sheet']]].getroot().find(
            f".//{{{ns}}}sheetData"
        )
        top, left, bottom, right = self._table_box(table)
        # Bottom-to-top so a cell is never overwritten before it is moved.
        for row_idx in range(bottom, top - 1, -1):
            row = sheet_data.find(f"{{{ns}}}row[@r='{row_idx + 1}']")
            if row is None:
                continue
            for col in range(left, right + 1):
                old_ref = indices_to_cell(row_idx, col)
                cell = row.find(f"{{{ns}}}c[@r='{old_ref}']")
                if cell is None:
                    continue
                row.remove(cell)
                new_ref = indices_to_cell(row_idx + delta, col)
                cell.set('r', new_ref)
                dest_row = self._row_get_or_create(sheet_data, row_idx + delta + 1)
                existing = dest_row.find(f"{{{ns}}}c[@r='{new_ref}']")
                if existing is not None:
                    dest_row.remove(existing)
                self._cell_insert_sorted(dest_row, cell)

        table['start_indices'] = (top + delta, left)
        table['range'] = [
            indices_to_cell(top + delta, left),
            indices_to_cell(bottom + delta, right),
        ]
        t_root = self.trees[table['xml_path']].getroot()
        t_root.set('ref', f"{table['range'][0]}:{table['range'][1]}")

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
        """Preservation Loop."""
        with zipfile.ZipFile(self.filename, 'r') as zin:
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in self.trees:
                        buf = io.BytesIO()
                        buf.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
                        self.trees[item.filename].write(
                            buf, encoding='utf-8', xml_declaration=False
                        )
                        zout.writestr(item.filename, buf.getvalue())
                    else:
                        zout.writestr(item, zin.read(item.filename))
