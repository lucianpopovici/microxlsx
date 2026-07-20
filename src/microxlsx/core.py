"""
Core functionality for MicroXLSX.
"""
import re
import zipfile
import io
import xml.etree.ElementTree as ET
from .utils import cell_to_indices, indices_to_cell

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
        self._set_full_calc_on_load()

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
