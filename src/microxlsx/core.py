import zipfile
import io
import xml.etree.ElementTree as ET
from .utils import cell_to_indices, indices_to_cell

class XLSXPackage:
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
            # Map Sheets
            wb_tree = self._get_tree(zin, 'xl/workbook.xml')
            sheets = wb_tree.getroot().find(f"{{{self.NS['main']}}}sheets")
            id_to_name = {s.get(f"{{{self.NS['r']}}}id"): s.get('name') for s in sheets}
            rel_tree = self._get_tree(zin, 'xl/_rels/workbook.xml.rels')
            for rel in rel_tree.getroot():
                rid, target = rel.get('Id'), rel.get('Target')
                if rid in id_to_name:
                    path = f"xl/{target}" if not target.startswith('xl/') else target
                    self.sheet_map[id_to_name[rid]] = path

            # Map Tables
            for sheet_name, sheet_path in self.sheet_map.items():
                rel_path = f"xl/worksheets/_rels/{sheet_path.split('/')[-1]}.rels"
                try:
                    with zin.open(rel_path) as f:
                        t_rel_tree = ET.parse(f)
                        for rel in t_rel_tree.getroot():
                            if "table" in rel.get('Type'):
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
                except (KeyError, FileNotFoundError): continue

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

    def update_cell(self, sheet_name, cell_ref, value=None, formula=None, style_id=None):
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
            if style_id is not None: cell.set('s', str(style_id))
            if formula:
                f_node = cell.find(f"{{{self.NS['main']}}}f") or ET.SubElement(cell, f"{{{self.NS['main']}}}f")
                f_node.text = formula.lstrip('=')
            if value is not None:
                if isinstance(value, (int, float)):
                    v = cell.find(f"{{{self.NS['main']}}}v") or ET.SubElement(cell, f"{{{self.NS['main']}}}v")
                    v.text = str(value)
                    cell.attrib.pop('t', None)
                else:
                    cell.set('t', 'inlineStr')
                    is_node = cell.find(f"{{{self.NS['main']}}}is") or ET.SubElement(cell, f"{{{self.NS['main']}}}is")
                    t_node = is_node.find(f"{{{self.NS['main']}}}t") or ET.SubElement(is_node, f"{{{self.NS['main']}}}t")
                    t_node.text = str(value)

    def update_table_cell(self, table_name, row_offset, col_name, value, style_id=None):
        """Updates table cell by column name and expands table range automatically."""
        t = self.table_map[table_name]
        col_idx = t['columns'][col_name]
        abs_row = t['start_indices'][0] + row_offset
        abs_col = t['start_indices'][1] + col_idx
        
        # Range Expander
        curr_end_row = cell_to_indices(t['range'][1])[0]
        if abs_row > curr_end_row:
            t['range'][1] = indices_to_cell(abs_row, t['start_indices'][1] + len(t['columns']) - 1)
            self.trees[t['xml_path']].getroot().set('ref', f"{t['range'][0]}:{t['range'][1]}")
            
        self.update_cell(t['sheet'], indices_to_cell(abs_row, abs_col), value, style_id=style_id)

    def save(self, output_path):
        """Preservation Loop."""
        with zipfile.ZipFile(self.filename, 'r') as zin:
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in self.trees:
                        buf = io.BytesIO()
                        buf.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
                        self.trees[item.filename].write(buf, encoding='utf-8', xml_declaration=False)
                        zout.writestr(item.filename, buf.getvalue())
                    else:
                        zout.writestr(item, zin.read(item.filename))
