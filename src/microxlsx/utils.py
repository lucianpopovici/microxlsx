import re

def cell_to_indices(cell_ref):
    """Converts 'B2' to (1, 1)."""
    match = re.match(r"([A-Z]+)([0-9]+)", cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_str, row_str = match.groups()
    col_idx = 0
    for char in col_str:
        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
    return int(row_str) - 1, col_idx - 1

def indices_to_cell(row_idx, col_idx):
    """Converts (1, 1) to 'B2'."""
    col_str = ""
    temp_col = col_idx + 1
    while temp_col > 0:
        temp_col, remainder = divmod(temp_col - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return f"{col_str}{row_idx + 1}"
