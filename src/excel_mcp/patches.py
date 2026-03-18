"""
Patch detection and ID management for Excel sheets.
A 'patch' is a contiguous rectangular block of non-empty cells, separated from
other patches by completely empty rows or columns.
"""

from __future__ import annotations
import re
from typing import Any


def col_to_letter(col_idx: int) -> str:
    """Convert 0-indexed column number to Excel letter notation (0→A, 25→Z, 26→AA)."""
    result = ""
    n = col_idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def letter_to_col(letters: str) -> int:
    """Convert Excel column letters to 0-indexed column number (A→0, Z→25, AA→26)."""
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def cell_notation(row: int, col: int) -> str:
    """Convert 0-indexed (row, col) to A1 notation."""
    return f"{col_to_letter(col)}{row + 1}"


def parse_a1(ref: str) -> tuple[int, int]:
    """Parse A1-notation cell reference to 0-indexed (row, col)."""
    m = re.match(r"^([A-Za-z]+)(\d+)$", ref.strip())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return int(m.group(2)) - 1, letter_to_col(m.group(1))


def parse_range(range_str: str) -> tuple[int, int, int, int]:
    """Parse 'A1:C10' into (min_row, max_row, min_col, max_col) 0-indexed."""
    parts = range_str.strip().split(":")
    if len(parts) != 2:
        raise ValueError(f"Invalid range: {range_str!r}")
    r1, c1 = parse_a1(parts[0])
    r2, c2 = parse_a1(parts[1])
    return min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)


def is_empty_value(val: Any) -> bool:
    """Return True if the cell value should be considered empty."""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    if isinstance(val, float):
        import math
        if math.isnan(val):
            return True
    return False


def detect_patches(data: list[list[Any]]) -> list[tuple[int, int, int, int]]:
    """
    Detect rectangular patches of non-empty cells.

    Algorithm (O(n*m)):
    1. Find non-empty rows; group contiguous non-empty row ranges.
    2. For each row group, find non-empty columns; group contiguous col ranges.
    3. Each (row_group, col_group) pair = one patch.

    Returns list of (min_row, max_row, min_col, max_col) tuples, 0-indexed,
    sorted by (min_row, min_col).
    """
    if not data:
        return []

    num_rows = len(data)

    def get_val(r: int, c: int) -> Any:
        row = data[r]
        return row[c] if c < len(row) else None

    # Pass 1: find non-empty rows
    non_empty_rows: list[int] = []
    for r in range(num_rows):
        if any(not is_empty_value(v) for v in data[r]):
            non_empty_rows.append(r)

    if not non_empty_rows:
        return []

    # Group contiguous non-empty rows
    row_groups: list[tuple[int, int]] = []
    rs = non_empty_rows[0]
    rp = non_empty_rows[0]
    for r in non_empty_rows[1:]:
        if r > rp + 1:
            row_groups.append((rs, rp))
            rs = r
        rp = r
    row_groups.append((rs, rp))

    patches: list[tuple[int, int, int, int]] = []
    for r_start, r_end in row_groups:
        # Collect all non-empty columns in this row band
        nonempty_cols: set[int] = set()
        for r in range(r_start, r_end + 1):
            for c, v in enumerate(data[r]):
                if not is_empty_value(v):
                    nonempty_cols.add(c)
        if not nonempty_cols:
            continue
        sorted_cols = sorted(nonempty_cols)
        # Group contiguous columns
        cs = sorted_cols[0]
        cp = sorted_cols[0]
        for c in sorted_cols[1:]:
            if c > cp + 1:
                patches.append((r_start, r_end, cs, cp))
                cs = c
            cp = c
        patches.append((r_start, r_end, cs, cp))

    return sorted(patches, key=lambda p: (p[0], p[2]))


def make_patch_id(sheet_name: str, patch_index: int,
                  min_row: int, max_row: int,
                  min_col: int, max_col: int) -> str:
    """
    Generate a deterministic, human-readable patch ID.

    Format: {SafeSheetName}_P{index:02d}_{TopLeftCell}_{BottomRightCell}
    Example: Sales_P01_A1_E15000
    """
    safe = re.sub(r"[^A-Za-z0-9]", "_", sheet_name)
    tl = cell_notation(min_row, min_col)
    br = cell_notation(max_row, max_col)
    return f"{safe}_P{patch_index:02d}_{tl}_{br}"


def parse_patch_id_cells(patch_id: str) -> tuple[str, str]:
    """
    Extract top-left and bottom-right cell references from a patch ID.

    The ID always ends in _{CellRef}_{CellRef}, e.g. '...P01_A1_E15000'.
    """
    m = re.search(r"_([A-Z]+\d+)_([A-Z]+\d+)$", patch_id.upper())
    if not m:
        raise ValueError(f"Cannot parse cell range from patch ID: {patch_id!r}")
    return m.group(1), m.group(2)
