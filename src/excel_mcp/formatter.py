"""
Markdown table formatting for Excel patch data.
Supports both full rendering and truncated rendering (top-N + bottom-N).
"""

from __future__ import annotations
import math
from typing import Any


def fmt_val(val: Any) -> str:
    """Format a cell value for markdown output."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return str(val)
        # Show as int if lossless
        if val == int(val) and abs(val) < 1e15:
            return str(int(val))
        return repr(val)
    if isinstance(val, str):
        # Escape pipe characters so markdown table doesn't break
        return val.replace("|", "\\|")
    return str(val)


def _make_md_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def _make_separator(n_cols: int) -> str:
    return "|" + "|".join("---" for _ in range(n_cols)) + "|"


def patch_to_markdown(
    data: list[list[Any]],
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    formulas: dict[tuple[int, int], str] | None = None,
    content: str = "values",
    truncate_rows_threshold: int = 10,
    truncate_cols_threshold: int = 10,
    top_n: int = 3,
) -> str:
    """
    Render a patch as a markdown table.

    If the patch exceeds truncate_rows_threshold rows or truncate_cols_threshold
    cols, it will be truncated: show top_n rows + marker + bottom_n rows
    (and same for cols).
    """
    n_rows = max_row - min_row + 1
    n_cols = max_col - min_col + 1

    trunc_rows = n_rows > truncate_rows_threshold
    trunc_cols = n_cols > truncate_cols_threshold

    def cell_content(r: int, c: int) -> str:
        if content == "hybrid" and formulas and (r, c) in formulas:
            return formulas[(r, c)]
        if r >= len(data):
            return ""
        row = data[r]
        val = row[c] if c < len(row) else None
        return fmt_val(val)

    # Build row/col index lists (None = truncation placeholder)
    if trunc_rows:
        row_idxs: list[int | None] = (
            list(range(min_row, min_row + top_n))
            + [None]
            + list(range(max_row - top_n + 1, max_row + 1))
        )
        hidden_rows = n_rows - 2 * top_n
    else:
        row_idxs = list(range(min_row, max_row + 1))
        hidden_rows = 0

    if trunc_cols:
        col_idxs: list[int | None] = (
            list(range(min_col, min_col + top_n))
            + [None]
            + list(range(max_col - top_n + 1, max_col + 1))
        )
        hidden_cols = n_cols - 2 * top_n
    else:
        col_idxs = list(range(min_col, max_col + 1))
        hidden_cols = 0

    lines: list[str] = []
    is_header = True

    for row_idx in row_idxs:
        cells: list[str] = []

        for col_idx in col_idxs:
            if row_idx is None and col_idx is None:
                # Intersection of truncated row + truncated col
                cells.append(f"... (truncated {hidden_rows} rows)")
            elif row_idx is None:
                # Truncated row, normal column
                cells.append("...")
            elif col_idx is None:
                if is_header:
                    # Header row, truncated col -> show col count
                    cells.append(f"... (truncated {hidden_cols} cols)")
                else:
                    # Data row, truncated col
                    cells.append("...")
            else:
                cells.append(cell_content(row_idx, col_idx))

        lines.append(_make_md_row(cells))
        if is_header:
            lines.append(_make_separator(len(cells)))
            is_header = False

    return "\n".join(lines)


def full_patch_to_markdown(
    data: list[list[Any]],
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    formulas: dict[tuple[int, int], str] | None = None,
    content: str = "values",
) -> str:
    """Render a patch as a full (non-truncated) markdown table."""
    def cell_content(r: int, c: int) -> str:
        if content == "hybrid" and formulas and (r, c) in formulas:
            return formulas[(r, c)]
        if r >= len(data):
            return ""
        row = data[r]
        val = row[c] if c < len(row) else None
        return fmt_val(val)

    lines: list[str] = []
    is_header = True
    for r in range(min_row, max_row + 1):
        cells = [cell_content(r, c) for c in range(min_col, max_col + 1)]
        lines.append(_make_md_row(cells))
        if is_header:
            lines.append(_make_separator(len(cells)))
            is_header = False
    return "\n".join(lines)
