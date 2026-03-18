"""
Excel MCP Server — 6 focused tools for fast multi-sheet Excel work.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Any, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from .formatter import _make_md_row, _make_separator, fmt_val, full_patch_to_markdown, patch_to_markdown
from .patches import (
    detect_patches,
    is_empty_value,
    make_patch_id,
    parse_a1,
    parse_patch_id_cells,
    parse_range,
)
from .reader import get_sheet_names as _get_sheet_names
from .reader import get_sheet_size as _get_sheet_size
from .reader import read_sheet_formulas, read_sheet_values
from .writer import write_excel as _write_excel

mcp = FastMCP(
    "excel-mcp-server",
    instructions=(
        "Excel MCP Server — 8 tools for reading and writing Excel workbooks.\n\n"
        "STANDARD WORKFLOW FOR ANALYSIS:\n"
        "  1. excel_list_sheets                  — discover available sheets\n"
        "  2. excel_get_sheet_size               — check rows × cols (O(1) for xlsx)\n"
        "  3. excel_get_sheet_patches_truncated  — overview of all data regions\n"
        "  4. excel_get_sheet_patches_by_id      — full data for key patches\n\n"
        "ALTERNATIVES:\n"
        "  • excel_get_sheet_full         — read a small sheet entirely (< ~500 rows)\n"
        "  • excel_get_sheet_cell_ranges  — read specific A1-notation ranges\n"
        "  • excel_write_workbook         — write data back to Excel\n\n"
        "Use excel_get_sheet_size first on unknown files to decide the read strategy. "
        "For quick overviews: excel_get_sheet_patches_truncated with top_n_patches=5."
    ),
)

# Shared content parameter description reused across tools
_CONTENT_DOC = (
    "Content mode:\n"
    "  'values' (default) — computed cell values. Fast.\n"
    "  'hybrid' — formula cells shown as '=SUM(A1:A10)'; hardcoded cells shown as values. "
    "Slightly slower (requires an extra openpyxl pass)."
)


# ──────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ──────────────────────────────────────────────────────────────────────────────

def _load(
    file_path: str, sheet_name: str, content: str
) -> tuple[list[list[Any]], dict[tuple[int, int], str] | None]:
    """Load sheet values (always calamine) and optionally formulas (openpyxl)."""
    path = str(Path(file_path).expanduser().resolve())
    data = read_sheet_values(path, sheet_name)
    formulas = read_sheet_formulas(path, sheet_name) if content == "hybrid" else None
    return data, formulas


def _err(e: Exception) -> str:
    import traceback
    return json.dumps({"error": str(e), "traceback": traceback.format_exc()})


# ──────────────────────────────────────────────────────────────────────────────
# Tool 1: excel_list_sheets
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_list_sheets(
    file_path: Annotated[str, Field(
        description="Path to the Excel file (.xlsx, .xls, .xlsb, .ods)."
    )],
) -> str:
    """
    Return all sheet names in an Excel workbook as a JSON array.

    Use this tool first — before any read or write operation — to discover what
    sheets exist. This is the required first step for any Excel analysis task.

    Returns a JSON array of strings, e.g. ["Sheet1", "Sales", "Config"].

    EXAMPLE:
        excel_list_sheets("/data/report.xlsx")
        → ["Sales", "Forecast", "Config", "Formulas"]
    """
    try:
        path = str(Path(file_path).expanduser().resolve())
        return json.dumps(_get_sheet_names(path))
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Tool 2: excel_get_sheet_size
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_sheet_size(
    file_path: Annotated[str, Field(description="Path to the Excel file (.xlsx, .xls, .xlsb, .ods).")],
    sheet_name: Annotated[str, Field(description="Sheet name as returned by excel_list_sheets.")],
) -> str:
    """
    Return the row and column count of a sheet. Fast even on huge files.

    Use this tool when you need to know the size of a sheet before deciding
    how to read it — e.g. to choose between excel_get_sheet_full (small sheets)
    and the truncated-patches workflow (large sheets).

    Speed guarantee:
      .xlsx / .xlsm  — O(1): reads ~10 KB from the ZIP (workbook metadata +
                       the first 8 KB of the sheet XML) regardless of sheet size.
                       A 200 MB sheet with 2 million rows takes the same time
                       as a 10-row sheet.
      .xls / .xlsb / .ods — O(n rows): iterates with calamine row-by-row
                       without materialising the full sheet in memory.

    Returns JSON with rows, cols, total cells, and the method used.

    EXAMPLE:
        excel_get_sheet_size("large_report.xlsx", "Sales")
        → {"sheet_name": "Sales", "rows": 150001, "cols": 12,
           "cells": 1800012, "method": "xml_dimension_tag"}
    """
    try:
        path = str(Path(file_path).expanduser().resolve())
        rows, cols, method = _get_sheet_size(path, sheet_name)
        return json.dumps({
            "sheet_name": sheet_name,
            "rows": rows,
            "cols": cols,
            "cells": rows * cols,
            "method": method,
        })
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Tool 4: excel_get_sheet_full
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_sheet_full(
    file_path: Annotated[str, Field(description="Path to the Excel file.")],
    sheet_name: Annotated[str, Field(
        description="Sheet name as returned by excel_list_sheets."
    )],
    content: Annotated[str, Field(description=_CONTENT_DOC)] = "values",
    to_drop_sparsity: Annotated[Optional[int], Field(
        description=(
            "Minimum number of non-empty cells a row or column must contain to be included. "
            "E.g. to_drop_sparsity=3 silently drops any row or column with fewer than 3 values. "
            "Default None = keep all rows and columns."
        )
    )] = None,
) -> str:
    """
    Return an entire sheet as a single markdown table.

    Use this tool when the sheet is small (< ~500 rows) or when the most
    thorough single-pass analysis is needed. For larger sheets use the
    2-step approach: excel_get_sheet_patches_truncated →
    excel_get_sheet_patches_by_id.

    Returns a markdown table string.

    EXAMPLE:
        excel_get_sheet_full("report.xlsx", "Config")
        excel_get_sheet_full("report.xlsx", "Config", to_drop_sparsity=2)
        excel_get_sheet_full("model.xlsx", "Forecast", content="hybrid")
    """
    try:
        data, formulas = _load(file_path, sheet_name, content)
        return _render_full(data, formulas, content, to_drop_sparsity)
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Tool 5: excel_get_sheet_patches_truncated
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_sheet_patches_truncated(
    file_path: Annotated[str, Field(description="Path to the Excel file.")],
    sheet_name: Annotated[str, Field(
        description="Sheet name as returned by excel_list_sheets."
    )],
    content: Annotated[str, Field(description=_CONTENT_DOC)] = "values",
    top_n_patches: Annotated[int, Field(
        description=(
            "How many patches to return, ranked by area (rows × cols).\n"
            "  -1 (default) = return all detected patches.\n"
            "  Positive integer N = return only the N largest patches.\n"
            "  Set to 5 for quick overviews of large sheets."
        )
    )] = -1,
    truncate_top_n: Annotated[int, Field(
        description=(
            "Number of rows and columns to show at the head and tail of each truncated patch. "
            "Default 3."
        )
    )] = 3,
    truncate_threshold: Annotated[int, Field(
        description=(
            "A patch dimension (rows or cols) must exceed this value to trigger truncation. "
            "Default 10."
        )
    )] = 10,
) -> str:
    """
    Auto-detect all data patches in a sheet and return them truncated with IDs.

    A 'patch' is a contiguous rectangular block of non-empty cells separated from
    other patches by empty rows/columns. Each patch gets a deterministic ID encoding
    its sheet and position, e.g. 'Sales_P01_A1_E15000'.

    Large patches are truncated: top N rows + '... (truncated X rows)' + bottom N rows,
    and similarly for columns.

    Use this tool for:
      • Quick sheet overviews (set top_n_patches=5)
      • STEP 1 of deep analysis: get patch IDs → pass relevant ones to
        excel_get_sheet_patches_by_id for the full data

    Returns JSON: {patch_id: truncated_markdown_table, ...}

    EXAMPLE:
        # Quick overview — top 5 largest tables
        excel_get_sheet_patches_truncated("data.xlsx", "Sales", top_n_patches=5)

        # Deep analysis step 1 — all patches
        result = excel_get_sheet_patches_truncated("data.xlsx", "Sales")
        # → {"Sales_P01_A1_E15000": "| Order ID | ...", "Sales_P02_G1_J50": "..."}
        # Then: excel_get_sheet_patches_by_id("data.xlsx", "Sales", ["Sales_P01_A1_E15000"])
    """
    try:
        data, formulas = _load(file_path, sheet_name, content)
        return _render_patches_truncated(
            data, formulas, content, sheet_name,
            top_n_patches, truncate_top_n, truncate_threshold,
        )
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Tool 6: excel_get_sheet_patches_by_id
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_sheet_patches_by_id(
    file_path: Annotated[str, Field(description="Path to the Excel file.")],
    sheet_name: Annotated[str, Field(
        description="Sheet name the patches belong to."
    )],
    patch_ids: Annotated[list[str], Field(
        description=(
            "List of patch ID strings from a prior excel_get_sheet_patches_truncated call. "
            "The cell bounding box is encoded inside the ID and extracted automatically — "
            "no state is needed between calls. "
            "Example: ['Sales_P01_A1_E15000', 'Sales_P02_G1_J50']"
        )
    )],
    content: Annotated[str, Field(description=_CONTENT_DOC)] = "values",
) -> str:
    """
    Return full (untruncated) data for specific patches identified by ID.

    Use this tool as STEP 2 of deep analysis, after calling
    excel_get_sheet_patches_truncated to identify which patches are relevant.
    The patch ID encodes the bounding-box cell range, so no state is required
    between the two calls.

    Returns JSON: {patch_id: full_markdown_table, ...}

    EXAMPLE:
        # After getting IDs from excel_get_sheet_patches_truncated:
        excel_get_sheet_patches_by_id(
            "data.xlsx",
            "Sales",
            patch_ids=["Sales_P01_A1_E15000"]
        )
        # → {"Sales_P01_A1_E15000": "| Order ID | Date | ... (all 15000 rows)"}
    """
    try:
        data, formulas = _load(file_path, sheet_name, content)
        return _render_patches_by_id(data, formulas, content, patch_ids)
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Tool 7: excel_get_sheet_cell_ranges
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_sheet_cell_ranges(
    file_path: Annotated[str, Field(description="Path to the Excel file.")],
    sheet_name: Annotated[str, Field(
        description="Sheet name as returned by excel_list_sheets."
    )],
    cell_ranges: Annotated[list[str], Field(
        description=(
            "List of cell ranges in A1 notation. "
            "Example: ['A1:C10', 'F5:H20', 'B2:B100']. "
            "Each range is returned as a separate entry in the output dict."
        )
    )],
    content: Annotated[str, Field(description=_CONTENT_DOC)] = "values",
) -> str:
    """
    Return data for explicit A1-notation cell ranges.

    Use this tool when you already know the exact location of data — either from
    a prior patch overview, from the user's instructions, or from domain knowledge.
    More targeted than patch retrieval when the precise range is known.

    Returns JSON: {"A1:C10": markdown_table, "F5:H20": markdown_table, ...}

    EXAMPLE:
        excel_get_sheet_cell_ranges("report.xlsx", "Dashboard", ["A1:D5", "H10:J20"])
        excel_get_sheet_cell_ranges("model.xlsx", "Forecast", ["B5:AZ10"], content="hybrid")
    """
    try:
        data, formulas = _load(file_path, sheet_name, content)
        return _render_cell_ranges(data, formulas, content, cell_ranges)
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Tool 8: excel_write_workbook
# ──────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_write_workbook(
    file_path: Annotated[str, Field(
        description=(
            "Output path for the .xlsx file. Parent directories are created automatically. "
            "Example: '/tmp/report.xlsx' or 'output/results.xlsx'."
        )
    )],
    sheets_data: Annotated[dict[str, str], Field(
        description=(
            "Dict mapping sheet names to markdown table strings. "
            "Each key becomes a sheet tab; each value is a markdown table whose first "
            "row becomes the spreadsheet header.\n\n"
            "Cell coercion rules applied automatically:\n"
            "  '=...'       → Excel formula\n"
            "  '42' / '3.14' → numeric\n"
            "  'TRUE/FALSE'  → boolean\n"
            "  '15%'         → 0.15\n"
            "  ''            → blank cell\n"
            "  anything else → string\n\n"
            "Example:\n"
            '  {"Sales": "| ID | Amount |\\n|---|---|\\n| 1 | 100 |",\n'
            '   "Summary": "| Total |\\n|---|\\n| =SUM(Sales!B2:B3) |"}'
        )
    )],
) -> str:
    """
    Write an Excel workbook to disk from markdown table data.

    Use this tool to create or overwrite an .xlsx file. Provide one markdown
    table per sheet. Cell values are coerced to native Excel types automatically
    (formulas, numbers, booleans, percentages).

    Returns JSON confirmation with output path and sheet count.

    EXAMPLE:
        excel_write_workbook(
            "/tmp/report.xlsx",
            {
                "Revenue": "| Quarter | Amount |\\n|---|---|\\n| Q1 | 10000 |\\n| Q2 | 12000 |",
                "Notes":   "| Author | Comment |\\n|---|---|\\n| Alice | First draft |",
            }
        )
    """
    try:
        path = str(Path(file_path).expanduser().resolve())
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        _write_excel(path, sheets_data)
        sheet_list = list(sheets_data.keys())
        return json.dumps({
            "status": "success",
            "file_path": path,
            "sheets_written": sheet_list,
            "total_sheets": len(sheet_list),
        })
    except Exception as e:
        return _err(e)


# ──────────────────────────────────────────────────────────────────────────────
# Private rendering implementations
# ──────────────────────────────────────────────────────────────────────────────

def _render_full(
    data: list[list[Any]],
    formulas: dict | None,
    content: str,
    to_drop_sparsity: int | None,
) -> str:
    if not data:
        return "*(empty sheet)*"

    n_rows = len(data)
    n_cols = max(len(r) for r in data) if data else 0

    if to_drop_sparsity is not None and to_drop_sparsity > 0:
        row_counts = [
            sum(1 for c in range(n_cols)
                if not is_empty_value(data[r][c] if c < len(data[r]) else None))
            for r in range(n_rows)
        ]
        col_counts = [
            sum(1 for r in range(n_rows)
                if not is_empty_value(data[r][c] if c < len(data[r]) else None))
            for c in range(n_cols)
        ]
        keep_rows = [r for r in range(n_rows) if row_counts[r] >= to_drop_sparsity]
        keep_cols = [c for c in range(n_cols) if col_counts[c] >= to_drop_sparsity]
    else:
        keep_rows = list(range(n_rows))
        keep_cols = list(range(n_cols))

    if not keep_rows or not keep_cols:
        return "*(no data after sparsity filtering)*"

    lines: list[str] = []
    is_header = True
    for r in keep_rows:
        cells: list[str] = []
        for c in keep_cols:
            if content == "hybrid" and formulas and (r, c) in formulas:
                cells.append(formulas[(r, c)])
            else:
                val = data[r][c] if c < len(data[r]) else None
                cells.append(fmt_val(val))
        lines.append(_make_md_row(cells))
        if is_header:
            lines.append(_make_separator(len(cells)))
            is_header = False

    return "\n".join(lines)


def _render_patches_truncated(
    data: list[list[Any]],
    formulas: dict | None,
    content: str,
    sheet_name: str,
    top_n_patches: int,
    truncate_top_n: int,
    truncate_threshold: int,
) -> str:
    patches = detect_patches(data)
    if not patches:
        return json.dumps({})

    patches_with_area = [(p, (p[1] - p[0] + 1) * (p[3] - p[2] + 1)) for p in patches]
    if top_n_patches > 0:
        patches_with_area.sort(key=lambda x: -x[1])
        patches_with_area = patches_with_area[:top_n_patches]
    # Re-sort by top-left position for deterministic ID numbering
    patches_sorted = sorted(patches_with_area, key=lambda x: (x[0][0], x[0][2]))

    result: dict[str, str] = {}
    for idx, (p, _) in enumerate(patches_sorted, start=1):
        min_row, max_row, min_col, max_col = p
        pid = make_patch_id(sheet_name, idx, min_row, max_row, min_col, max_col)
        md = patch_to_markdown(
            data, min_row, max_row, min_col, max_col,
            formulas=formulas,
            content=content,
            truncate_rows_threshold=truncate_threshold,
            truncate_cols_threshold=truncate_threshold,
            top_n=truncate_top_n,
        )
        result[pid] = md

    return json.dumps(result, ensure_ascii=False)


def _render_patches_by_id(
    data: list[list[Any]],
    formulas: dict | None,
    content: str,
    patch_ids: list[str],
) -> str:
    result: dict[str, str] = {}
    for pid in patch_ids:
        try:
            tl, br = parse_patch_id_cells(pid)
            min_row, min_col = parse_a1(tl)
            max_row, max_col = parse_a1(br)
            result[pid] = full_patch_to_markdown(
                data, min_row, max_row, min_col, max_col,
                formulas=formulas,
                content=content,
            )
        except Exception as e:
            result[pid] = f"ERROR: {e}"
    return json.dumps(result, ensure_ascii=False)


def _render_cell_ranges(
    data: list[list[Any]],
    formulas: dict | None,
    content: str,
    cell_ranges: list[str],
) -> str:
    result: dict[str, str] = {}
    for rng in cell_ranges:
        try:
            min_row, max_row, min_col, max_col = parse_range(rng)
            result[rng] = full_patch_to_markdown(
                data, min_row, max_row, min_col, max_col,
                formulas=formulas,
                content=content,
            )
        except Exception as e:
            result[rng] = f"ERROR: {e}"
    return json.dumps(result, ensure_ascii=False)


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
