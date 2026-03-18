# Excel MCP Server — Claude Guidance

## Project layout

```
src/excel_mcp/
├── server.py     — 6 MCP tools (FastMCP); all business logic delegates to helpers below
├── patches.py    — patch detection (O(n·m) connected-components), ID generation, A1 parsing
├── reader.py     — calamine for values; openpyxl for formula extraction (hybrid mode only)
├── formatter.py  — markdown rendering: full tables and truncated (top-N + bottom-N)
└── writer.py     — markdown table parser + openpyxl workbook writer
tests/
├── create_test_excel.py   — generates tests/test_workbook.xlsx (5-sheet fixture)
└── test_tools.py          — 13 tests; run with: python tests/test_tools.py
```

## Development commands

```bash
source .venv/bin/activate      # activate venv (created during install)
python tests/create_test_excel.py   # regenerate test fixture
python tests/test_tools.py          # run all tests
excel-mcp-server                    # start MCP server (stdio)
```

## Key design rules

- **Speed first.** Always use calamine (`reader.read_sheet_values`) for reading values.
  Only fall back to openpyxl for `content="hybrid"` (formula extraction) or writing.
- **6 single-purpose tools.** Do not recombine them into a mode-based tool. Each tool
  has only the parameters that apply to it — no conditional/dead parameters.
- **Patch IDs are self-contained.** Format: `{Sheet}_P{n}_{TopLeft}_{BottomRight}`.
  The bounding box is in the ID; `excel_get_sheet_patches_by_id` parses it directly
  — no server state is required between calls.
- The `_render_*` private helpers contain all rendering logic. The `@mcp.tool()`
  functions are thin: load data → call helper → return string.

---

## How to use the MCP tools (for analysis tasks)

### Standard workflow — large or unknown file

```
1. excel_list_sheets(file_path)
   → ["Sales", "Forecast", "Config", ...]

2. excel_get_sheet_size(file_path, sheet_name)          ← O(1) for .xlsx
   → {"rows": 150001, "cols": 12, "method": "xml_dimension_tag"}
   Use this to decide the read strategy before loading any cell data.

3. excel_get_sheet_patches_truncated(file_path, sheet_name)
   → {"Sales_P01_A1_E15000": "| Order ID | ... (truncated) ...", ...}
   Use top_n_patches=5 when only a quick overview is needed.

4. excel_get_sheet_patches_by_id(file_path, sheet_name, patch_ids=[...])
   → {"Sales_P01_A1_E15000": "| Order ID | ... (all rows) ..."}
   Pick only the patches relevant to the question.
```

### Small sheet or targeted work

```
excel_get_sheet_full(file_path, sheet_name)            # whole sheet
excel_get_sheet_cell_ranges(file_path, sheet_name, ["A1:D20"])   # known range
```

### Seeing formulas

Pass `content="hybrid"` to any read tool. Formula cells render as `=SUM(A1:A10)`;
hardcoded cells render as their value.

### Writing back

```
excel_write_workbook(file_path, {
    "Sheet1": "| Col A | Col B |\n|---|---|\n| 1 | =A2*2 |"
})
```

### When to use each tool

| Situation | Tool(s) |
|---|---|
| Don't know what sheets exist | `excel_list_sheets` |
| Check sheet dimensions before reading | `excel_get_sheet_size` |
| Quick overview of a large sheet | `excel_get_sheet_patches_truncated` with `top_n_patches=5` |
| Thorough analysis of a large sheet | `excel_get_sheet_patches_truncated` → `excel_get_sheet_patches_by_id` |
| Small sheet (< ~500 rows) | `excel_get_sheet_full` |
| Already know the exact cell range | `excel_get_sheet_cell_ranges` |
| Need to see formulas, not values | Any read tool with `content="hybrid"` |
| Create or update an Excel file | `excel_write_workbook` |
