# Excel MCP Server — Advanced

Fast, multi-sheet Excel retrieval and writing via the Model Context Protocol.
Built on [python-calamine](https://github.com/dimastbk/python-calamine) for speed; openpyxl only where calamine cannot reach (formula extraction, writing).

---

## Tools

| Tool | Purpose |
|---|---|
| `excel_list_sheets` | List all sheet names in a workbook |
| `excel_get_sheet_size` | Get rows × cols — O(1) for .xlsx, O(n) otherwise |
| `excel_get_sheet_full` | Read an entire sheet as a markdown table |
| `excel_get_sheet_patches_truncated` | Auto-detect data regions (patches), return them truncated with IDs |
| `excel_get_sheet_patches_by_id` | Return full data for specific patch IDs |
| `excel_get_sheet_cell_ranges` | Return data for explicit A1-notation ranges |
| `excel_write_workbook` | Write `{sheet: markdown_table}` data to an .xlsx file |

---

## Typical Workflows

### Deep analysis (recommended for large files)
```
excel_list_sheets(file)
→ excel_get_sheet_size(file, sheet)                    # check dimensions first (O(1))
→ excel_get_sheet_patches_truncated(file, sheet)       # get patch IDs + shape
→ excel_get_sheet_patches_by_id(file, sheet, [ids…])   # pull full data for relevant patches
```

### Quick overview
```
excel_get_sheet_patches_truncated(file, sheet, top_n_patches=5)
```

### Small sheet
```
excel_get_sheet_full(file, sheet)
```

### Targeted extraction
```
excel_get_sheet_cell_ranges(file, sheet, ["A1:D20", "G5:J15"])
```

---

## Sheet Size

`excel_get_sheet_size` returns `{rows, cols, cells, method}`:

```json
{"sheet_name": "Sales", "rows": 150001, "cols": 12, "cells": 1800012, "method": "xml_dimension_tag"}
```

**Speed:**
- `.xlsx` / `.xlsm`: reads `workbook.xml` (~2 KB) + the first 8 KB of the sheet XML to extract the `<dimension ref="A1:E150001"/>` element. A 200 MB file with 2 million rows takes the same time as a 10-row file.
- `.xls` / `.xlsb` / `.ods`: calamine row iteration — O(n rows) but never loads the full sheet into memory.

---

## Patch IDs

A **patch** is a contiguous rectangular block of non-empty cells, separated from other patches by completely empty rows or columns.

Patch IDs encode their location: `{Sheet}_P{n}_{TopLeft}_{BottomRight}`
- Example: `Sales_P01_A1_E15000`, `Forecast_P03_B5_BD10`
- The bounding box is embedded in the ID, so `excel_get_sheet_patches_by_id` requires no shared state between calls.

---

## Content Modes

Both `values` and `hybrid` are supported on all read tools via the `content` parameter.

| Mode | Behaviour |
|---|---|
| `values` (default) | Computed cell values. Fast — calamine only. |
| `hybrid` | Formula cells show the raw formula (`=SUM(A1:A10)`); hardcoded cells show values. Slightly slower — requires an extra openpyxl pass. |

---

## Truncation

`excel_get_sheet_patches_truncated` truncates large patches to keep context manageable:

```
| Order ID | Date       | Customer   | ... (truncated 50 cols) | Amount  |
|----------|------------|------------|-------------------------|---------|
| 10001    | 2024-01-01 | Apex Corp  | ...                     | $150.00 |
| 10002    | 2024-01-02 | Beta LLC   | ...                     | $200.00 |
| ...      | ...        | ...        | ... (truncated 14,995)  | ...     |
| 24995    | 2024-12-30 | Charlie Inc| ...                     | $150.00 |
| 24996    | 2024-12-31 | Delta Co   | ...                     | $350.00 |
```

Control via:
- `truncate_top_n` — rows/cols shown at head and tail (default 3)
- `truncate_threshold` — dimension must exceed this to trigger truncation (default 10)

---

## Writing

`excel_write_workbook` accepts a dict of `{sheet_name: markdown_table}` and coerces cell values:

| Input string | Excel type |
|---|---|
| `=SUM(A1:A10)` | Formula |
| `42`, `3.14` | Numeric |
| `TRUE` / `FALSE` | Boolean |
| `15%` | 0.15 (numeric) |
| `` (empty) | Blank cell |
| anything else | String |

---

## Installation

```bash
git clone <repo>
cd Excel-MCP-Server-Advanced
python -m venv .venv && source .venv/bin/activate
pip install -e .
```

**Supported formats:** `.xlsx`, `.xls`, `.xlsb`, `.ods` (read); `.xlsx` (write).

---

## MCP Client Configuration

Add to your MCP client config (e.g. Claude Desktop `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "excel": {
      "command": "/path/to/.venv/bin/excel-mcp-server"
    }
  }
}
```

---

## Development

```bash
# Generate test workbook
python tests/create_test_excel.py

# Run tests
python tests/test_tools.py
```
