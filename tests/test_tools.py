"""
Tests for the Excel MCP Server tools.
Run: python tests/test_tools.py
"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from excel_mcp.server import (
    excel_get_sheet_cell_ranges,
    excel_get_sheet_full,
    excel_get_sheet_patches_by_id,
    excel_get_sheet_patches_truncated,
    excel_get_sheet_size,
    excel_list_sheets,
    excel_write_workbook,
)

TEST_FILE = os.path.join(os.path.dirname(__file__), "test_workbook.xlsx")


def run(label: str, fn):
    print(f"\n{'='*60}")
    print(f"TEST: {label}")
    print("=" * 60)
    try:
        result = fn()
        try:
            parsed = json.loads(result)
            if isinstance(parsed, dict) and "error" in parsed:
                print(f"ERROR: {parsed['error']}")
                if "traceback" in parsed:
                    print(parsed["traceback"])
            else:
                output = json.dumps(parsed, indent=2, ensure_ascii=False)
                print(output[:2000] + ("\n... (truncated)" if len(output) > 2000 else ""))
        except (json.JSONDecodeError, TypeError):
            print(str(result)[:2000])
    except Exception as e:
        import traceback
        print(f"EXCEPTION: {e}")
        traceback.print_exc()


# ── Individual test functions ──────────────────────────────────────────────────

def test_list_sheets():
    result = excel_list_sheets(file_path=TEST_FILE)
    parsed = json.loads(result)
    assert isinstance(parsed, list), f"Expected list, got {type(parsed)}"
    assert "Sales" in parsed
    assert "Config" in parsed
    print("✓ excel_list_sheets OK:", parsed)
    return result


def test_get_sheet_size():
    # Sales sheet: 100 data rows + 1 header = 101 rows, 5 cols
    result = excel_get_sheet_size(file_path=TEST_FILE, sheet_name="Sales")
    parsed = json.loads(result)
    assert parsed["rows"] == 101, f"Expected 101 rows, got {parsed['rows']}"
    assert parsed["cols"] == 5,   f"Expected 5 cols, got {parsed['cols']}"
    assert parsed["cells"] == 101 * 5
    assert parsed["method"] == "xml_dimension_tag", f"Expected fast path, got {parsed['method']}"
    print(f"✓ excel_get_sheet_size OK — {parsed['rows']}×{parsed['cols']} via {parsed['method']}")
    return result


def test_get_sheet_size_forecast():
    # Forecast patch starts at B5, not A1 — dimension tag should still cover it
    result = excel_get_sheet_size(file_path=TEST_FILE, sheet_name="Forecast")
    parsed = json.loads(result)
    assert parsed["rows"] > 0 and parsed["cols"] > 0
    print(f"✓ Forecast size: {parsed['rows']}×{parsed['cols']} via {parsed['method']}")
    return result


def test_get_sheet_size_roundtrip():
    """Write a known-size sheet and verify size tool matches."""
    tmp = "/tmp/size_test.xlsx"
    # 4 data rows + 1 header = 5 rows, 3 cols
    md = "| A | B | C |\n|---|---|---|\n| 1 | 2 | 3 |\n| 4 | 5 | 6 |\n| 7 | 8 | 9 |\n| 10 | 11 | 12 |"
    excel_write_workbook(file_path=tmp, sheets_data={"Sheet1": md})
    result = excel_get_sheet_size(file_path=tmp, sheet_name="Sheet1")
    parsed = json.loads(result)
    assert parsed["rows"] == 5, f"Expected 5 rows, got {parsed['rows']}"
    assert parsed["cols"] == 3, f"Expected 3 cols, got {parsed['cols']}"
    print(f"✓ Size roundtrip OK — {parsed['rows']}×{parsed['cols']} via {parsed['method']}")
    return result


def test_get_sheet_full():
    result = excel_get_sheet_full(file_path=TEST_FILE, sheet_name="Config")
    assert "Setting" in result
    print("✓ excel_get_sheet_full OK")
    return result


def test_get_sheet_full_sparsity():
    result = excel_get_sheet_full(
        file_path=TEST_FILE, sheet_name="Config", to_drop_sparsity=2
    )
    assert "Setting" in result
    print("✓ excel_get_sheet_full with sparsity OK")
    return result


def test_patches_truncated():
    result = excel_get_sheet_patches_truncated(
        file_path=TEST_FILE, sheet_name="Sales"
    )
    parsed = json.loads(result)
    assert isinstance(parsed, dict), "Expected dict"
    assert len(parsed) > 0, "Expected at least one patch"
    for pid in parsed:
        assert "_P0" in pid or "_P1" in pid, f"Bad patch ID format: {pid}"
    print(f"✓ excel_get_sheet_patches_truncated: {len(parsed)} patch(es) — {list(parsed.keys())}")
    return result


def test_patches_truncated_top_n():
    result = excel_get_sheet_patches_truncated(
        file_path=TEST_FILE, sheet_name="Sales", top_n_patches=1
    )
    parsed = json.loads(result)
    assert len(parsed) == 1, f"Expected 1 patch, got {len(parsed)}"
    print("✓ top_n_patches=1 works")
    return result


def test_patches_by_id():
    # Step 1: get truncated patches to obtain IDs
    result1 = excel_get_sheet_patches_truncated(
        file_path=TEST_FILE, sheet_name="Sales"
    )
    patches = json.loads(result1)
    patch_id = list(patches.keys())[0]
    print(f"  Using patch ID: {patch_id}")

    # Step 2: retrieve full data for that patch
    result2 = excel_get_sheet_patches_by_id(
        file_path=TEST_FILE, sheet_name="Sales", patch_ids=[patch_id]
    )
    parsed = json.loads(result2)
    assert patch_id in parsed, f"Expected {patch_id} in result"
    # Full data should be larger than the truncated version
    assert len(parsed[patch_id]) > len(patches[patch_id])
    print(f"✓ excel_get_sheet_patches_by_id OK — full rows returned")
    return result2


def test_cell_ranges():
    result = excel_get_sheet_cell_ranges(
        file_path=TEST_FILE, sheet_name="Sales", cell_ranges=["A1:E5", "A1:B3"]
    )
    parsed = json.loads(result)
    assert "A1:E5" in parsed
    assert "A1:B3" in parsed
    print("✓ excel_get_sheet_cell_ranges OK")
    return result


def test_sparse_sheet():
    result = excel_get_sheet_patches_truncated(
        file_path=TEST_FILE, sheet_name="Sparse"
    )
    parsed = json.loads(result)
    assert len(parsed) == 2, f"Expected 2 patches for Sparse sheet, got {len(parsed)}"
    print(f"✓ Sparse sheet: {len(parsed)} patches — {list(parsed.keys())}")
    return result


def test_non_a1_patch():
    # Forecast data starts at B5, not A1
    result = excel_get_sheet_patches_truncated(
        file_path=TEST_FILE, sheet_name="Forecast"
    )
    parsed = json.loads(result)
    keys = list(parsed.keys())
    # Patch should start at B5, not A1
    assert any("B5" in k for k in keys), f"Expected B5 in patch IDs, got {keys}"
    print(f"✓ Non-A1 patch detected: {keys}")
    return result


def test_hybrid_mode():
    result = excel_get_sheet_full(
        file_path=TEST_FILE, sheet_name="Formulas", content="hybrid"
    )
    # Should contain at least one formula string
    assert "=SUM" in result or "=B2" in result, "Expected formula in hybrid output"
    print(f"✓ Hybrid mode OK — formula detected in output")
    return result


def test_hybrid_patches():
    result = excel_get_sheet_patches_truncated(
        file_path=TEST_FILE, sheet_name="Formulas", content="hybrid"
    )
    parsed = json.loads(result)
    combined = " ".join(parsed.values())
    assert "=SUM" in combined or "=B2" in combined, "Expected formula in hybrid patch output"
    print("✓ Hybrid mode with patches OK")
    return result


def test_write_workbook(tmp_path="/tmp/test_output_v2.xlsx"):
    sheets_data = {
        "Results": "| Name | Score | Grade |\n|---|---|---|\n| Alice | 95 | A |\n| Bob | 82 | B |\n| Carol | 71 | C |",
        "Summary": "| Metric | Value |\n|---|---|\n| Average | =AVERAGE(Results!B2:B4) |\n| Max | =MAX(Results!B2:B4) |",
    }
    result = excel_write_workbook(file_path=tmp_path, sheets_data=sheets_data)
    parsed = json.loads(result)
    assert parsed.get("status") == "success"
    assert os.path.exists(tmp_path)
    assert parsed["total_sheets"] == 2
    print(f"✓ excel_write_workbook OK → {tmp_path}")
    return result


def test_roundtrip():
    """Write then read back to verify data integrity."""
    tmp = "/tmp/roundtrip_v2.xlsx"
    sheets_data = {
        "Data": "| ID | Value |\n|---|---|\n| 1 | 100 |\n| 2 | 200 |\n| 3 | 300 |",
    }
    excel_write_workbook(file_path=tmp, sheets_data=sheets_data)
    result = excel_get_sheet_full(file_path=tmp, sheet_name="Data")
    assert "100" in result and "200" in result and "300" in result
    print("✓ Round-trip write → read OK")
    return result


# ── Runner ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Excel MCP Server — Tool Tests")
    print("=" * 60)

    if not os.path.exists(TEST_FILE):
        print(f"Test workbook not found: {TEST_FILE}")
        print("Run: python tests/create_test_excel.py")
        sys.exit(1)

    run("excel_list_sheets", test_list_sheets)
    run("excel_get_sheet_size (Sales)", test_get_sheet_size)
    run("excel_get_sheet_size (Forecast)", test_get_sheet_size_forecast)
    run("excel_get_sheet_size roundtrip", test_get_sheet_size_roundtrip)
    run("excel_get_sheet_full", test_get_sheet_full)
    run("excel_get_sheet_full + sparsity", test_get_sheet_full_sparsity)
    run("excel_get_sheet_patches_truncated", test_patches_truncated)
    run("excel_get_sheet_patches_truncated top_n=1", test_patches_truncated_top_n)
    run("excel_get_sheet_patches_by_id (2-step)", test_patches_by_id)
    run("excel_get_sheet_cell_ranges", test_cell_ranges)
    run("sparse sheet detection", test_sparse_sheet)
    run("non-A1 patch (Forecast at B5)", test_non_a1_patch)
    run("hybrid mode (full)", test_hybrid_mode)
    run("hybrid mode (patches)", test_hybrid_patches)
    run("excel_write_workbook", test_write_workbook)
    run("roundtrip write → read", test_roundtrip)

    print("\n" + "=" * 60)
    print("All tests completed.")
