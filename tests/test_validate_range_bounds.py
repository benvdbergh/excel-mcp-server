"""Excel-limit range bounds (used range is not the ceiling)."""

from __future__ import annotations

import os
import sys

from openpyxl import Workbook

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.validation import validate_range_bounds  # noqa: E402


def test_validate_range_bounds_uses_excel_limits() -> None:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1
    ok, _ = validate_range_bounds(ws, 100, 1)
    assert ok is True
    bad_row, _ = validate_range_bounds(ws, 1_048_577, 1)
    assert bad_row is False
    bad_col, _ = validate_range_bounds(ws, 1, 16_385)
    assert bad_col is False
