"""Shared worksheet-table export payload helpers (openpyxl- and COM-free)."""

from __future__ import annotations

from typing import Any, Dict, List

from excel_mcp.exceptions import DataError

DEFAULT_EXPORT_MAX_ROWS = 10000


def normalize_export_max_rows(max_rows: int | None) -> int:
    if max_rows is None:
        return DEFAULT_EXPORT_MAX_ROWS
    if max_rows < 1:
        raise DataError("max_rows must be a positive integer")
    return max_rows


def export_read_end_row(start_row: int, end_row: int, cap: int) -> int:
    """Last row to read when exporting: header row plus up to ``cap`` data rows."""
    if end_row < start_row:
        return start_row
    total_data_rows = end_row - start_row
    return start_row + min(total_data_rows, cap)


def build_worksheet_table_payload(
    sheet_name: str,
    range_str: str,
    matrix: List[List[Any]],
    *,
    max_rows: int = DEFAULT_EXPORT_MAX_ROWS,
    total_data_rows: int | None = None,
) -> Dict[str, Any]:
    """Build compact table JSON from a rectangular cell matrix (first row = headers)."""
    cap = normalize_export_max_rows(max_rows)
    if not matrix:
        total = 0 if total_data_rows is None else max(0, total_data_rows)
        return {
            "sheet_name": sheet_name,
            "range": range_str,
            "headers": [],
            "rows": [],
            "row_count": total,
            "truncated": total > cap,
            "max_rows": cap,
        }
    headers = list(matrix[0])
    data_rows = [list(row) for row in matrix[1:]]
    total = total_data_rows if total_data_rows is not None else len(data_rows)
    truncated = total > cap
    return {
        "sheet_name": sheet_name,
        "range": range_str,
        "headers": headers,
        "rows": data_rows[:cap],
        "row_count": total,
        "truncated": truncated,
        "max_rows": cap,
    }
