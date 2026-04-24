"""Map registered MCP tool function names to ``RoutedWorkbookOperations`` method names."""

from __future__ import annotations

# MCP handler name (FastMCP / ``server.py`` function name) -> contract method name.
_MCP_TO_CONTRACT: dict[str, str] = {
    "read_data_from_excel": "read_range_with_metadata",
    "write_data_to_excel": "write_cell_grid",
    "get_workbook_metadata": "workbook_metadata",
    "get_merged_cells": "read_merged_cell_ranges",
    "get_data_validation_info": "read_worksheet_data_validation",
    "validate_excel_range": "validate_sheet_range",
    "create_chart": "create_chart_in_sheet",
    "create_pivot_table": "create_pivot_table_in_sheet",
    "create_table": "create_excel_table",
    "copy_range": "copy_cell_range",
    "delete_range": "delete_cell_range",
}


def contract_operation_name_for_mcp_tool(mcp_tool_name: str) -> str:
    """Return ``RoutedWorkbookOperations`` method name for the given MCP tool."""
    return _MCP_TO_CONTRACT.get(mcp_tool_name, mcp_tool_name)
