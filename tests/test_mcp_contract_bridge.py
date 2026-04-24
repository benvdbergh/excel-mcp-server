"""MCP tool name → contract operation name (Epic 5)."""

from __future__ import annotations

from excel_mcp.routing.mcp_contract_bridge import contract_operation_name_for_mcp_tool
from excel_mcp.routing.tool_inventory import MCP_TOOL_INVENTORY
from excel_mcp.routing.workbook_operation_contract import ROUTED_WORKBOOK_OPERATION_NAMES


def test_overrides() -> None:
    assert (
        contract_operation_name_for_mcp_tool("read_data_from_excel")
        == "read_range_with_metadata"
    )
    assert contract_operation_name_for_mcp_tool("apply_formula") == "apply_formula"


def test_every_inventory_tool_maps_to_contract_name() -> None:
    allowed = frozenset(ROUTED_WORKBOOK_OPERATION_NAMES)
    for name in MCP_TOOL_INVENTORY:
        op = contract_operation_name_for_mcp_tool(name)
        assert op in allowed, f"{name} -> {op} not in contract"
