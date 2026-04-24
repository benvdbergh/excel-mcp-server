"""Routing helpers (transport backend selection)."""

from excel_mcp.routing.tool_inventory import (
    MCP_TOOL_INVENTORY,
    ToolInventoryEntry,
    ToolKind,
    get_tool_kind,
)
from excel_mcp.routing.file_workbook_service import FileWorkbookService
from excel_mcp.routing.mcp_contract_bridge import contract_operation_name_for_mcp_tool
from excel_mcp.routing.routed_dispatch import (
    execute_routed_workbook_operation,
    redact_workbook_path_for_logs,
)
from excel_mcp.routing.routing_backend import (
    RoutingBackend,
    WorkbookBackendResolution,
    WorkbookTransport,
)
from excel_mcp.routing.routing_env import (
    EXCEL_MCP_COM_ALLOW_FILE_FALLBACK,
    EXCEL_MCP_COM_STRICT,
    EXCEL_MCP_SAVE_AFTER_WRITE_DEFAULT,
    EXCEL_MCP_TRANSPORT,
    effective_com_strict,
    effective_save_after_write,
    read_com_allow_file_fallback,
    read_com_strict,
    read_save_after_write_default,
    read_workbook_transport,
    resolve_workbook_transport,
)
from excel_mcp.routing.routing_errors import ComExecutionNotImplementedError, ComRoutingError
from excel_mcp.routing.workbook_open_detection import (
    StubWorkbookOpenInExcel,
    WorkbookOpenInExcelPort,
)
from excel_mcp.routing.workbook_operation_contract import (
    ROUTED_WORKBOOK_OPERATION_NAMES,
    RoutedWorkbookOperations,
    WorkbookOperationMetadata,
    WorkbookReadOperations,
    WorkbookWriteOperations,
)

__all__ = [
    "ComExecutionNotImplementedError",
    "ComRoutingError",
    "EXCEL_MCP_COM_ALLOW_FILE_FALLBACK",
    "EXCEL_MCP_COM_STRICT",
    "EXCEL_MCP_SAVE_AFTER_WRITE_DEFAULT",
    "EXCEL_MCP_TRANSPORT",
    "contract_operation_name_for_mcp_tool",
    "effective_com_strict",
    "effective_save_after_write",
    "execute_routed_workbook_operation",
    "FileWorkbookService",
    "MCP_TOOL_INVENTORY",
    "StubWorkbookOpenInExcel",
    "redact_workbook_path_for_logs",
    "ROUTED_WORKBOOK_OPERATION_NAMES",
    "RoutingBackend",
    "RoutedWorkbookOperations",
    "read_com_allow_file_fallback",
    "read_com_strict",
    "read_save_after_write_default",
    "read_workbook_transport",
    "resolve_workbook_transport",
    "ToolInventoryEntry",
    "ToolKind",
    "WorkbookBackendResolution",
    "WorkbookOpenInExcelPort",
    "WorkbookOperationMetadata",
    "WorkbookReadOperations",
    "WorkbookTransport",
    "WorkbookWriteOperations",
    "get_tool_kind",
]
