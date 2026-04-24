"""Routing helpers (transport backend selection)."""

from excel_mcp.routing.tool_inventory import (
    MCP_TOOL_INVENTORY,
    ToolInventoryEntry,
    ToolKind,
    get_tool_kind,
)
from excel_mcp.routing.file_workbook_service import FileWorkbookService
from excel_mcp.routing.routed_dispatch import (
    execute_routed_workbook_operation,
    redact_workbook_path_for_logs,
)
from excel_mcp.routing.routing_backend import (
    RoutingBackend,
    WorkbookBackendResolution,
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
    "execute_routed_workbook_operation",
    "FileWorkbookService",
    "MCP_TOOL_INVENTORY",
    "StubWorkbookOpenInExcel",
    "redact_workbook_path_for_logs",
    "ROUTED_WORKBOOK_OPERATION_NAMES",
    "RoutingBackend",
    "RoutedWorkbookOperations",
    "ToolInventoryEntry",
    "ToolKind",
    "WorkbookBackendResolution",
    "WorkbookOpenInExcelPort",
    "WorkbookOperationMetadata",
    "WorkbookReadOperations",
    "WorkbookWriteOperations",
    "get_tool_kind",
]
