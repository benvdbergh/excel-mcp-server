"""Authoritative MCP tool inventory for transport/routing (single source of truth).

Cites:
- PRD: ``docs/specs/PRD-excel-mcp-transport-routing.md``
- Blueprint: ``docs/excel-mcp-fork-com-vs-file-routing.md``
- ADR 0008 (COM-first default; read-class tools use the same transport branches as writes)
- ADR 0004 (chart/pivot v1 exception: tool-forced file backend)
- ``docs/architecture/pre-fork-architecture.md`` (25 tools)

Classifications:
- ``READ``: safe read-style tools; default routing is COM-first when ``auto``/``com`` and COM is viable (same as ``WRITE``).
- ``WRITE``: mutating tools.
- ``V1_FILE_FORCED``: v1 routing must use the file backend regardless of
  auto→COM (ADR 0004 policy 1).
- ``SESSION``: Excel host lifecycle (``Workbooks.Open`` / close); does not use
  :meth:`RoutingBackend.resolve_workbook_backend` — COM entry points only (ADR 0008).
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from types import MappingProxyType
from typing import Final, Mapping


class ToolKind(str, Enum):
    """MCP tool classification for routing."""

    READ = "read"
    WRITE = "write"
    V1_FILE_FORCED = "v1_file_forced"
    SESSION = "session"


@dataclass(frozen=True, slots=True)
class ToolInventoryEntry:
    """One registered MCP tool (FastMCP: Python function name) and its routing kind."""

    kind: ToolKind
    notes: str = ""


_RAW_INVENTORY: dict[str, ToolInventoryEntry] = {
    "apply_formula": ToolInventoryEntry(ToolKind.WRITE),
    "validate_formula_syntax": ToolInventoryEntry(ToolKind.READ),
    "format_range": ToolInventoryEntry(ToolKind.WRITE),
    "read_data_from_excel": ToolInventoryEntry(ToolKind.READ),
    "write_data_to_excel": ToolInventoryEntry(ToolKind.WRITE),
    "create_workbook": ToolInventoryEntry(ToolKind.WRITE),
    "save_workbook": ToolInventoryEntry(
        ToolKind.WRITE,
        notes="ADR 0003: explicit persist before file reads when using COM without per-write save.",
    ),
    "create_worksheet": ToolInventoryEntry(ToolKind.WRITE),
    "create_chart": ToolInventoryEntry(
        ToolKind.V1_FILE_FORCED,
        notes="ADR 0004: chart uses file backend in v1 regardless of auto→COM.",
    ),
    "create_pivot_table": ToolInventoryEntry(
        ToolKind.V1_FILE_FORCED,
        notes="ADR 0004: pivot uses file backend in v1 regardless of auto→COM.",
    ),
    "create_table": ToolInventoryEntry(ToolKind.WRITE),
    "copy_worksheet": ToolInventoryEntry(ToolKind.WRITE),
    "delete_worksheet": ToolInventoryEntry(ToolKind.WRITE),
    "rename_worksheet": ToolInventoryEntry(ToolKind.WRITE),
    "get_workbook_metadata": ToolInventoryEntry(ToolKind.READ),
    "merge_cells": ToolInventoryEntry(ToolKind.WRITE),
    "unmerge_cells": ToolInventoryEntry(ToolKind.WRITE),
    "get_merged_cells": ToolInventoryEntry(ToolKind.READ),
    "copy_range": ToolInventoryEntry(ToolKind.WRITE),
    "delete_range": ToolInventoryEntry(ToolKind.WRITE),
    "validate_excel_range": ToolInventoryEntry(ToolKind.READ),
    "get_data_validation_info": ToolInventoryEntry(ToolKind.READ),
    "insert_rows": ToolInventoryEntry(ToolKind.WRITE),
    "insert_columns": ToolInventoryEntry(ToolKind.WRITE),
    "delete_sheet_rows": ToolInventoryEntry(ToolKind.WRITE),
    "delete_sheet_columns": ToolInventoryEntry(ToolKind.WRITE),
    "excel_open_workbook": ToolInventoryEntry(
        ToolKind.SESSION,
        notes="ADR 0008: Workbooks.Open; not resolved via file/openpyxl backend.",
    ),
    "excel_close_workbook": ToolInventoryEntry(
        ToolKind.SESSION,
        notes="ADR 0008: close workbook in Excel host with optional save.",
    ),
    "excel_list_open_workbooks": ToolInventoryEntry(
        ToolKind.SESSION,
        notes="ADR 0009: enumerate Application.Workbooks; no filepath routing.",
    ),
}

MCP_TOOL_INVENTORY: Final[Mapping[str, ToolInventoryEntry]] = MappingProxyType(
    _RAW_INVENTORY
)


def get_tool_kind(name: str) -> ToolKind:
    """Return routing kind for a registered MCP tool name.

    Args:
        name: FastMCP tool name (Python handler function name).

    Raises:
        KeyError: If ``name`` is not in the inventory.
    """
    return MCP_TOOL_INVENTORY[name].kind
