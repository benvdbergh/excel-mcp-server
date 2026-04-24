"""Authoritative MCP tool inventory for transport/routing (single source of truth).

Cites:
- PRD: ``docs/specs/PRD-excel-mcp-transport-routing.md``
- Blueprint: ``docs/excel-mcp-fork-com-vs-file-routing.md``
- ADR 0003 (file-backed reads; read defaults align with file backend)
- ADR 0004 (chart/pivot v1 exception: tool-forced file backend)
- ``docs/architecture/pre-fork-architecture.md`` (25 tools)

Classifications:
- ``READ``: safe read-style tools; default routing aligns with file-backed reads.
- ``WRITE``: mutating tools.
- ``V1_FILE_FORCED``: v1 routing must use the file backend regardless of
  auto→COM (ADR 0004 policy 1).
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
