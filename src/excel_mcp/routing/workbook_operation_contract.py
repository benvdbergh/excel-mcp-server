"""Shared workbook operation contract for workbook backends (FR-4, FR-5).

Operation-oriented names are **not** MCP wire transport names (``stdio`` /
``sse`` / ``streamable-http``); see ADR 0001. They also differ from FastMCP
Python handler names where a clearer domain verb exists.

Traceability:
- PRD ``docs/specs/PRD-excel-mcp-transport-routing.md`` (FR-4, FR-5): narrow
  shared API for ``FileWorkbookService`` and ``ComWorkbookService``.
- ``docs/architecture/target-architecture.md`` §6 (``ComWorkbookService``):
  same method surface as the file façade for routed operations.
- Epic-3 façade: handlers delegate path resolution + routing, then call this
  contract on the selected backend.

``operation_metadata`` (keyword-only) carries ``tool_kind`` / routing hints for
default file-backed reads and future **opt-in COM reads** (ADR 0003). It must
not be confused with MCP client connection mode.

This module intentionally does **not** import ``excel_mcp.server`` (avoid
cycles and keep routing free of handler wiring).
"""

from __future__ import annotations

from typing import Any, Dict, List, Mapping, Optional, Protocol, TypedDict


class WorkbookOperationMetadata(TypedDict, total=False):
    """Routing context attached by ``RoutingBackend`` / façade (ADR 0003).

    All keys optional. Prefer value strings aligned with
    ``excel_mcp.routing.tool_inventory.ToolKind`` (``read``, ``write``,
    ``v1_file_forced``) for ``tool_kind`` when populated.

    Attributes:
        tool_kind: Read vs write vs v1-file-forced classification for routing.
        mcp_tool_name: Registered MCP tool function name for logs (wire name).
        operation_kind: Finer-grained label for metrics/logs (follow-up).
        com_read_opt_in: Phase-2 opt-in for COM-backed reads per call/env.
    """

    tool_kind: str
    mcp_tool_name: str
    operation_kind: str
    com_read_opt_in: bool


class WorkbookReadOperations(Protocol):
    """Read-class workbook operations (default file-backed; ADR 0003)."""

    def read_range_with_metadata(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        preview_only: bool = False,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Range read including per-cell validation metadata (MCP: ``read_data_from_excel``)."""
        ...

    def workbook_metadata(
        self,
        filepath: str,
        include_ranges: bool = False,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Workbook structure / info (MCP: ``get_workbook_metadata``)."""
        ...

    def read_merged_cell_ranges(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """List merged ranges on a sheet (MCP: ``get_merged_cells``)."""
        ...

    def read_worksheet_data_validation(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Worksheet-level validation rules (MCP: ``get_data_validation_info``)."""
        ...

    def validate_sheet_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Structural range validation (MCP: ``validate_excel_range``)."""
        ...

    def validate_formula_syntax(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Syntax check without mutating the cell (MCP: ``validate_formula_syntax``)."""
        ...


class WorkbookWriteOperations(Protocol):
    """Write-class and lifecycle operations routed to file or COM."""

    def apply_formula(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Validate then write a formula to a single cell."""
        ...

    def format_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: Optional[int] = None,
        font_color: Optional[str] = None,
        bg_color: Optional[str] = None,
        border_style: Optional[str] = None,
        border_color: Optional[str] = None,
        number_format: Optional[str] = None,
        alignment: Optional[str] = None,
        wrap_text: bool = False,
        merge_cells: bool = False,
        protection: Optional[Dict[str, Any]] = None,
        conditional_format: Optional[Dict[str, Any]] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Formatting and optional merge-in-pass for a rectangular range."""
        ...

    def write_cell_grid(
        self,
        filepath: str,
        sheet_name: str,
        data: List[List[Any]],
        start_cell: str = "A1",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Write a 2D grid of values starting at ``start_cell``."""
        ...

    def create_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Create a new workbook file at ``filepath``."""
        ...

    def create_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Add a worksheet to an existing workbook."""
        ...

    def create_chart_in_sheet(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        chart_type: str,
        target_cell: str,
        title: str = "",
        x_axis: str = "",
        y_axis: str = "",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Insert a chart from a data range (``excel_mcp.chart`` implementation).

        **v1 routing exception (ADR 0004):** chart creation is tool-forced to
        the **file** backend for deterministic openpyxl behavior; operation name
        matches the file stack so ``ComWorkbookService`` can defer or no-op
        until COM parity exists.
        """
        ...

    def create_pivot_table_in_sheet(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        rows: List[str],
        values: List[str],
        columns: Optional[List[str]] = None,
        agg_func: str = "mean",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Create pivot-style output (``excel_mcp.pivot``; not native COM pivot in v1).

        **v1 routing exception (ADR 0004):** same as chart — file-forced in v1;
        name aligned with file implementation for future COM adapter.
        """
        ...

    def create_excel_table(
        self,
        filepath: str,
        sheet_name: str,
        data_range: str,
        table_name: Optional[str] = None,
        table_style: str = "TableStyleMedium9",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Create a native ListObject / table over a range."""
        ...

    def copy_worksheet(
        self,
        filepath: str,
        source_sheet: str,
        target_sheet: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Duplicate a sheet inside the same workbook."""
        ...

    def delete_worksheet(
        self,
        filepath: str,
        sheet_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Remove a worksheet."""
        ...

    def rename_worksheet(
        self,
        filepath: str,
        old_name: str,
        new_name: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Rename a worksheet."""
        ...

    def merge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Merge a rectangular range."""
        ...

    def unmerge_cells(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Unmerge a rectangular range."""
        ...

    def copy_cell_range(
        self,
        filepath: str,
        sheet_name: str,
        source_start: str,
        source_end: str,
        target_start: str,
        target_sheet: Optional[str] = None,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Copy values/formulas from one range to another (optionally cross-sheet)."""
        ...

    def delete_cell_range(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        shift_direction: str = "up",
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Delete a range and shift remaining cells."""
        ...

    def insert_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Insert blank rows."""
        ...

    def insert_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Insert blank columns (1-based column index; follow-up: confirm vs Excel UI)."""
        ...

    def delete_sheet_rows(
        self,
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Delete rows from a worksheet."""
        ...

    def delete_sheet_columns(
        self,
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Delete columns from a worksheet."""
        ...

    def save_workbook(
        self,
        filepath: str,
        *,
        operation_metadata: Optional[Mapping[str, Any]] = None,
    ) -> str:
        """Persist host/disk workbook state (**future MCP tool**; ADR 0003).

        Not yet exposed as a FastMCP handler in the baseline inventory; kept on
        the contract so ``RoutingBackend`` and both backends share one surface
        when the tool lands.
        """
        ...


class RoutedWorkbookOperations(WorkbookReadOperations, WorkbookWriteOperations, Protocol):
    """Full routed surface implemented by ``FileWorkbookService`` / ``ComWorkbookService``."""

    ...


ROUTED_WORKBOOK_OPERATION_NAMES: tuple[str, ...] = (
    "read_range_with_metadata",
    "workbook_metadata",
    "read_merged_cell_ranges",
    "read_worksheet_data_validation",
    "validate_sheet_range",
    "validate_formula_syntax",
    "apply_formula",
    "format_range",
    "write_cell_grid",
    "create_workbook",
    "create_worksheet",
    "create_chart_in_sheet",
    "create_pivot_table_in_sheet",
    "create_excel_table",
    "copy_worksheet",
    "delete_worksheet",
    "rename_worksheet",
    "merge_cells",
    "unmerge_cells",
    "copy_cell_range",
    "delete_cell_range",
    "insert_rows",
    "insert_columns",
    "delete_sheet_rows",
    "delete_sheet_columns",
    "save_workbook",
)

__all__ = [
    "ROUTED_WORKBOOK_OPERATION_NAMES",
    "RoutedWorkbookOperations",
    "WorkbookOperationMetadata",
    "WorkbookReadOperations",
    "WorkbookWriteOperations",
]
