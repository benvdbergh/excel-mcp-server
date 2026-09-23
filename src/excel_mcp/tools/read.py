from __future__ import annotations

import logging
from typing import Callable, Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_mcp.exceptions import (
    CalculationError,
    ValidationError,
    WorkbookError,
)
from excel_mcp.fileio.service import FileWorkbookService
from excel_mcp.routing.routing_errors import (
    ComExecutionNotImplementedError,
    ComRoutingError,
)
from excel_mcp.value_mode import validate_metadata_mode, validate_value_mode

logger = logging.getLogger("excel-mcp")


def register_read(
    mcp: FastMCP,
    *,
    file_service: FileWorkbookService,
    workbook_dispatch: Callable[..., str],
    com_dispatch: Callable[..., Callable[[str], str] | None],
) -> Dict[str, Callable]:
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Read Data from Excel",
            readOnlyHint=True,
        ),
    )
    def read_data_from_excel(
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        workbook_transport: Optional[str] = None,
        include_routing_metadata: bool = False,
        value_mode: str = "value",
        metadata_mode: str = "full",
    ) -> str:
        """
        Read data from Excel worksheet with cell metadata including validation rules.

        Args:
            filepath: Absolute local path, or for COM/cloud workbooks the exact ``https://…``
                SharePoint-style URL matching Excel ``Workbook.FullName`` (ADR 0006). Call
                ``excel_list_open_workbooks`` and copy a ``full_name`` when unsure (ADR 0009).
            sheet_name: Name of worksheet
            start_cell: Starting cell (default A1)
            end_cell: Ending cell (optional, auto-expands if not provided)
            workbook_transport: Optional execution mode ``auto`` | ``file`` | ``com`` (ADR 0001;
                default from ``EXCEL_MCP_TRANSPORT`` env). Not the MCP wire transport.
                With ``auto``/``com``, reads the live Excel grid (ADR 0008), not necessarily
                the last saved file on disk.
            include_routing_metadata: When true, wrap JSON in ADR 0010 envelope with
                _meta (workbook_transport, workbook_backend, routing_reason, duration_ms)
                and optional warnings (e.g. file_backend_formula_not_evaluated on .xlsm).
            value_mode: ``value`` (raw cell values) or ``text`` (display text; COM uses
                Range.Text; file backend is best-effort formatted string)
            metadata_mode: ``full`` (per-cell validation metadata, default) or ``compact``
                (omit per-cell validation to shrink large-range payloads)

        File backend (openpyxl) does not evaluate formulas; ``.xlsm`` reads may return
        null for formula cells. Prefer ``workbook_transport=auto`` or ``com`` when the
        workbook is open in Excel. With ``include_routing_metadata=true``, that limit
        surfaces as warning code ``file_backend_formula_not_evaluated``.

        Returns:
            JSON string containing structured cell data with validation metadata.
            Each cell includes: address, value, row, column, and validation info (if any).
            Root JSON includes ``value_mode`` and ``metadata_mode``. When include_routing_metadata is true, the
            payload is wrapped per ADR 0010.
        """
        try:
            validate_value_mode(value_mode)
            validate_metadata_mode(metadata_mode)
            routing_warnings: list = []

            def _file_read(fp: str) -> str:
                op_meta = (
                    {"_response_warnings": routing_warnings}
                    if include_routing_metadata
                    else None
                )
                return file_service.read_range_with_metadata(
                    fp,
                    sheet_name,
                    start_cell,
                    end_cell,
                    value_mode=value_mode,
                    metadata_mode=metadata_mode,
                    operation_metadata=op_meta,
                )

            return workbook_dispatch(
                "read_data_from_excel",
                filepath,
                workbook_transport,
                _file_read,
                com_do_op=com_dispatch(
                    lambda c, fp: c.read_range_with_metadata(
                        fp,
                        sheet_name,
                        start_cell,
                        end_cell,
                        value_mode=value_mode,
                        metadata_mode=metadata_mode,
                    )
                ),
                include_routing_metadata=include_routing_metadata,
                response_warnings=routing_warnings if include_routing_metadata else None,
            )
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error reading data: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Export Worksheet Table",
            readOnlyHint=True,
        ),
    )
    def export_worksheet_table(
        filepath: str,
        sheet_name: str,
        start_cell: str = "A1",
        end_cell: Optional[str] = None,
        max_rows: int = 10000,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """
        Export worksheet data as a compact table (header row + data rows).

        Args:
            filepath: Path to Excel file, or exact https SharePoint-style URL matching
                Excel Workbook.FullName when using COM (see excel_list_open_workbooks).
            sheet_name: Name of worksheet
            start_cell: Starting cell (default A1); used range when end_cell omitted
            end_cell: Optional ending cell
            max_rows: Maximum data rows returned (default 10000); sets truncated when exceeded
            workbook_transport: Workbook execution mode (auto, file, com)

        Returns:
            JSON string with sheet_name, range, headers, rows, row_count, truncated, max_rows.
            First row of the range is treated as column headers; remaining rows are data.
        """
        try:
            return workbook_dispatch(
                "export_worksheet_table",
                filepath,
                workbook_transport,
                lambda fp: file_service.export_worksheet_table(
                    fp,
                    sheet_name,
                    start_cell,
                    end_cell,
                    max_rows,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.export_worksheet_table(
                        fp,
                        sheet_name,
                        start_cell,
                        end_cell,
                        max_rows,
                    )
                ),
            )
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error exporting worksheet table: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Workbook Metadata",
            readOnlyHint=True,
        ),
    )
    def get_workbook_metadata(
        filepath: str,
        include_ranges: bool = False,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Get metadata about workbook including sheets, ranges, etc."""
        try:
            return workbook_dispatch(
                "get_workbook_metadata",
                filepath,
                workbook_transport,
                lambda fp: file_service.workbook_metadata(
                    fp, include_ranges=include_ranges
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.workbook_metadata(
                        fp, include_ranges=include_ranges
                    )
                ),
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error getting workbook metadata: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Merged Cells",
            readOnlyHint=True,
        ),
    )
    def get_merged_cells(
        filepath: str,
        sheet_name: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Get merged cells in a worksheet."""
        try:
            return workbook_dispatch(
                "get_merged_cells",
                filepath,
                workbook_transport,
                lambda fp: file_service.read_merged_cell_ranges(
                    fp, sheet_name
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.read_merged_cell_ranges(fp, sheet_name)
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error getting merged cells: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Get Data Validation Info",
            readOnlyHint=True,
        ),
    )
    def get_data_validation_info(
        filepath: str,
        sheet_name: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """
        Get all data validation rules in a worksheet.
    
        This tool helps identify which cell ranges have validation rules
        and what types of validation are applied.
    
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
        
        Returns:
            JSON string containing all validation rules in the worksheet
        """
        try:
            return workbook_dispatch(
                "get_data_validation_info",
                filepath,
                workbook_transport,
                lambda fp: file_service.read_worksheet_data_validation(
                    fp, sheet_name
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.read_worksheet_data_validation(fp, sheet_name)
                ),
            )
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error getting validation info: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Validate Excel Range",
            readOnlyHint=True,
        ),
    )
    def validate_excel_range(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Validate if a range exists and is properly formatted."""
        try:
            return workbook_dispatch(
                "validate_excel_range",
                filepath,
                workbook_transport,
                lambda fp: file_service.validate_sheet_range(
                    fp, sheet_name, start_cell, end_cell
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.validate_sheet_range(
                        fp, sheet_name, start_cell, end_cell
                    )
                ),
            )
        except ValidationError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error validating range: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Validate Formula Syntax",
            readOnlyHint=True,
        ),
    )
    def validate_formula_syntax(
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Validate Excel formula syntax without applying it."""
        try:
            return workbook_dispatch(
                "validate_formula_syntax",
                filepath,
                workbook_transport,
                lambda fp: file_service.validate_formula_syntax(
                    fp, sheet_name, cell, formula
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.validate_formula_syntax(
                        fp, sheet_name, cell, formula
                    )
                ),
            )
        except (ValidationError, CalculationError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error validating formula: {e}")
            raise


    return {
        "read_data_from_excel": read_data_from_excel,
        "export_worksheet_table": export_worksheet_table,
        "get_workbook_metadata": get_workbook_metadata,
        "get_merged_cells": get_merged_cells,
        "get_data_validation_info": get_data_validation_info,
        "validate_excel_range": validate_excel_range,
        "validate_formula_syntax": validate_formula_syntax,
    }
