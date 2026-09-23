from __future__ import annotations

import logging
from typing import Any, Callable, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_mcp.exceptions import (
    CalculationError,
    DataError,
    FormattingError,
    ValidationError,
)
from excel_mcp.fileio.service import FileWorkbookService
from excel_mcp.routing.routing_errors import (
    ComExecutionNotImplementedError,
    ComRoutingError,
)

logger = logging.getLogger("excel-mcp")


def register_write(
    mcp: FastMCP,
    *,
    file_service: FileWorkbookService,
    workbook_dispatch: Callable[..., str],
    com_dispatch: Callable[..., Callable[[str], str] | None],
) -> Dict[str, Callable]:
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Apply Formula",
            destructiveHint=True,
        ),
    )
    def apply_formula(
        filepath: str,
        sheet_name: str,
        cell: str,
        formula: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """
        Apply Excel formula to cell.
        Excel formula will write to cell with verification.
        """
        try:
            return workbook_dispatch(
                "apply_formula",
                filepath,
                workbook_transport,
                lambda fp: file_service.apply_formula(
                    fp, sheet_name, cell, formula
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.apply_formula(fp, sheet_name, cell, formula)
                ),
            )
        except (ValidationError, CalculationError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error applying formula: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Range",
            destructiveHint=True,
        ),
    )
    def format_range(
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
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Apply formatting to a range of cells."""
        try:
            return workbook_dispatch(
                "format_range",
                filepath,
                workbook_transport,
                lambda fp: file_service.format_range(
                    fp,
                    sheet_name,
                    start_cell,
                    end_cell,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    font_size=font_size,
                    font_color=font_color,
                    bg_color=bg_color,
                    border_style=border_style,
                    border_color=border_color,
                    number_format=number_format,
                    alignment=alignment,
                    wrap_text=wrap_text,
                    merge_cells=merge_cells,
                    protection=protection,
                    conditional_format=conditional_format,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.format_range(
                        fp,
                        sheet_name,
                        start_cell,
                        end_cell,
                        bold=bold,
                        italic=italic,
                        underline=underline,
                        font_size=font_size,
                        font_color=font_color,
                        bg_color=bg_color,
                        border_style=border_style,
                        border_color=border_color,
                        number_format=number_format,
                        alignment=alignment,
                        wrap_text=wrap_text,
                        merge_cells=merge_cells,
                        protection=protection,
                        conditional_format=conditional_format,
                    )
                ),
            )
        except (ValidationError, FormattingError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error formatting range: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Write Data to Excel",
            destructiveHint=True,
        ),
    )
    def write_data_to_excel(
        filepath: str,
        sheet_name: str,
        data: List[List],
        start_cell: str = "A1",
        workbook_transport: Optional[str] = None,
    ) -> str:
        """
        Write data to Excel worksheet.
        Excel formula will write to cell without any verification.

        PARAMETERS:  
        filepath: Path to Excel file
        sheet_name: Name of worksheet to write to
        data: List of lists containing data to write to the worksheet, sublists are assumed to be rows
        start_cell: Cell to start writing to, default is "A1"
  
        """
        try:
            return workbook_dispatch(
                "write_data_to_excel",
                filepath,
                workbook_transport,
                lambda fp: file_service.write_cell_grid(
                    fp, sheet_name, data, start_cell
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.write_cell_grid(fp, sheet_name, data, start_cell)
                ),
            )
        except (ValidationError, DataError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error writing data: {e}")
            raise


    return {
        "apply_formula": apply_formula,
        "format_range": format_range,
        "write_data_to_excel": write_data_to_excel,
    }
