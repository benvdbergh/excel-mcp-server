from __future__ import annotations

import logging
from typing import Callable, Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_mcp.exceptions import SheetError, ValidationError
from excel_mcp.fileio.service import FileWorkbookService
from excel_mcp.routing.routing_errors import (
    ComExecutionNotImplementedError,
    ComRoutingError,
)

logger = logging.getLogger("excel-mcp")


def register_sheet(
    mcp: FastMCP,
    *,
    file_service: FileWorkbookService,
    workbook_dispatch: Callable[..., str],
    com_dispatch: Callable[..., Callable[[str], str] | None],
) -> Dict[str, Callable]:
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Worksheet",
            destructiveHint=True,
        ),
    )
    def create_worksheet(
        filepath: str,
        sheet_name: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Create new worksheet in workbook."""
        try:
            return workbook_dispatch(
                "create_worksheet",
                filepath,
                workbook_transport,
                lambda fp: file_service.create_worksheet(fp, sheet_name),
                com_do_op=com_dispatch(
                    lambda c, fp: c.create_worksheet(fp, sheet_name)
                ),
            )
        except (ValidationError, WorkbookError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error creating worksheet: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Copy Worksheet",
            destructiveHint=True,
        ),
    )
    def copy_worksheet(
        filepath: str,
        source_sheet: str,
        target_sheet: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Copy worksheet within workbook."""
        try:
            return workbook_dispatch(
                "copy_worksheet",
                filepath,
                workbook_transport,
                lambda fp: file_service.copy_worksheet(
                    fp, source_sheet, target_sheet
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.copy_worksheet(fp, source_sheet, target_sheet)
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error copying worksheet: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Worksheet",
            destructiveHint=True,
        ),
    )
    def delete_worksheet(
        filepath: str,
        sheet_name: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Delete worksheet from workbook."""
        try:
            return workbook_dispatch(
                "delete_worksheet",
                filepath,
                workbook_transport,
                lambda fp: file_service.delete_worksheet(fp, sheet_name),
                com_do_op=com_dispatch(
                    lambda c, fp: c.delete_worksheet(fp, sheet_name)
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error deleting worksheet: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Rename Worksheet",
            destructiveHint=True,
        ),
    )
    def rename_worksheet(
        filepath: str,
        old_name: str,
        new_name: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Rename worksheet in workbook."""
        try:
            return workbook_dispatch(
                "rename_worksheet",
                filepath,
                workbook_transport,
                lambda fp: file_service.rename_worksheet(
                    fp, old_name, new_name
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.rename_worksheet(fp, old_name, new_name)
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error renaming worksheet: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Merge Cells",
            destructiveHint=True,
        ),
    )
    def merge_cells(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Merge a range of cells."""
        try:
            return workbook_dispatch(
                "merge_cells",
                filepath,
                workbook_transport,
                lambda fp: file_service.merge_cells(
                    fp, sheet_name, start_cell, end_cell
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.merge_cells(
                        fp, sheet_name, start_cell, end_cell
                    )
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error merging cells: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Unmerge Cells",
            destructiveHint=True,
        ),
    )
    def unmerge_cells(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Unmerge a range of cells."""
        try:
            return workbook_dispatch(
                "unmerge_cells",
                filepath,
                workbook_transport,
                lambda fp: file_service.unmerge_cells(
                    fp, sheet_name, start_cell, end_cell
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.unmerge_cells(
                        fp, sheet_name, start_cell, end_cell
                    )
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error unmerging cells: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Copy Range",
            destructiveHint=True,
        ),
    )
    def copy_range(
        filepath: str,
        sheet_name: str,
        source_start: str,
        source_end: str,
        target_start: str,
        target_sheet: Optional[str] = None,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Copy a range of cells to another location."""
        try:
            return workbook_dispatch(
                "copy_range",
                filepath,
                workbook_transport,
                lambda fp: file_service.copy_cell_range(
                    fp,
                    sheet_name,
                    source_start,
                    source_end,
                    target_start,
                    target_sheet=target_sheet,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.copy_cell_range(
                        fp,
                        sheet_name,
                        source_start,
                        source_end,
                        target_start,
                        target_sheet=target_sheet,
                    )
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error copying range: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Range",
            destructiveHint=True,
        ),
    )
    def delete_range(
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: str,
        shift_direction: str = "up",
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Delete a range of cells and shift remaining cells."""
        try:
            return workbook_dispatch(
                "delete_range",
                filepath,
                workbook_transport,
                lambda fp: file_service.delete_cell_range(
                    fp,
                    sheet_name,
                    start_cell,
                    end_cell,
                    shift_direction,
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.delete_cell_range(
                        fp,
                        sheet_name,
                        start_cell,
                        end_cell,
                        shift_direction,
                    )
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error deleting range: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert Rows",
            destructiveHint=True,
        ),
    )
    def insert_rows(
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Insert one or more rows starting at the specified row."""
        try:
            return workbook_dispatch(
                "insert_rows",
                filepath,
                workbook_transport,
                lambda fp: file_service.insert_rows(
                    fp, sheet_name, start_row, count
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.insert_rows(fp, sheet_name, start_row, count)
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error inserting rows: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Insert Columns",
            destructiveHint=True,
        ),
    )
    def insert_columns(
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Insert one or more columns starting at the specified column."""
        try:
            return workbook_dispatch(
                "insert_columns",
                filepath,
                workbook_transport,
                lambda fp: file_service.insert_columns(
                    fp, sheet_name, start_col, count
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.insert_columns(fp, sheet_name, start_col, count)
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error inserting columns: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Rows",
            destructiveHint=True,
        ),
    )
    def delete_sheet_rows(
        filepath: str,
        sheet_name: str,
        start_row: int,
        count: int = 1,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Delete one or more rows starting at the specified row."""
        try:
            return workbook_dispatch(
                "delete_sheet_rows",
                filepath,
                workbook_transport,
                lambda fp: file_service.delete_sheet_rows(
                    fp, sheet_name, start_row, count
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.delete_sheet_rows(
                        fp, sheet_name, start_row, count
                    )
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error deleting rows: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Delete Columns",
            destructiveHint=True,
        ),
    )
    def delete_sheet_columns(
        filepath: str,
        sheet_name: str,
        start_col: int,
        count: int = 1,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Delete one or more columns starting at the specified column."""
        try:
            return workbook_dispatch(
                "delete_sheet_columns",
                filepath,
                workbook_transport,
                lambda fp: file_service.delete_sheet_columns(
                    fp, sheet_name, start_col, count
                ),
                com_do_op=com_dispatch(
                    lambda c, fp: c.delete_sheet_columns(
                        fp, sheet_name, start_col, count
                    )
                ),
            )
        except (ValidationError, SheetError) as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error deleting columns: {e}")
            raise


    return {
        "create_worksheet": create_worksheet,
        "copy_worksheet": copy_worksheet,
        "delete_worksheet": delete_worksheet,
        "rename_worksheet": rename_worksheet,
        "merge_cells": merge_cells,
        "unmerge_cells": unmerge_cells,
        "copy_range": copy_range,
        "delete_range": delete_range,
        "insert_rows": insert_rows,
        "insert_columns": insert_columns,
        "delete_sheet_rows": delete_sheet_rows,
        "delete_sheet_columns": delete_sheet_columns,
    }
