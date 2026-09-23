from __future__ import annotations

import logging
from typing import Callable, Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_mcp.com.service import ComWorkbookService
from excel_mcp.exceptions import WorkbookError
from excel_mcp.fileio.service import FileWorkbookService
from excel_mcp.routing.routing_env import resolve_workbook_transport
from excel_mcp.routing.routing_errors import (
    ComExecutionNotImplementedError,
    ComRoutingError,
)

logger = logging.getLogger("excel-mcp")


def register_session(
    mcp: FastMCP,
    *,
    file_service: FileWorkbookService,
    get_com_service: Callable[[], ComWorkbookService | None],
    workbook_dispatch: Callable[..., str],
    com_dispatch: Callable[..., Callable[[str], str] | None],
    get_excel_path: Callable[[str], str],
) -> Dict[str, Callable]:
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Create Workbook",
            destructiveHint=True,
        ),
    )
    def create_workbook(
        filepath: str,
        workbook_transport: Optional[str] = None,
        open_in_excel: bool = False,
    ) -> str:
        """Create new Excel workbook.

        When ``open_in_excel`` is true and COM is available, opens the file in Excel
        after creation (ADR 0008 lifecycle).
        """
        try:
            resolved = get_excel_path(filepath)
            out = workbook_dispatch(
                "create_workbook",
                filepath,
                workbook_transport,
                lambda fp: file_service.create_workbook(fp),
                com_do_op=com_dispatch(lambda c, fp: c.create_workbook(fp)),
            )
            if open_in_excel:
                com_service = get_com_service()
                if com_service is None:
                    return (
                        f"{out}\n"
                        "Note: open_in_excel was ignored (Excel COM is not available)."
                    )
                if not str(out).lstrip().lower().startswith("error:"):
                    extra = com_service.open_workbook_in_excel(resolved)
                    if str(extra).lstrip().lower().startswith("error:"):
                        return f"Error: {out}\n{extra}"
                    return f"{out}\n{extra}"
            return out
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error creating workbook: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Open Workbook in Excel",
            destructiveHint=True,
        ),
    )
    def excel_open_workbook(filepath: str) -> str:
        """Open an existing workbook in the Excel host (``Workbooks.Open``).

        Binds the workbook for subsequent COM-first routing when identity matches.
        Requires Windows Excel COM (ADR 0008).
        """
        try:
            resolved = get_excel_path(filepath)
            com_service = get_com_service()
            if com_service is None:
                return "Error: Excel COM automation is not available on this host."
            return com_service.open_workbook_in_excel(resolved)
        except ValueError as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"excel_open_workbook: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Close Workbook in Excel",
            destructiveHint=True,
        ),
    )
    def excel_close_workbook(filepath: str, save: bool = False) -> str:
        """Close a workbook in the Excel host. Optionally save to disk first (ADR 0008)."""
        try:
            resolved = get_excel_path(filepath)
            com_service = get_com_service()
            if com_service is None:
                return "Error: Excel COM automation is not available on this host."
            return com_service.close_workbook_in_excel(resolved, save=save)
        except ValueError as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"excel_close_workbook: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Open Workbooks",
            readOnlyHint=True,
        ),
    )
    def excel_list_open_workbooks(detail: Optional[str] = None) -> str:
        """Enumerate workbooks open in Excel; returns JSON with ``full_name``, ``name``, ``is_active``.

        Use each ``full_name`` (disk path or https SharePoint-style URL per Excel) with
        ``get_workbook_metadata`` and other filepath-based tools. COM-only; no ``filepath``.

        ``detail``: ``minimal`` (default) lists workbooks only; ``active_context`` also
        returns ``active_workbook``, ``active_sheet``, and ``selection``.

        Requires Windows with ``excel-com-mcp[com]`` and a running Excel instance.
        """
        try:
            com_service = get_com_service()
            if com_service is None:
                return "Error: Excel COM automation is not available on this host."
            return com_service.list_open_workbooks(detail=detail)
        except Exception as e:
            logger.error(f"excel_list_open_workbooks: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Evaluate Range (COM Recalc)",
            destructiveHint=True,
        ),
    )
    def evaluate_range(
        filepath: str,
        sheet_name: str,
        start_cell: Optional[str] = None,
        end_cell: Optional[str] = None,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Force Excel to recalculate a worksheet or range before COM reads.

        COM-only host side effect: updates in-memory Excel state only. Does **not** write
        to disk — call ``save_workbook`` when subsequent ``workbook_transport=file`` reads
        must see recalculated values. Omit ``start_cell`` to recalc the entire sheet.

        ``workbook_transport=file`` is rejected (recalc cannot affect openpyxl reads).
        """
        try:
            transport = resolve_workbook_transport(workbook_transport)
            if transport == "file":
                return (
                    "Error: evaluate_range requires COM (Excel host recalc). "
                    "workbook_transport=file is not supported; open the workbook in Excel "
                    "and omit transport or use com/auto."
                )
            resolved = get_excel_path(filepath)
            com_service = get_com_service()
            if com_service is None:
                return "Error: Excel COM automation is not available on this host."
            return com_service.evaluate_range(
                resolved, sheet_name, start_cell, end_cell
            )
        except ValueError as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"evaluate_range: {e}")
            raise

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Save Workbook",
            destructiveHint=True,
        ),
    )
    def save_workbook(
        filepath: str,
        workbook_transport: Optional[str] = None,
    ) -> str:
        """Persist the workbook to disk (file backend or COM host save).

        Use this before ``read_data_from_excel`` when mutations ran via COM so
        on-disk state matches Excel (ADR 0003).
        """
        try:
            return workbook_dispatch(
                "save_workbook",
                filepath,
                workbook_transport,
                lambda fp: file_service.save_workbook(fp),
                com_do_op=com_dispatch(lambda c, fp: c.save_workbook(fp)),
            )
        except WorkbookError as e:
            return f"Error: {str(e)}"
        except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
            return f"Error: {str(e)}"
        except Exception as e:
            logger.error(f"Error saving workbook: {e}")
            raise

    return {
        "create_workbook": create_workbook,
        "excel_open_workbook": excel_open_workbook,
        "excel_close_workbook": excel_close_workbook,
        "excel_list_open_workbooks": excel_list_open_workbooks,
        "evaluate_range": evaluate_range,
        "save_workbook": save_workbook,
    }
