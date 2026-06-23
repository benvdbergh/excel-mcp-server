"""COM-based detection of whether a workbook path is open in Excel (FR-2).

Used by :class:`excel_mcp.routing.routing_backend.RoutingBackend` when the COM
runtime is available. Enumeration runs on :class:`excel_mcp.com_executor.ComThreadExecutor`
so pywin32 is never touched from arbitrary threads (FR-10).
"""

from __future__ import annotations

from excel_mcp.com_executor import ComThreadExecutor
from excel_mcp.path_resolution import normalize_workbook_target_for_com
from excel_mcp.routing.workbook_host_identity import count_workbook_collection_matches


def _count_workbook_matches_worker(resolved_path: str) -> int:
    """Return how many ``Application.Workbooks`` entries match ``resolved_path``."""
    import win32com.client  # lazy: COM thread only

    try:
        target = normalize_workbook_target_for_com(resolved_path)
    except ValueError:
        return 0
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return 0
    return count_workbook_collection_matches(xl, target)


class ComWorkbookOpenInExcel:
    """True when exactly one normal (non–Protected View) window holds ``resolved_path``.

    Protected View workbooks are not counted here so ``auto`` does not route
    mutations to COM until the user clicks **Enable Editing** (then the workbook
    appears in ``Workbooks`` and matches).
    """

    def __init__(self, executor: ComThreadExecutor) -> None:
        self._executor = executor

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        count = self._executor.submit(_count_workbook_matches_worker, resolved_path)
        return count == 1


__all__ = ["ComWorkbookOpenInExcel"]
