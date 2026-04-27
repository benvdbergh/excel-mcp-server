"""Port for detecting whether a workbook path is open in Excel (FR-2).

When the COM stack is available, :mod:`excel_mcp.server` wires
:class:`excel_mcp.routing.com_workbook_open_detection.ComWorkbookOpenInExcel`
(executor-backed enumeration). Otherwise :class:`StubWorkbookOpenInExcel` is used;
it never starts Excel (FR-10).
"""

from __future__ import annotations

from typing import Protocol, runtime_checkable


@runtime_checkable
class WorkbookOpenInExcelPort(Protocol):
    """Detects if a normalized file path corresponds to an open Excel workbook."""

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        """Return whether ``resolved_path`` is open in Excel.

        ``resolved_path`` must already be normalized (e.g. via ``os.path.realpath``)
        so callers and implementations agree on a single canonical string form.
        """
        ...


class StubWorkbookOpenInExcel:
    """Port implementation used when COM is unavailable: never reports open.

    Does not start Excel, spawn subprocesses, or import COM bindings.
    """

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        del resolved_path
        return False
