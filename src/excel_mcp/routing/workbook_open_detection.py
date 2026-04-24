"""Port for detecting whether a workbook path is open in Excel (FR-2).

Epic 6 may supply a COM-backed implementation that compares ``resolved_path``
to Excel's open workbooks. The default stub never starts Excel (FR-10).
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
    """Default port implementation: never reports a workbook as open in Excel.

    Always returns ``False`` on every OS until a real COM detector exists (Epic 6).
    Does not start Excel, spawn subprocesses, or import COM bindings.
    """

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        del resolved_path
        return False
