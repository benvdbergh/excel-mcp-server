"""Tests for ``WorkbookOpenInExcelPort`` / ``StubWorkbookOpenInExcel`` (STORY-4-1)."""

import os
import sys

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.workbook_open_detection import (  # noqa: E402
    StubWorkbookOpenInExcel,
    WorkbookOpenInExcelPort,
)


def _is_open_under_test(port: WorkbookOpenInExcelPort, resolved_path: str) -> bool:
    """Thin consumer to prove the port is injectable without wiring ``server.py``."""
    return port.is_workbook_open_in_excel(resolved_path)


class _FakeOpenForPath:
    """Test double: True only for one normalized path string."""

    def __init__(self, open_path: str) -> None:
        self._open_path = open_path

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        return resolved_path == self._open_path


def test_stub_always_false() -> None:
    stub = StubWorkbookOpenInExcel()
    assert stub.is_workbook_open_in_excel("/any/normalized/path.xlsx") is False
    assert stub.is_workbook_open_in_excel("") is False


def test_fake_injectable_true_only_for_specific_normalized_path() -> None:
    normalized = "/tmp/project/Book1.xlsx"
    fake = _FakeOpenForPath(normalized)
    assert isinstance(fake, WorkbookOpenInExcelPort)

    port: WorkbookOpenInExcelPort = fake
    assert _is_open_under_test(port, normalized) is True
    assert _is_open_under_test(port, "/tmp/project/Other.xlsx") is False
    assert _is_open_under_test(port, normalized.upper()) is False


def test_routing_package_exports_workbook_open_port_and_stub() -> None:
    from excel_mcp.routing import (  # noqa: E402
        StubWorkbookOpenInExcel as Stub,
        WorkbookOpenInExcelPort as Port,
    )

    assert Stub is StubWorkbookOpenInExcel
    assert Port is WorkbookOpenInExcelPort
