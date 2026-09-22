"""stdio entrypoint must keep stdout JSON-RPC-only."""

from __future__ import annotations

import os
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.__main__ import stdio  # noqa: E402


def test_stdio_runtime_error_goes_to_stderr(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def boom() -> None:
        raise RuntimeError("boom")

    monkeypatch.setattr("excel_mcp.__main__.run_stdio", boom)
    stdio()
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "boom" in captured.err
    assert "Service stopped." in captured.err


def test_stdio_keyboard_interrupt_goes_to_stderr(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def stop() -> None:
        raise KeyboardInterrupt

    monkeypatch.setattr("excel_mcp.__main__.run_stdio", stop)
    stdio()
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "Shutting down" in captured.err
    assert "Service stopped." in captured.err
