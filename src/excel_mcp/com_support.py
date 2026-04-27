"""COM stack availability for optional Windows Excel automation (pywin32).

On non-Windows platforms this module never imports pywin32 or win32com so
Linux CI and headless installs stay free of COM DLLs.
"""

from __future__ import annotations

import sys

if sys.platform == "win32":
    try:
        import win32com.client  # noqa: F401

        _PYWIN32_AVAILABLE = True
    except ImportError:
        _PYWIN32_AVAILABLE = False
else:
    _PYWIN32_AVAILABLE = False

COM_STACK_AVAILABLE: bool = _PYWIN32_AVAILABLE


def is_com_runtime_supported() -> bool:
    """True when running on Windows and pywin32 (win32com) is importable."""
    return COM_STACK_AVAILABLE
