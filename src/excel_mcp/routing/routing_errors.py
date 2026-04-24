"""Routing-specific errors (COM strict mode, ADR 0005)."""

from __future__ import annotations


class ComRoutingError(Exception):
    """Raised when COM routing is required (strict) but cannot be honored.

    The message always contains :data:`STABLE_TOKEN` so integrators can match
    without depending on full prose text across releases.
    """

    STABLE_TOKEN: str = "EXCEL_MCP_COM_STRICT_UNAVAILABLE"

    def __init__(self, *, reason_code: str, message: str | None = None) -> None:
        self.reason_code = reason_code
        body = message or reason_code
        super().__init__(f"{self.STABLE_TOKEN}: {body}")


class ComExecutionNotImplementedError(Exception):
    """Raised when routing selects COM but the Epic 4 dispatcher only runs file I/O.

    Real COM execution is deferred to later epics; this error prevents silently
    mutating the on-disk workbook when the matrix chose ``com``.
    """

    STABLE_TOKEN: str = "EXCEL_MCP_COM_EXECUTION_NOT_IMPLEMENTED"

    def __init__(self, *, message: str | None = None) -> None:
        body = message or "COM workbook execution is not wired (Epic 4 file-only dispatcher)"
        super().__init__(f"{self.STABLE_TOKEN}: {body}")
