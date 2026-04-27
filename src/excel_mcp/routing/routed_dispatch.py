"""Routed workbook dispatch with timing and structured logs (Epics 4–6, NFR-3)."""

from __future__ import annotations

import json
import logging
import os
import time
from collections.abc import Callable

from excel_mcp.routing.routing_backend import (
    RoutingBackend,
    WorkbookBackend,
    WorkbookTransport,
)
from excel_mcp.routing.routing_errors import ComExecutionNotImplementedError
from excel_mcp.routing.tool_inventory import ToolKind
from excel_mcp.routing.workbook_operation_contract import (
    ROUTED_WORKBOOK_OPERATION_NAMES,
    RoutedWorkbookOperations,
)


def redact_workbook_path_for_logs(resolved_path: str) -> str:
    """Return a log-safe workbook path segment (basename by default).

    Directory structure is stripped so logs do not leak layout of the host.
    Set ``EXCEL_MCP_LOG_FULL_PATHS=1`` to log the full normalized path for
    break-glass debugging (operators only).

    TODO: consider tightening (hash-only) if full paths prove too sensitive even
    when opt-in.
    """
    flag = os.environ.get("EXCEL_MCP_LOG_FULL_PATHS", "").strip().lower()
    if flag in ("1", "true", "yes"):
        return resolved_path
    return os.path.basename(resolved_path)


def execute_routed_workbook_operation(
    routing_backend: RoutingBackend,
    file_workbook_service: RoutedWorkbookOperations,
    *,
    resolved_path: str,
    workbook_transport: WorkbookTransport,
    tool_kind: ToolKind | str,
    com_strict: bool,
    operation_name: str,
    operation_callable: Callable[[], str],
    com_operation_callable: Callable[[], str] | None = None,
    mcp_tool_name: str | None = None,
    logger: logging.Logger | None = None,
) -> tuple[str, WorkbookBackend]:
    """Resolve backend, run file or COM I/O, emit one structured log line.

    When resolution is ``backend="file"``, ``operation_callable`` runs (typically
    closes over ``FileWorkbookService``).

    When resolution is ``backend="com"``, ``com_operation_callable`` runs if
    provided; if it is ``None``, logs then raises
    :class:`ComExecutionNotImplementedError` (no silent file fallback).

    Returns ``(result_text, executed_backend)`` where ``executed_backend`` is
    ``"file"`` or ``"com"``.

    ``file_workbook_service`` is required for handler wiring consistency; callers
    typically close over it inside ``operation_callable``. This module does not
    invoke methods on it directly.

    Log line: a single ``logger.info`` with ``json.dumps`` of a dict using ADR
    0001-aligned field names (``workbook_transport``, ``workbook_backend``,
    ``routing_reason``, ``duration_ms``, ``workbook_path``, ``operation_name``,
    optional ``mcp_tool_name``). Uses logger ``excel-mcp.routing`` by default
    (stdio-safe: no ``print``).
    """
    _ = file_workbook_service
    if operation_name not in ROUTED_WORKBOOK_OPERATION_NAMES:
        allowed = ", ".join(sorted(ROUTED_WORKBOOK_OPERATION_NAMES))
        raise ValueError(f"operation_name must be one of ROUTED_WORKBOOK_OPERATION_NAMES; got {operation_name!r}. ({allowed})")

    log = logger if logger is not None else logging.getLogger("excel-mcp.routing")
    t0 = time.perf_counter()
    resolution = None
    pending_com: ComExecutionNotImplementedError | None = None
    result: str | None = None
    executed: WorkbookBackend | None = None
    try:
        resolution = routing_backend.resolve_workbook_backend(
            resolved_path=resolved_path,
            transport=workbook_transport,
            tool_kind=tool_kind,
            com_strict=com_strict,
        )
        if resolution.backend == "com":
            if com_operation_callable is None:
                pending_com = ComExecutionNotImplementedError()
            else:
                result = com_operation_callable()
                executed = "com"
        else:
            result = operation_callable()
            executed = "file"
    finally:
        if resolution is not None:
            duration_ms = (time.perf_counter() - t0) * 1000.0
            payload: dict[str, object] = {
                "workbook_transport": workbook_transport,
                "workbook_backend": resolution.backend,
                "routing_reason": resolution.reason,
                "duration_ms": round(duration_ms, 3),
                "workbook_path": redact_workbook_path_for_logs(resolved_path),
                "operation_name": operation_name,
            }
            if mcp_tool_name is not None:
                payload["mcp_tool_name"] = mcp_tool_name
            log.info(json.dumps(payload, separators=(",", ":"), ensure_ascii=True))
    if pending_com is not None:
        raise pending_com
    assert result is not None and executed is not None
    return result, executed
