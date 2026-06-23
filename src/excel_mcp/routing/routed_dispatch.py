"""Routed workbook dispatch with timing and structured logs (Epics 4–6, NFR-3)."""

from __future__ import annotations

import json
import logging
import os
import time
from collections.abc import Callable, Mapping, Sequence
from typing import Any, TypedDict

from excel_mcp.routing.routing_backend import (
    RoutingBackend,
    WorkbookBackend,
    WorkbookTransport,
)
from excel_mcp.routing.routing_errors import ComExecutionNotImplementedError
from excel_mcp.routing.tool_inventory import ToolKind
from excel_mcp.path_resolution import is_cloud_workbook_locator
from excel_mcp.routing.workbook_operation_contract import (
    ROUTED_WORKBOOK_OPERATION_NAMES,
    RoutedWorkbookOperations,
)

# Stable user-visible error when openpyxl/file backend is selected for HTTPS locators;
# file-based tools cannot open remote URLs.
_FILE_BACKEND_CLOUD_LOCATOR_ERROR = (
    "Error: The file (openpyxl) backend cannot use cloud HTTPS workbook URLs. "
    "Open the workbook in Excel and use workbook_transport auto or com (workbook must be open in Excel for COM routing)."
)


class RoutedDispatchMeta(TypedDict):
    """Routing metadata for MCP response envelope (ADR 0010)."""

    workbook_transport: WorkbookTransport
    workbook_backend: WorkbookBackend
    routing_reason: str
    duration_ms: float


class RoutedResponseWarning(TypedDict):
    code: str
    message: str


def build_routed_response_envelope(
    result_text: str,
    meta: RoutedDispatchMeta | Mapping[str, Any],
    warnings: Sequence[RoutedResponseWarning | Mapping[str, str]] | None = None,
) -> str:
    """Serialize ADR 0010 success envelope around ``result_text``."""
    try:
        parsed = json.loads(result_text)
        result: Any = parsed if isinstance(parsed, (dict, list)) else result_text
    except json.JSONDecodeError:
        result = result_text
    envelope = {
        "result": result,
        "_meta": dict(meta),
        "warnings": list(warnings or ()),
    }
    return json.dumps(envelope, ensure_ascii=False)


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
    # `os.path.basename` on POSIX ignores `\`; workbook paths can still be
    # Windows-shaped (tests, or resolved paths forwarded across contexts).
    normalized = resolved_path.replace("\\", "/")
    return os.path.basename(normalized)


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
) -> tuple[str, WorkbookBackend, RoutedDispatchMeta]:
    """Resolve backend, run file or COM I/O, emit one structured log line.

    When resolution is ``backend="file"``, ``operation_callable`` runs (typically
    closes over ``FileWorkbookService``).

    When resolution is ``backend="com"``, ``com_operation_callable`` runs if
    provided; if it is ``None``, logs then raises
    :class:`ComExecutionNotImplementedError` (no silent file fallback).

    Returns ``(result_text, executed_backend, routing_meta)`` where
    ``executed_backend`` is ``"file"`` or ``"com"`` and ``routing_meta`` carries
    ADR 0010 field names for optional MCP envelopes.

    ``file_workbook_service`` is required for handler wiring consistency; callers
    typically close over it inside ``operation_callable``. This module does not
    invoke methods on it directly.

    Log line: a single ``logger.info`` with ``json.dumps`` of a dict using ADR
    0001-aligned field names (``workbook_transport``, ``workbook_backend``,
    ``routing_reason``, ``duration_ms``, ``workbook_path``, ``operation_name``,
    optional ``mcp_tool_name``, and ``v1_file_forced``: ``true`` when ADR 0004
    applies). Uses logger ``excel-mcp.routing`` by default
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
    routing_meta: RoutedDispatchMeta = {
        "workbook_transport": workbook_transport,
        "workbook_backend": "file",
        "routing_reason": "",
        "duration_ms": 0.0,
    }
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
            if is_cloud_workbook_locator(resolved_path):
                result = _FILE_BACKEND_CLOUD_LOCATOR_ERROR
            else:
                result = operation_callable()
            executed = "file"
    finally:
        if resolution is not None:
            duration_ms = (time.perf_counter() - t0) * 1000.0
            backend_for_meta = executed if executed is not None else resolution.backend
            routing_meta = {
                "workbook_transport": workbook_transport,
                "workbook_backend": backend_for_meta,
                "routing_reason": resolution.reason,
                "duration_ms": round(duration_ms, 3),
            }
            payload: dict[str, object] = {
                "workbook_transport": workbook_transport,
                "workbook_backend": resolution.backend,
                "routing_reason": resolution.reason,
                "duration_ms": routing_meta["duration_ms"],
                "workbook_path": redact_workbook_path_for_logs(resolved_path),
                "operation_name": operation_name,
            }
            if mcp_tool_name is not None:
                payload["mcp_tool_name"] = mcp_tool_name
            if resolution.reason == "v1_file_forced":
                payload["v1_file_forced"] = True
            log.info(json.dumps(payload, separators=(",", ":"), ensure_ascii=True))
    if pending_com is not None:
        raise pending_com
    assert result is not None and executed is not None
    return result, executed, routing_meta
