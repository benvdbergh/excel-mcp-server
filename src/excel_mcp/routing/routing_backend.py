"""Workbook backend selection (FR-3) — file vs COM with machine-readable reasons."""

from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Literal

from excel_mcp.routing.routing_errors import ComRoutingError
from excel_mcp.routing.tool_inventory import ToolKind
from excel_mcp.routing.workbook_open_detection import WorkbookOpenInExcelPort

WorkbookBackend = Literal["file", "com"]
WorkbookTransport = Literal["auto", "file", "com"]


@dataclass(frozen=True, slots=True)
class WorkbookBackendResolution:
    """Result of :meth:`RoutingBackend.resolve_workbook_backend`."""

    backend: WorkbookBackend
    reason: str
    requested_transport: WorkbookTransport | None = None


def _normalize_tool_kind(tool_kind: ToolKind | str) -> ToolKind:
    if isinstance(tool_kind, ToolKind):
        return tool_kind
    return ToolKind(tool_kind)


class RoutingBackend:
    """Selects file vs COM workbook backend from transport mode and policy."""

    def __init__(
        self,
        workbook_open: WorkbookOpenInExcelPort,
        *,
        com_execution_available: bool = False,
        runtime_platform: str | None = None,
    ) -> None:
        """Args:
            workbook_open: Injectable open-detection port (tests use fakes).
            com_execution_available: When ``False`` (Epic 4 default), COM runtime
                is treated as absent even on Windows — strict ``com``/``auto``
                paths that require COM raise :class:`ComRoutingError`.
            runtime_platform: Override ``sys.platform`` (tests). ``None`` uses
                real ``sys.platform`` so Linux CI stays COM-free without importing
                COM bindings.
        """
        self._workbook_open = workbook_open
        self._com_execution_available = com_execution_available
        self._runtime_platform = runtime_platform

    def _platform(self) -> str:
        return self._runtime_platform if self._runtime_platform is not None else sys.platform

    def _com_runtime_viable(self) -> bool:
        """True if this process could execute COM-backed workbook operations."""
        return self._platform() == "win32" and self._com_execution_available

    def _strict_com_failure(
        self,
        *,
        reason_code: str,
        detail: str,
    ) -> ComRoutingError:
        return ComRoutingError(reason_code=reason_code, message=detail)

    def resolve_workbook_backend(
        self,
        *,
        resolved_path: str,
        transport: WorkbookTransport,
        tool_kind: ToolKind | str,
        com_strict: bool,
    ) -> WorkbookBackendResolution:
        """Choose ``file`` or ``com`` and a stable ``reason`` for logs/metrics.

        Reason strings align with ``docs/architecture/target-architecture.md``
        (e.g. ``full_name_match``, ``forced_file``, ``forced_com``).

        **``transport="file"``** — always file, reason ``forced_file``.

        **``transport="auto"``** — file when the workbook is not reported open;
        reason ``auto_workbook_not_open_file``. When open and ``tool_kind`` is
        :attr:`~ToolKind.V1_FILE_FORCED`, file is forced with ``v1_file_forced``
        (ADR 0004) even though COM would otherwise apply. When open and COM is
        not viable (non-Windows or no COM executor), behavior matches unavailable
        COM: ``com_strict`` → :class:`ComRoutingError`; else file with
        ``com_unsupported_non_windows`` or ``com_unavailable_file_fallback``.
        When open and COM is viable → ``com`` with ``full_name_match``.

        **``transport="com"``** — :attr:`~ToolKind.V1_FILE_FORCED` forces file
        with ``v1_file_forced``. If COM is not viable, ``com_strict`` raises
        :class:`ComRoutingError`; otherwise file with
        ``com_unavailable_file_fallback`` / ``com_unsupported_non_windows``.
        If COM is viable but the workbook is not open, ``com_strict`` raises
        :class:`ComRoutingError` (ADR 0005); non-strict falls back to file with
        ``com_workbook_not_open_file_fallback``. Otherwise ``com`` with
        ``forced_com``.

        **Non-strict COM unavailable fallback:** documented here — returns file
        backend with ``com_unavailable_file_fallback`` (Windows, executor off)
        or ``com_unsupported_non_windows`` when ``sys.platform`` (or override)
        is not ``win32``.
        """
        tk = _normalize_tool_kind(tool_kind)
        req: WorkbookTransport | None = transport

        if transport == "file":
            return WorkbookBackendResolution(
                backend="file",
                reason="forced_file",
                requested_transport=req,
            )

        open_in_excel = self._workbook_open.is_workbook_open_in_excel(resolved_path)
        com_viable = self._com_runtime_viable()

        if transport == "auto":
            if not open_in_excel:
                return WorkbookBackendResolution(
                    backend="file",
                    reason="auto_workbook_not_open_file",
                    requested_transport=req,
                )
            if tk == ToolKind.V1_FILE_FORCED:
                return WorkbookBackendResolution(
                    backend="file",
                    reason="v1_file_forced",
                    requested_transport=req,
                )
            if not com_viable:
                if com_strict:
                    raise self._strict_com_unavailable()
                reason = (
                    "com_unsupported_non_windows"
                    if self._platform() != "win32"
                    else "com_unavailable_file_fallback"
                )
                return WorkbookBackendResolution(
                    backend="file",
                    reason=reason,
                    requested_transport=req,
                )
            return WorkbookBackendResolution(
                backend="com",
                reason="full_name_match",
                requested_transport=req,
            )

        # transport == "com"
        if tk == ToolKind.V1_FILE_FORCED:
            return WorkbookBackendResolution(
                backend="file",
                reason="v1_file_forced",
                requested_transport=req,
            )
        if not com_viable:
            if com_strict:
                raise self._strict_com_unavailable()
            reason = (
                "com_unsupported_non_windows"
                if self._platform() != "win32"
                else "com_unavailable_file_fallback"
            )
            return WorkbookBackendResolution(
                backend="file",
                reason=reason,
                requested_transport=req,
            )
        if not open_in_excel:
            if com_strict:
                raise self._strict_com_failure(
                    reason_code="com_workbook_not_open",
                    detail="workbook_transport=com requires workbook open in Excel",
                )
            return WorkbookBackendResolution(
                backend="file",
                reason="com_workbook_not_open_file_fallback",
                requested_transport=req,
            )
        return WorkbookBackendResolution(
            backend="com",
            reason="forced_com",
            requested_transport=req,
        )

    def _strict_com_unavailable(self) -> ComRoutingError:
        plat = self._platform()
        if plat != "win32":
            return self._strict_com_failure(
                reason_code="com_unsupported_non_windows",
                detail="COM workbook backend is not supported on this platform",
            )
        return self._strict_com_failure(
            reason_code="com_execution_unavailable",
            detail="COM workbook execution is not available (strict mode)",
        )
