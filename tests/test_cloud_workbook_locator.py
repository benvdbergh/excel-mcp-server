"""Cloud workbook locator parsing, get_excel_path, com normalization, and allowlist (STORY-9-1/9-2)."""

from __future__ import annotations

import os
import sys

import pytest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

import excel_mcp.server as server  # noqa: E402
from excel_mcp.path_policy import (  # noqa: E402
    assert_cloud_workbook_url_allowlist,
    cloud_workbook_url_allowed_by_prefix_list,
)
from excel_mcp.path_resolution import (  # noqa: E402
    is_cloud_workbook_locator,
    normalize_workbook_target_for_com,
    parse_cloud_workbook_locator,
)
from excel_mcp.routing.routed_dispatch import execute_routed_workbook_operation  # noqa: E402
from excel_mcp.routing.routing_backend import RoutingBackend  # noqa: E402
from excel_mcp.routing.tool_inventory import ToolKind  # noqa: E402

_VALID_HTTPS = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/book.xlsx"
_VALID_HTTPS_HOST_CASE = "HTTPS://TENANT.SharePoint.com/sites/s/Shared%20Documents/book.xlsx"


class _FakeOpen:
    def __init__(self, open_paths: frozenset[str]) -> None:
        self._open_paths = open_paths

    def is_workbook_open_in_excel(self, resolved_path: str) -> bool:
        return resolved_path in self._open_paths


class _DummyFileSvc:
    pass


_DUMMY = _DummyFileSvc()


def test_valid_https_accepted_and_canonicalized() -> None:
    assert is_cloud_workbook_locator(_VALID_HTTPS)
    out = parse_cloud_workbook_locator(_VALID_HTTPS)
    assert out.startswith("https://tenant.sharepoint.com/")
    assert "book.xlsx" in out


def test_http_scheme_rejected() -> None:
    assert not is_cloud_workbook_locator("http://example.com/a.xlsx")
    with pytest.raises(ValueError, match="https"):
        parse_cloud_workbook_locator("http://example.com/a.xlsx")


def test_malformed_missing_host() -> None:
    assert not is_cloud_workbook_locator("https:///nope.xlsx")
    with pytest.raises(ValueError, match="host"):
        parse_cloud_workbook_locator("https:///nope.xlsx")


def test_nul_rejected() -> None:
    assert not is_cloud_workbook_locator("https://x\x00y.com/a")
    with pytest.raises(ValueError, match="NUL"):
        parse_cloud_workbook_locator("https://x\x00y.com/a")


def test_get_excel_path_stdio_returns_normalized_without_allowlist() -> None:
    server.EXCEL_FILES_PATH = None
    prev = os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
    try:
        out = server.get_excel_path(_VALID_HTTPS)
        assert out == parse_cloud_workbook_locator(_VALID_HTTPS)
        assert out.startswith("https://")
    finally:
        if prev is not None:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = prev


def test_get_excel_path_sse_jail_rejects_cloud() -> None:
    server.EXCEL_FILES_PATH = os.getcwd()
    try:
        with pytest.raises(ValueError, match="EXCEL_FILES_PATH"):
            server.get_excel_path(_VALID_HTTPS)
    finally:
        server.EXCEL_FILES_PATH = None


def test_get_excel_path_stdio_allowlist_fail_closed() -> None:
    server.EXCEL_FILES_PATH = None
    prev_paths = os.environ.get("EXCEL_MCP_ALLOWED_PATHS")
    prev_urls = os.environ.pop("EXCEL_MCP_ALLOWED_URL_PREFIXES", None)
    try:
        os.environ["EXCEL_MCP_ALLOWED_PATHS"] = os.getcwd()
        with pytest.raises(ValueError, match="EXCEL_MCP_ALLOWED_URL_PREFIXES"):
            server.get_excel_path(_VALID_HTTPS)
    finally:
        if prev_paths is None:
            os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
        else:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = prev_paths
        if prev_urls is not None:
            os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = prev_urls


def test_get_excel_path_stdio_allowlist_allows_cloud_with_url_prefix() -> None:
    server.EXCEL_FILES_PATH = None
    prev_paths = os.environ.get("EXCEL_MCP_ALLOWED_PATHS")
    prev_urls = os.environ.get("EXCEL_MCP_ALLOWED_URL_PREFIXES")
    can = parse_cloud_workbook_locator(_VALID_HTTPS)
    try:
        os.environ["EXCEL_MCP_ALLOWED_PATHS"] = os.getcwd()
        os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = "https://tenant.sharepoint.com/sites/s/"
        out = server.get_excel_path(_VALID_HTTPS)
        assert out == can
    finally:
        if prev_paths is None:
            os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
        else:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = prev_paths
        if prev_urls is None:
            os.environ.pop("EXCEL_MCP_ALLOWED_URL_PREFIXES", None)
        else:
            os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = prev_urls


def test_get_excel_path_stdio_allowlist_rejects_url_not_under_prefix() -> None:
    server.EXCEL_FILES_PATH = None
    prev_paths = os.environ.get("EXCEL_MCP_ALLOWED_PATHS")
    prev_urls = os.environ.get("EXCEL_MCP_ALLOWED_URL_PREFIXES")
    try:
        os.environ["EXCEL_MCP_ALLOWED_PATHS"] = os.getcwd()
        os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = "https://other.example.com/"
        with pytest.raises(ValueError, match="not under any prefix"):
            server.get_excel_path(_VALID_HTTPS)
    finally:
        if prev_paths is None:
            os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
        else:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = prev_paths
        if prev_urls is None:
            os.environ.pop("EXCEL_MCP_ALLOWED_URL_PREFIXES", None)
        else:
            os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = prev_urls


def test_normalize_workbook_target_for_com_https_equivalent_forms() -> None:
    a = normalize_workbook_target_for_com(_VALID_HTTPS)
    b = normalize_workbook_target_for_com(_VALID_HTTPS_HOST_CASE)
    assert a == b
    assert a.startswith("https://tenant.sharepoint.com/")


def test_normalize_workbook_target_for_com_unquoted_path_matches_encoded() -> None:
    spaced = "https://tenant.sharepoint.com/sites/s/Shared Documents/book.xlsx"
    encoded = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/book.xlsx"
    assert normalize_workbook_target_for_com(spaced) == normalize_workbook_target_for_com(encoded)


def test_url_prefix_list_accepts_host_case_variants() -> None:
    can = parse_cloud_workbook_locator(_VALID_HTTPS)
    assert cloud_workbook_url_allowed_by_prefix_list(can) is False
    os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = "https://TENANT.SHAREPOINT.COM/sites/s/"
    try:
        assert cloud_workbook_url_allowed_by_prefix_list(can) is True
    finally:
        os.environ.pop("EXCEL_MCP_ALLOWED_URL_PREFIXES", None)


def test_parse_cloud_prefix_host_canonicalization() -> None:
    p1 = parse_cloud_workbook_locator("https://TENANT.SHAREPOINT.COM/sites/s/")
    p2 = parse_cloud_workbook_locator("https://tenant.sharepoint.com/sites/s/")
    assert p1 == p2


def test_assert_cloud_workbook_url_allowlist_noop_without_path_allowlist() -> None:
    prev_paths = os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
    os.environ["EXCEL_MCP_ALLOWED_URL_PREFIXES"] = "https://evil.com/"
    try:
        assert_cloud_workbook_url_allowlist("https://tenant.sharepoint.com/x")
    finally:
        if prev_paths is not None:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = prev_paths
        os.environ.pop("EXCEL_MCP_ALLOWED_URL_PREFIXES", None)


def test_dispatch_file_forced_cloud_returns_error_not_callable() -> None:
    called: list[bool] = []

    def op() -> str:
        called.append(True)
        return "file-read"

    rb = RoutingBackend(_FakeOpen(frozenset({_VALID_HTTPS})), com_execution_available=True, runtime_platform="win32")
    normalized = parse_cloud_workbook_locator(_VALID_HTTPS)
    out, backend = execute_routed_workbook_operation(
        rb,
        _DUMMY,
        resolved_path=normalized,
        workbook_transport="file",
        tool_kind=ToolKind.READ,
        com_strict=True,
        operation_name="workbook_metadata",
        operation_callable=op,
        mcp_tool_name="get_workbook_metadata",
    )
    assert not called
    assert backend == "file"
    assert out.startswith("Error: ")
    assert "openpyxl" in out.lower() or "https" in out.lower()


def test_dispatch_auto_closed_workbook_cloud_returns_error_not_callable() -> None:
    """auto + workbook not open → file backend; cloud URL must not reach openpyxl."""
    called: list[bool] = []

    def op() -> str:
        called.append(True)
        return "write"

    rb = RoutingBackend(_FakeOpen(frozenset()), com_execution_available=True, runtime_platform="win32")
    normalized = parse_cloud_workbook_locator(_VALID_HTTPS)
    out, backend = execute_routed_workbook_operation(
        rb,
        _DUMMY,
        resolved_path=normalized,
        workbook_transport="auto",
        tool_kind=ToolKind.WRITE,
        com_strict=False,
        operation_name="apply_formula",
        operation_callable=op,
        mcp_tool_name="apply_formula",
    )
    assert not called
    assert backend == "file"
    assert out.startswith("Error: ")
