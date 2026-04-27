"""Tests for operator env: workbook transport and COM strict / fallback (Story 5-1)."""

from __future__ import annotations

import pytest

import importlib

from excel_mcp.routing.routing_env import (
    EXCEL_MCP_COM_ALLOW_FILE_FALLBACK,
    EXCEL_MCP_COM_STRICT,
    EXCEL_MCP_TRANSPORT,
    effective_com_strict,
    read_com_allow_file_fallback,
    read_com_strict,
    read_workbook_transport,
    resolve_workbook_transport,
)


@pytest.fixture
def clean_routing_env(monkeypatch: pytest.MonkeyPatch) -> None:
    for key in (
        EXCEL_MCP_TRANSPORT,
        EXCEL_MCP_COM_STRICT,
        EXCEL_MCP_COM_ALLOW_FILE_FALLBACK,
    ):
        monkeypatch.delenv(key, raising=False)


def test_read_workbook_transport_default_auto(clean_routing_env: None) -> None:
    assert read_workbook_transport() == "auto"


@pytest.mark.parametrize(
    "value,expected",
    [
        ("AUTO", "auto"),
        ("Auto", "auto"),
        ("FILE", "file"),
        ("file", "file"),
        ("COM", "com"),
        ("  com  ", "com"),
    ],
)
def test_read_workbook_transport_values(
    monkeypatch: pytest.MonkeyPatch,
    value: str,
    expected: str,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_TRANSPORT, value)
    assert read_workbook_transport() == expected


def test_read_workbook_transport_empty_whitespace_is_auto(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_TRANSPORT, "   ")
    assert read_workbook_transport() == "auto"


def test_read_workbook_transport_invalid_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(EXCEL_MCP_TRANSPORT, "stdio")
    with pytest.raises(ValueError) as exc:
        read_workbook_transport()
    msg = str(exc.value)
    assert "EXCEL_MCP_TRANSPORT" in msg
    assert "workbook" in msg.lower() or "COM" in msg or "file" in msg
    assert "wire" in msg.lower() or "stdio" in msg.lower() or "HTTP" in msg


def test_read_com_strict_default_true(clean_routing_env: None) -> None:
    assert read_com_strict() is True


@pytest.mark.parametrize("v", ["1", "true", "TRUE", "yes", " Yes "])
def test_read_com_strict_truthy(monkeypatch: pytest.MonkeyPatch, v: str) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_STRICT, v)
    assert read_com_strict() is True


@pytest.mark.parametrize("v", ["0", "false", "FALSE", "no", " NO "])
def test_read_com_strict_falsy(monkeypatch: pytest.MonkeyPatch, v: str) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_STRICT, v)
    assert read_com_strict() is False


def test_read_com_strict_empty_string_default_strict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_STRICT, "")
    assert read_com_strict() is True


def test_read_com_strict_invalid(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_STRICT, "maybe")
    with pytest.raises(ValueError, match="EXCEL_MCP_COM_STRICT"):
        read_com_strict()


def test_read_com_allow_file_fallback_default_false(clean_routing_env: None) -> None:
    assert read_com_allow_file_fallback() is False


@pytest.mark.parametrize("v", ["1", "true", "yes"])
def test_read_com_allow_file_fallback_truthy(
    monkeypatch: pytest.MonkeyPatch,
    v: str,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_ALLOW_FILE_FALLBACK, v)
    assert read_com_allow_file_fallback() is True


@pytest.mark.parametrize("v", ["0", "false", "no"])
def test_read_com_allow_file_fallback_falsy(
    monkeypatch: pytest.MonkeyPatch,
    v: str,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_ALLOW_FILE_FALLBACK, v)
    assert read_com_allow_file_fallback() is False


def test_effective_com_strict_defaults(clean_routing_env: None) -> None:
    assert effective_com_strict() is True


def test_effective_com_strict_explicit_non_strict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_STRICT, "0")
    monkeypatch.delenv(EXCEL_MCP_COM_ALLOW_FILE_FALLBACK, raising=False)
    assert effective_com_strict() is False


def test_effective_com_strict_fallback_forces_non_strict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_COM_STRICT, "1")
    monkeypatch.setenv(EXCEL_MCP_COM_ALLOW_FILE_FALLBACK, "true")
    assert effective_com_strict() is False


def test_effective_com_strict_fallback_with_default_strict_still_non_strict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv(EXCEL_MCP_COM_STRICT, raising=False)
    monkeypatch.setenv(EXCEL_MCP_COM_ALLOW_FILE_FALLBACK, "yes")
    assert effective_com_strict() is False


def test_save_after_write_api_removed() -> None:
    """STORY-11-2: persistence is only via save_workbook; env flag API is gone."""
    re = importlib.import_module("excel_mcp.routing.routing_env")
    for name in (
        "EXCEL_MCP_SAVE_AFTER_WRITE_DEFAULT",
        "read_save_after_write_default",
        "effective_save_after_write",
    ):
        assert not hasattr(re, name), f"expected {name} to be removed"


def test_resolve_workbook_transport_uses_env_when_override_none(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv(EXCEL_MCP_TRANSPORT, "file")
    assert resolve_workbook_transport(None) == "file"
    assert resolve_workbook_transport("  ") == "file"


def test_resolve_workbook_transport_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(EXCEL_MCP_TRANSPORT, "file")
    assert resolve_workbook_transport("COM") == "com"


def test_resolve_workbook_transport_invalid_override() -> None:
    with pytest.raises(ValueError, match="workbook_transport"):
        resolve_workbook_transport("stdio")


def test_read_functions_use_explicit_environ_mapping() -> None:
    """Call-time mapping overrides process env for these readers."""
    env = {
        EXCEL_MCP_TRANSPORT: "FILE",
        EXCEL_MCP_COM_STRICT: "false",
        EXCEL_MCP_COM_ALLOW_FILE_FALLBACK: "0",
    }
    assert read_workbook_transport(env) == "file"
    assert read_com_strict(env) is False
    assert read_com_allow_file_fallback(env) is False
    assert effective_com_strict(env) is False
