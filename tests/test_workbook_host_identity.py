"""Tests for shared workbook host identity helpers (BEN-131 / ADR 0006)."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from excel_mcp.path_resolution import normalize_workbook_target_for_com
from excel_mcp.routing.workbook_host_identity import (
    count_workbook_collection_matches,
    normalized_workbook_fullname,
    protected_view_candidate_paths,
    protected_view_matches_target,
    workbook_in_protected_view,
)

_HTTPS = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/book.xlsx"
_HTTPS_SPACED = "https://tenant.sharepoint.com/sites/s/Shared Documents/book.xlsx"


def test_normalized_workbook_fullname_https_equivalent() -> None:
    wb = MagicMock()
    wb.FullName = _HTTPS_SPACED
    got = normalized_workbook_fullname(wb)
    assert got == normalize_workbook_target_for_com(_HTTPS)


def test_protected_view_candidate_paths_https_source_path_and_name() -> None:
    pv = MagicMock()
    pv.Workbook = MagicMock()
    pv.Workbook.FullName = _HTTPS_SPACED
    pv.SourcePath = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/"
    pv.SourceName = "book.xlsx"

    cands = protected_view_candidate_paths(pv)
    target = normalize_workbook_target_for_com(_HTTPS)
    assert target in cands


def test_protected_view_candidate_paths_https_source_path_no_trailing_slash() -> None:
    pv = MagicMock()
    pv.Workbook = MagicMock()
    pv.Workbook.FullName = ""
    pv.SourcePath = "https://tenant.sharepoint.com/sites/s/Shared%20Documents"
    pv.SourceName = "book.xlsx"

    cands = protected_view_candidate_paths(pv)
    assert normalize_workbook_target_for_com(_HTTPS) in cands


def test_protected_view_candidate_paths_disk_source() -> None:
    pv = MagicMock()
    pv.Workbook = MagicMock()
    pv.Workbook.FullName = r"C:\sync\book.xlsx"
    pv.SourcePath = r"C:\sync"
    pv.SourceName = "book.xlsx"

    cands = protected_view_candidate_paths(pv)
    assert len(cands) >= 1
    assert cands[0] == normalize_workbook_target_for_com(r"C:\sync\book.xlsx")


def test_count_workbook_collection_matches_single() -> None:
    wb = MagicMock()
    wb.FullName = _HTTPS
    xl = MagicMock()
    xl.Workbooks.Count = 1
    xl.Workbooks.Item = MagicMock(side_effect=lambda i: wb)

    target = normalize_workbook_target_for_com(_HTTPS_SPACED)
    assert count_workbook_collection_matches(xl, target) == 1


def test_count_workbook_collection_matches_ignores_protected_view() -> None:
    """Workbooks collection only — PV windows are not counted (FR-9)."""
    xl = MagicMock()
    xl.Workbooks.Count = 0

    pvw = MagicMock()
    pv = MagicMock()
    pv.Workbook = MagicMock()
    pv.Workbook.FullName = _HTTPS
    pv.SourcePath = "https://tenant.sharepoint.com/sites/s/Shared%20Documents/"
    pv.SourceName = "book.xlsx"
    pvw.Count = 1
    pvw.Item = MagicMock(side_effect=lambda i: pv)
    xl.ProtectedViewWindows = pvw

    target = normalize_workbook_target_for_com(_HTTPS)
    assert count_workbook_collection_matches(xl, target) == 0
    assert protected_view_matches_target(xl, target) is True


def test_workbook_in_protected_view_true() -> None:
    wb = MagicMock()
    wb.FullName = _HTTPS
    pv_wb = MagicMock()
    pv_wb.FullName = _HTTPS_SPACED
    pv = MagicMock()
    pv.Workbook = pv_wb
    pvw = MagicMock()
    pvw.Count = 1
    pvw.Item = MagicMock(side_effect=lambda i: pv)
    xl = MagicMock()
    xl.ProtectedViewWindows = pvw

    assert workbook_in_protected_view(xl, wb) is True


@pytest.mark.parametrize(
    "operator_url,excel_fullname",
    [
        (
            "HTTPS://tenant.SharePoint.com/sites/s/Shared%20Documents/book.xlsx",
            "https://tenant.sharepoint.com/sites/s/Shared Documents/book.xlsx",
        ),
        (
            "https://tenant.sharepoint.com/sites/s/Shared%20Documents/book.xlsx",
            "https://tenant.sharepoint.com/sites/s/Shared%20Documents/book.xlsx",
        ),
    ],
)
def test_operator_and_excel_fullname_normalize_same(
    operator_url: str, excel_fullname: str
) -> None:
    assert normalize_workbook_target_for_com(operator_url) == normalize_workbook_target_for_com(
        excel_fullname
    )
