"""Tests for ``FileWorkbookService`` (STORY-3-1): delegation and contract surface."""

import json
import os
import sys
from unittest.mock import MagicMock, patch

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.routing.file_workbook_service import FileWorkbookService  # noqa: E402
from excel_mcp.routing.workbook_operation_contract import (  # noqa: E402
    ROUTED_WORKBOOK_OPERATION_NAMES,
)


def test_file_workbook_service_has_all_routed_operation_names() -> None:
    svc = FileWorkbookService()
    for name in ROUTED_WORKBOOK_OPERATION_NAMES:
        assert hasattr(svc, name), name
        assert callable(getattr(svc, name)), name


@patch("excel_mcp.routing.file_workbook_service.read_excel_range_with_metadata")
def test_read_range_with_metadata_json(mock_read: MagicMock) -> None:
    payload = {
        "range": "A1:A1",
        "sheet_name": "S",
        "cells": [{"address": "A1", "value": 1, "row": 1, "column": 1}],
    }
    mock_read.return_value = payload
    svc = FileWorkbookService()
    out = svc.read_range_with_metadata("/abs/book.xlsx", "S", "A1", "A1")
    mock_read.assert_called_once_with("/abs/book.xlsx", "S", "A1", "A1")
    assert json.loads(out) == payload
    assert out.startswith("{")


@patch("excel_mcp.routing.file_workbook_service.read_excel_range_with_metadata")
def test_read_range_with_metadata_empty(mock_read: MagicMock) -> None:
    mock_read.return_value = {"cells": []}
    svc = FileWorkbookService()
    assert svc.read_range_with_metadata("/abs/b.xlsx", "S") == "No data found in specified range"


@patch("excel_mcp.routing.file_workbook_service.get_all_validation_ranges")
@patch("excel_mcp.routing.file_workbook_service.load_workbook")
def test_read_worksheet_data_validation_closes_workbook_on_get_all_failure(
    mock_load_workbook: MagicMock, mock_get_all: MagicMock
) -> None:
    ws = MagicMock()
    wb = MagicMock()
    wb.sheetnames = ["Data"]
    wb.__getitem__.return_value = ws
    mock_load_workbook.return_value = wb
    mock_get_all.side_effect = RuntimeError("validation scan failed")

    svc = FileWorkbookService()
    try:
        svc.read_worksheet_data_validation("/abs/w.xlsx", "Data")
    except RuntimeError:
        pass
    else:
        raise AssertionError("expected RuntimeError from get_all_validation_ranges")

    mock_load_workbook.assert_called_once_with("/abs/w.xlsx", read_only=False)
    wb.close.assert_called_once()


@patch("excel_mcp.routing.file_workbook_service.get_all_validation_ranges")
@patch("excel_mcp.routing.file_workbook_service.load_workbook")
def test_read_worksheet_data_validation(
    mock_load_workbook: MagicMock, mock_get_all: MagicMock
) -> None:
    ws = MagicMock()
    wb = MagicMock()
    wb.sheetnames = ["Data"]
    wb.__getitem__.return_value = ws
    mock_load_workbook.return_value = wb
    mock_get_all.return_value = [{"cells": "A1:A2", "type": "list"}]

    svc = FileWorkbookService()
    out = svc.read_worksheet_data_validation("/abs/w.xlsx", "Data")

    mock_load_workbook.assert_called_once_with("/abs/w.xlsx", read_only=False)
    mock_get_all.assert_called_once_with(ws)
    wb.close.assert_called_once()
    data = json.loads(out)
    assert data["sheet_name"] == "Data"
    assert data["validation_rules"] == mock_get_all.return_value


@patch("excel_mcp.routing.file_workbook_service.validate_range_in_sheet_operation")
def test_validate_sheet_range(mock_validate: MagicMock) -> None:
    mock_validate.return_value = {"message": "Range 'A1' is valid. Sheet contains data in range 'A1:B2'"}
    svc = FileWorkbookService()
    assert svc.validate_sheet_range("/abs/f.xlsx", "Sh", "A1") == mock_validate.return_value["message"]
    mock_validate.assert_called_once_with("/abs/f.xlsx", "Sh", "A1")

    svc.validate_sheet_range("/abs/f.xlsx", "Sh", "A1", "B2")
    mock_validate.assert_called_with("/abs/f.xlsx", "Sh", "A1:B2")


@patch("excel_mcp.routing.file_workbook_service.write_data")
def test_write_cell_grid(mock_write: MagicMock) -> None:
    mock_write.return_value = {"message": "Data written successfully to Sheet1"}
    svc = FileWorkbookService()
    grid = [[1, 2]]
    assert (
        svc.write_cell_grid("/abs/w.xlsx", "Sheet1", grid, "B2")
        == "Data written successfully to Sheet1"
    )
    mock_write.assert_called_once_with("/abs/w.xlsx", "Sheet1", grid, "B2")


@patch("excel_mcp.routing.file_workbook_service.wb_create_workbook")
def test_create_workbook(mock_create: MagicMock) -> None:
    svc = FileWorkbookService()
    out = svc.create_workbook("/abs/new.xlsx")
    mock_create.assert_called_once_with("/abs/new.xlsx")
    assert out == "Created workbook at /abs/new.xlsx"


@patch("excel_mcp.routing.file_workbook_service.load_workbook")
def test_save_workbook(mock_load_workbook: MagicMock) -> None:
    wb = MagicMock()
    mock_load_workbook.return_value = wb
    svc = FileWorkbookService()
    out = svc.save_workbook("/abs/existing.xlsx")
    mock_load_workbook.assert_called_once_with("/abs/existing.xlsx")
    wb.save.assert_called_once_with("/abs/existing.xlsx")
    wb.close.assert_called_once()
    assert out == "Workbook saved: /abs/existing.xlsx"


def test_routing_package_exports_file_workbook_service() -> None:
    from excel_mcp.routing import FileWorkbookService as FWS  # noqa: E402

    assert FWS is FileWorkbookService
