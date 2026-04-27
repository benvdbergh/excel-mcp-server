"""Integration-style routing through server handlers (Epic 5 Story 5-3)."""

from __future__ import annotations

import json
import logging
from pathlib import Path

import pytest
from openpyxl import Workbook


def test_read_path_logs_routing_json(
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO, logger="excel-mcp.routing")
    p = tmp_path / "route_read.xlsx"
    Workbook().save(p)
    path = str(p.resolve())

    from excel_mcp import server as srv

    out = srv.get_workbook_metadata(path, workbook_transport="file")
    assert "Sheet" in out or "Error" not in out[:20]

    payloads = []
    for rec in caplog.records:
        if rec.name == "excel-mcp.routing" and rec.levelno == logging.INFO:
            try:
                payloads.append(json.loads(rec.getMessage()))
            except json.JSONDecodeError:
                continue
    assert payloads, "expected JSON routing log"
    last = payloads[-1]
    assert last["workbook_transport"] == "file"
    assert last["workbook_backend"] == "file"
    assert last["routing_reason"] == "read_class_file_backed"
    assert last["operation_name"] == "workbook_metadata"
    assert last["mcp_tool_name"] == "get_workbook_metadata"


def test_write_path_with_save_after_write(tmp_path: Path) -> None:
    p = tmp_path / "route_write.xlsx"
    Workbook().save(p)
    path = str(p.resolve())

    from excel_mcp import server as srv

    msg = srv.write_data_to_excel(
        path,
        "Sheet",
        [["x"]],
        start_cell="A1",
        workbook_transport="file",
        save_after_write=True,
    )
    assert "success" in msg.lower() or "written" in msg.lower() or "Error" not in msg
