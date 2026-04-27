import logging
import os
from typing import Any, Callable, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

# Import exceptions
from excel_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError
)

from excel_mcp import com_support
from excel_mcp.com_executor import ComThreadExecutor
from excel_mcp.routing import (
    ComWorkbookService,
    FileWorkbookService,
    RoutingBackend,
    StubWorkbookOpenInExcel,
    ToolKind,
    contract_operation_name_for_mcp_tool,
    effective_com_strict,
    effective_save_after_write,
    execute_routed_workbook_operation,
    get_tool_kind,
    resolve_workbook_transport,
)
from excel_mcp.routing.com_workbook_open_detection import ComWorkbookOpenInExcel
from excel_mcp.routing.routing_errors import (
    ComExecutionNotImplementedError,
    ComRoutingError,
)
from excel_mcp.path_policy import (
    allowlist_enforced,
    assert_cloud_workbook_url_allowlist,
    assert_path_allowed,
    resolved_path_is_within as _resolved_path_is_within,
)
from excel_mcp.path_resolution import (
    is_cloud_workbook_locator,
    parse_cloud_workbook_locator,
    resolve_target,
)

# Get project root directory path for log file path.
# When using the stdio transmission method,
# relative paths may cause log files to fail to create
# due to the client's running location and permission issues,
# resulting in the program not being able to run.
# Thus using os.path.join(ROOT_DIR, "excel-mcp.log") instead.

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_FILE = os.path.join(ROOT_DIR, "excel-mcp.log")

# Initialize EXCEL_FILES_PATH variable without assigning a value
EXCEL_FILES_PATH = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        # Referring to https://github.com/modelcontextprotocol/python-sdk/issues/409#issuecomment-2816831318
        # The stdio mode server MUST NOT write anything to its stdout that is not a valid MCP message.
        logging.FileHandler(LOG_FILE)
    ],
)
logger = logging.getLogger("excel-mcp")
_FILE_WORKBOOK_SERVICE = FileWorkbookService()
com_execution_available = com_support.is_com_runtime_supported()
_COM_WORKBOOK_SERVICE: ComWorkbookService | None = None
if com_execution_available:
    _COM_EXECUTOR = ComThreadExecutor()
    _COM_WORKBOOK_SERVICE = ComWorkbookService(_COM_EXECUTOR)
else:
    _COM_EXECUTOR = None
_workbook_open = (
    ComWorkbookOpenInExcel(_COM_EXECUTOR)
    if com_execution_available and _COM_EXECUTOR is not None
    else StubWorkbookOpenInExcel()
)
_ROUTING_BACKEND = RoutingBackend(
    _workbook_open,
    com_execution_available=com_execution_available,
)


def _workbook_dispatch(
    mcp_tool_name: str,
    filepath: str,
    workbook_transport: Optional[str],
    save_after_write: Optional[bool],
    do_op: Callable[[str], str],
    com_do_op: Callable[[str], str] | None = None,
) -> str:
    """Resolve path, route transport, run one contract op, optional explicit save."""
    full_path = get_excel_path(filepath)
    transport = resolve_workbook_transport(workbook_transport)
    com_strict = effective_com_strict()
    tool_kind = get_tool_kind(mcp_tool_name)
    operation_name = contract_operation_name_for_mcp_tool(mcp_tool_name)
    com_callable: Callable[[], str] | None = None
    if _COM_WORKBOOK_SERVICE is not None and com_do_op is not None:
        com_callable = lambda: com_do_op(full_path)
    out, backend = execute_routed_workbook_operation(
        _ROUTING_BACKEND,
        _FILE_WORKBOOK_SERVICE,
        resolved_path=full_path,
        workbook_transport=transport,
        tool_kind=tool_kind,
        com_strict=com_strict,
        operation_name=operation_name,
        operation_callable=lambda: do_op(full_path),
        com_operation_callable=com_callable,
        mcp_tool_name=mcp_tool_name,
    )
    if tool_kind != ToolKind.READ and effective_save_after_write(save_after_write):
        if backend == "file":
            _FILE_WORKBOOK_SERVICE.save_workbook(full_path)
        elif backend == "com" and _COM_WORKBOOK_SERVICE is not None:
            _COM_WORKBOOK_SERVICE.save_workbook(full_path)
    return out


def _com_dispatch(com_fn: Callable[[ComWorkbookService, str], str]) -> Callable[[str], str] | None:
    """Build ``com_do_op`` for :func:`_workbook_dispatch` when COM service is enabled."""
    if _COM_WORKBOOK_SERVICE is None:
        return None
    svc = _COM_WORKBOOK_SERVICE
    return lambda fp: com_fn(svc, fp)


# Initialize FastMCP server
mcp = FastMCP(
    "excel-mcp",
    host=os.environ.get("FASTMCP_HOST", "0.0.0.0"),
    port=int(os.environ.get("FASTMCP_PORT", "8017")),
    instructions=(
        "Excel MCP server: create/read/edit .xlsx workbooks (openpyxl; optional Windows COM). "
        "Parameter filepath: absolute disk path, OR for COM/cloud workbooks the exact https SharePoint-style "
        "URL that matches Excel Workbook.FullName (in VBA Immediate use ? ActiveWorkbook.FullName). "
        "If Excel reports https but you pass only a local synced path, COM may not match. "
        "M365 sign-in is via Excel/Office, not this server. "
        "Optional on tools: workbook_transport (auto|file|com), save_after_write. "
        "Env: EXCEL_MCP_TRANSPORT, EXCEL_MCP_ALLOWED_PATHS, EXCEL_MCP_ALLOWED_URL_PREFIXES (with path allowlist). "
        "Full operator docs: repository README and TOOLS.md; local Cursor MCP: README section on uv run --project."
    ),
)


def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.

    Args:
        filename: Name of Excel file

    Returns:
        Full path to Excel file
    """
    if not filename or "\x00" in filename:
        raise ValueError(f"Invalid filename: {filename}")

    # Cloud workbook locators (HTTPS): avoid resolve_target / os.path.realpath (Story 9-1 / ADR 0006).
    if is_cloud_workbook_locator(filename):
        if EXCEL_FILES_PATH is not None:
            raise ValueError(
                "Cloud workbook URLs (HTTPS) are not supported when EXCEL_FILES_PATH is set "
                "(SSE/HTTP jail). Use filesystem paths under the jail root, or run the server "
                "without EXCEL_FILES_PATH for HTTPS workbook identity strings."
            )
        canonical = parse_cloud_workbook_locator(filename)
        if allowlist_enforced():
            assert_cloud_workbook_url_allowlist(canonical)
        return canonical

    if EXCEL_FILES_PATH is None:
        if not os.path.isabs(filename):
            raise ValueError(f"Invalid filename: {filename}, must be an absolute path when not in SSE mode")
        if not allowlist_enforced():
            return os.path.normpath(filename)
        resolved = resolve_target(filename)
        assert_path_allowed(resolved, jail_realpath=None)
        return resolved

    if os.path.isabs(filename):
        raise ValueError(f"Invalid filename: {filename}, must be relative to EXCEL_FILES_PATH")

    base = os.path.realpath(EXCEL_FILES_PATH)
    candidate = resolve_target(filename, cwd=base)
    assert_path_allowed(candidate, jail_realpath=base)
    return candidate

@mcp.tool(
    annotations=ToolAnnotations(
        title="Apply Formula",
        destructiveHint=True,
    ),
)
def apply_formula(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """
    Apply Excel formula to cell.
    Excel formula will write to cell with verification.
    """
    try:
        return _workbook_dispatch(
            "apply_formula",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.apply_formula(
                fp, sheet_name, cell, formula
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.apply_formula(fp, sheet_name, cell, formula)
            ),
        )
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error applying formula: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Validate Formula Syntax",
        readOnlyHint=True,
    ),
)
def validate_formula_syntax(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Validate Excel formula syntax without applying it."""
    try:
        return _workbook_dispatch(
            "validate_formula_syntax",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.validate_formula_syntax(
                fp, sheet_name, cell, formula
            ),
        )
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating formula: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Format Range",
        destructiveHint=True,
    ),
)
def format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    border_style: Optional[str] = None,
    border_color: Optional[str] = None,
    number_format: Optional[str] = None,
    alignment: Optional[str] = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Optional[Dict[str, Any]] = None,
    conditional_format: Optional[Dict[str, Any]] = None,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Apply formatting to a range of cells."""
    try:
        return _workbook_dispatch(
            "format_range",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.format_range(
                fp,
                sheet_name,
                start_cell,
                end_cell,
                bold=bold,
                italic=italic,
                underline=underline,
                font_size=font_size,
                font_color=font_color,
                bg_color=bg_color,
                border_style=border_style,
                border_color=border_color,
                number_format=number_format,
                alignment=alignment,
                wrap_text=wrap_text,
                merge_cells=merge_cells,
                protection=protection,
                conditional_format=conditional_format,
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.format_range(
                    fp,
                    sheet_name,
                    start_cell,
                    end_cell,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    font_size=font_size,
                    font_color=font_color,
                    bg_color=bg_color,
                    border_style=border_style,
                    border_color=border_color,
                    number_format=number_format,
                    alignment=alignment,
                    wrap_text=wrap_text,
                    merge_cells=merge_cells,
                    protection=protection,
                    conditional_format=conditional_format,
                )
            ),
        )
    except (ValidationError, FormattingError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error formatting range: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Read Data from Excel",
        readOnlyHint=True,
    ),
)
def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """
    Read data from Excel worksheet with cell metadata including validation rules.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell (default A1)
        end_cell: Ending cell (optional, auto-expands if not provided)
        preview_only: Whether to return preview only
    
    Returns:  
    JSON string containing structured cell data with validation metadata.
    Each cell includes: address, value, row, column, and validation info (if any).
    """
    try:
        return _workbook_dispatch(
            "read_data_from_excel",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.read_range_with_metadata(
                fp,
                sheet_name,
                start_cell,
                end_cell,
                preview_only,
            ),
        )
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Write Data to Excel",
        destructiveHint=True,
    ),
)
def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1",
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.

    PARAMETERS:  
    filepath: Path to Excel file
    sheet_name: Name of worksheet to write to
    data: List of lists containing data to write to the worksheet, sublists are assumed to be rows
    start_cell: Cell to start writing to, default is "A1"
  
    """
    try:
        return _workbook_dispatch(
            "write_data_to_excel",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.write_cell_grid(
                fp, sheet_name, data, start_cell
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.write_cell_grid(fp, sheet_name, data, start_cell)
            ),
        )
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Workbook",
        destructiveHint=True,
    ),
)
def create_workbook(
    filepath: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Create new Excel workbook."""
    try:
        return _workbook_dispatch(
            "create_workbook",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.create_workbook(fp),
            com_do_op=_com_dispatch(lambda c, fp: c.create_workbook(fp)),
        )
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Save Workbook",
        destructiveHint=True,
    ),
)
def save_workbook(
    filepath: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Persist the workbook to disk (file backend or COM host save).

    Use this before ``read_data_from_excel`` when mutations ran via COM with
    ``save_after_write=false`` so on-disk state matches Excel (ADR 0003).

    The optional ``save_after_write`` follows the same env default as other
    tools: when ``true``, the server runs a second explicit save after the
    primary ``save_workbook`` operation. That is idempotent (save then save).
    When omitted or ``false`` (default), only one save runs.
    """
    try:
        return _workbook_dispatch(
            "save_workbook",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.save_workbook(fp),
            com_do_op=_com_dispatch(lambda c, fp: c.save_workbook(fp)),
        )
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error saving workbook: {e}")
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Worksheet",
        destructiveHint=True,
    ),
)
def create_worksheet(
    filepath: str,
    sheet_name: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Create new worksheet in workbook."""
    try:
        return _workbook_dispatch(
            "create_worksheet",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.create_worksheet(fp, sheet_name),
            com_do_op=_com_dispatch(
                lambda c, fp: c.create_worksheet(fp, sheet_name)
            ),
        )
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating worksheet: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Chart",
        destructiveHint=True,
    ),
)
def create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = "",
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Create chart in worksheet."""
    try:
        return _workbook_dispatch(
            "create_chart",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.create_chart_in_sheet(
                fp,
                sheet_name,
                data_range,
                chart_type,
                target_cell,
                title=title,
                x_axis=x_axis,
                y_axis=y_axis,
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.create_chart_in_sheet(
                    fp,
                    sheet_name,
                    data_range,
                    chart_type,
                    target_cell,
                    title=title,
                    x_axis=x_axis,
                    y_axis=y_axis,
                )
            ),
        )
    except (ValidationError, ChartError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating chart: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Pivot Table",
        destructiveHint=True,
    ),
)
def create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: Optional[List[str]] = None,
    agg_func: str = "mean",
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Create pivot table in worksheet."""
    try:
        return _workbook_dispatch(
            "create_pivot_table",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.create_pivot_table_in_sheet(
                fp,
                sheet_name,
                data_range,
                rows,
                values,
                columns=columns,
                agg_func=agg_func,
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.create_pivot_table_in_sheet(
                    fp,
                    sheet_name,
                    data_range,
                    rows,
                    values,
                    columns=columns,
                    agg_func=agg_func,
                )
            ),
        )
    except (ValidationError, PivotError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating pivot table: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Table",
        destructiveHint=True,
    ),
)
def create_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str] = None,
    table_style: str = "TableStyleMedium9",
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Creates a native Excel table from a specified range of data."""
    try:
        return _workbook_dispatch(
            "create_table",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.create_excel_table(
                fp,
                sheet_name,
                data_range,
                table_name=table_name,
                table_style=table_style,
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.create_excel_table(
                    fp,
                    sheet_name,
                    data_range,
                    table_name=table_name,
                    table_style=table_style,
                )
            ),
        )
    except DataError as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating table: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Copy Worksheet",
        destructiveHint=True,
    ),
)
def copy_worksheet(
    filepath: str,
    source_sheet: str,
    target_sheet: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Copy worksheet within workbook."""
    try:
        return _workbook_dispatch(
            "copy_worksheet",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.copy_worksheet(
                fp, source_sheet, target_sheet
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.copy_worksheet(fp, source_sheet, target_sheet)
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying worksheet: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Worksheet",
        destructiveHint=True,
    ),
)
def delete_worksheet(
    filepath: str,
    sheet_name: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Delete worksheet from workbook."""
    try:
        return _workbook_dispatch(
            "delete_worksheet",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.delete_worksheet(fp, sheet_name),
            com_do_op=_com_dispatch(
                lambda c, fp: c.delete_worksheet(fp, sheet_name)
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting worksheet: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Rename Worksheet",
        destructiveHint=True,
    ),
)
def rename_worksheet(
    filepath: str,
    old_name: str,
    new_name: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Rename worksheet in workbook."""
    try:
        return _workbook_dispatch(
            "rename_worksheet",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.rename_worksheet(
                fp, old_name, new_name
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.rename_worksheet(fp, old_name, new_name)
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error renaming worksheet: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Workbook Metadata",
        readOnlyHint=True,
    ),
)
def get_workbook_metadata(
    filepath: str,
    include_ranges: bool = False,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Get metadata about workbook including sheets, ranges, etc."""
    try:
        return _workbook_dispatch(
            "get_workbook_metadata",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.workbook_metadata(
                fp, include_ranges=include_ranges
            ),
        )
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting workbook metadata: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Merge Cells",
        destructiveHint=True,
    ),
)
def merge_cells(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Merge a range of cells."""
    try:
        return _workbook_dispatch(
            "merge_cells",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.merge_cells(
                fp, sheet_name, start_cell, end_cell
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.merge_cells(
                    fp, sheet_name, start_cell, end_cell
                )
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error merging cells: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Unmerge Cells",
        destructiveHint=True,
    ),
)
def unmerge_cells(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Unmerge a range of cells."""
    try:
        return _workbook_dispatch(
            "unmerge_cells",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.unmerge_cells(
                fp, sheet_name, start_cell, end_cell
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.unmerge_cells(
                    fp, sheet_name, start_cell, end_cell
                )
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error unmerging cells: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Merged Cells",
        readOnlyHint=True,
    ),
)
def get_merged_cells(
    filepath: str,
    sheet_name: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Get merged cells in a worksheet."""
    try:
        return _workbook_dispatch(
            "get_merged_cells",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.read_merged_cell_ranges(
                fp, sheet_name
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting merged cells: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Copy Range",
        destructiveHint=True,
    ),
)
def copy_range(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: Optional[str] = None,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Copy a range of cells to another location."""
    try:
        return _workbook_dispatch(
            "copy_range",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.copy_cell_range(
                fp,
                sheet_name,
                source_start,
                source_end,
                target_start,
                target_sheet=target_sheet,
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.copy_cell_range(
                    fp,
                    sheet_name,
                    source_start,
                    source_end,
                    target_start,
                    target_sheet=target_sheet,
                )
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying range: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Range",
        destructiveHint=True,
    ),
)
def delete_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up",
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Delete a range of cells and shift remaining cells."""
    try:
        return _workbook_dispatch(
            "delete_range",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.delete_cell_range(
                fp,
                sheet_name,
                start_cell,
                end_cell,
                shift_direction,
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.delete_cell_range(
                    fp,
                    sheet_name,
                    start_cell,
                    end_cell,
                    shift_direction,
                )
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting range: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Validate Excel Range",
        readOnlyHint=True,
    ),
)
def validate_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Validate if a range exists and is properly formatted."""
    try:
        return _workbook_dispatch(
            "validate_excel_range",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.validate_sheet_range(
                fp, sheet_name, start_cell, end_cell
            ),
        )
    except ValidationError as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating range: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Data Validation Info",
        readOnlyHint=True,
    ),
)
def get_data_validation_info(
    filepath: str,
    sheet_name: str,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """
    Get all data validation rules in a worksheet.
    
    This tool helps identify which cell ranges have validation rules
    and what types of validation are applied.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        
    Returns:
        JSON string containing all validation rules in the worksheet
    """
    try:
        return _workbook_dispatch(
            "get_data_validation_info",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.read_worksheet_data_validation(
                fp, sheet_name
            ),
        )
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting validation info: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Insert Rows",
        destructiveHint=True,
    ),
)
def insert_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Insert one or more rows starting at the specified row."""
    try:
        return _workbook_dispatch(
            "insert_rows",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.insert_rows(
                fp, sheet_name, start_row, count
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.insert_rows(fp, sheet_name, start_row, count)
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error inserting rows: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Insert Columns",
        destructiveHint=True,
    ),
)
def insert_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Insert one or more columns starting at the specified column."""
    try:
        return _workbook_dispatch(
            "insert_columns",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.insert_columns(
                fp, sheet_name, start_col, count
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.insert_columns(fp, sheet_name, start_col, count)
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error inserting columns: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Rows",
        destructiveHint=True,
    ),
)
def delete_sheet_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Delete one or more rows starting at the specified row."""
    try:
        return _workbook_dispatch(
            "delete_sheet_rows",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.delete_sheet_rows(
                fp, sheet_name, start_row, count
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.delete_sheet_rows(
                    fp, sheet_name, start_row, count
                )
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting rows: {e}")
        raise

@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Columns",
        destructiveHint=True,
    ),
)
def delete_sheet_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1,
    workbook_transport: Optional[str] = None,
    save_after_write: Optional[bool] = None,
) -> str:
    """Delete one or more columns starting at the specified column."""
    try:
        return _workbook_dispatch(
            "delete_sheet_columns",
            filepath,
            workbook_transport,
            save_after_write,
            lambda fp: _FILE_WORKBOOK_SERVICE.delete_sheet_columns(
                fp, sheet_name, start_col, count
            ),
            com_do_op=_com_dispatch(
                lambda c, fp: c.delete_sheet_columns(
                    fp, sheet_name, start_col, count
                )
            ),
        )
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except (ComRoutingError, ComExecutionNotImplementedError, ValueError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting columns: {e}")
        raise

def run_sse():
    """Run Excel MCP server in SSE mode."""
    # Assign value to EXCEL_FILES_PATH in SSE mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with SSE transport (files directory: {EXCEL_FILES_PATH})")
        mcp.run(transport="sse")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

def run_streamable_http():
    """Run Excel MCP server in streamable HTTP mode."""
    # Assign value to EXCEL_FILES_PATH in streamable HTTP mode
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    # Create directory if it doesn't exist
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
    
    try:
        logger.info(f"Starting Excel MCP server with streamable HTTP transport (files directory: {EXCEL_FILES_PATH})")
        mcp.run(transport="streamable-http")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

def run_stdio():
    """Run Excel MCP server in stdio mode."""
    # No need to assign EXCEL_FILES_PATH in stdio mode
    
    try:
        logger.info("Starting Excel MCP server with stdio transport")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")