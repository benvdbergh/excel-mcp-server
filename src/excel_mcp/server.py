import logging
import os
from typing import Callable, Optional

from mcp.server.fastmcp import FastMCP

from excel_mcp.com import support as com_support
from excel_mcp.com.executor import ComThreadExecutor
from excel_mcp.com.open_detection import ComWorkbookOpenInExcel
from excel_mcp.com.service import ComWorkbookService
from excel_mcp.fileio.service import FileWorkbookService
from excel_mcp.path.policy import (
    allowlist_enforced,
    assert_cloud_workbook_url_allowlist,
    assert_path_allowed,
    resolved_path_is_within as _resolved_path_is_within,
)
from excel_mcp.path.resolution import (
    is_cloud_workbook_locator,
    parse_cloud_workbook_locator,
    resolve_target,
)
from excel_mcp.routing.mcp_contract_bridge import contract_operation_name_for_mcp_tool
from excel_mcp.routing.routed_dispatch import (
    build_routed_response_envelope,
    execute_routed_workbook_operation,
)
from excel_mcp.routing.routing_backend import RoutingBackend
from excel_mcp.routing.routing_env import effective_com_strict, resolve_workbook_transport
from excel_mcp.routing.tool_inventory import get_tool_kind
from excel_mcp.routing.workbook_open_detection import StubWorkbookOpenInExcel
from excel_mcp.tools.read import register_read
from excel_mcp.tools.session import register_session
from excel_mcp.tools.sheet import register_sheet
from excel_mcp.tools.tables import register_tables
from excel_mcp.tools.write import register_write

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
    do_op: Callable[[str], str],
    com_do_op: Callable[[str], str] | None = None,
    *,
    include_routing_metadata: bool = False,
    response_warnings: list | None = None,
) -> str:
    """Resolve path, route transport, run one contract op."""
    full_path = get_excel_path(filepath)
    transport = resolve_workbook_transport(workbook_transport)
    com_strict = effective_com_strict()
    tool_kind = get_tool_kind(mcp_tool_name)
    operation_name = contract_operation_name_for_mcp_tool(mcp_tool_name)
    com_callable: Callable[[], str] | None = None
    if _COM_WORKBOOK_SERVICE is not None and com_do_op is not None:
        com_callable = lambda: com_do_op(full_path)
    out, _backend, routing_meta = execute_routed_workbook_operation(
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
    if include_routing_metadata:
        return build_routed_response_envelope(
            out,
            routing_meta,
            warnings=list(response_warnings or ()),
        )
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
        "Optional on tools: workbook_transport (auto|file|com). "
        "Discovery: excel_list_open_workbooks (COM) for open workbook FullName locators (ADR 0009). "
        "Lifecycle: excel_open_workbook, excel_close_workbook (COM). "
        "Recalc: evaluate_range (COM-only; does not save to disk). create_workbook optional open_in_excel. "
        "Env: EXCEL_MCP_TRANSPORT, EXCEL_MCP_ALLOWED_PATHS, EXCEL_MCP_ALLOWED_URL_PREFIXES (with path allowlist). "
        "Before first tool call in a session, inspect host MCP tool schemas for filepath, workbook_transport, and sheet_name; "
        "full contract in TOOLS.md. "
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


_registered_read = register_read(
    mcp,
    file_service=_FILE_WORKBOOK_SERVICE,
    workbook_dispatch=_workbook_dispatch,
    com_dispatch=_com_dispatch,
)
_registered_write = register_write(
    mcp,
    file_service=_FILE_WORKBOOK_SERVICE,
    workbook_dispatch=_workbook_dispatch,
    com_dispatch=_com_dispatch,
)
_registered_sheet = register_sheet(
    mcp,
    file_service=_FILE_WORKBOOK_SERVICE,
    workbook_dispatch=_workbook_dispatch,
    com_dispatch=_com_dispatch,
)
_registered_tables = register_tables(
    mcp,
    file_service=_FILE_WORKBOOK_SERVICE,
    workbook_dispatch=_workbook_dispatch,
    com_dispatch=_com_dispatch,
)
_registered_session = register_session(
    mcp,
    file_service=_FILE_WORKBOOK_SERVICE,
    get_com_service=lambda: _COM_WORKBOOK_SERVICE,
    workbook_dispatch=_workbook_dispatch,
    com_dispatch=_com_dispatch,
    get_excel_path=get_excel_path,
)

# Bind tool functions as module globals for tests that call excel_mcp.server.<tool>.
read_data_from_excel = _registered_read["read_data_from_excel"]
export_worksheet_table = _registered_read["export_worksheet_table"]
get_workbook_metadata = _registered_read["get_workbook_metadata"]
get_merged_cells = _registered_read["get_merged_cells"]
get_data_validation_info = _registered_read["get_data_validation_info"]
validate_excel_range = _registered_read["validate_excel_range"]
validate_formula_syntax = _registered_read["validate_formula_syntax"]

apply_formula = _registered_write["apply_formula"]
format_range = _registered_write["format_range"]
write_data_to_excel = _registered_write["write_data_to_excel"]

create_worksheet = _registered_sheet["create_worksheet"]
copy_worksheet = _registered_sheet["copy_worksheet"]
delete_worksheet = _registered_sheet["delete_worksheet"]
rename_worksheet = _registered_sheet["rename_worksheet"]
merge_cells = _registered_sheet["merge_cells"]
unmerge_cells = _registered_sheet["unmerge_cells"]
copy_range = _registered_sheet["copy_range"]
delete_range = _registered_sheet["delete_range"]
insert_rows = _registered_sheet["insert_rows"]
insert_columns = _registered_sheet["insert_columns"]
delete_sheet_rows = _registered_sheet["delete_sheet_rows"]
delete_sheet_columns = _registered_sheet["delete_sheet_columns"]

create_table = _registered_tables["create_table"]
create_chart = _registered_tables["create_chart"]
create_pivot_table = _registered_tables["create_pivot_table"]
list_tables = _registered_tables["list_tables"]
query_table = _registered_tables["query_table"]
map_sheet_layout = _registered_tables["map_sheet_layout"]
query_region = _registered_tables["query_region"]
apply_table_view = _registered_tables["apply_table_view"]
clear_table_view = _registered_tables["clear_table_view"]

create_workbook = _registered_session["create_workbook"]
excel_open_workbook = _registered_session["excel_open_workbook"]
excel_close_workbook = _registered_session["excel_close_workbook"]
excel_list_open_workbooks = _registered_session["excel_list_open_workbooks"]
evaluate_range = _registered_session["evaluate_range"]
save_workbook = _registered_session["save_workbook"]


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
