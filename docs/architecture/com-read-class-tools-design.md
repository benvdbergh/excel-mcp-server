# COM read-class tools: design note (software architecture)

**Status:** **[ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md)** is the **accepted** product direction: **COM-first** reads when viable, **file** fallback otherwise. [ADR 0007](adr/0007-com-read-class-tools-routing.md) (**superseded**) described **file-first + opt-in COM reads**—the **gap list in §1** below records the **pre–Epic-11** state; mainline code now implements COM-first routing, `com_do_op` on read handlers, and **non-stub** `ComWorkbookService` read paths (see [Epic-11](../../plan/transport-routing/Epics/Epic-11-com-first-session-and-lifecycle.md)).

For session semantics, lifecycle tools, SSE jail, and routing matrix, see **[COM-first workbook session design](com-first-workbook-session-design.md)**.

This note records what it takes for **all read-class MCP tools** to execute against the **COM (Excel host)** backend as well as the **file / openpyxl** path—especially for **cloud HTTPS workbook locators** ([ADR 0006](adr/0006-cloud-workbook-locator-sharepoint-urls.md)) and **workbooks open in Excel**. It complements [ADR 0003](adr/0003-read-path-com-parity.md) (historical file-default Phase 1).

**Scope:** Routing, handler wiring, `ComWorkbookService` parity with `FileWorkbookService`, risks, and **migration** from today’s code to ADR 0008. Backward compatibility is **not** a product constraint pre-release.

---

## 1. **Historical** pre–Epic-11 behavior vs **current** (ADR 0008) implementation

The following subsections document the **old** state (read→file only, no `com_do_op`, COM read stubs) for **traceability**. **Current** `main` after **Epic 11** implements **COM-first** for `ToolKind.READ`, `com_do_op` on all read-class handlers, and **COM** implementations for the read contract methods in [`com_workbook_service.py`](../../src/excel_mcp/routing/com_workbook_service.py) (with `V1_FILE_FORCED` unchanged for chart/pivot per [ADR 0004](adr/0004-chart-pivot-com-parity-scope.md)).

### 1.1 (Historical) Routing: read-class always used the file backend

Previously, `read_class_file_backed` short-circuited `ToolKind.READ` to the file backend. **Current:** that branch is **removed**; `READ` uses the same resolution as `WRITE` (see `routing_backend.py`).

### 1.2 (Historical) Read handlers omitted `com_do_op`

**Current:** all read-class tools in the inventory pass `com_do_op` (see `server.py` and `tests/test_read_class_com_wiring.py`).

### 1.3 Cloud HTTPS locators and the file backend

[`execute_routed_workbook_operation`](../../src/excel_mcp/routing/routed_dispatch.py) still returns a **fixed error** when the **file** backend is selected for a pure `https` locator. **With COM-first reads**, when routing selects **COM** and the identity matches Excel, read tools can target **cloud** workbooks; when routing falls back to **file**, the HTTPS+file case remains invalid for openpyxl.

### 1.4 (Historical) `ComWorkbookService` read methods were stubs

**Current:** read methods run on the COM executor and return shapes aligned with the file façade (see implementation in `com_workbook_service.py`). `WorkbookOperationMetadata` / `com_read_opt_in` remain in the contract for forward use; default COM reads are **not** “opt-in only” (ADR 0008 supersedes that product gate).

---

## 2. Read vs mutate (authoritative inventory)

Source: [`tool_inventory.py`](../../src/excel_mcp/routing/tool_inventory.py) (single source of truth) and [`server.py`](../../src/excel_mcp/server.py) (tool registration).

| `ToolKind` | MCP tools (handler function name) |
|------------|-----------------------------------|
| **READ** | `validate_formula_syntax`, `read_data_from_excel`, `get_workbook_metadata`, `get_merged_cells`, `validate_excel_range`, `get_data_validation_info` |
| **WRITE** | All other mutating tools in the inventory except the chart/pivot exception below (e.g. `write_data_to_excel`, `apply_formula`, `format_range`, `create_workbook`, `save_workbook`, sheet/range operations, etc.) |
| **V1_FILE_FORCED** | `create_chart`, `create_pivot_table` ([ADR 0004](adr/0004-chart-pivot-com-parity-scope.md)) |

**Note:** `save_workbook` is **WRITE** in the inventory. **[ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md)** makes **`save_workbook` the sole persistence control** by **removing `save_after_write`** from mutating tools ([ADR 0003](adr/0003-read-path-com-parity.md) explicit-save intent, extended).

---

## 3. Gap summary

| Area | Gap |
|------|-----|
| **Routing** | `ToolKind.READ` forces file; COM cannot be selected for reads. |
| **Handlers** | Read tools omit `com_do_op`; COM branch would not run even if routing allowed it. |
| **Cloud URLs** | File backend rejects HTTPS locators; without COM reads, read tools cannot target SharePoint-style identities. |
| **COM implementation** | Read methods on `ComWorkbookService` are stubs; must mirror file semantics where required. |

---

## 4. What `ComWorkbookService` would need (mirror file contract)

[`FileWorkbookService`](../../src/excel_mcp/routing/file_workbook_service.py) delegates reads to:

- **`read_range_with_metadata`** → [`read_excel_range_with_metadata`](../../src/excel_mcp/data.py) → JSON with `range`, `sheet_name`, `cells[]` (`address`, `value`, `row`, `column`, optional `validation` per cell).
- **`workbook_metadata`** → [`get_workbook_info`](../../src/excel_mcp/workbook.py).
- **`read_merged_cell_ranges`** → [`get_merged_ranges`](../../src/excel_mcp/sheet.py).
- **`read_worksheet_data_validation`** → openpyxl + [`get_all_validation_ranges`](../../src/excel_mcp/cell_validation.py).
- **`validate_sheet_range`** → [`validate_range_in_sheet_operation`](../../src/excel_mcp/validation.py).
- **`validate_formula_syntax`** → [`validate_formula_in_cell_operation`](../../src/excel_mcp/validation.py) (file path loads workbook).

For **COM parity**, each operation needs a **threaded COM implementation** (existing pattern: `self._executor.submit(...)` on the Excel STA thread per [ADR 0002](adr/0002-com-automation-stack.md)) that:

1. **Resolves the workbook** using the same `_get_open_workbook_com` / `FullName` matching and errors as writes (not open, multiple match, Protected View, read-only where relevant)—see existing helpers in [`com_workbook_service.py`](../../src/excel_mcp/routing/com_workbook_service.py).
2. **`read_range_with_metadata`:** Build the same JSON shape as the file path. Likely use `Range.Value2` (or `Value` if formula text is required—**product choice**: `Value2` matches “calculated value” behavior; formula bar text needs `Formula` / `FormulaR1C1`). Per-cell **validation** via COM `Validation` API, mapped to the same schema consumers expect from openpyxl (`cell_validation` shape) or document intentional deltas. **`preview_only`:** today the file façade ignores this parameter; COM should either implement truncation consistently or keep documented parity with file.
3. **`workbook_metadata`:** Enumerate sheets, names, and optional ranges via COM (`Workbook.Worksheets`, `UsedRange`, etc.) and align string/JSON shape with `get_workbook_info` output or document differences.
4. **`read_merged_cell_ranges`:** `Range.MergeCells` / merged areas collection on the COM worksheet.
5. **`read_worksheet_data_validation`:** Walk `Worksheet.Validation` or per-range validation objects; map to the existing JSON structure where feasible.
6. **`validate_sheet_range`:** Structural checks (sheet exists, range parseable, within bounds) using COM objects.
7. **`validate_formula_syntax`:** Can reuse **pure** validation where possible (same `validate_formula` stack as `_apply_formula_com`); if the contract requires “cell context” from the file, mirror that or document COM-only behavior.

**Performance:** Large ranges over COM can be slower than openpyxl streaming; consider batch `Value2` on rectangular ranges (already one COM call per rectangle) vs per-cell loops.

---

## 5. Routing direction (ADR 0008)

[ADR 0007](adr/0007-com-read-class-tools-routing.md) options **(b)** separate read kinds and **(c)** env opt-in are **deprioritized**: the chosen approach is **(a)—extend `resolve_workbook_backend` for READ** using **COM-first defaults** (same matrix as writes). Optional **`ToolKind.SESSION`** applies to **open/close** tools only—see [com-first-workbook-session-design.md](com-first-workbook-session-design.md) §7.

| Area | Target |
|------|--------|
| **READ routing** | Same COM-first / fallback as WRITE; drop unconditional `read_class_file_backed`. |
| **Opt-in env for reads** | **Not** required for defaults ([ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md)). |
| **Inventory** | Existing **READ** list unchanged; COM execution behind routing. |

---

## 6. Risks

| Risk | Notes |
|------|--------|
| **COM threading** | All Excel automation must stay on the dedicated COM executor ([ADR 0002](adr/0002-com-automation-stack.md)); reads add more queued work. |
| **Performance** | COM round-trips for large ranges; possible timeouts vs fast disk reads. |
| **Formula vs value** | `Value2` vs displayed text vs formula string; must match documented MCP contract and file behavior. |
| **Protected View / read-only** | Already surfaced for COM writes; reads may need different policy (e.g. allow read in Protected View for inspection—**product decision**). |
| **SharePoint `FullName` matching** | [ADR 0006](adr/0006-cloud-workbook-locator-sharepoint-urls.md); normalized URL must match operator `filepath`. |
| **Strict COM** | [ADR 0005](adr/0005-com-strict-and-fallback-controls.md): if COM read fails, fallback to file may return **stale** data—dangerous; fail-closed may be preferable for reads. |

---

## 7. Behavior change: disk path while Excel has the file open

**Today:** For a **disk path**, read tools use **openpyxl on the file** even if Excel has the workbook open. The on-disk snapshot may **lag** unsaved Excel state ([ADR 0003](adr/0003-read-path-com-parity.md)).

**After ADR 0008 implementation:** **`transport=auto`** read tools prefer **COM** when the workbook is **open and matched**—**live Excel state** vs file snapshot. **`transport=file`** remains **always file** for deterministic openpyxl semantics.

**Persistence:** With **`save_after_write` removed**, agents call **`save_workbook`** when they need disk to reflect Excel before relying on file transport or external tools.

---

## 8. Cross-links (existing ADRs and docs)

| Document | Relevance |
|----------|-----------|
| [ADR 0001](adr/0001-workbook-transport-vs-mcp-wire-transport.md) | Workbook transport naming vs MCP wire transport. |
| [ADR 0002](adr/0002-com-automation-stack.md) | COM executor and threading. |
| [ADR 0003](adr/0003-read-path-com-parity.md) | v1 file-only reads; `save_workbook`; Phase 2 COM reads. |
| [ADR 0004](adr/0004-chart-pivot-com-parity-scope.md) | Chart/pivot file-forced routing. |
| [ADR 0005](adr/0005-com-strict-and-fallback-controls.md) | `com_strict` and file fallback. |
| [ADR 0006](adr/0006-cloud-workbook-locator-sharepoint-urls.md) | HTTPS locators; file backend rejection. |
| [ADR 0007](adr/0007-com-read-class-tools-routing.md) | Superseded; see ADR 0008. |
| [ADR 0008](adr/0008-com-first-default-and-file-lifecycle-tools.md) | COM-first default, lifecycle tools, `save_after_write` removal. |
| [com-first-workbook-session-design.md](com-first-workbook-session-design.md) | Routing matrix, session, SSE, threading, inventory. |
| [Target architecture](target-architecture.md) | High-level workbook routing story. |

---

## 9. Implementation checklist (ADR 0008)

1. Update **`resolve_workbook_backend`**: **READ** uses COM-first / fallback; remove read-only file gate.
2. Add **`com_do_op`** to each read handler in **`server.py`** mirroring file lambdas.
3. Replace **stubs** in **`ComWorkbookService`** with real COM implementations; align JSON with **`FileWorkbookService`**.
4. **Strip `save_after_write`** (and env plumbing as decided) from mutating tools per ADR 0008.
5. Add lifecycle tools (**open in Excel**, **close in Excel** with save flag; clarify **create** vs `create_workbook`) per [com-first-workbook-session-design.md](com-first-workbook-session-design.md).
6. Extend **integration tests**: cloud locator + COM, open workbook + **auto** read = COM, strict/fallback ([ADR 0005](adr/0005-com-strict-and-fallback-controls.md)).
7. Update **operator docs** (README, TOOLS.md): COM-first defaults, explicit save, lifecycle tools.
