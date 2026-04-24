---
kind: epic
id: EPIC-4
title: RoutingBackend and open-workbook detection (file-backed execution)
status: draft
depends_on:
  - EPIC-3
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
  - path: src/excel_mcp/routing/workbook_operation_contract.py
slice: vertical
acceptance_criteria:
  - resolve_workbook_backend implements auto/file/com selection with injectable workbook_open_in_excel (FR-2, FR-3, FR-10, FR-12).
  - Golden-path unit tests with mocks prove COM vs file *selection* for auto mode; executed I/O in this epic remains FileWorkbookService only (PRD AC1, NFR-1, NFR-6).
  - Each operation dispatched through RoutingBackend emits structured log fields transport, reason, duration_ms, workbook_path per policy (NFR-3).
  - transport=com when COM execution is not yet available follows ADR 0005 (clear error under strict; no silent file write where the matrix forbids it).
created: "2026-04-24"
updated: "2026-04-24"
---

# Epic-4: RoutingBackend and open-workbook detection (file-backed execution)

## Description

Add **`RoutingBackend`** and **`workbook_open_in_excel`** behind an injectable port. **Selection** (auto / file / com, reason codes, strict semantics) is fully exercised with **mocks**; **executed** workbook I/O in this epic remains **`FileWorkbookService`** only so CI stays Excel-free (`NFR-6`). This matches target-architecture **migration step 3** (router before real COM).

## Prerequisites (Epic 3 baseline — done)

Epic 3 should be **closed** before starting Epic 4. The following is assumed in the codebase:

- **`FileWorkbookService`** (`src/excel_mcp/routing/file_workbook_service.py`) implements **`RoutedWorkbookOperations`** and centralizes file-side workbook operations.
- **`server.py`** resolves paths then calls that façade (e.g. module-level service instance); **no** inline **`load_workbook`** in handlers for routed tools.
- Shared contract and operation names live in **`src/excel_mcp/routing/workbook_operation_contract.py`** (with **`operation_metadata`** / `tool_kind` hooks for later read parity per **ADR 0003**).

Epic 4 **does not** redesign the façade; it **sits between** path resolution / policy and **`FileWorkbookService`**, and later **`ComWorkbookService`**.

## Objectives

- Implement the **transport selection matrix** from the blueprint / PRD with explicit **reason codes** for logs.
- Enforce **no auto-start of Excel** for routing (`FR-10`).
- On non-Windows, COM module skipped and `com` mode surfaces a clear unsupported error (`FR-12`).
- Carry **`tool_kind`** (read vs write, inventory alignment) through routing context for ADR 0003 / future COM reads without another contract break.

## Scope boundary vs Epic 5

| Epic | Responsibility |
|------|----------------|
| **Epic 4** | **`RoutingBackend`** implementation, injectable **`workbook_open_in_excel`**, file-only **execution**, **structured logging** on router dispatch, **tests** proving the matrix and logging. |
| **Epic 5** | Operator defaults (**`EXCEL_MCP_TRANSPORT`**, **`EXCEL_MCP_COM_STRICT`** per ADR 0005), **optional MCP tool parameters** (`workbook_transport`, `save_after_write` per **ADR 0001** / FR-7–8), and **Story-5-3**: refactor **`server.py`** so every routed tool uses **`RoutingBackend`** as the single gate instead of calling **`FileWorkbookService`** directly. |

Epic 4 may introduce a small **dispatch helper** or test-only entry points so Story-4-2/4-3 can prove behavior end-to-end without duplicating Epic 5’s full handler + schema work.

## User stories (links)

- [Story-4-1](../Stories/Story-4-1-injectable-workbook-open-in-excel-port-with-test-doubles.md)
- [Story-4-2](../Stories/Story-4-2-routingbackend-selection-matrix-with-file-only-execution-path.md)
- [Story-4-3](../Stories/Story-4-3-structured-logging-for-every-routed-operation.md)

## Dependencies (narrative)

Requires **`FileWorkbookService`** and consolidated handler → façade usage from **EPIC-3**. Epic 5 depends on this epic for router behavior, logging, and detection hooks.

## Related sources

- `docs/architecture/target-architecture.md` — §3–4 `workbook_open_in_excel`, `RoutingBackend`, layered view, migration step 3.
- `docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md` — vocabulary for logs and config.
- `docs/architecture/adr/0003-read-path-com-parity.md` — default file-backed reads; router metadata for future COM reads.
- `docs/architecture/adr/0005-com-strict-and-fallback-controls.md` — `transport=com` vs strict when COM execution is not available.
