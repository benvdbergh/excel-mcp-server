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
slice: vertical
acceptance_criteria:
  - resolve_workbook_backend implements auto/file/com selection with injectable workbook_open_in_excel (FR-2, FR-3, FR-10, FR-12).
  - Golden-path unit tests with mocks prove COM vs file selection for auto mode (PRD AC1, NFR-1).
  - Each routed call emits structured log fields transport, reason, duration_ms, workbook_path per policy (NFR-3).
created: "2026-04-24"
updated: "2026-04-24"
---

# Epic-4: RoutingBackend and open-workbook detection (file-backed execution)

## Description

Add **`RoutingBackend`** and **`workbook_open_in_excel`** behind an injectable port. Initially all **executed** I/O remains **file**-based; COM selection is exercised via **mocks** so CI stays Excel-free (`NFR-6`).

## Objectives

- Implement the **transport selection matrix** from the blueprint / PRD with explicit reason codes for logs.
- Enforce **no auto-start of Excel** for routing (`FR-10`).
- On non-Windows, COM module skipped and `com` mode surfaces a clear unsupported error (`FR-12`).

## User stories (links)

- [Story-4-1](../Stories/Story-4-1-injectable-workbook-open-in-excel-port-with-test-doubles.md)
- [Story-4-2](../Stories/Story-4-2-routingbackend-selection-matrix-with-file-only-execution-path.md)
- [Story-4-3](../Stories/Story-4-3-structured-logging-for-every-routed-operation.md)

## Dependencies (narrative)

Requires **`FileWorkbookService`** from **EPIC-3** as the default backend implementation.

## Related sources

- `docs/architecture/target-architecture.md` — §3–4 `workbook_open_in_excel`, `RoutingBackend`.
- `docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md` — vocabulary for logs and config.
