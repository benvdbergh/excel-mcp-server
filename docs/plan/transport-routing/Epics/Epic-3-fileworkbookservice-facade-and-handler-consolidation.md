---
kind: epic
id: EPIC-3
title: FileWorkbookService façade and handler consolidation
status: done
depends_on:
  - EPIC-2
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/pre-fork-architecture.md
slice: vertical
acceptance_criteria:
  - All file-based workbook operations for routed tools flow through FileWorkbookService (or documented exceptions pending COM).
  - Inline load_workbook usage in server.py is eliminated or justified with a follow-up story.
created: "2026-04-24"
updated: "2026-04-25"
---

# Epic-3: FileWorkbookService façade and handler consolidation

## Delivery summary

`FileWorkbookService` implements the full routed contract and `server.py` delegates routed tools to it (no inline `load_workbook` in handlers for those paths). Story 3-3 lifecycle hardening beyond the façade is **documented as deferred debt** in the module docstring of `src/excel_mcp/routing/file_workbook_service.py` (deeper modules still manage their own workbook handles).

## Description

Introduce **`FileWorkbookService`** as a thin façade over existing `workbook`, `sheet`, `data`, `formatting`, `calculations`, `chart`, `pivot`, `tables`, and `validation` modules (`FR-4`, target architecture §5).

## Objectives

- Centralize workbook lifecycle patterns where low-risk (target migration §1–3).
- Prepare **one injection point** for `RoutingBackend` to delegate to file vs COM.

## User stories (links)

- [Story-3-1](../Stories/Story-3-1-introduce-fileworkbookservice-delegating-to-existing-modules.md)
- [Story-3-2](../Stories/Story-3-2-consolidate-server-py-inline-workbook-loads-through-fileworkbookservice.md)
- [Story-3-3](../Stories/Story-3-3-workbook-lifecycle-consistency-close-on-reads-where-low-risk.md)

## Dependencies (narrative)

Requires normalized paths and policy from **EPIC-2** and the operation contract from **EPIC-1**.

## Related sources

- `docs/architecture/target-architecture.md` — §5 `FileWorkbookService`, migration step 2.
