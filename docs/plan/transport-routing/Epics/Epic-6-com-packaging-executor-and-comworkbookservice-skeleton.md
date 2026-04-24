---
kind: epic
id: EPIC-6
title: COM packaging, executor, and ComWorkbookService skeleton
status: draft
depends_on:
  - EPIC-5
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0002-com-automation-stack.md
slice: vertical
acceptance_criteria:
  - Optional com extra in packaging; default install remains lightweight (NFR-5).
  - COM entry points guarded; non-Windows skips COM import paths (FR-12, NFR-6).
  - All COM calls run on a single dedicated thread compatible with sync MCP handlers (FR-6).
  - ADR 0002 records the chosen library (pywin32 vs xlwings) and rationale in-repo.
created: "2026-04-24"
updated: "2026-04-24"
---

# Epic-6: COM packaging, executor, and ComWorkbookService skeleton

## Description

Add the **Windows-only COM stack** behind optional dependencies, implement the **single-thread COM worker**, and introduce **`ComWorkbookService`** implementing the shared contract for a minimal vertical slice (e.g. open workbook + noop or single write stub) to prove threading and packaging.

## Objectives

- Satisfy **FR-5**, **FR-6**, **NFR-4**, **NFR-5** at the infrastructure layer before broad tool parity.
- Keep Linux CI green with **mocks** only for COM behavior (`NFR-6`).

## User stories (links)

- [Story-6-1](../Stories/Story-6-1-optional-com-extra-and-import-guards-in-pyproject.md)
- [Story-6-2](../Stories/Story-6-2-single-thread-com-executor-compatible-with-sync-mcp-handlers.md)
- [Story-6-3](../Stories/Story-6-3-comworkbookservice-skeleton-bound-to-executor.md)

## Dependencies (narrative)

Requires **operator and schema wiring** from **EPIC-5** so COM can be selected safely in `auto` and `com` modes.

## Related sources

- `docs/architecture/adr/0002-com-automation-stack.md` — COM library decision.
- `docs/architecture/target-architecture.md` — §6–7 `ComWorkbookService`, COM execution model.
