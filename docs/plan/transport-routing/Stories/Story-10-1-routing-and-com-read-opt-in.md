---
kind: story
id: STORY-10-1
title: Read-class routing and COM read opt-in
status: draft
parent: EPIC-10
depends_on: []
traces_to:
  - path: docs/architecture/com-read-class-tools-design.md
  - path: docs/architecture/adr/0007-com-read-class-tools-routing.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: src/excel_mcp/routing/routing_backend.py
  - path: src/excel_mcp/routing/workbook_operation_contract.py
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - When COM read opt-in is off, ToolKind.READ behavior matches current v1 (file backend, reason read_class_file_backed or equivalent documented path).
  - When opt-in is on, resolve_workbook_backend can return com for READ per ADR 0007 candidate rules (transport=com, auto+open workbook, and/or cloud https locator that cannot use openpyxl), subject to COM viability and path/URL policy.
  - Inventory and WorkbookOperationMetadata reflect com_read_opt_in (or agreed equivalent) so operators and code share one source of truth.
  - Unit tests cover routing matrix for READ with opt-in on/off and cloud locator vs disk path stubs (mocks acceptable for COM viability).
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-10-1: Read-class routing and COM read opt-in

## Description

Extend **`RoutingBackend.resolve_workbook_backend`** and related metadata so **`ToolKind.READ`** is no longer forced to the file backend in all cases. Introduce the **explicit opt-in** for COM reads (tool parameter and/or environment variable) aligned with **`WorkbookOperationMetadata.com_read_opt_in`** and [ADR 0007](../../../architecture/adr/0007-com-read-class-tools-routing.md).

## User story

As an **operator**, I want **predictable routing for read tools** so that I can **keep default file-backed reads** or **enable COM reads** when I need cloud URLs or live Excel state, without silent behavior changes.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Revisit ordering: today **READ** may short-circuit before open-workbook detection; adjust per design note §1 and ADR 0007.
- Reuse **Epic-9** cloud locator detection and allowlist ([ADR 0006](../../../architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md)).
- Document interaction with **ADR 0005** strict mode for read failure modes in STORY-10-5 if not resolved here.

## Dependencies (narrative)

Builds on **Epic-6** (routing + contract) and **Epic-9** (cloud locators). Unblocks **STORY-10-2** (handlers need a routable COM path).
