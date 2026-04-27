---
kind: story
id: STORY-11-1
title: COM-first routing for READ and WRITE, file fallback, and V1_FILE_FORCED
status: draft
parent: EPIC-11
depends_on: []
traces_to:
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/com-first-workbook-session-design.md
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
  - path: src/excel_mcp/routing/routing_backend.py
  - path: src/excel_mcp/routing/routed_dispatch.py
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - RoutingBackend removes unconditional READ → file short-circuit; ToolKind.READ follows the same COM-first / file fallback rules as writes when transport is auto, subject to COM viability and workbook identity match per ADR 0008 and com-first-workbook-session-design.
  - ToolKind.V1_FILE_FORCED continues to route to file per ADR 0004; tests or documented assertions prevent accidental COM selection for chart/pivot-forced tools.
  - routed_dispatch and reason codes support observable distinction between COM-selected reads and file fallback reads where applicable.
  - Unit and integration tests cover routing matrix (auto/com/file, READ vs WRITE, open vs not open workbook, cloud locator stubs as feasible).
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-1: COM-first routing for READ and WRITE, file fallback, and V1_FILE_FORCED

## Description

Implement **COM-first default routing** for **both reads and writes** when `transport=auto` and COM is viable, with **file/openpyxl fallback** when the workbook is not matched open or COM is unavailable. Remove **READ-specific unconditional file backing**. Preserve **chart/pivot file-forced** behavior per **ADR 0004**. Align structured logs and dispatch reasons with **[COM-first workbook session design](../../../architecture/com-first-workbook-session-design.md)** §1.

## User story

As an **operator**, I want **reads and writes to follow one routing philosophy** so that **Excel-hosted workbooks** use the host when possible without a separate “COM read opt-in.”

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Supersedes **Epic-10 / Story-10-1** opt-in model; trace decisions to **ADR 0008**, not ADR 0007 defaults.
- Coordinate with **ADR 0005** for strict mode when COM is requested but not viable.

## Dependencies (narrative)

Builds on **Epic-6** and **Epic-9**. Unblocks **Story-11-3** and COM read implementation stories.
