---
kind: story
id: STORY-12-1
title: COM — enumerate open workbooks and return stable JSON contract
status: draft
parent: EPIC-12
depends_on: []
traces_to:
  - path: docs/architecture/adr/0009-open-workbook-discovery-tool.md
  - path: docs/architecture/adr/0002-com-automation-stack.md
  - path: docs/architecture/com-first-workbook-session-design.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Code on the COM executor enumerates Application.Workbooks and produces a list of objects with at least full_name, name, and is_active (field names finalized in implementation; semantics per ADR 0009).
  - Active workbook detection is correct when multiple books are open; order of list is deterministic (e.g. same as Excel’s collection index) and documented.
  - When no workbooks are open, returns an empty list with no spurious errors; when Excel automation is unavailable, error is clear and consistent with existing COM failure patterns.
  - Unit or contract tests exercise enumeration with mocked Workbooks (no Excel required on Linux CI); any Windows-only tests are opt-in or marked.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-12-1: COM — enumerate open workbooks and stable JSON contract

## Description

Add a **COM-only** implementation that walks **`Workbooks`**, compares **active workbook** identity, and returns a **JSON-serializable** structure matching **[ADR 0009](../../../architecture/adr/0009-open-workbook-discovery-tool.md)** minimum payload. Prefer placement in **`ComWorkbookService`** or a small companion module imported only from COM code paths—**no** file-backend equivalent beyond a clear “use COM” error if someone calls the wrong layer.

## User story

As a **maintainer**, I want **enumerated open workbooks as tested JSON** so **MCP wiring** in Story 12-2 does not duplicate COM plumbing.

## Technical notes

- Reuse the **same single-thread COM executor** as other host operations ([com-first workbook session design §5](../../../architecture/com-first-workbook-session-design.md)).
- **Out of scope:** URL allowlist filtering of *returned* locators unless product later requires “redact disallowed URLs from list”—default is **report host truth**; subsequent tools enforce allowlist on *use* (ADR 0009 / session design §6).

## Dependencies (narrative)

**Epic-11** context assumed. No hard story-on-story dependency; **12-2** blocks on this.
