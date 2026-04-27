---
kind: story
id: STORY-11-6
title: Lifecycle — file create, excel_open_workbook, close workbook (save optional)
status: draft
parent: EPIC-11
depends_on:
  - STORY-11-1
traces_to:
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/com-first-workbook-session-design.md
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - File create semantics are aligned with create_workbook per ADR 0008 (single clarified tool and optional open-in-excel flag, or an explicitly documented split); no implicit save after create when save_after_write is gone.
  - Explicit open-in-Excel tool exists with MCP name documented (e.g. excel_open_workbook per ADR 0008); Workbooks.Open semantics bind the host session for subsequent COM-first operations on allowed paths or identities.
  - Close workbook in Excel exists with save optional (save-then-close vs discard); does not delete files; COM-thread sequencing per session design §5.
  - Path and URL allowlists apply; security notes from com-first-workbook-session-design §6 reflected in TOOLS.md.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-6: Lifecycle — file create, `excel_open_workbook`, close workbook (save optional)

## Description

Implement **ADR 0008** §2 **lifecycle** roles: **clarified file create** versus **`create_workbook`**, **`excel_open_workbook`** (name per ADR 0008; exact manifest string finalized in implementation), and **close workbook** with **optional save**. Optional **`ToolKind.SESSION`** or manifest-only tagging is decided during implementation per **[com-first-workbook-session-design](../../../architecture/com-first-workbook-session-design.md)** §7.

## User story

As an **agent author**, I want to **open and close workbooks in Excel explicitly** so **COM-first routing** applies without relying on pre-opened windows.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- **SSE jail** and HTTPS constraints from session design §4 must be reflected in error messages and docs.
- Sequencing with **Story-11-2**: creates must not reintroduce implicit save.

## Dependencies (narrative)

Depends on **Story-11-1** for consistent routing context; **Story-11-2** should be understood before final create/save behavior is frozen.
