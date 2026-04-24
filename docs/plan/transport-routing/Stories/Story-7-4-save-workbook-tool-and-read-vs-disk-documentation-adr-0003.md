---
kind: story
id: STORY-7-4
title: save_workbook tool and read vs disk documentation (ADR 0003)
status: draft
parent: EPIC-7
depends_on:
  - STORY-7-1
traces_to:
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: docs/architecture/target-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: false
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - New save_workbook MCP tool registered in server.py, manifest.json, and TOOLS.md per ADR 0003 inventory note.
  - README explains file-backed reads and agent pattern save_workbook before read_data when using COM without per-write save.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-7-4: save_workbook tool and read vs disk documentation (ADR 0003)

## Description

Add **`save_workbook`** routed like other **write-class** tools so agents can **persist Excel host state** before file reads (**ADR 0003**, target architecture tool table).

## User story

As an **agent author**, I want an **explicit save** after COM mutations so that **subsequent file reads** reflect what Excel flushed to disk.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- COM path: `Workbook.Save` / `SaveAs` with threading and errors from **STORY-7-2**.
- File path: document idempotent or no-op semantics when already persisted.

## Dependencies (narrative)

Depends on **STORY-7-1** for routed COM writes baseline.
