---
kind: story
id: STORY-11-3
title: Wire com_do_op for all read-class MCP handlers
status: draft
parent: EPIC-11
depends_on:
  - STORY-11-1
traces_to:
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/com-read-class-tools-design.md
  - path: src/excel_mcp/routing/routed_dispatch.py
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - Every read-class handler passes a non-null com_do_op into the workbook dispatch path when COM execution is required by contract; no ComExecutionNotImplementedError for missing wiring on the read path after this story.
  - Parity with com-read-class-tools-design for handler list coverage; inventory or checklist updated if new read tools are added.
  - Tests or contract checks ensure new read handlers cannot register without com_do_op where ToolKind.READ applies.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-3: Wire `com_do_op` for all read-class MCP handlers

## Description

Complete **handler wiring** so **read-class tools** execute **`com_do_op`** on the COM path, matching the direction in **[com-read-class-tools-design.md](../../../architecture/com-read-class-tools-design.md)** and **ADR 0008** read-class execution. This story focuses on **plumbing**; **Story-11-4** and **Story-11-5** cover **ComWorkbookService** implementation depth.

## User story

As a **maintainer**, I want **every read tool** to be **routable to COM** without ad-hoc exceptions in dispatch.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Depends on **Story-11-1** so COM routing for READ is real before wiring cost is justified.
- Overlaps **Epic-10 Story-10-2** intent; Epic-11 assumes COM-first default, not opt-in.

## Dependencies (narrative)

**Story-11-1** required. Unblocks **Story-11-4** and **Story-11-5**.
