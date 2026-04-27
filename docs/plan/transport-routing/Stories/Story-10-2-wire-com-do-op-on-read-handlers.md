---
kind: story
id: STORY-10-2
title: Wire com_do_op for all read-class MCP handlers
status: draft
parent: EPIC-10
depends_on:
  - STORY-10-1
traces_to:
  - path: docs/architecture/com-read-class-tools-design.md
  - path: src/excel_mcp/server.py
  - path: src/excel_mcp/routing/routed_dispatch.py
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Every read-class tool registration path that uses _workbook_dispatch passes a com_do_op analogous to write tools (e.g. via _com_dispatch), so execute_routed_workbook_operation never sees a null COM callable solely because the tool is a read.
  - Read tools remain file-first by default; COM branch executes only when routing selects com and opt-in allows (STORY-10-1).
  - Integration or unit tests assert com_do_op is passed for each read handler name in tool_inventory ToolKind.READ (table-driven).
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-10-2: Wire com_do_op for all read-class MCP handlers

## Description

Today **read** tools pass only the file `do_op` lambda into **`_workbook_dispatch`**, so the COM branch cannot run even if routing changes. Mirror the **write** pattern: supply **`com_do_op`** that delegates to **`ComWorkbookService`** for the same contract operation as the file façade.

## User story

As a **maintainer**, I want **symmetric dispatch** for read and write tools so that **routing decisions** actually drive execution without **`ComExecutionNotImplementedError`** from a missing callable.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Tools: `read_data_from_excel`, `get_workbook_metadata`, `get_merged_cells`, `validate_excel_range`, `get_data_validation_info`, `validate_formula_syntax` (per [tool inventory design](../../../architecture/com-read-class-tools-design.md) §2).
- COM method bodies may remain stubs until **STORY-10-3** / **10-4**; this story completes the **wiring** end-to-end.

## Dependencies (narrative)

Depends on **STORY-10-1** (read routing must be coherent before end-to-end tests). Enables parallel **STORY-10-3** and **STORY-10-4**.
