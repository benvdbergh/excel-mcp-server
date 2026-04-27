---
kind: story
id: STORY-10-3
title: COM implementation for read_range_with_metadata
status: draft
parent: EPIC-10
depends_on:
  - STORY-10-2
traces_to:
  - path: docs/architecture/com-read-class-tools-design.md
  - path: src/excel_mcp/routing/com_workbook_service.py
  - path: src/excel_mcp/routing/file_workbook_service.py
  - path: src/excel_mcp/data.py
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - ComWorkbookService.read_range_with_metadata returns the same JSON shape as the file path (range, sheet_name, cells with address/value/row/column and validation when applicable) for representative cases, or documents intentional deltas in STORY-10-5 deliverables.
  - Workbook resolution uses the same COM open/match/error behavior as write operations (not open, multiple match, protected view policy as decided).
  - Large ranges use batch COM access where practical; no cross-thread COM access outside the executor.
  - Automated tests cover at least one COM path or a fully mocked COM boundary; manual Windows note if required.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-10-3: COM implementation for read_range_with_metadata

## Description

Replace the **read_range_with_metadata** stub in **`ComWorkbookService`** with a real **threaded COM** implementation that mirrors **`FileWorkbookService`** → **`read_excel_range_with_metadata`**, per the design note §4.1–4.2 (Value2 vs formula text, per-cell validation mapping).

## User story

As an **operator**, I want **`read_data_from_excel`** to return **accurate cell data** from **Excel-hosted** workbooks (including **https** locators) so that **agents read the live grid** when COM reads are enabled.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Align **`preview_only`** with file façade (today may be ignored; keep consistent).
- Performance: prefer rectangular **Value2** reads; avoid per-cell COM in tight loops unless required for validation metadata.
- Coordinate with **STORY-10-4** on shared helpers (e.g. workbook get, sheet resolution).

## Dependencies (narrative)

Depends on **STORY-10-2** (COM dispatch wired). Can run in parallel with **STORY-10-4** after **10-2**.
