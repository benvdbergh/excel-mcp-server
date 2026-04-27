---
kind: story
id: STORY-12-3
title: Tests polish, TOOLS.md, README, changelog, manual validation
status: draft
parent: EPIC-12
depends_on:
  - STORY-12-2
traces_to:
  - path: docs/architecture/adr/0009-open-workbook-discovery-tool.md
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - TOOLS.md documents the discovery workflow (list → copy FullName → call get_workbook_metadata / reads), HTTPS vs local path reminder, COM-only requirement, and empty list when no workbooks open.
  - README “operator” or “quick start” section mentions the tool for agents replacing VBA Immediate / external scripts.
  - CHANGELOG.md entry follows project policy; SemVer bump plan noted for the release that ships the feature.
  - MANUAL-WINDOWS-RC-CHECKLIST.md (or equivalent) includes a short line to run discovery with multiple workbooks open and verify active flag.
  - Any new tests from 12-1/12-2 are integrated into the default CI path; coverage not regressed on Linux.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-12-3: Tests polish, TOOLS.md, README, changelog, manual validation

## Description

Finish **release-grade** documentation and validation for Epic-12: operator-facing prose, changelog, optional checklist updates, and consolidation of automated tests so **Epic-12** meets **Definition of Done**.

## User story

As an **operator**, I know **when and how** to use discovery with **SharePoint URLs** and existing **allowlist** behavior without reading source code.

## Technical notes

- Cross-link **[ADR 0009](../../../architecture/adr/0009-open-workbook-discovery-tool.md)** from TOOLS.md (short pointer).
- Align wording with **[target-architecture.md](../../../architecture/target-architecture.md)** §3 companion bullet once behavior ships.

## Dependencies (narrative)

Requires **STORY-12-2** so tool name and semantics are final.
