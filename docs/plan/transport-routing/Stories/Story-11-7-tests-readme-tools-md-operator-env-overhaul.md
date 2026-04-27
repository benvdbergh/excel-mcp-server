---
kind: story
id: STORY-11-7
title: Tests, README, TOOLS.md, and operator environment overhaul
status: draft
parent: EPIC-11
depends_on:
  - STORY-11-1
  - STORY-11-2
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
  - README and TOOLS.md document COM-first default, file fallback, removed save_after_write, new lifecycle tools, env vars (allowlist, URL prefixes, transport, strict/fallback), and breaking changes for agent authors.
  - MCP tool instructions and server use strings match shipped behavior; cloud URL and jail behavior are consistent with path_policy and server entrypoints.
  - Test suite coverage spans routing, dispatch, save contract, COM reads (where CI allows), and lifecycle tools; documented manual Windows checklist for gaps.
  - Operator environment section lists deprecated or removed variables and replacements.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-7: Tests, README, TOOLS.md, and operator environment overhaul

## Description

Consolidate **documentation** and **tests** for the **Epic-11** program: end-to-end operator clarity, **manifest** accuracy, and **regression** coverage for the **largest behavior change** since transport routing was introduced. This story absorbs the **Epic-10 Story-10-5** intent at Epic-11 scope.

## User story

As an **operator**, I can **configure and reason about** the server **from docs alone** after COM-first and lifecycle changes ship.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Continuous doc updates during **11-1**–**11-6** reduce end load; this story still owns final consistency pass and **LintPlan** validation if the project uses it.
- **ADR 0004** and **ADR 0005** cross-links in README for chart/pivot and strict mode.

## Dependencies (narrative)

Gates on substantive completion of **Story-11-1** and **Story-11-2**; finishes after or in parallel with **11-3**–**11-6** for final integration.
