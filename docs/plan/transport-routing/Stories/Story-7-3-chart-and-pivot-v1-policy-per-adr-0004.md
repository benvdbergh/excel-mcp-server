---
kind: story
id: STORY-7-3
title: Chart and pivot v1 policy per ADR 0004
status: draft
parent: EPIC-7
depends_on:
  - STORY-7-1
traces_to:
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - README and tool descriptions state whether create_chart and create_pivot_table force file backend or allow routed COM with documented differences (per chosen ADR 0004 policy).
  - Router or tool layer encodes the chosen policy with tests preventing silent drift.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-7-3: Chart and pivot v1 policy per ADR 0004

## Description

Apply **ADR 0004** for **chart** and **pivot** tools: default recommendation is **tool-forced file** in v1 to preserve deterministic behavior.

## User story

As a **maintainer**, I want **explicit v1 scope** for heavy COM adapters so that **routing delivery** is not blocked by chart/pivot parity.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- If policy forces file, logs should still indicate `forced_file` or dedicated reason for transparency (**NFR-3**).

## Dependencies (narrative)

Depends on **STORY-7-1** so default write routing exists to compare against exceptions.
