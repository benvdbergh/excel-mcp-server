---
kind: story
id: STORY-8-2
title: Reusable validate workflow and PR CI
status: done
parent: EPIC-8
depends_on:
  - STORY-8-1
traces_to:
  - path: docs/architecture/ci-cd-packaging-governance.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - "`.github/workflows/reusable-validate-and-test.yml` exists with `workflow_call`, optional `checkout_ref`, pinned Python minor, install dev deps, pytest, `hatch build`, and `twine check dist/*` (or documented equivalent)."
  - "`.github/workflows/ci.yml` (or named consistently) runs on push/PR to protected branches and calls the reusable workflow with stable job names suitable for branch protection."
  - "Workflow uses least-privilege default permissions (`contents: read`); timeouts set per governance doc."
  - "No COM/Excel required on ubuntu runner for default job."
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-8-2: Reusable validate workflow and PR CI

## As delivered

- **`.github/workflows/reusable-validate-and-test.yml`** — `workflow_call`, **`checkout_ref`**, Python **3.12**, **`pip install -e ".[dev]"`**, **pytest**, **hatch build**, **`twine check dist/*`**, Actions pinned to commit SHAs.
- **`.github/workflows/ci.yml`** — **CI** workflow on **`push` / `pull_request`** to **`main` / `master`**, job name **`validate-and-test`** (for branch protection: typically **`CI / validate-and-test`** in GitHub UI).
- **`permissions: contents: read`**, job **`timeout-minutes`**, PR **concurrency** with cancel-in-progress per governance.

## Description

Add the **reusable validation** unit and a thin **PR CI** entry workflow so every merge request runs the same gates as release paths, per [ci-cd-packaging-governance.md](../../../architecture/ci-cd-packaging-governance.md) §3.

## User story

As a **maintainer**, I want **PR checks that mirror release quality gates** so that **regressions are caught before tag or publish**.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Align displayed job `name:` fields with governance **§5 stable names** once branch protection is updated (document exact strings in PR or governance doc follow-up).
- **Story-7-5** lists `depends_on: STORY-8-2` for PRD-aligned “default CI passes.”

## Dependencies (narrative)

Requires **STORY-8-1** so install steps have a single source of truth in `pyproject.toml`.
