---
kind: story
id: STORY-8-3
title: Manual release packaging workflow and artifacts
status: done
parent: EPIC-8
depends_on:
  - STORY-8-2
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
  - "`workflow_dispatch` workflow exists with input `release_ref` (and optional `artifact_retention_days` defaulting to short alpha-style retention)."
  - "Job order: call reusable validate-and-test with `checkout_ref` → `hatch build` → upload `dist/` artifacts (`*.whl`, `*.tar.gz`)."
  - "`if-no-files-found: error` on artifact upload; concurrency group documented."
  - "Governance doc workflow map (§3) updated if filenames differ from the documented placeholders."
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-8-3: Manual release packaging workflow and artifacts

## As delivered

- **`.github/workflows/release-packaging.yml`** — **`workflow_dispatch`** with **`release_ref`**, **`artifact_retention_days`** (default **7**), **`validate-and-test`** (reusable) → **`package-distributions`** (**hatch build** → **`upload-artifact`**, `if-no-files-found: error`, **`dist/*.whl`**, **`dist/*.tar.gz`**), non-cancelling **concurrency** on **`release_ref`**.

## Description

Implement **manual release packaging** so operators can produce **inspectable distributions** from an explicit ref before publishing, matching [ci-cd-packaging-governance.md](../../../architecture/ci-cd-packaging-governance.md) §3 and §8.

## User story

As a **release operator**, I want **to build and download wheels/sdists from a chosen ref** so that **I can verify artifacts before PyPI upload**.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Naming suggestion: `.github/workflows/release-packaging.yml` (adjust governance doc in same PR if renamed).

## Dependencies (narrative)

Requires **STORY-8-2** so packaging reuses the same reusable gates.
