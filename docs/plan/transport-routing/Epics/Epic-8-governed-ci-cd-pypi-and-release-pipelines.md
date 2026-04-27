---
kind: epic
id: EPIC-8
title: Governed CI/CD, PyPI packaging, and release pipelines
status: draft
depends_on: []
traces_to:
  - path: docs/architecture/ci-cd-packaging-governance.md
  - path: docs/architecture/release-versioning-policy.md
  - path: docs/architecture/target-architecture.md
slice: vertical
acceptance_criteria:
  - PR CI runs the same quality gates as documented for local development (pytest, hatch build, twine check or equivalent).
  - Manual release packaging produces versioned sdist/wheel artifacts with configurable retention.
  - Manual PyPI publish path exists with least-privilege permissions and optional alignment to GitHub Release publish via shared reusable gates.
  - Operator can follow governance docs to configure branch protection check names and PyPI trusted publishing.
  - Documentation allows configuring the MCP from a PyPI-based no-install pattern (e.g. `uvx` + distribution name) in agentic clients; README and `manifest.json` stay aligned with `[project].name` (see Story-8-4).
created: "2026-04-27"
updated: "2026-04-27"
---

# Epic-8: Governed CI/CD, PyPI packaging, and release pipelines

## Description

Implement the **target workflow layout** and **dependency hygiene** described in `docs/architecture/ci-cd-packaging-governance.md` and `docs/architecture/release-versioning-policy.md`: reusable validation, PR CI, optional manual packaging and publish workflows, and hardened **`publish.yml`** so release automation matches **documented** SemVer and PyPI practices.

This epic is **orthogonal** to workbook transport delivery (Epics 1–7) and may proceed in parallel; **Story-7-5** depends on **Story-8-2** so “default CI passes” in the PRD release gate reuses the governed pipeline.

## Objectives

- One **reusable** workflow defines validate-and-test for PRs, packaging, and publish.
- **Dev dependencies** are declared in `pyproject.toml` for reproducible CI and contributor onboarding.
- **Manual** `workflow_dispatch` flows support explicit `release_ref` (tag/branch/SHA) for auditability.
- **PyPI trusted publishing** remains the default posture; publish jobs are the only jobs requiring `id-token: write`.

## User stories (links)

- [Story-8-1](../Stories/Story-8-1-declare-dev-dependencies-and-local-validate-commands.md)
- [Story-8-2](../Stories/Story-8-2-reusable-validate-workflow-and-pr-ci.md)
- [Story-8-3](../Stories/Story-8-3-manual-release-packaging-workflow-and-artifacts.md)
- [Story-8-4](../Stories/Story-8-4-manual-pypi-publish-and-harden-release-published-path.md)

## Dependencies (narrative)

None on prior epics; **Story-7-5** consumes **Story-8-2** for CI acceptance.

## Related sources

- `docs/architecture/README.md` — architecture index
- Reference patterns: external **`workflows`** repo (reusable validate, manual packaging, manual npm publish mapped to PyPI in governance doc)
