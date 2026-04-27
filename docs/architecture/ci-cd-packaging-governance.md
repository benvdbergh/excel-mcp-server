# CI/CD, packaging, and PyPI governance

**Last reviewed:** 2026-04-27

This document defines how **continuous integration**, **release packaging**, and **PyPI publishing** should operate for `excel-mcp-server`. It is aligned with the same structural ideas used in the reference **`workflows`** repository (reusable quality gates, manual release paths, least-privilege permissions, documented check names), adapted for **Python**, **Hatch**, and **PyPI** instead of npm.

## 1) Design goals

| Goal | Rationale |
|------|-----------|
| **One definition of “green”** | The same commands (tests, build checks) run on every PR and before any publish or packaging job, so CI never drifts from release gates. |
| **Separate validate, package, publish** | Packaging produces inspectable artifacts; publishing is a deliberate, auditable step with minimal extra permissions. |
| **Explicit release ref** | Manual workflows accept a **tag, branch, or SHA** so hotfixes and release audits are reproducible. |
| **Least privilege** | Default read-only repo access; **OIDC** (`id-token: write`) only on the job that talks to PyPI. |
| **Cost and hygiene** | Timeouts, concurrency where safe, short artifact retention for non-production packaging. |

## 2) Reference model (transferable from `workflows`)

The `workflows` repo separates:

1. **PR / main validation** — a thin workflow that calls a **reusable** workflow.
2. **Manual release packaging** — `workflow_dispatch` → quality gates → build tarball → **upload-artifact**.
3. **Manual trusted publish** — same gates → **registry publish** with provenance-style trust (npm: `--provenance`; PyPI: **trusted publishing** via OIDC).

**What transfers directly:** reusable workflow pattern, stable job names for branch protection, permissions map, timeouts, concurrency, retention, manual-only publish loops, runbooks.

**What differs for PyPI:** there are no npm **dist-tags** (`alpha` / `latest`) on the same version string. **Release channels** are expressed with **PEP 440** pre-release versions (e.g. `0.2.0a1`) and/or a **TestPyPI** workflow. See [release-versioning-policy.md](release-versioning-policy.md).

## 3) Target workflow map (repository layout)

These filenames are the **intended** layout once implemented under `.github/workflows/`:

| Workflow | Trigger | Role |
|----------|---------|------|
| `ci.yml` (or equivalent) | `push` / `pull_request` to protected branches | Calls reusable validation (fast feedback for contributors). |
| `reusable-validate-and-test.yml` | `workflow_call` only | Checkout (optional ref), set up Python, install project + dev deps, **pytest**, optional linters, **`hatch build`**, **`twine check dist/*`**. |
| `release-packaging.yml` | `workflow_dispatch` | Inputs: `release_ref`, optional `artifact_retention_days` → reusable gates → `hatch build` → **upload-artifact** (`dist/*.whl`, `*.tar.gz`). |
| `release-pypi-publish.yml` | `workflow_dispatch` | Inputs: `release_ref`, target (`pypi` / `testpypi` or env-based) → reusable gates → **`pypa/gh-action-pypi-publish`** with trusted publishing. |
| `publish.yml` (existing or merged) | `release: types: [published]` *optional* | Tag-aligned publish after a GitHub Release; must **reuse the same validation** job or workflow as manual publish. |

**Reuse rule:** `ci.yml`, `release-packaging.yml`, and `release-pypi-publish.yml` must all call the **same** reusable validation workflow so the command set matches what maintainers document for local runs.

## 4) npm → PyPI / Hatch command mapping

| `workflows` (npm) | `excel-mcp-server` (Python) |
|-------------------|-----------------------------|
| `npm ci` | `uv sync` / `pip install -e ".[dev]"` with lock or pinned constraints (project policy). |
| `npm test` | `pytest` (declare dev deps in `pyproject.toml`). |
| `npm run validate-workflows` / conformance | Project-specific: e.g. inventory tests, routing tests; add scripts as `tool.hatch.envs` or documented `pytest` markers. |
| `npm pack --dry-run` | `hatch build` + inspect `dist/`; optionally **`twine check dist/*`**. |
| `npm pack` + upload artifact | `hatch build` → `actions/upload-artifact` on `dist/`. |
| `npm publish --provenance` + dist-tag | **`pypa/gh-action-pypi-publish`** with **PyPI trusted publishing**; channel via **version** (pre-releases) or **TestPyPI**. |

## 5) Required checks and stable names (branch protection)

When workflows exist, configure branch protection to require checks with **stable names**. Example mapping (adjust to match exact `name:` fields in YAML):

| Policy | Example check name pattern |
|--------|----------------------------|
| Merge gate (PR) | `CI / validate-and-test` (or the reusable job’s displayed name in GitHub). |
| Release evidence | Manual workflow job names: `release-quality-gates`, `package-distributions`, `publish-to-pypi`. |

**Guidance:** treat **PR validation** as required for merge. **Manual release jobs** stay optional on day-to-day PRs but are used for release evidence and audits. If job `name:` strings change, update branch protection in the same change to avoid silent drift.

## 6) Permissions map (least privilege)

| Scope | `contents` | `id-token` | Notes |
|-------|------------|------------|--------|
| Default workflow | `read` | omit | PR CI, reusable caller. |
| Packaging job | `read` | omit | Build + upload-artifact only. |
| PyPI publish job | `read` | **`write`** | Required for OIDC token exchange with PyPI trusted publishing. |

Avoid granting `packages: write`, `actions: write`, or broad `pull-requests: write` unless a workflow truly needs them.

## 7) PyPI trusted publishing

Prerequisites (operator checklist):

1. Package **`excel-mcp-server`** on PyPI is configured for **trusted publishing** from this GitHub repo (and environment, if using GitHub Environments).
2. The workflow’s **`environment:`** name matches the PyPI trusted publisher configuration (commonly `pypi` or `release`).
3. Publish job retains **`permissions: id-token: write`**.
4. Version in **`pyproject.toml`** matches the intended release and has **not** already been uploaded.

**Troubleshooting:**

- **Authentication / forbidden:** verify trusted publisher mapping and that only the publish job requests `id-token: write`.
- **File already exists:** bump version per [release-versioning-policy.md](release-versioning-policy.md); PyPI immutability prevents overwrites.

## 8) Retention, cache, cost (defaults)

- **Artifact retention:** short default (e.g. 7 days) for packaging artifacts built from manual `release-packaging`; override via workflow input when needed.
- **Python setup:** pin a **minor** Python version in CI (e.g. `3.12`) instead of `3.x` for reproducibility.
- **Concurrency:** allow cancel-in-progress for PR CI on the same ref; use **non-canceling** concurrency for release publish if two operators might dispatch concurrently (match policy to risk).
- **Timeouts:** set `timeout-minutes` on reusable and release jobs.

## 9) Current repository state (as of last review)

Until the target workflows are added:

- **`.github/workflows/publish.yml`** publishes on **`release: published`**, builds with **Hatch**, uses **`pypa/gh-action-pypi-publish`**, and sets **`id-token: write`** (appropriate for trusted publishing).
- There is **no** PR CI workflow yet; **tests are not run** in Actions before publish; **`pytest`** should be declared under **dev / optional** dependencies in `pyproject.toml` for reproducible local and CI runs.

Treat this section as a **gap list** to close by implementing §3 and declaring dev dependencies.

## 10) Related documents

- [release-versioning-policy.md](release-versioning-policy.md) — SemVer, tags, changelog, automation options.
- [target-architecture.md](target-architecture.md) — product architecture; packaging bullet cross-links here.
- [adr/README.md](adr/README.md) — ADRs for workbook transport and COM behavior.
- Implementation plan: `docs/plan/transport-routing/Epics/Epic-8-governed-ci-cd-pypi-and-release-pipelines.md` and linked stories under `docs/plan/transport-routing/Stories/`.
