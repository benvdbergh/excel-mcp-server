# Implementation roadmap — workbook transport routing

This roadmap decomposes `docs/specs/PRD-excel-mcp-transport-routing.md` into **epics and stories** under `docs/plan/transport-routing/`, aligned with `docs/architecture/target-architecture.md` and ADRs `docs/architecture/adr/`.

**Status (2026-04-27):** Epics **1–8** are implemented. **Epic 8** (governed CI/CD, manual packaging/publish, PyPI **excel-com-mcp** via OIDC) is **delivered**; see [Epic-8](Epics/Epic-8-governed-ci-cd-pypi-and-release-pipelines.md). Transport epics **1–7**: COM write parity, FR-9 COM guards, ADR 0004 tests/logging, `save_workbook` MCP tool, README/CI/manual checklist, NFR-2 note.

## Phasing (execution order)

| Phase | Epic | Summary |
|-------|------|---------|
| 1 | [Epic-1](Epics/Epic-1-tool-inventory-and-shared-workbook-operation-contract.md) | Tool inventory and shared workbook operation contract |
| 2 | [Epic-2](Epics/Epic-2-path-normalization-and-unified-allowlist.md) | Path normalization and unified allowlist *(delivered)* |
| 3 | [Epic-3](Epics/Epic-3-fileworkbookservice-facade-and-handler-consolidation.md) | `FileWorkbookService` façade and handler consolidation *(delivered)* |
| 4 | [Epic-4](Epics/Epic-4-routingbackend-and-open-workbook-detection-file-backed-execution.md) | `RoutingBackend`, injectable open-workbook detection, structured logs *(delivered)* |
| 5 | [Epic-5](Epics/Epic-5-operator-controls-and-mcp-tool-wiring.md) | Operator controls: env vars, tool params, handler wiring *(delivered)* |
| 6 | [Epic-6](Epics/Epic-6-com-packaging-executor-and-comworkbookservice-skeleton.md) | COM packaging, single-thread executor, `ComWorkbookService` skeleton *(delivered)* |
| 7 | [Epic-7](Epics/Epic-7-com-write-parity-edge-policies-save-workbook-and-release-hardening.md) | COM write parity, edge policies, `save_workbook`, docs, CI, manual checklist *(delivered)* |
| 8 | [Epic-8](Epics/Epic-8-governed-ci-cd-pypi-and-release-pipelines.md) | Governed CI/CD, reusable gates, manual packaging/publish, PyPI (**`excel-com-mcp`**) *(delivered)* |

## Architecture traceability

| Theme | Architecture source |
|-------|---------------------|
| Layering (`RoutingBackend`, services, path policy) | `docs/architecture/target-architecture.md` |
| Workbook vs MCP wire naming | `docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md` |
| COM stack choice | `docs/architecture/adr/0002-com-automation-stack.md` |
| Read path + `save_workbook` | `docs/architecture/adr/0003-read-path-com-parity.md` |
| Chart/pivot v1 scope | `docs/architecture/adr/0004-chart-pivot-com-parity-scope.md` |
| Strict mode and fallback | `docs/architecture/adr/0005-com-strict-and-fallback-controls.md` |
| Baseline coupling | `docs/architecture/pre-fork-architecture.md` |
| CI/CD, PyPI, release gates | `docs/architecture/ci-cd-packaging-governance.md` |
| Versioning and changelog | `docs/architecture/release-versioning-policy.md` |

## Validate planning artifacts

Using the **project-planning** skill’s `LintPlan.ts` (requires [Bun](https://bun.sh/)):

```bash
bun run LintPlan.ts --root <repo-root>
```

Run `LintPlan.ts` from the skill’s `scripts/` directory, passing this repository as `--root`.
