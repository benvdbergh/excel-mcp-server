# NFR-2: Routing overhead (p95)

**NFR-2** in the transport-routing PRD calls out bounded routing overhead (for example, p95 latency for resolve plus dispatch).

That **p95 is not measured continuously in CI**. Ubuntu runners in GitHub Actions are shared and noisy; the default pipeline only proves correctness (tests + packaging), not production latency SLOs.

**Local micro-benchmark (optional):** on a developer machine, time `RoutingBackend.resolve_workbook_backend` plus a no-op or mocked dispatch in a tight loop (warmup then many iterations) and inspect percentiles with a small script or `timeit`. A dedicated benchmark harness in-repo is **deferred** until product asks for regression-gated timings.
