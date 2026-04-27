# Manual Windows RC checklist (transport routing)

Use this list before RC sign-off for **workbook transport** (file vs COM) on **Windows with Microsoft Excel**. It extends the blueprint testing matrix in [`docs/excel-mcp-fork-com-vs-file-routing.md`](../../excel-mcp-fork-com-vs-file-routing.md) §7.

**Reference:** [COM vs file routing blueprint](../../excel-mcp-fork-com-vs-file-routing.md) — architecture, matrix, and operator notes.

---

## Automated (no Excel)

- [ ] Default CI passes on the target revision (Linux: router file path; COM skipped or mocked on non-Windows).

---

## Manual Windows (Excel required)

### From blueprint §7 (testing matrix)

- [ ] **Closed workbook:** `write` (or equivalent mutating tool) → open Excel → verify on-disk content matches expectations (file path).
- [ ] **Open workbook in Excel:** `write` with `workbook_transport=auto` (or unset / default) → cell updates in the UI **without** closing Excel.
- [ ] **Open in Excel and forced file:** pass `workbook_transport=file` → file on disk updates; Excel may show stale values until reload (documents risk; expected).
- [ ] **OneDrive-style path (if applicable):** workbook open in Excel, `write` via COM, **Save** → no surprising duplicate / conflict behavior vs parallel file-only edits (subjective; capture notes or screenshots).

### FR-9 / host state errors

- [ ] **Protected View:** workbook opens in Protected View → operation returns a **clear, fail-closed** error (no silent wrong backend).
- [ ] **Read-only:** workbook open read-only → **clear, fail-closed** error when a COM write/save is not allowed.

### Duplicate instances / path ambiguity

- [ ] **Duplicate Excel instances** (same logical path open twice): server returns **fail-closed** error; operator closes duplicate instance or consolidates to a **single** Excel instance.

### ADR 0003 — save then read

- [ ] **`save_workbook` after COM write, then read:** With COM mutations and **no** per-write save (`save_after_write` false or default), call **`save_workbook`**, then **`read_data_from_excel`** → read reflects flushed on-disk state as intended.

### ADR 0004 — chart / pivot v1 file-forced

- [ ] **`create_chart`:** With workbook open in Excel (`auto` would otherwise prefer COM for writes), tool still uses **file-forced** path; logs show a stable file-forced reason (e.g. `v1_file_forced` / documented `routing_reason` for chart).
- [ ] **`create_pivot_table`:** Same as chart — **file-forced** v1 behavior and expected **`routing_reason`** in `excel-mcp.routing` logs.

---

## Sign-off

| Date | Operator | Notes |
|------|----------|-------|
|      |          |       |
