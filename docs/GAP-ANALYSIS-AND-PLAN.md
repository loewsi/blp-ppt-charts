# SlideChart — gap analysis & migration plan

Classification vs the architecture brief. Legend: ✅ aligned · 🟡 partial · ❌ not aligned ·
🔬 needs an Office.js proof-of-concept (can't be settled without PowerPoint testing).

## Gap analysis (by requirement area)

| # | Requirement | Status | Notes |
|---|---|---|---|
| 1 | Cross-platform Office.js, stacked-column first | ✅ | Column/bar × stacked/clustered/100% engine; Mac not runtime-verified 🔬 |
| 2 | Free/anonymous, no backend for chart use | ✅ | No backend exists; only static hosting |
| 3 | Local-first (all ops local) | ✅ | Zero remote functions; all ops in TS |
| 4.1 | Charts visible offline as native shapes | ✅ | Native rects/text/lines in the .pptx |
| 4.2 | Editing after offline cold start | ❌/🔬 | Pane loads from HTTPS; documented as future research |
| 5 | No presentation data in the cloud | ✅ | Model stored in shape tags inside the .pptx |
| 6 | Model → render plan → adapter separation | ✅ | Engine is Office-free; `render.ts` is the sole adapter |
| 7 | Modular repo | 🟡 | `src/*` modules separate concerns; no `packages/*` (deliberate, see ADR) |
| 8 | Model stored in presentation + schema versioning | 🟡 | Shape tags ✅, `schemaVersion` ✅; per-shape semantic tags pending; chunking pending |
| 9 | Stable IDs, copy/duplicate handling | ❌ | `chartId` only; per-object IDs in render plan but not persisted; duplicate-ID repair unbuilt 🔬 |
| 10 | Native shapes; update-in-place re-render | 🟡 | Native shapes ✅; currently delete-and-recreate, not in-place 🔬 |
| 11 | Local spreadsheet-style grid | 🟡 | Grid with add/remove/reorder/transpose/paste ✅; no keyboard-nav/inline-preview; custom grid, no library |
| 12 | Stacked-column feature scope | 🟡 | Colors/segment labels/totals/legend?/axis? — colors+labels+totals ✅; legend/axis/gridlines pending |
| 13 | Manual label offsets (model support) | ❌ | Not in model yet; planned as `LabelPlacement` |
| 14 | Future chart features possible | ✅ | Registry + render plan designed for it |
| 15 | CAGR annotation model (semantic anchors) | ❌ | Annotation model not yet added |
| 16 | UI decoupled from model/renderer | ✅ | Taskpane depends on engine, not vice versa |
| 17 | Release notifications | ❌ | Not built (optional, later) |
| 18 | Keep VBA tool separate | ✅ | Separate product |
| 19 | One productivity command PoC | ❌ | Not built (later) |
| 20 | Capability layer | 🟡 | `capabilities.ts` added; Mac behaviours unverified 🔬 |
| 21 | Undo strategy | ❌/🔬 | Delete-recreate today; in-place update needed |
| 22 | Selection recognition | 🟡 | Chart-level ✅; sub-object pending (needs persisted objectType) |
| 23 | Corruption prevention / transactional render | ❌ | No guard yet |
| 24 | Tests for pure logic | 🟡 | 23 unit tests (format/layout/migration); no render-plan snapshots/adapter mocks yet |
| 25 | Telemetry/privacy | ✅ | None collected |

## Migration recommendation

- **Keep:** the whole `src/engine/*` (model, format, layout, primitives-as-render-plan,
  layout registry), `persistence.ts`, `render.ts` (adapter), `capabilities.ts`, taskpane UI,
  the GitHub Pages pipeline. These already match the target shape.
- **Evolve (behind existing interfaces):**
  - `primitives.ts` → treat as the formal **render plan**; already carries `ShapeMeta`.
  - `render.ts` → persist per-shape `objectType`/indices and move to **update-in-place** diff.
  - `chartModel.ts` → add `LabelPlacement` and an `annotations` array (CAGR/difference/value
    lines) using **semantic category/series anchors**, not slide coordinates.
- **Add:** duplicate-`chartId` detection/repair on load; model chunking for large charts;
  render-plan snapshot tests; a transactional render guard.
- **Do NOT do now:** monorepo split, accounts/billing/orgs, external Excel links, AI, the VBA
  feature set, a broad shortcut system. (Explicitly out of scope.)

## Phased plan (acceptance criteria)

- **Phase 0 — Architecture extraction.** ✅ *Done.* Model/layout/render-plan/adapter separated;
  `normalizeModel` + `schemaVersion`; number-format engine; capabilities module; 23 unit tests.
- **Phase 1 — Persistence PoC (🔬 needs PowerPoint, Win+Mac).** Insert a stacked column; store
  model in the .pptx; close/reopen PowerPoint and reopen the chart; copy to another slide and
  another presentation with a **duplicate-ID repair** step. *Accept:* chart re-editable after
  reopen; copies get fresh `chartId`s; verified on Windows and Mac.
- **Phase 2 — Local data editor.** ✅ mostly (grid, add/remove/reorder/transpose, paste). *Add:*
  keyboard navigation, explicit Apply/Cancel draft state, inline validation messages.
- **Phase 3 — Core formatting.** Series colors ✅, segment colors 🟡 (per-series done, per-segment
  pending), segment labels ✅, total labels ✅, **legend** (pending), **axis + gridlines**
  (pending). *Accept:* legend on/off + position; optional value axis + gridlines; per-segment
  color override stored in the model.
- **Phase 4 — Interaction quality (🔬).** Update-in-place re-render; manual label offsets;
  sub-object selection; undo behaviour documented/tested; duplicate-ID repair; error recovery.
- **Phase 5 — Offline & release.** Verify editing after network loss; document restart-offline
  behaviour on Win+Mac; add non-blocking release-notification check.
- **Phase 6 — Variants.** Simple column, clustered, 100%, horizontal bar (engine already covers
  these), negative-value handling.

## Files expected to change vs keep

- **Change/extend:** `src/engine/render.ts` (per-shape tags, in-place update), `src/engine/
  persistence.ts` (duplicate-ID repair, chunking), `src/model/chartModel.ts` (LabelPlacement,
  annotations), `src/taskpane/*` (legend/axis controls, Apply/Cancel).
- **Add:** `src/engine/annotations/*`, `src/engine/render-plan snapshot tests`, `services/
  release-check` (later).
- **Keep as-is:** `format.ts`, `layout.ts`, `layoutBarColumn.ts` (extend, don't rewrite),
  `primitives.ts`, `capabilities.ts`, build/deploy pipeline, manifest.

## Open questions (can't be answered from the repo)

1. Do shape tags reliably persist and travel on **Mac** and when copying across presentations?
2. Actual tag **value-length limit** per host (drives when chunking is needed)?
3. Does `DocumentSelectionChanged` fire for shape selection on Windows and Mac?
4. How many Office ops does one `Update` cost in **undo** on each host?
5. Minimum Office versions to support (affects requirement-set assumptions)?

## Smallest PoC to validate the architecture

**Persistence + identity round-trip (Phase 1):** insert one stacked column → close/reopen
PowerPoint → reopen chart from its stored model → copy to another slide and another
presentation → confirm the copy is re-editable and gets a fresh `chartId`. This single flow
exercises native-shape rendering, in-file storage, reopen, and copy/ID handling — the
foundations everything else depends on. It requires hands-on PowerPoint testing on Windows and
Mac (cannot be validated headlessly).
