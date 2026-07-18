# SlideChart — architecture

This answers the review requested in the architecture brief: current architecture, storage
strategy, cross-platform capability risks, and the offline / undo / selection / reliability
positions. See `ADR-0001-core-architecture.md` for decisions and `GAP-ANALYSIS-AND-PLAN.md`
for the requirement-by-requirement gap analysis and phased plan.

## 1. Current architecture (as-built)

- **Host / stack:** Office.js PowerPoint task-pane add-in. TypeScript + Vite. **No React** —
  plain DOM in the task pane. Cross-platform by design (Windows/Mac/web).
- **Backend:** **none.** No server, no API, no database. The only network dependency is
  **static hosting** of the task-pane assets (GitHub Pages, auto-deployed by GitHub Actions).
- **Remote vs local:** there are **no remotely executed functions.** Every chart operation
  runs locally in the task pane after load — create, edit data, add/remove categories &
  series, geometry, stacked totals, percentages, number formatting, rendering, reading and
  writing chart data, reopening. (So the brief's concern that we "rely too much on remotely
  executed functions" does not apply — the only remote thing is serving the static UI.)
- **Pipeline (already matches the target):**
  `ChartModel → computeLayout() → Primitive[] (platform-neutral render plan) → render.ts
  (PowerPoint adapter, the only Office.js-dependent module) → native shapes`.
  - `src/model/chartModel.ts` — domain model, defaults, `normalizeModel` (migration).
  - `src/engine/format.ts` — number formatting (pure).
  - `src/engine/layoutBarColumn.ts` — geometry for column|bar × stacked|clustered|100% (pure).
  - `src/engine/layout.ts` — chart-type registry.
  - `src/engine/primitives.ts` — the render-plan shapes (`rect|text|line`) + `ShapeMeta`
    (`objectType`, `seriesIndex`, `categoryIndex`).
  - `src/engine/render.ts` — **the adapter**; the only file importing Office.js for drawing.
  - `src/engine/persistence.ts` — read/find/delete charts + selection + live bounds.
  - `src/office/capabilities.ts` — host capability detection.
  - `src/taskpane/taskpane.ts` — UI/state (grid, options, insert/update/load).
- **Chart data storage:** the full `ChartModel` is JSON-serialized into a **PowerPoint shape
  tag** on the chart's group (`BLPCHARTMODEL`), with `BLPCHARTID` on the group so shapes can
  be found/deleted. Everything lives **inside the .pptx** — nothing is uploaded.
- **Rendering:** native rectangles (segments), text boxes (labels/totals/categories), lines
  (baseline). Grouped into one object when `addGroup` (PowerPointApi 1.8) is available.
- **Reopening:** `getSlideCharts` reads tags → `normalizeModel` → editor. Selecting a chart on
  the slide auto-loads it (`getSelectedChartId` + `DocumentSelectionChanged`), with an
  "Edit selected chart" button as fallback.
- **IDs:** a per-chart `chartId` (`crypto.randomUUID`). Render-plan shapes now carry semantic
  `meta` (objectType + indices); this is **not yet persisted per shape** (see gaps).
- **Dependencies:** dev-only (typescript, vite, vitest, @types/office-js, office-addin-dev-certs,
  office-addin-manifest). **No runtime dependencies** — office.js loads from CDN. Bundle ≈ 17 KB.
- **Tests:** 23 unit tests (number format, layout geometry incl. 100%/clustered/zero/bar,
  model migration). Run with `npm test`.

## 2. Storage strategy

- **Chosen: shape tags** (PowerPointApi 1.3). Reasons: travel with the shape, survive copy
  between slides/presentations, per-chart granularity, cross-platform.
- **Custom XML parts:** in PowerPoint, Office.js **does not** expose `customXmlParts` (that is
  Word/Excel only) — so this is **not** an option here.
- **Document settings** (`Office.context.document.settings`): document-scoped, not per-chart,
  and doesn't travel when a chart is copied to another file — poor fit.
- **Known limit:** tag values have a size cap. Small models are fine; for large charts we must
  **chunk** the JSON across multiple tags (planned, not yet needed).
- **Preferred future design:** one canonical serialized model per chart (done) + per-shape
  semantic metadata (objectType/seriesId/categoryId) so re-render can update in place and
  selection can map to sub-objects (render plan ready; persistence pending).

## 3. Cross-platform capability risks (needs Mac verification)

| Capability | Requirement set | Risk / fallback |
|---|---|---|
| Shape tags (persist model) | PowerPointApi 1.3 | Core dependency; widely supported. |
| `addGroup` (one movable object) | PowerPointApi 1.8 | Newer; **feature-detected**, falls back to per-shape tags. |
| `getSelectedShapes` (edit selected) | PowerPointApi 1.5 | Fallback: pick from the on-slide list. |
| `DocumentSelectionChanged` | host-dependent | Fallback: "Edit selected chart" button. |
| Tag value length | — | Chunk large models. |
| Copy to another presentation | — | Tags should travel, but **duplicate chartId** must be detected/repaired (unbuilt). |

All of the above are **verified to compile against the official typings** but **not yet
runtime-verified on Mac**. `capabilities.ts` centralizes detection; the UI warns when grouping
is unavailable.

## 4. Offline behaviour (honest position)

- **Presentation visibility (mandatory): met.** Charts are native shapes — visible, printable,
  exportable, copyable **without internet and without the add-in**. Data is in the .pptx.
- **Add-in editing after an offline cold start: not guaranteed.** The task pane is loaded from
  HTTPS; if PowerPoint starts offline, the pane may fail to load. We do **not** claim browser
  caching solves this. What we do promise: *once loaded, editing does not depend on the
  network* (all operations are local). True cold-start offline is a **future research item**.

## 5. Undo / reliability (current limitations)

- **Undo:** `Update` currently **deletes the group and redraws**, which is several Office
  operations and likely several undo steps — and delete-recreate is unfriendly to undo and
  selection. Target: **update shapes in place** keyed by semantic id (render plan supports it;
  adapter change pending, needs PowerPoint testing).
- **Reliability:** no transactional guard yet — a mid-render failure could leave a chart
  partially removed. Target sequence: validate → build plan → persist previous → render →
  verify → persist model → clean up → roll back on failure.

## 6. Selection recognition

`getSelectedChartId` maps the selected shape/group back to its `chartId` via tags. Sub-object
recognition (segment vs label vs legend) becomes possible once per-shape `objectType` tags are
persisted (render plan already carries them).
