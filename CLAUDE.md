# SlideChart — handoff / context for Claude

**What this is:** a cross-platform (Windows + Mac + web) PowerPoint add-in that inserts
**live-linked charts** — think-cell-style. Built on **Office.js + TypeScript + Vite**.
Separate project from **QuickTools** (the VBA `.ppam` add-in in `../01 PPT`), because VBA
UserForms/ribbons don't exist on Mac.

**Owner:** Silvan Loew (BLP Digital). Read `README.md` first for run/sideload steps.

## Decisions locked in this session (2026-07-16)
- **Cross-platform ⇒ Office.js**, not VBA. XML manifest (JSON manifest isn't cross-platform yet).
- **v1 = stacked column.** **Waterfall is the eventual goal** — architecture is built for it.
- **Behaviour = live-linked:** model is JSON-stamped on a shape tag (`BLPCHARTMODEL`),
  saved in the `.pptx`. Update = read model → delete tagged shapes → redraw.
- **Deploy = GitHub Pages** (`https://loewsi.github.io/blp-ppt-charts/`, auto-deploys on push to
  `main` via `.github/workflows/deploy.yml`). Sideload from the shared-folder catalog for testing.
  Shared-runtime ribbon/auto-pane is reverted for now (trusted-catalog sideload didn't list it) —
  code paths are guarded no-ops; revisit with central deploy / AppSource.

## Architecture (extensibility is the whole point)
`taskpane.ts` (grid UI) → `ChartModel` → `layout.ts` registry → a `layout*.ts` returns a
generic `Primitive[]` (rect/text/line) → `render.ts` draws native shapes + tags →
`persistence.ts` reads/deletes by tag.

**To add waterfall:** add `ChartType` `"waterfall"` in `chartModel.ts`, write
`layoutWaterfall.ts` (running totals → rise/fall rects + connector `line` primitives +
labels — reuses the same primitives), register it in `layout.ts`. **`render.ts` and
`persistence.ts` don't change.** Then let the task pane pick a chart type.

## Status (updated 2026-07-21 pm)
- `npm run build` green; **116 unit tests** pass (`npx vitest run`). Live deploy auto-serves latest.
- **Chart types (all built):** column/bar, **line**, **combination** (per-series line + secondary
  axis), **pie/doughnut**, **scatter/bubble**, **mekko**, **waterfall** (with "e" total columns).
- **Annotations:** difference arrows, CAGR arrows (horizontal above chart), reference line, connectors.
  Arrowheads are custom triangles (Office.js PowerPoint LineFormat has no arrowhead API); dashed
  guides use real connectors (LineFormat.dashStyle).
- **Also:** duplicate-id repair, orphan-legend cleanup, agenda slide generator, auto-shade colors,
  relevant-only options panel, legend drawn outside the plot.
- **New primitives:** arrow, triangle, ellipse (render.ts `makeShapes` returns Shape[]).
- **Still open:** plot-anchored box model so axis labels don't shrink the plot (#8; needs the resize
  poll to subtract decoration insets — do carefully); sync-axis across charts; per-segment color;
  waterfall connector-controlled totals; line area fill (needs a polygon primitive). Each feature is
  its own commit (revertable). See `docs/FEATURES.md` + `docs/TESTING.md`.
- **`docs/FEATURES.md`** = the living status list (✅ tested / 🔵 built-not-verified / ⬜ open).
  **`docs/TESTING.md`** = the PowerPoint checklist for Silvan to flip 🔵 → ✅.
- **Blind-rendering constraint:** Claude can't see rendered charts; everything 🔵 is unit-tested
  only. Verify in PowerPoint before trusting visuals. (An earlier over-claim — marking combination
  line + waterfall subtotals as built when they weren't — is why FEATURES.md is now kept honest.)

## Watch-outs / open risks (verify in-app)
- `getSelectedSlides()` needs a recent host build; falls back to slide 1.
- Shape **tags** are upper-cased by PowerPoint and have a value-length limit — fine for small
  models; if data grows, move the model to a CustomXmlPart.
- `addLine(ConnectorType.straight, {left,top,width,height})` and
  `addGeometricShape`/`addTextBox` signatures assume current PowerPointApi — confirm on Mac.
- Icons in `public/assets/` are solid-blue placeholders (`npm run gen-icons`) — swap for real BLP icons.
- **Location**: `C:\Users\silva\dev\blp-ppt-charts` — moved out of OneDrive (git + OneDrive risks corruption). GitHub is the backup.

## Next steps
1. **Silvan: walk `docs/TESTING.md` in PowerPoint** and report — that unblocks flipping the 🔵
   backlog to ✅ and tells us what actually looks wrong.
2. **Difference & CAGR arrows** (think-cell signature, next in FEATURES.md order). Needs two
   things first: a real arrowhead in `render.ts` (today lines become plain rotated rects — no
   head; add `LineFormat` arrowhead styling on a genuine line/connector shape) and a decision on
   which points to compare (sensible default: first ↔ last category total, toggleable).
3. **Combination line series** — add `Series.kind: "bar" | "line"`; line series excluded from
   stacking/totals, folded into the axis range, drawn as a polyline+markers. Column orientation
   first. Diagonal-line rendering (rotated rect) already supports the segments.
4. Per-segment color override / single-bar highlight; waterfall subtotal columns; error bars.
5. Cross-machine + Mac verification; real BLP icons (still placeholders in `public/assets/`).
