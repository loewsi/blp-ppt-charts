# BLP Charts — handoff / context for Claude

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
- **Deploy = shared-folder catalog** (internal, no AppSource).

## Architecture (extensibility is the whole point)
`taskpane.ts` (grid UI) → `ChartModel` → `layout.ts` registry → a `layout*.ts` returns a
generic `Primitive[]` (rect/text/line) → `render.ts` draws native shapes + tags →
`persistence.ts` reads/deletes by tag.

**To add waterfall:** add `ChartType` `"waterfall"` in `chartModel.ts`, write
`layoutWaterfall.ts` (running totals → rise/fall rects + connector `line` primitives +
labels — reuses the same primitives), register it in `layout.ts`. **`render.ts` and
`persistence.ts` don't change.** Then let the task pane pick a chart type.

## Status
- Scaffold complete; `npm run build` (tsc + vite) is **green**; dev server serves 200 over HTTPS.
- **Not yet runtime-verified inside PowerPoint** — nothing here can drive PowerPoint, so
  Silvan verifies live behaviour (same constraint as QuickTools).

## Watch-outs / open risks (verify in-app)
- `getSelectedSlides()` needs a recent host build; falls back to slide 1.
- Shape **tags** are upper-cased by PowerPoint and have a value-length limit — fine for small
  models; if data grows, move the model to a CustomXmlPart.
- `addLine(ConnectorType.straight, {left,top,width,height})` and
  `addGeometricShape`/`addTextBox` signatures assume current PowerPointApi — confirm on Mac.
- Icons in `public/assets/` are solid-blue placeholders (`npm run gen-icons`) — swap for real BLP icons.
- **Location**: `C:\Users\silva\dev\blp-ppt-charts` — moved out of OneDrive (git + OneDrive risks corruption). GitHub is the backup.

## Next steps
1. Runtime-verify insert/update/load in PowerPoint (Windows first, then Mac).
2. Real hosting for `dist/` (SharePoint/Azure) + repoint manifest for cross-machine use.
3. Polish stacked column (legend, number formatting options, chart resize handling).
4. Implement **waterfall** per the recipe above.
