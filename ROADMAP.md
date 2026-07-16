# SlideChart roadmap — prioritized (bar/column family)

Prioritization of the full think-cell-style feature list. Scope for now: the **bar/column
family**, but everything is built to be reused by future chart types (waterfall, Mekko…).

## The governing constraint

We render charts as **native PowerPoint shapes** (no chart API exists). Consequences:
- ✅ **Visual output** (types, colors, labels, axes, legends, annotations) is broadly achievable.
- ⚠️ Some interactions are **partial** (native drag works, but the add-in re-syncing to it is fragile).
- ❌ A few think-cell things are **not possible** in Office.js — see the bottom.

Feasibility legend: ✅ doable · ⚠️ partial/hard · ❌ not possible (substitute noted).

## Reusable foundations (build once, every chart type uses them)

- **Number-format engine** — decimals, %, k/M/bn, prefix/suffix, +/− signs, parentheses,
  hide-zeros, dash-for-zero. Used by segment labels, totals, axis, arrows, value lines (§17).
- **Style/theme model** — palettes (BLP + custom), fills, outlines, fonts — one schema.
- **Layout → primitives → render** (already built) — each chart type is one `layout*.ts`.
- **Annotation layer** — value lines, connectors, arrows, CAGR operate on computed geometry,
  so they work across column/bar/waterfall unchanged.
- **Label engine** — placement + formatting shared by all label kinds.

---

## Phase A — Data editing that feels great  *(your priority #1)*

The "datasheet" + re-open experience. Maps to §2, §18 (data parts).
- ✅ Paste a range from Excel/clipboard (TSV) straight into the grid.
- ✅ Add / insert / delete / reorder categories & series; rename labels; transpose;
  toggle rows-as-series; negative / zero / percentage values.
- ✅ Re-open on click: selecting a chart loads its data into the editor (already working) —
  harden this so it always reflects the selected chart.
- ❌ PowerPoint's *embedded Excel* datasheet — not reachable via Office.js. **Substitute:** our
  task-pane grid + paste-from-Excel is the datasheet.
- ⚠️/❌ *External live Excel links* (§2) — would need Graph API + stored file handles; heavy and
  fragile cross-platform. **Deferred** (revisit only if truly needed).

## Phase B — Visual formatting: colors, fonts, numbers  *(your priority #2)*

Maps to §4, §17, and label/total formatting in §5/§6/§7.
- ✅ Chart color schemes (BLP + custom); reverse mapping; theme colors.
- ✅ Per-series fill + outline (color/width/style); per-segment color override & highlight.
- ⚠️ Pattern fills (limited support in Office.js).
- ✅ The number-format engine (above) surfaced in the UI.
- ✅ Label/total fonts, size, color, bold/italic, alignment; **totals formatted independently
  from segment labels** (§6 — an explicit benchmark row).

## Phase C — Structure & the bar/column family (engine payoff)

Maps to §1 (insertion/config), §3, §16.
- ✅ Switch type **preserving data**: simple column, stacked, 100% stacked, clustered,
  horizontal bar (rotate, labels stay horizontal), combination (series → line).
- ✅ Absolute ↔ percentage; reverse category/series order; flip value axis.
- ✅ Spacing: category gap, cluster gap, column width; explicit category gaps for grouping.
- ✅ Positive above / negative below baseline; separate stacking.
- ✅ Move / resize whole chart (already grouped); resize re-lays-out.
- ⚠️/❌ Separate plot-area vs chart-area resizing, edge locking, auto-grow, align-to-object
  (§1 geometry) — hard; approximate later, not core.

## Phase D — Chart furniture: legend, axis, gridlines, labels

Maps to §5, §7, §8, §9, §10.
- ✅ Legend (add/remove, position around/inside, format).
- ✅ Value axis (auto/manual min/max/tick, %/absolute, tick labels, title). ⚠️ axis breaks.
- ✅ Gridlines + baseline (color/width/style).
- ✅ Segment / series / category labels (value / % / datasheet text, show-hide small & zero).
- ✅ Automatic label placement (incl. simple overlap avoidance). ⚠️ **manual drag** of a single
  label that survives re-layout (native drag possible; re-sync is the hard part).

## Phase E — Data-driven annotations (think-cell signature)

Maps to §11, §12, §13, §14, §15.
- ✅ Series connectors (also essential for waterfall).
- ✅ Value / reference lines (data- or manually-anchored).
- ✅ Difference arrows (level & total, absolute & %). ⚠️ drag endpoints.
- ✅ CAGR arrows (auto-calc from ends & interval; chart-area expands to fit).
- ✅ Error bars / ranges (incl. football-field via rotated chart).
- All recompute on data change. ⚠️ dragging endpoints/labels is the constrained part.

## Phase F — Interaction & feel  *(your priority #3)*

Maps to §18. **Key decision — the UI surface:**
- **Task pane (side)** — robust, always available, resizable. ✅ Our current approach.
- **Office Dialog** — a **floating window over PowerPoint**, closest to think-cell's floating
  panel *feel*. ✅ possible, but it is **not anchored to the chart** and behaves like a separate
  window. Good for a focused "edit this chart" popup.
- **A toolbar anchored on top of the chart on the canvas** — ❌ **not possible** in Office.js.
- ✅ Contextual controls (select a series/segment → its controls appear in the surface).
- ✅ Overall polish, keyboard, responsiveness — "look and feel good."

## After the bar family
Reuse the engine for **waterfall** (your original end goal), then Mekko/others.

---

## Not possible in Office.js (be clear with stakeholders)
- ❌ Embedded Excel OLE datasheet → **substitute:** task-pane grid + paste-from-Excel.
- ❌/⚠️ Live external Excel links → deferred (Graph API, heavy).
- ❌ Toolbar anchored on the chart canvas → task pane or floating dialog instead.
- ⚠️ Add-in-tracked dragging of individual labels / arrow endpoints (native drag works;
  keeping the model in sync across re-layout is the hard, partly-solvable part).

## Recommended build order
1. **Phase A** (data editing + re-open) — your #1, unblocks everything.
2. **Number-format engine** (start of Phase B) — reused everywhere, so build early.
3. **Phase B** (colors/visuals) — your #2.
4. Decide the **interaction surface** (Phase F decision) early, since it frames the UI.
5. Then C → D → E, then waterfall.
