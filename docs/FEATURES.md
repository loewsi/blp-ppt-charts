# SlideChart — features & status

Living checklist, reconciled with the original think-cell-style spec. Organized as
**General** (cross-chart — built in the shared engine, so every chart type inherits it)
and **Per chart type**.

**Status:**
- ✅ **Tested** — confirmed working in PowerPoint by Silvan
- 🔵 **Implemented** — built + unit-tested in code; NOT yet confirmed in PowerPoint (Claude is blind to the rendered chart)
- ⬜ **Open** — not built yet
- ⚠️ **Deferred / partial** — blocked by the platform, or built with a caveat to verify

**Doc process (living, not a graveyard):** this file + [TESTING.md](TESTING.md) are the two
maintained docs and are updated in the same commit as the code they describe (or at the end of a
work round). FEATURES = status of every capability; TESTING = the PowerPoint check-off list. The
old GAP-ANALYSIS doc was removed (2026-07-21) as duplicative — its architecture-gap items live in
"Robustness / architecture backlog" below. Last reconciled: **2026-07-21 pm**.

Architecture is shared: model → layout → render-plan → PowerPoint adapter. Anything under
**General** applies to all chart types once that type has a layout function.

---

## General (applies across chart types)

### Data & editing
- ✅ Excel-like grid (edit categories/series/values, keyboard nav, range select, Shift+Space, Ctrl +/−, copy/paste)
- ⚠️ Ctrl+Space (column select) is captured by a Windows OS shortcut — can't be intercepted from the pane
- ✅ Add / remove / transpose series & categories (remove targets the cursor; **insert-at-cursor: verify**)
- ✅ Live apply — typing a value commits on Enter/Tab/blur
- ✅ Insert a fresh chart; select loads it; click-away deselects
- ✅ Positive / negative values · 🔵 zero label shown in stack order when zeros aren't hidden
- ✅ Duplicate-chartId repair — same-slide copy-paste splits the copy onto a fresh id · 🔵 next-free name + orphan-legend cleanup
- ⬜ Dates as categories; custom "value = 100%" base

### Colors
- ✅ Presets + per-series swatch (Blue confirmed) · 🔵 Green / Red / Orange / Multi-color / Master accents
- 🔵 Auto-shades: pick a base color → light→dark ramp across the series
- ⬜ Per-series picker showing the master scheme; per-segment color override + highlight; series outline; pattern fill

### Segment / value labels
- ✅ Inside chip with collision offset (Silvan: "labels inside are fine, I like them")
- ✅ Show/hide, value or %, box sized to text, in front of bars, vertical-centered
- 🔵 Wide labels get a colored chip (fixes white-on-white); overlap-spread now moves the anchor label too
- 🔵 Global inside/outside placement — "Outside" moves ALL value labels out, not just small ones
- ⬜ Custom prefix/suffix text per label; datasheet text; footnotes; manual drag + leader line

### Total labels
- ✅ Totals above stacks, independent font size, sized to text
- ⬜ Totals for selected categories only; custom total text; drag a total

### Legend
- ✅ Add/remove, position top/bottom/left/right, autofit box, own movable group
- 🔵 Drawn OUTSIDE the plot, so toggling/repositioning never shifts the bars *(fixed #7 — verify)*
- ⬜ Reorder entries independently; format one entry; border

### Axis, gridlines, baseline
- ✅ Baseline; optional axis labels (auto scale, ticks); optional gridlines; axis line toggle
- ✅ Manual axis min/max
- ⚠️ Turning on axis labels still shrinks the plot (plot-anchored box model pending, #8)
- ⬜ Sync axis (same scale) across multiple charts; manual tick step; axis break; axis title

### Number formatting (shared engine)
- ✅ Decimals, prefix, suffix, group-thousands, negatives in parens, plus sign, hide-zero
- 🔵 Magnitude as a power of ten (×10⁻³/⁻⁶/⁻⁹, ×10³/⁶ — no auto letter); zero shown in stack order
- 🔵 Separator style: `1,234.56` / `1.234,56` / `1'234.56` / `1 234,56` / system
- ⬜ Currency presets; per-label-type formats

### Fonts / text
- ✅ Font family; independent segment & total sizes; middle+centered
- ⬜ Bold/italic/color per label type

### Interaction / lifecycle
- ✅ Click a chart → loads; move (no reload); resize → re-lays-out on release; tiny sizes stay visible
- ✅ Data stored in the .pptx (reopen/travel); same-slide copy-paste → independent copy
- ⬜ Undo-step quality (delete-and-recreate today)
- ⚠️ Ribbon insert + auto open/close pane (needs shared runtime; blocked on trusted-catalog sideload → revisit via central deploy / AppSource)

### Annotations (data-driven)
- ✅ Series connectors (stacked columns); difference arrows (totals or a series, signed delta + %)
- 🔵 Difference-arrow placement (diffPos); CAGR arrow — horizontal above the plot with a white % bubble
- ✅ Value / reference line (fixed value, labelled) · 🔵 configurable color
- 🔵 Arrows use triangular heads (Silvan: heads visible but seams show); guide lines are real dashed connectors
- ⬜ Error bars / ranges

---

## Per chart type

### Column / bar family
- ✅ Simple column, stacked, clustered, 100% stacked, horizontal bar
- ✅ Negative values below baseline; reverse series; 100% stacked shows the absolute total on top
- ✅ Combination — mark any series as a line (per-series toggle); secondary axis works
- 🔵 Secondary-axis manual min/max + axis line both sides
- ⬜ Category gaps / visual grouping; line area fill (needs a polygon primitive)

### Waterfall — 🔵 reworked
- 🔵 Running totals, rise/fall bars, connectors, zero baseline, signed labels
- 🔵 think-cell **"e" cell** → a computed total/subtotal column (baseline→running sum, distinct color)
- 🔵 **Multi-series steps** — each step sums its series, stacked as per-series sub-segments *(new)*
- ⬜ Connector-controlled totals (delete a connector to start a new sum) — interactive, next step
- ⬜ Bar (horizontal) orientation; per-bar color override; series legend

### Line — 🔵
- 🔵 Multi-series lines with markers + value labels *(new)*
- ⬜ Connect-gaps; area fill (needs a polygon primitive)

### Combination — ✅
- ✅ Bars + any series as a line (per-series toggle); optional secondary axis (manual min/max: verify)

### Pie / doughnut — 🔵
- 🔵 First series → slices; % inside, category outside; doughnut hole + centre total *(new)*
- ⬜ Leader lines; explode a slice; per-slice color pickers; multi-series (rings)
- ⚠️ Slices are a ~105-facet fan (Office.js has no arc geometry); geometry unit-verified — **eyeball roundness**

### Scatter / bubble — 🔵
- 🔵 Rows = X, Y, optional Size (bubble); points labelled by category; fitted axes; quadrant lines *(new)*
- ⬜ Per-point color; trend line; axis titles

### Mekko / Marimekko — 🔵
- 🔵 Column width ∝ category total; 100%-stacked segments; % labels, column totals *(new)*
- ⬜ Width-axis (%) scale labels along the top

---

## Non-chart think-cell features

### Agenda — 🔵 (v1)
- 🔵 Pane chapter list → auto-numbered "Agenda" slide; re-run updates it in place (tagged) *(new)*
- ⬜ Auto-sync from chapter markers across the deck; current-chapter highlight

### Standard slide elements — ⬜ (see "Where should slide elements live?" below)
- ⬜ Harvey balls, checkmarks/crosses, stoplights, arrows/pointers, brackets, ticks
- ⬜ Text boxes with the house style; process-flow / chevron shapes
- ⬜ These overlap with the **QuickTools** VBA ribbon add-in — decide one home (discussion below)

---

## Robustness / architecture backlog (from the former gap analysis)
- ⬜ In-place update instead of delete-and-recreate (better undo; fewer flicker/id churn)
- ⬜ Transactional render / corruption guard (roll back if a draw half-fails)
- ⬜ Persist per-object semantic tags → sub-object selection (click one segment)
- ⬜ Model chunking / CustomXmlPart if a chart's JSON outgrows the shape-tag length limit
- ⬜/🔬 Offline cold-start editing (pane loads over HTTPS; native shapes already render offline)
- 🔬 Mac runtime verification (all APIs feature-detected; not yet run on Mac)
- ✅ Local-only / no backend / no telemetry; model lives in the .pptx

## Deferred / not feasible in Office.js
- ⚠️ Shared-runtime ribbon + auto-pane (deployment-path dependent)
- ❌ Embedded Excel datasheet (grid + paste is the substitute)
- ❌ Live external Excel links (would need Graph API)
- ❌ On-canvas anchored toolbar

---

## Testing backlog (needs Silvan in PowerPoint)
Everything marked 🔵 is unit-tested (116 tests) but not visually confirmed. See [TESTING.md](TESTING.md).
Newest, never seen rendered: pie/doughnut, scatter/bubble, mekko, waterfall "e" totals, agenda,
CAGR-above-chart, diff-arrow placement, auto-shades.

## Still open
- ⬜ Plot-anchored box model so axis labels don't shrink the plot (#8 — needs resize-poll rework)
- ⬜ **Pie jumps to initial position on edit** (Silvan; likely tied to #8 box handling)
- ⬜ Sync-axis across charts; per-segment color override; per-series master-scheme color picker
- ⬜ Waterfall connector-controlled totals; line area fill (needs polygon primitive)
- ⬜ Standard slide elements (pending the "where should they live" decision)
- ❓ **Data-grid direction** (see TESTING.md bottom): Silvan wants the grid to feel more like Excel
  out of the box (formulas, hide/fold rows-cols, …). Not building piecemeal yet — reviewing options together.
  (Multi-series waterfall: ✅ built.)
