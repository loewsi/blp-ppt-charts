# SlideChart — features & status

Living checklist, reconciled with the original think-cell-style spec. Organized as
**General** (cross-chart — built in the shared engine, so every chart type inherits it)
and **Per chart type**.

**Status:**
- ✅ **Tested** — confirmed working in PowerPoint by Silvan
- 🔵 **Implemented** — built + unit-tested in code; NOT yet confirmed in PowerPoint (Claude is blind to the rendered chart)
- ⬜ **Open** — not built yet
- ⚠️ **Deferred** — blocked by the Office.js platform / deployment path

Architecture is shared: model → layout → render-plan → PowerPoint adapter. Anything under
**General** applies to all chart types once that type has a layout function.

---

## General (applies across chart types)

### Data & editing
- ✅ Excel-like grid (edit categories/series/values, keyboard nav, range select, Ctrl+Space, Ctrl +/−, copy/paste)
- 🔵 Add / remove / transpose series & categories — **remove now targets the cursor's row/col**
- 🔵 Live apply — **typing a value now commits on Enter/Tab/blur** (was: needed a transpose)
- 🔵 Insert a fresh chart; select loads it; **click-away deselects**
- 🔵 Positive / zero / negative values (zero label shown at baseline when zeros aren't hidden)
- ⬜ **Duplicate-chartId repair** — copy-paste on the same slide clones the id (works across slides). **HIGH**
- ⬜ Dates as categories; custom "value = 100%" base

### Colors
- 🔵 Master theme accents (live) · Blue · Grayscale presets (more blue shades); per-series swatch
- ⬜ Color picker + auto-generated shades; per-series picker showing the master scheme
- ⬜ Per-segment color override + highlight; series outline; pattern fill

### Segment / value labels
- ✅ Inside chip with collision offset (Silvan: "labels inside are fine, I like them")
- 🔵 Show/hide, value or %, box sized to text, in front of bars, vertical-centered
- 🔵 Wide labels get a colored chip (fixes white-on-white); overlap-spread now moves the anchor label too
- ⬜ Global inside/outside placement for ALL labels (today "Move outside" only affects small segments)
- ⬜ Custom prefix/suffix text per label; datasheet text; footnotes; manual drag + leader line

### Total labels
- 🔵 Totals above stacks, independent font size, sized to text
- ⬜ Totals for selected categories only; custom total text; drag a total

### Legend
- 🔵 Add/remove, position top/bottom/left/right, autofit box, own movable group
- 🔵 Picking a position snaps the legend there; a manual drag stays until you pick a new position *(fixed)*
- ⬜ Reorder entries independently; format one entry; border

### Axis, gridlines, baseline
- 🔵 Baseline; optional axis labels (auto scale, ticks); optional gridlines; **axis line toggle** *(new)*
- 🔵 Manual axis min/max
- ⬜ Sync axis (same scale) across multiple charts
- ⬜ Manual tick step; secondary axis; axis break; axis title

### Number formatting (shared engine)
- 🔵 Decimals, magnitude ÷1e3/÷1e6 (no auto letter — add your own suffix), prefix, suffix
- 🔵 Hide-zero (toggle respected), percent, plus sign, negatives in parentheses
- 🔵 Separator style: `1,234.56` / `1.234,56` / `1'234.56` / `1 234,56` / system  *(new)*
- ⬜ Currency presets; per-label-type formats

### Fonts / text
- 🔵 Font family; independent segment & total sizes; middle+centered
- ⬜ Bold/italic/color per label type

### Interaction / lifecycle
- 🔵 Click a chart → loads into pane; move (no reload); resize → re-lays-out on release; tiny sizes stay visible; legend stays put
- 🔵 Data stored in the .pptx (reopen/travel)
- ⬜ Copy-to-slide/deck duplicate-chartId repair; undo-step quality
- ⚠️ Ribbon insert + auto open/close pane (needs shared runtime; blocked on trusted-catalog sideload → revisit via central deploy / AppSource)

### Annotations (data-driven)
- 🔵 Series connectors (stacked columns)  *(new)*
- ⬜ Difference arrows (level & total, absolute & %)
- ⬜ CAGR arrows (auto growth over a period)
- 🔵 Value / reference line (target line at a fixed value, labelled, **color configurable**)
- ⬜ Error bars / ranges

---

## Per chart type

### Column / bar family
- 🔵 Simple column, stacked, clustered, 100% stacked, horizontal bar
- 🔵 Negative values below baseline
- 🔵 Reverse series (stacking/cluster order), separate from reverse categories *(new)*
- 🔵 100% stacked shows the absolute total on top *(fixed)*
- ⬜ Combination — one series as a line
- ⬜ Line markers, secondary axis; category gaps / visual grouping

### Waterfall — ⚠️ needs rework (Silvan: "logic does not work yet")
- ⬜ Use the same categories/series grid + naming as the column family
- ⬜ think-cell "e" cell = a calculated subtotal/total column
- ⬜ Connector-controlled totals (choose where the total lands) — shared mechanic with CAGR from→to
- ⬜ Bar (horizontal) orientation; per-bar color override

### Line — ⬜
- ⬜ Single/multi-series line; markers; value labels; connect-gaps; area fill option

### Combination — ⬜
- ⬜ Bars + one or more series drawn as a line (shares the column engine's axis)
- ⬜ Secondary axis for the line series

### Pie / doughnut — ⬜
- ⬜ Single series → slices; % or value labels; doughnut hole size; leader lines; explode a slice

### Scatter / bubble — ⬜
- ⬜ XY scatter (x,y per point); bubble adds a size dimension; point labels; quadrant lines

### Mekko / Marimekko — ⬜
- ⬜ Variable column widths (weighted by a total) + 100% stacked segments

---

## Non-chart think-cell features

### Agenda — ⬜
- ⬜ Table-of-contents / agenda slides that auto-number and stay in sync across the deck;
      current-chapter highlight; regenerate when sections change. (think-cell's "Agenda")

### Standard slide elements — ⬜ (see "Where should slide elements live?" below)
- ⬜ Harvey balls, checkmarks/crosses, stoplights, arrows/pointers, brackets, ticks
- ⬜ Text boxes with the house style; process-flow / chevron shapes
- ⬜ These overlap with the **QuickTools** VBA ribbon add-in — decide one home (discussion below)

---

## Deferred / not feasible in Office.js
- ⚠️ Shared-runtime ribbon + auto-pane (deployment-path dependent)
- ❌ Embedded Excel datasheet (grid + paste is the substitute)
- ❌ Live external Excel links (would need Graph API)
- ❌ On-canvas anchored toolbar

---

## Testing backlog (needs Silvan in PowerPoint)
Everything marked 🔵 is unit-tested but not visually confirmed. Highest-value to verify first:
waterfall render, negative values, connectors, reference line, legend positions, colors, resize/move, number formats.
