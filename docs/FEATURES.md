# SlideChart — features & status

Living checklist, reconciled with the original think-cell-style spec. Organized as
**General** (cross-chart — built in the shared engine, so every chart type inherits it)
and **Per chart type**. Status: ✅ done · 🟡 partial · ⬜ todo · ⚠️ constrained/deferred.

The architecture is shared: model → layout → render-plan → PowerPoint adapter. Anything under
**General** applies to all chart types once that type has a layout function.

---

## General (applies across chart types)

### Data & editing
- ✅ Excel-like grid: editable categories (row 0) + series (col 0) + values
- ✅ Keyboard nav, range select (mouse + Shift+arrows), Ctrl+Space/Shift+Space, Ctrl +/−
- ✅ Copy / paste from Excel (native in the grid)
- ✅ Add / remove / reorder / transpose series & categories
- ✅ Live apply — any edit re-renders the selected chart automatically
- ✅ Insert a fresh chart; duplicate by copy-paste on the slide
- ✅ Positive / zero values · 🟡 negative values (below-baseline handling not done)
- ⬜ Use dates as categories; custom "value = 100%" base

### Colors
- ✅ Master theme accents (live) · BLP (blue) · grayscale presets
- ✅ Per-series color (swatch), applied across the whole series
- ⬜ Per-segment color override (one segment a different color) + highlight
- ⬜ Series outline (color/width/style), pattern fill

### Segment / value labels
- ✅ Show/hide, inside chip with collision offset, or move outside
- ✅ Value or % (100% stacked); box sized to text; in front of bars; vertical-centered
- ⬜ Custom prefix/suffix text per label, datasheet text, footnotes
- ⬜ Manual drag of a single label + leader line (model supports offset later)

### Total labels
- ✅ Totals above stacks, independent font size, sized to text
- ⬜ Show totals only for selected categories; custom total text; drag a total

### Legend
- ✅ Add/remove, position top/bottom/left/right, box auto-fits the name, own group (movable)
- ⬜ Reorder legend entries independently; format one entry; border

### Axis, gridlines, baseline
- ✅ Baseline; optional value axis (auto scale, tick labels); optional gridlines
- ⬜ Manual min/max/tick; secondary axis; axis break; axis title

### Number formatting (shared engine)
- ✅ Decimals, k/M scale, prefix, suffix, hide-zero
- ⬜ Percent/currency presets, thousands separators toggle, parentheses for negatives, per-type formats

### Fonts / text
- ✅ Font family; independent segment & total sizes; middle+centered in box
- ⬜ Bold/italic/color per label type; per-type independent everything

### Interaction / lifecycle
- ✅ Click a chart → loads into the pane; edits live
- ✅ Move chart (no reload); resize → re-lays-out on release; tiny sizes stay visible
- ✅ Legend stays where you move it across redraws
- ✅ Data stored in the .pptx (survives reopen, travels with the file)
- ⬜ Copy-to-another-slide/deck **duplicate-chartId repair**; undo-step quality
- ⚠️ Ribbon "insert type" menu + auto open/close pane — needs shared runtime; blocked on
  trusted-catalog sideload. Revisit via M365 central deploy / AppSource.

### Annotations (data-driven) — none built yet
- ⬜ **Series connectors** (between stacked segments across categories) — also enables waterfall
- ⬜ **Difference arrows** (level & total, absolute & %)
- ⬜ **CAGR arrows** (auto growth over a period; chart-area grows to fit)
- ⬜ Value / reference lines
- ⬜ Error bars / ranges (football field)

---

## Per chart type

### Column / bar family — ✅ core done
- ✅ Simple column (1 series), stacked, clustered, 100% stacked
- ✅ Horizontal bar variants (rotate; labels stay horizontal)
- ⬜ Combination (a series as a line), line markers, secondary axis
- 🟡 Negative values (render path exists; below-baseline layout not handled)
- ⬜ Category gaps / visual grouping; per-cluster gaps

### Waterfall — ⬜ NEXT
- ⬜ Running totals with rise/fall bars, connectors, subtotal columns
- Reuses labels/colors/legend/axis/number-format from General

### Mekko / Marimekko — ⬜ later
### Pie / doughnut — ⬜ later (out of the bar family; only if wanted)

---

## Explicitly not feasible in Office.js (documented)
- ❌ Embedded Excel datasheet (grid + paste is the substitute)
- ❌ Live external Excel links (would need Graph API; deferred)
- ❌ On-canvas anchored toolbar; ⚠️ shared-runtime ribbon/auto-pane (deploy-path dependent)

---

## Recommended order from here
1. **Series connectors** (small, general, and a prerequisite for waterfall)
2. **Waterfall** chart type (your original end goal)
3. **Difference arrows**, then **CAGR arrows**
4. Per-segment color, negative-value handling, richer number formats
5. Combination (line) series
