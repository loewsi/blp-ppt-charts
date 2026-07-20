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
- 🔵 Add / remove / reorder / transpose series & categories
- 🔵 Live apply (any edit re-renders the selected chart)
- 🔵 Insert a fresh chart; duplicate by copy-paste
- 🔵 Positive / zero / negative values
- ⬜ Dates as categories; custom "value = 100%" base

### Colors
- 🔵 Master theme accents (live) · BLP · grayscale presets; per-series swatch
- ⬜ Per-segment color override + highlight; series outline; pattern fill

### Segment / value labels
- ✅ Inside chip with collision offset (Silvan: "labels inside are fine, I like them")
- 🔵 Show/hide, move-outside option, value or %, box sized to text, in front of bars, vertical-centered
- ⬜ Custom prefix/suffix text per label; datasheet text; footnotes; manual drag + leader line

### Total labels
- 🔵 Totals above stacks, independent font size, sized to text
- ⬜ Totals for selected categories only; custom total text; drag a total

### Legend
- 🔵 Add/remove, position top/bottom/left/right, autofit box, own movable group
- ⬜ Reorder entries independently; format one entry; border

### Axis, gridlines, baseline
- 🔵 Baseline; optional value axis (auto scale, ticks); optional gridlines
- ⬜ Manual min/max/tick; secondary axis; axis break; axis title

### Number formatting (shared engine)
- 🔵 Decimals, k/M scale, prefix, suffix, hide-zero
- 🔵 Percent, thousands separator, negatives in parentheses, plus sign  *(new)*
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
- 🔵 Value / reference line (target line at a fixed value, labelled)  *(new)*
- ⬜ Error bars / ranges

---

## Per chart type

### Column / bar family
- 🔵 Simple column, stacked, clustered, 100% stacked, horizontal bar
- 🔵 Negative values below baseline  *(new)*
- ⬜ Combination — one series as a line
- ⬜ Line markers, secondary axis; category gaps / visual grouping

### Waterfall
- 🔵 Running totals, rise/fall bars, connectors, zero baseline, signed labels
- ⬜ Subtotal / total columns anchored to baseline
- ⬜ Bar (horizontal) orientation; per-bar color override

### Mekko / Marimekko — ⬜
### Pie / doughnut — ⬜ (only if wanted; outside the bar family)

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
