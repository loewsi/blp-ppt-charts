# SlideChart — testing & feedback log

**How this works.** Two parts:
- **Part A — Needs your test:** built but not yet confirmed in PowerPoint (Claude renders blind). Tick `[x]` as you go.
- **Part B — Tested & your feedback:** what you've confirmed, plus your notes. Your feedback here is the
  input for my next changes and for [FEATURES.md](FEATURES.md). When an A-item passes, it moves to B.

**Setup after a code change:** close PowerPoint → clear the webview cache
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` → reopen. The add-in loads from the GitHub Pages URL.
Last updated: **2026-07-21 pm**.

---

# Part A — Needs your test 🔲

## A1. Fixes from your 2026-07-21 pm feedback — please re-verify
- [ ] Grid: `+ Series` inserts **below** the cursor (top when in the header row); `+ Category` **left** of the cursor, consistently; **multi-row/col select deletes all** selected.
- [ ] Excel paste larger than the grid **grows the table**.
- [ ] Segment labels near the top **offset sideways** so they don't collide with the total *(if I got to it — else still stacking)*.
- [ ] **Arrows**: solid lines (no dashed); the line no longer pokes past the arrow tip.
- [ ] **CAGR arrow**: now **sloped** (from above one total to above the other), rounded bubble, **number-only** (e.g. `+50%`, no "CAGR").
- [ ] **Line chart**: no totals; difference/CAGR compare **series** (not totals).
- [ ] **Combination**: the per-series **"line"** checkbox and **2nd-axis** controls show **only** for Combination; line axis min/max disappear when the 2nd axis is off.
- [ ] **Legend**: a **line** series shows a **line marker**, not a box.
- [ ] **Label shows** selector: Value / Percent / Value (%) / % (Value) — on bar + mekko.
- [ ] **Mekko**: columns are **contiguous** (no gap); totals show.
- [ ] **Scatter**: rows auto-named **X / Y / Size**; **Scatter axes** toggle hides the axes.
- [ ] **Pie**: changing the color **recolors** it (slices are shades of the series color); **white hairlines** gone.

## A2. Brand-new — first render
- [ ] **Waterfall "e"** total column (baseline → running sum). *Note: v1 is single-series; multi-series is a separate decision (below).*
- [ ] **Agenda** (bottom of pane): list chapters → *Create / update agenda slide*; re-run updates it.
- [ ] **Doughnut hole %** > 0 → ring with the total in the middle.

## A3. Known-open — don't test (not built / needs your decision)
- Axis labels shrink the plot (#8) · sync-axis across charts · per-segment color override · line area fill · standard slide elements.
- **Pie jumps to its initial position on edit** — logged; may need the box-model work.
- **Decisions needed** (see bottom of this file): Excel-style **formulas**, a **foldable/Excel-like grid**, **multi-series waterfall**.
See [FEATURES.md](FEATURES.md) for full status.

---

# Part B — Tested & confirmed ✅ (+ your feedback)

Legend: **✅** you confirmed it works · **✎** your feedback → what I did (→ **fixed**, re-verify in A · → **open** · → **future**).

## Core & lifecycle
- ✅ Insert · blank-click deselect · click-to-load · live cell edit · move (no reload) · resize · shrink stays visible · save/reopen · same-slide copy-paste gives an independent copy.

## Data grid
- ✅ Edit categories/series/values · arrow-nav · Excel paste · transpose · remove series/category · Shift+Space.
- ✎ `− Series`/`− Category` removed the *last*, not the cursor's → **fixed**. `+` appended at the end → **fixed** (re-verify A1).
- ✎ **Ctrl+Space** opens an OS menu → **open** (Windows OS shortcut; can't be intercepted from the pane).

## Chart styles
- ✅ Orientation column/bar · arrangements · gap · totals · value labels · reverse categories · reverse series · gridlines · value axis · connectors · reference line · reference color · axis min/max.
- ✎ 100% stacked should show the absolute total → **fixed** (✅). Reverse only affected categories → **fixed** (reverse series added, ✅).
- ✎ Legend positions worked but the **graph jumped** when switching → **fixed** (legend drawn outside the plot; re-verify A1).
- ✎ Turning on **axis labels resizes the graph** → **open** (#8, plot-anchored box model).

## Negatives
- ✅ Bars extend below the baseline; axis spans below zero.

## Number format
- ✅ Decimals · prefix/suffix · group thousands · hide-zero · negatives ( ) · plus sign.
- ✎ Don't append a k/M letter for scale — want a magnitude like ×10⁻³ → **fixed** (re-verify A1). Set group-thousands on by default → **done**. Hidden-zero label was pinned to the axis → **fixed** (stack order; re-verify A1).

## Colors
- ✅ Per-series swatch · Blue scheme.
- ✎ Call it "Blue" not "BLP" → **done**. Want more colors + a master-based picker + auto-generated shades → **partly done** (Green/Red/Orange/Multi + Auto-shades added, re-verify A1; per-series master-scheme picker still **open**).

## Fonts
- ✅ Font family; independent segment/total sizes.

## Combination & line
- ✅ Combination works (mark a series as a line); secondary axis works; line chart works (lightly tested).
- ✎ Want to set the **secondary-axis scale** + axis lines on both sides → **fixed** (re-verify A1).

## Arrows
- ✅ Difference arrow works; positioning works.
- ✎ Want to **choose placement** → **done** (Position field). Line rectangle shows at the arrow tip / dashed looked off → **fixed** (solid line, body stops at the head). **CAGR** should be **sloped** with a round bubble and number-only → **fixed** (re-verify A1). Heads stay custom triangles (Office.js has no PowerPoint arrowhead API).

## Grid (this round)
- ✎ `+ Category` inconsistent (right when in a series row); `+ Series` skipped a row when in the header; multi-select deleted only one → **fixed** (re-verify A1). Excel paste bigger than the grid should enlarge it → **works** (re-verify A1).

## Panel / legend (this round)
- ✎ Show the "line" option only for Combination; line min/max should vanish when 2nd axis is off; a line series should show a **line** in the legend → **fixed** (re-verify A1).

## Chart types (this round)
- ✎ Line should have **no totals**; diff/CAGR only between series → **fixed**. Scatter unclear + want to hide axes + rename rows → **fixed** (auto X/Y/Size names + Scatter-axes toggle). Mekko gap between columns → **removed**; want value/%/combo labels → **done** (Label shows). Pie color didn't change + white lines → **fixed** (shades of series color + facet overlap). Pie **jumps on edit** → **open**.

## Labels
- ✅ Inside color chips (you liked them) — "works, but not everywhere".
- ✎ Large numbers went white-on-white; small labels near the axis didn't move → **fixed** (wide-label chip + anchor spread; re-verify A1).

## Waterfall & doughnut (were broken/wrong → reworked)
- ✎ Waterfall "logic doesn't work yet" → **reworked** with running totals + the **"e" total column** (test A2).
- ✎ Doughnut "not round" → **denser facets** + geometry proven by unit test (test A2).

---

# Decisions needed (big items — not built yet)
These are sizable and worth your steer before I build:
1. **Excel-style formulas in the grid** (`=A1+B1`, `=SUM(...)`). Useful but means a formula parser + a
   cell-reference model. Worth it, or keep values-only and rely on paste-from-Excel?
2. **Foldable / true-Excel grid** (hide/collapse rows/columns for calculations). This likely means
   replacing the lightweight grid with a heavier component. Do you want that, or is the current grid enough?
3. **Multi-series waterfall** ("e still shows one line, not all"). Today waterfall reads one series of
   deltas. Multi-series (stacked steps per category) is a real rebuild — do you want it, and how should
   several series combine per step?
