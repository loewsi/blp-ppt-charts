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

## A1. Fixes from your last feedback — please re-verify
- [ ] `+ Series` / `+ Category` now insert **at the cursor** (below the row / left of the column), not at the end.
- [ ] Number **Magnitude** reads **×10⁻³ / ×10⁻⁶ / ×10³** etc. (no k/M letter); **Group thousands** on by default.
- [ ] A `0` value's label sits **in stack order**, not pinned to the axis.
- [ ] New color schemes **Green / Red / Orange / Multi-color**; **Auto-shades from base** ramps one color across series.
- [ ] Pasted chart gets the **next free "Chart N"** name; deleting a chart also clears its **orphan legend**.
- [ ] **Difference arrow**: the **Position** field places it at a chosen slot boundary (blank = auto).
- [ ] **CAGR**: now a **horizontal arrow above the chart** with a white % bubble (not a vertical span).
- [ ] **Secondary axis**: **Line axis min/max** fields; axis line on **both sides** when active.
- [ ] Panel shows **only relevant options** per chart type (no CAGR fields unless CAGR on, no doughnut hole unless pie…).
- [ ] **Legend**: repositioning it no longer makes the graph jump (legend now sits outside the plot).
- [ ] Labels: **Outside** moves *all* value labels out; wide labels get a colored chip; overlapping small labels spread both ways.

## A2. Brand-new — never rendered
- [ ] **Line** chart → all series as connected lines with markers + labels.
- [ ] **Scatter / bubble** → rows = X, Y, optional Size; each column a point; **Quadrant lines** toggle. (Pane shows a hint.)
- [ ] **Mekko** → column width ∝ category total; 100%-stacked; % + column totals.
- [ ] **Waterfall** → type **"e"** in a value cell → a computed total column from the baseline to the running sum.
- [ ] **Pie / doughnut** → slices from the first series; % inside, category outside; **Doughnut hole %** > 0 → ring with the total. (Was "not round" — denser facets now; confirm.)
- [ ] **Agenda** (bottom of pane) → list chapters → *Create / update agenda slide*; run again → updates the same slide.

## A3. Known-open — don't test (not built)
Axis labels shrink the plot (#8) · sync-axis across charts · per-segment color override ·
waterfall connector-controlled totals · line area fill · standard slide elements. See [FEATURES.md](FEATURES.md).

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
- ✅ Difference arrow works.
- ✎ Want to **choose where the difference arrow sits** → **fixed** (Position field, re-verify A1). **CAGR** should sit **above the chart** with a % bubble → **fixed** (re-verify A1). Arrowheads look hand-made / seams show → **partly**: guides are now real dashed connectors, but heads stay custom triangles (Office.js has no PowerPoint arrowhead API).

## Labels
- ✅ Inside color chips (you liked them) — "works, but not everywhere".
- ✎ Large numbers went white-on-white; small labels near the axis didn't move → **fixed** (wide-label chip + anchor spread; re-verify A1).

## Waterfall & doughnut (were broken/wrong → reworked)
- ✎ Waterfall "logic doesn't work yet" → **reworked** with running totals + the **"e" total column** (test A2).
- ✎ Doughnut "not round" → **denser facets** + geometry proven by unit test (test A2).
