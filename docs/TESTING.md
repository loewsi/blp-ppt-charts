# SlideChart — PowerPoint testing checklist

Everything below is **built + unit-tested** but **not yet visually confirmed in PowerPoint**
(Claude renders blind). Walk this list in PowerPoint and mark each ✅ / ✍️ (notes). Then I flip
the corresponding 🔵 → ✅ in [FEATURES.md](FEATURES.md).

**Setup once per code change:** close PowerPoint → clear the webview cache
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` → reopen. The add-in loads from the GitHub Pages URL.

---

## ⭐ Re-verify first — fixed from your 2026-07-20 feedback
- [ ] **Typing a value now updates the chart** (no transpose needed) — commit with Enter/Tab or click away.
- [ ] **Click blank area** → pane resets to "No chart selected".
- [ ] **− Series / − Category** now removes the row/column your cursor is in (not the last).
- [ ] **Ctrl+Space** selects the column without opening a menu (may still hit an OS language shortcut — tell me if so).
- [ ] **100% stacked** now shows the **absolute** total on top.
- [ ] **Reverse series** checkbox (separate from Reverse categories).
- [ ] **Legend position** switching works directly now (no need to remove + re-add); a manually-dragged legend still stays put until you pick a new position.
- [ ] **Axis line** checkbox draws the y-axis (column) / value-scale line (bar).
- [ ] **Reference color** picker next to Reference line.
- [ ] **Number magnitude** (÷1,000 / ÷1,000,000) no longer appends a "k"/"M" — add your own suffix.
- [ ] **Separators** dropdown: pick `1,234.56` / `1.234,56` / `1'234.56` / `1 234,56` / System.
- [ ] **Hide zeros** toggle is respected (off → a `0` shows at the baseline; on → hidden).
- [ ] **Large value labels** that are wider than the bar now get a colored chip (no more white-on-white).
- [ ] **Overlapping small labels** now spread to both sides — including the one nearest the axis.
- [ ] **Colors:** scheme renamed **Blue** (was "BLP"), with more shades.

---

## ⭐⭐⭐ Latest (2026-07-21 pm) — feedback fixes + all remaining chart types

**Your feedback fixes**
- [ ] `+ Series` inserts **below** the cursor's row; `+ Category` **left** of its column.
- [ ] Number **Magnitude** now reads ×10⁻³ / ×10⁻⁶ / ×10³ etc. (no k/M letter); **Group thousands** on by default.
- [ ] A `0` value's label sits **in stack order**, not pinned to the axis.
- [ ] More color schemes: **Green / Red / Orange / Multi-color**.
- [ ] Pasted chart gets the **next free "Chart N"** name; deleting a chart also removes its **orphan legend**.
- [ ] **Difference arrow**: "Position" field places it at a chosen slot boundary (blank = auto).
- [ ] **CAGR** now shows as a **horizontal arrow above the chart** with a white % bubble.
- [ ] Guide lines are **real dashed connectors**.
- [ ] **Secondary axis**: Line axis min/max fields; axis line on **both sides** when active.
- [ ] **Doughnut** should look **round** now (denser facets) — please confirm.
- [ ] Panel shows **only relevant options** (no CAGR fields unless CAGR on, no doughnut hole unless pie, etc.).
- [ ] **Legend**: changing position no longer makes the graph jump (legend sits outside the plot).
      ⚠️ Still open: turning on **axis labels** shrinks the plot (plot-anchored box model pending).

**New chart types** (Chart dropdown)
- [ ] **Scatter / bubble**: rows = X, Y, optional Size; each column a point; "Quadrant lines" toggle. (Hint shows in pane.)
- [ ] **Mekko**: column width ∝ category total; 100%-stacked; % + column totals.
- [ ] **Waterfall**: type **"e"** in a value cell → a computed total column to the baseline. (Hint shows in pane.)

**Agenda** (bottom of pane)
- [ ] Type chapters (one per line) → **Create / update agenda slide** adds a numbered "Agenda" slide;
      run it again → updates the same slide in place.

**Colors**
- [ ] **Auto-shades from base**: pick a base color → series become a light→dark ramp of it.

---

## ⭐⭐ Overnight (2026-07-21 am) — brand-new, test with fresh eyes
These are freshly built and **never seen rendered**. Expect rough edges; note anything off.

**Bug fixes**
- [ ] **Same-slide copy-paste** now yields an independently editable copy (select the copy → edits only affect it).

**Arrows** (Chart style → Arrows section)
- [ ] **Difference arrow**: set "Between totals", From cat 1, To cat 2 → double-headed arrow with signed delta (e.g. `+20`); tick "Show %" → `+20 (+50%)`. Try "Between a series" + a series #.
- [ ] Arrows have real **triangular heads** (not plain lines).
- [ ] **CAGR arrow**: set "Of totals", From/To, Periods (or leave 0 = auto) → label like `CAGR +12.5%`.

**Combination**
- [ ] In the Colors list, tick **line** next to a series → that series draws as a line + markers over the bars.
- [ ] Chart type **Combination** in the dropdown.
- [ ] **Line 2nd axis** checkbox → line series scale on their own right-hand axis (turn on Axis labels to see right-side ticks).

**New chart types** (Chart dropdown)
- [ ] **Line** → all series as connected lines with markers + labels.
- [ ] **Pie / doughnut** → slices sized by the first series; % inside, category outside. Set **Doughnut hole %** > 0 → ring with the total in the middle. ⚠️ Slices are a facet fan — check they look round.

---

> `[x]` = confirmed by Silvan in PowerPoint. `[ ]` = still to verify (new or changed since last check).

## 1. Core lifecycle
- [x] **Insert** a chart → default stacked column appears.
- [x] Click a **blank area** → pane says "No chart selected".
- [x] Click the **chart** → its data + options load back into the pane.
- [x] Edit a **cell** → chart updates automatically (no Update button), ~0.4s later.
- [x] **Move** the chart → it does *not* reload/flicker.
- [x] **Resize** the chart → it re-lays-out once you release (labels/axis reflow).
- [x] Shrink it **very small** → chart stays visible (doesn't disappear).
- [x] **Copy-paste** the chart → the copy is independently editable.
- [x] Save, close, reopen the .pptx → chart still editable (data survived in the file).

## 2. Data grid
- [x] Type in categories (top row) and series names (left column).
- [x] Arrow-key navigation; Enter/Tab commit.
- [x] **Paste** a block from Excel → grid fills.
- [x] `− Series` / `− Category` / **Transpose** · [ ] `+ Series`/`+ Category` now insert at the cursor.
- [x] Range-select with mouse and **Shift+arrows**; **Shift+Space**.
- [ ] **Ctrl+Space** — likely captured by a Windows OS shortcut (report if it still opens a menu).

## 3. Chart styles (Chart style panel)
- [x] Orientation **Column ↔ Bar**.
- [x] Arrangement **Stacked / Clustered / 100% stacked** (100% shows the absolute total).
- [x] **Gap %** widens/narrows bars.
- [x] **Totals**, **Value labels**, **Reverse categories**, **Reverse series** toggles.
- [x] **Legend** positions work · [ ] re-verify: moving the legend no longer makes the graph jump.
- [x] **Gridlines** and **Value axis** toggles · ⚠️ axis labels still shrink the plot (#8, open).
- [x] **Connectors** (stacked column only) draw between adjacent stacks.
- [x] **Reference line** at a value · [x] **Reference color** · [x] **Axis min / max**.

## 4. Negative values
- [x] Enter negative numbers → bars extend **below** the zero baseline; axis spans below zero.

## 5. Labels
- [x] Small segments keep their label **inside** as a color chip (Silvan: works, but not everywhere).
- [ ] Switch to **Outside** → ALL value labels sit outside, no chip.
- [ ] Wide labels get a chip (no white-on-white); overlapping small labels spread both ways.

## 6. Number format
- [x] Decimals, **Prefix**, **Suffix**, **Hide zeros**, **Group thousands**, **Negatives ( )**, **Plus sign**.
- [ ] **Magnitude** now reads ×10⁻³/×10³ etc.; a `0` sits in stack order; **Separators** dropdown.

## 7. Fonts & colors
- [x] Font family applies to all labels; segment vs total sizes independent.
- [x] Color scheme **Blue** + per-series swatch · [ ] Green/Red/Orange/Multi, Master accents, Auto-shades.

## 8. Waterfall — reworked, re-verify
- [ ] Deltas render rise/fall with running total + connectors + baseline.
- [ ] Type **"e"** in a value cell → a computed total column from the baseline to the running sum.

---

## Known-open (don't test yet)
- **Axis labels shrink the plot** (#8) — plot-anchored box model pending.
- **Sync-axis across charts**, **per-segment color override**, **waterfall connector-controlled totals**,
  **line area fill**, **standard slide elements** (pending the "where should they live" decision).
See [FEATURES.md](FEATURES.md) for the full status.
