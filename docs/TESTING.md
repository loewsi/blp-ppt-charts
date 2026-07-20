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

## 1. Core lifecycle
- [ ] **Insert** a chart → default stacked column appears.
- [ ] Click a **blank area** → pane says "No chart selected".
- [ ] Click the **chart** → its data + options load back into the pane.
- [ ] Edit a **cell** → chart updates automatically (no Update button), ~0.4s later.
- [ ] **Move** the chart → it does *not* reload/flicker.
- [ ] **Resize** the chart → it re-lays-out once you release (labels/axis reflow).
- [ ] Shrink it **very small** → chart stays visible (doesn't disappear).
- [ ] **Copy-paste** the chart → the copy is independently editable.
- [ ] Save, close, reopen the .pptx → chart still editable (data survived in the file).

## 2. Data grid
- [ ] Type in categories (top row) and series names (left column).
- [ ] Arrow-key navigation; Enter/Tab commit.
- [ ] **Paste** a block from Excel → grid fills.
- [ ] `+ Series` / `− Series` / `+ Category` / `− Category` / **Transpose**.
- [ ] Range-select with mouse and **Shift+arrows**; **Ctrl+Space** / **Shift+Space**.

## 3. Chart styles (Chart style panel)
- [ ] Orientation **Column ↔ Bar**.
- [ ] Arrangement **Stacked / Clustered / 100% stacked**.
- [ ] **Gap %** widens/narrows bars.
- [ ] **Totals**, **Value labels**, **Reverse order** toggles.
- [ ] **Legend** on + each position **top/bottom/left/right**; legend text word-wraps to its box;
      drag the legend, then edit data → legend stays where you put it.
- [ ] **Gridlines** and **Value axis** toggles.
- [ ] **Connectors** (stacked column only) draw between adjacent stacks.
- [ ] **Reference line**: type a value → red labelled line at that value (horizontal for column,
      vertical for bar); clear it → line disappears; a value above the axis max shows nothing.
- [ ] **Axis min / max**: fix either end → bars rescale to the fixed range; blank = auto.

## 4. Negative values
- [ ] Enter negative numbers → bars extend **below** the zero baseline; axis spans below zero.

## 5. Labels
- [ ] Small segments keep their label **inside** as a color chip, nudged to avoid overlap.
- [ ] Switch small labels to **Move outside** → labels sit outside, no chip.
- [ ] Label boxes are only **as wide as the text**, in **front** of bars, vertically centered.

## 6. Number format
- [ ] Decimals, **Scale** k/M, **Prefix**, **Suffix**, **Hide zeros**.
- [ ] **Thousands ,** (note: your de-CH locale may show `1'234`), **Negatives ( )**, **Plus sign**.

## 7. Fonts & colors
- [ ] Font family applies to all labels; segment vs total sizes independent.
- [ ] Color scheme **Master accents** (reads the deck theme) / **BLP** / **Grayscale** → Apply.
- [ ] Per-series color swatch → recolors that whole series.

## 8. Waterfall
- [ ] Chart type **Waterfall** → first series = deltas; rise/fall bars; running total; connectors;
      signed (+/−) labels; zero baseline.

---

## Known-open (not built / broken — don't test yet)
- **Copy-paste on the SAME slide** duplicates the internal id → selecting/editing gets confused
  (works across different slides). Duplicate-id repair is the next fix.
- **Waterfall** needs a rework (categories/series naming + a think-cell "e" total cell +
  connector-controlled totals) — current version is placeholder-level.
- Difference arrows, CAGR arrows (both need real arrowheads), combination line series,
  per-segment color override, error bars, sync-axis-across-charts, color picker + auto-shades,
  global inside/outside label placement. See [FEATURES.md](FEATURES.md).
