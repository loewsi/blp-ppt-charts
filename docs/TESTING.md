# SlideChart — PowerPoint testing checklist

Everything below is **built + unit-tested** but **not yet visually confirmed in PowerPoint**
(Claude renders blind). Walk this list in PowerPoint and mark each ✅ / ✍️ (notes). Then I flip
the corresponding 🔵 → ✅ in [FEATURES.md](FEATURES.md).

**Setup once per code change:** close PowerPoint → clear the webview cache
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` → reopen. The add-in loads from the GitHub Pages URL.

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

## Known-open (not built — don't test yet)
Difference arrows, CAGR arrows, combination line series, waterfall subtotal columns,
per-segment color override, error bars. See [FEATURES.md](FEATURES.md).
