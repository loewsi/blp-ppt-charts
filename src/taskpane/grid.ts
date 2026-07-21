// Lightweight Excel-like data grid (no dependency, full control).
// Full-sheet layout: row 0 = category names, column 0 = series names, interior
// = values. Excel-style two modes: "navigate" (arrows move the selection,
// typing starts an edit) and "edit" (arrows move the caret). Range selection by
// mouse drag or Shift+arrows; Ctrl+Space / Shift+Space select column / row;
// Ctrl+"+" / Ctrl+"-" insert / remove; Excel copy & paste (TSV).
//
// Pure helpers (cellsFromData/dataFromCells/pasteInto/insert*/remove*) are unit
// tested; the DOM/keyboard layer wraps them.
import type { ChartData, ChartSheet, Series } from "../model/chartModel";
import { PALETTE } from "../model/chartModel";
import { evaluateGrid } from "./formula";

export type Cells = string[][]; // cells[row][col]; row 0 = categories, col 0 = series

// ---- pure helpers --------------------------------------------------------
/** True when a cell is the waterfall "total" marker ("e" for equals, or "="). */
export function isTotalMarker(raw: string): boolean {
  const t = (raw ?? "").trim().toLowerCase();
  return t === "e" || t === "=";
}

export function cellsFromData(data: ChartData): Cells {
  const rows: Cells = [["", ...data.categories]];
  data.series.forEach((s, si) => {
    rows.push([
      s.name,
      // The first row carries the waterfall "e" markers so they round-trip.
      ...data.categories.map((_, i) => (si === 0 && data.totalFlags?.[i] ? "e" : String(s.values[i] ?? 0))),
    ]);
  });
  return rows;
}

export function dataFromCells(cells: Cells, colors: string[]): ChartData {
  const header = cells[0] ?? [""];
  const categories = header.slice(1).map((v, i) => (v.trim() ? v.trim() : `Cat ${i + 1}`));
  const series: Series[] = cells.slice(1).map((row, ri) => ({
    name: (row[0] ?? "").trim() || `Series ${ri + 1}`,
    color: colors[ri] || PALETTE[ri % PALETTE.length],
    values: categories.map((_, i) => {
      const n = Number((row[i + 1] ?? "").replace(/[,\s%]/g, ""));
      return isFinite(n) ? n : 0;
    }),
  }));
  // "e" / "=" in the first data row marks a waterfall total column.
  const firstRow = cells[1] ?? [];
  const totalFlags = categories.map((_, i) => isTotalMarker(firstRow[i + 1] ?? ""));
  return { type: "barColumn", categories, series, totalFlags };
}

/** Write a pasted TSV block into `cells` starting at (r0,c0), growing as needed. */
export function pasteInto(cells: Cells, r0: number, c0: number, tsv: string): Cells {
  const block = tsv.replace(/\r/g, "").replace(/\n+$/, "").split("\n").map((line) => line.split("\t"));
  const out = cells.map((row) => [...row]);
  const needRows = r0 + block.length;
  const needCols = c0 + Math.max(...block.map((b) => b.length));
  while (out.length < needRows) out.push([]);
  for (const row of out) while (row.length < needCols) row.push("");
  block.forEach((line, i) => line.forEach((val, j) => (out[r0 + i][c0 + j] = val)));
  return out;
}

export function insertColumn(cells: Cells, at: number): Cells {
  return cells.map((row) => {
    const r = [...row];
    r.splice(at, 0, "");
    return r;
  });
}
export function removeColumn(cells: Cells, at: number): Cells {
  if ((cells[0]?.length ?? 0) <= 2) return cells; // keep >=1 category
  return cells.map((row) => {
    const r = [...row];
    r.splice(at, 1);
    return r;
  });
}
export function insertRow(cells: Cells, at: number): Cells {
  const width = cells[0]?.length ?? 1;
  const out = cells.map((row) => [...row]);
  out.splice(at, 0, Array(width).fill(""));
  return out;
}
export function removeRow(cells: Cells, at: number): Cells {
  if (cells.length <= 2) return cells; // keep >=1 series
  const out = cells.map((row) => [...row]);
  out.splice(at, 1);
  return out;
}

// ---- DOM layer -----------------------------------------------------------
let host: HTMLElement;
let onChange: (() => void) | undefined;
let cells: Cells = [[""]];
let seriesColors: string[] = [];
let seriesKinds: (("bar" | "line") | undefined)[] = [];
let hiddenRows = new Set<number>(); // cell-row indices (≥1) hidden from view + chart
let hiddenCols = new Set<number>(); // cell-col indices (≥1)
let active = { r: 1, c: 1 };
let anchor = { r: 1, c: 1 };
let editing = false;
let dragging = false;
let computed = { display: [] as string[][], values: [] as number[][] }; // formula results, refreshed on render

export function mountGrid(container: HTMLElement, changed?: () => void): void {
  host = container;
  onChange = changed;
  document.addEventListener("mouseup", () => (dragging = false));
}

export function setGridData(data: ChartData): void {
  cells = cellsFromData(data);
  seriesColors = data.series.map((s) => s.color);
  seriesKinds = data.series.map((s) => s.kind);
  hiddenRows = new Set();
  hiddenCols = new Set();
  clampActive();
  render();
}

/** Restore the full editable surface (raw cells + colors/kinds + hidden flags). */
export function setSheet(sheet: ChartSheet): void {
  cells = sheet.cells.map((r) => [...r]);
  seriesColors = [...sheet.colors];
  seriesKinds = [...sheet.kinds];
  hiddenRows = new Set(sheet.hiddenRows ?? []);
  hiddenCols = new Set(sheet.hiddenCols ?? []);
  clampActive();
  render();
}

/** The editable surface to persist in the model. */
export function getSheet(): ChartSheet {
  const nSer = Math.max(0, cells.length - 1);
  return {
    cells: cells.map((r) => [...r]),
    colors: Array.from({ length: nSer }, (_, i) => seriesColors[i] || PALETTE[i % PALETTE.length]),
    kinds: Array.from({ length: nSer }, (_, i) => seriesKinds[i]),
    hiddenRows: [...hiddenRows],
    hiddenCols: [...hiddenCols],
  };
}

export function getGridData(): ChartData {
  // Chart data uses COMPUTED values and EXCLUDES hidden rows/cols (helper scaffolding).
  const { values } = evaluateGrid(cells);
  const header = cells[0] ?? [];
  const catCols: number[] = []; // visible category cell-columns
  const categories: string[] = [];
  for (let c = 1; c < header.length; c++) {
    if (hiddenCols.has(c)) continue;
    catCols.push(c);
    categories.push(header[c].trim() || `Cat ${c}`);
  }
  const series: Series[] = [];
  for (let r = 1; r < cells.length; r++) {
    if (hiddenRows.has(r)) continue;
    const ri = r - 1;
    series.push({
      name: (cells[r][0] ?? "").trim() || `Series ${r}`,
      color: seriesColors[ri] || PALETTE[ri % PALETTE.length],
      values: catCols.map((c) => values[r]?.[c] ?? 0),
      ...(seriesKinds[ri] ? { kind: seriesKinds[ri] } : {}),
    });
  }
  return { type: "barColumn", categories, series };
}

/** Hide the selected rows / columns (helper scaffolding) from view + chart. */
export function hideSelectedRows(): void {
  const { r0, r1 } = selRect();
  for (let r = Math.max(1, r0); r <= r1; r++) hiddenRows.add(r);
  render();
  emit();
}
export function hideSelectedCols(): void {
  const { c0, c1 } = selRect();
  for (let c = Math.max(1, c0); c <= c1; c++) hiddenCols.add(c);
  render();
  emit();
}
export function unhideAll(): void {
  hiddenRows.clear();
  hiddenCols.clear();
  render();
  emit();
}
export function hasHidden(): boolean {
  return hiddenRows.size > 0 || hiddenCols.size > 0;
}

// ---- structural edits (operate on the raw cells, so formulas are preserved) ----
// Structural edits clear hidden flags to avoid index-shift bugs (documented).
function clearHidden(): void {
  hiddenRows.clear();
  hiddenCols.clear();
}

export function gridAddSeries(): void {
  clearHidden();
  const at = active.r === 0 ? 1 : Math.min(cells.length, active.r + 1); // below the active series; top if in header
  const width = cells[0]?.length ?? 1;
  const row = Array(width).fill("");
  row[0] = `Series ${cells.length}`;
  cells.splice(at, 0, row);
  seriesColors.splice(at - 1, 0, PALETTE[(at - 1) % PALETTE.length]);
  seriesKinds.splice(at - 1, 0, undefined);
  clampActive();
  render();
  emit();
}

export function gridRemoveSeries(): void {
  clearHidden();
  const { r0, r1 } = selRect();
  const from = Math.max(1, r0);
  const count = Math.min(Math.min(cells.length - 1, r1) - from + 1, cells.length - 2); // keep ≥1 series
  if (count <= 0) return;
  cells.splice(from, count);
  seriesColors.splice(from - 1, count);
  seriesKinds.splice(from - 1, count);
  clampActive();
  render();
  emit();
}

export function gridAddCategory(): void {
  clearHidden();
  const width = cells[0]?.length ?? 1;
  const at = Math.max(1, active.c); // left of the active category
  cells.forEach((row, r) => row.splice(at, 0, r === 0 ? `Cat ${width}` : ""));
  clampActive();
  render();
  emit();
}

export function gridRemoveCategory(): void {
  clearHidden();
  const width = cells[0]?.length ?? 1;
  const { c0, c1 } = selRect();
  const from = Math.max(1, c0);
  const count = Math.min(Math.min(width - 1, c1) - from + 1, width - 2); // keep ≥1 category
  if (count <= 0) return;
  cells.forEach((row) => row.splice(from, count));
  clampActive();
  render();
  emit();
}

export function gridTranspose(): void {
  clearHidden();
  const R = cells.length;
  const C = cells[0]?.length ?? 1;
  const out: Cells = [];
  for (let c = 0; c < C; c++) {
    const row: string[] = [];
    for (let r = 0; r < R; r++) row.push(cells[r][c] ?? "");
    out.push(row);
  }
  cells = out;
  const nSer = Math.max(0, out.length - 1);
  seriesColors = Array.from({ length: nSer }, (_, i) => PALETTE[i % PALETTE.length]);
  seriesKinds = Array(nSer).fill(undefined);
  clampActive();
  render();
  emit();
}

export function setSeriesColor(index: number, color: string): void {
  seriesColors[index] = color;
}

export function setSeriesKind(index: number, kind: "bar" | "line"): void {
  seriesKinds[index] = kind;
}

/** Recolor all series from a palette (cycled) without touching the cells/formulas. */
export function applyPalette(palette: string[]): void {
  const nSer = Math.max(0, cells.length - 1);
  for (let i = 0; i < nSer; i++) seriesColors[i] = palette[i % palette.length];
  render();
  emit();
}

/** The active cell (row 0 = categories, col 0 = series). Used by remove-at-cursor. */
export function getActive(): { r: number; c: number } {
  return { r: active.r, c: active.c };
}

/** The current selection rectangle (inclusive), for multi-row/col operations. */
export function getSelectionRange(): { r0: number; r1: number; c0: number; c1: number } {
  return selRect();
}

// ---- rendering -----------------------------------------------------------
function render(): void {
  if (!host) return;
  computed = evaluateGrid(cells); // refresh formula results
  host.innerHTML = "";
  const table = document.createElement("table");
  table.className = "xgrid";
  cells.forEach((row, r) => {
    if (hiddenRows.has(r)) return; // hidden helper row
    const tr = document.createElement("tr");
    row.forEach((val, c) => {
      if (hiddenCols.has(c)) return; // hidden helper column
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.className = "xcell" + (r === 0 || c === 0 ? " xhead" : "");
      // Show the computed result; the raw formula appears only while editing.
      inp.value = computed.display[r]?.[c] ?? val;
      if ((val ?? "").trim().startsWith("=")) inp.classList.add("xformula");
      inp.readOnly = true;
      inp.dataset.r = String(r);
      inp.dataset.c = String(c);
      inp.addEventListener("input", () => (cells[r][c] = inp.value));
      inp.addEventListener("mousedown", (e) => onMouseDown(r, c, e));
      inp.addEventListener("mouseover", () => onMouseOver(r, c));
      inp.addEventListener("dblclick", () => beginEdit(r, c));
      inp.addEventListener("keydown", (e) => onKey(e, r, c));
      inp.addEventListener("blur", () => {
        // Focus left this cell while still editing (e.g. clicked a pane control): commit.
        if (editing && active.r === r && active.c === c) {
          editing = false;
          render(); // refresh the computed display
          emit();
        }
      });
      inp.addEventListener("copy", onCopy);
      inp.addEventListener("paste", onPaste);
      td.appendChild(inp);
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  host.appendChild(table);
  paint();
}

function inputAt(r: number, c: number): HTMLInputElement | null {
  return host.querySelector<HTMLInputElement>(`input.xcell[data-r="${r}"][data-c="${c}"]`);
}

function selRect() {
  return {
    r0: Math.min(anchor.r, active.r),
    r1: Math.max(anchor.r, active.r),
    c0: Math.min(anchor.c, active.c),
    c1: Math.max(anchor.c, active.c),
  };
}

function paint(): void {
  const { r0, r1, c0, c1 } = selRect();
  host.querySelectorAll<HTMLInputElement>("input.xcell").forEach((inp) => {
    const r = Number(inp.dataset.r);
    const c = Number(inp.dataset.c);
    inp.classList.toggle("sel", r >= r0 && r <= r1 && c >= c0 && c <= c1);
    inp.classList.toggle("active", r === active.r && c === active.c);
  });
}

function clampActive(): void {
  const R = cells.length;
  const C = cells[0]?.length ?? 1;
  active.r = Math.min(Math.max(0, active.r), R - 1);
  active.c = Math.min(Math.max(0, active.c), C - 1);
  anchor.r = Math.min(Math.max(0, anchor.r), R - 1);
  anchor.c = Math.min(Math.max(0, anchor.c), C - 1);
}

function moveTo(r: number, c: number, extend: boolean): void {
  const wasEditing = editing; // committing an edit must re-render the chart
  const R = cells.length;
  const C = cells[0].length;
  active = { r: Math.min(Math.max(0, r), R - 1), c: Math.min(Math.max(0, c), C - 1) };
  if (!extend) anchor = { ...active };
  editing = false;
  if (wasEditing) render(); // recompute formulas so the just-edited cell shows its result
  const el = inputAt(active.r, active.c);
  if (el) {
    el.readOnly = true;
    el.focus();
  }
  paint();
  if (wasEditing) emit();
}

function beginEdit(r: number, c: number, replaceWith?: string): void {
  active = { r, c };
  anchor = { r, c };
  editing = true;
  const el = inputAt(r, c);
  if (!el) return;
  el.readOnly = false;
  // Editing reveals the raw entry (the formula), not the computed value.
  el.value = replaceWith !== undefined ? replaceWith : cells[r][c] ?? "";
  cells[r][c] = el.value;
  el.focus();
  const end = el.value.length;
  el.setSelectionRange(end, end);
  paint();
}

// ---- mouse ---------------------------------------------------------------
function onMouseDown(r: number, c: number, e: MouseEvent): void {
  if (editing && r === active.r && c === active.c) return; // clicking inside the cell being edited
  const wasEditing = editing; // clicking away from an edit commits it
  dragging = true;
  if (e.shiftKey) {
    active = { r, c }; // extend from existing anchor
  } else {
    anchor = { r, c };
    active = { r, c };
  }
  editing = false;
  if (wasEditing) render(); // show the committed cell's computed result
  const el = inputAt(r, c);
  if (el) {
    el.readOnly = true;
    el.focus();
  }
  paint();
  if (wasEditing) emit();
}

function onMouseOver(r: number, c: number): void {
  if (!dragging) return;
  active = { r, c };
  paint();
}

// ---- keyboard ------------------------------------------------------------
function onKey(e: KeyboardEvent, r: number, c: number): void {
  if (editing) {
    if (e.key === "Enter") {
      e.preventDefault();
      moveTo(r + 1, c, false);
    } else if (e.key === "Tab") {
      e.preventDefault();
      moveTo(r, e.shiftKey ? c - 1 : c + 1, false);
    } else if (e.key === "Escape") {
      e.preventDefault();
      const el = inputAt(r, c);
      if (el) el.value = cells[r][c]; // input listener already committed; keep as-is
      moveTo(r, c, false);
    }
    return; // arrows etc. edit the caret
  }

  // ---- navigate mode ----
  if (e.code === "Space" && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    e.stopPropagation(); // don't let the host webview treat Ctrl+Space as its own shortcut
    anchor = { r: 0, c };
    active = { r: cells.length - 1, c };
    paint();
    return;
  }
  if (e.code === "Space" && e.shiftKey) {
    e.preventDefault();
    e.stopPropagation();
    anchor = { r, c: 0 };
    active = { r, c: cells[0].length - 1 };
    paint();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && (e.key === "+" || e.key === "=")) {
    e.preventDefault();
    const isCol = selRect().r0 === 0 && selRect().r1 === cells.length - 1;
    cells = isCol ? insertColumn(cells, Math.max(1, c)) : insertRow(cells, Math.max(1, r));
    seriesColors = syncColors();
    render();
    emit();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && (e.key === "-" || e.key === "_")) {
    e.preventDefault();
    const isCol = selRect().r0 === 0 && selRect().r1 === cells.length - 1;
    cells = isCol ? removeColumn(cells, Math.max(1, c)) : removeRow(cells, Math.max(1, r));
    seriesColors = syncColors();
    clampActive();
    render();
    emit();
    return;
  }
  if (e.key === "ArrowUp") return nav(e, r - 1, c);
  if (e.key === "ArrowDown" || e.key === "Enter") return nav(e, r + 1, c);
  if (e.key === "ArrowLeft") return nav(e, r, c - 1);
  if (e.key === "ArrowRight") return nav(e, r, c + 1);
  if (e.key === "Tab") return nav(e, r, e.shiftKey ? c - 1 : c + 1, false);
  if (e.key === "F2") {
    e.preventDefault();
    beginEdit(r, c);
    return;
  }
  if (e.key === "Delete" || e.key === "Backspace") {
    e.preventDefault();
    const { r0, r1, c0, c1 } = selRect();
    for (let rr = r0; rr <= r1; rr++) for (let cc = c0; cc <= c1; cc++) cells[rr][cc] = "";
    render();
    emit();
    return;
  }
  // A printable character starts editing (Excel: overwrite).
  if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
    e.preventDefault();
    beginEdit(r, c, e.key);
  }
}

function nav(e: KeyboardEvent, r: number, c: number, extendAllowed = true): void {
  e.preventDefault();
  moveTo(r, c, extendAllowed && e.shiftKey);
}

// ---- clipboard -----------------------------------------------------------
function onCopy(e: ClipboardEvent): void {
  const { r0, r1, c0, c1 } = selRect();
  const lines: string[] = [];
  for (let r = r0; r <= r1; r++) lines.push(cells[r].slice(c0, c1 + 1).join("\t"));
  e.preventDefault();
  e.clipboardData?.setData("text/plain", lines.join("\n"));
}

function onPaste(e: ClipboardEvent): void {
  const text = e.clipboardData?.getData("text") ?? "";
  if (!text) return;
  if (!text.includes("\t") && !text.includes("\n")) {
    // single value → let it edit the active cell normally
    if (!editing) beginEdit(active.r, active.c, text);
    e.preventDefault();
    if (editing) cells[active.r][active.c] = text;
    render();
    emit();
    return;
  }
  e.preventDefault();
  const { r0, c0 } = selRect();
  cells = pasteInto(cells, r0, c0, text);
  seriesColors = syncColors();
  render();
  emit();
}

function syncColors(): string[] {
  const nSeries = Math.max(0, cells.length - 1);
  const out: string[] = [];
  for (let i = 0; i < nSeries; i++) out.push(seriesColors[i] || PALETTE[i % PALETTE.length]);
  return out;
}

function emit(): void {
  paint();
  if (onChange) onChange();
}
