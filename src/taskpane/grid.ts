// Lightweight Excel-like data grid (no dependency, full control).
// Full-sheet layout: row 0 = category names, column 0 = series names, interior
// = values. Excel-style two modes: "navigate" (arrows move the selection,
// typing starts an edit) and "edit" (arrows move the caret). Range selection by
// mouse drag or Shift+arrows; Ctrl+Space / Shift+Space select column / row;
// Ctrl+"+" / Ctrl+"-" insert / remove; Excel copy & paste (TSV).
//
// Pure helpers (cellsFromData/dataFromCells/pasteInto/insert*/remove*) are unit
// tested; the DOM/keyboard layer wraps them.
import type { ChartData, Series } from "../model/chartModel";
import { PALETTE } from "../model/chartModel";

export type Cells = string[][]; // cells[row][col]; row 0 = categories, col 0 = series

// ---- pure helpers --------------------------------------------------------
export function cellsFromData(data: ChartData): Cells {
  const rows: Cells = [["", ...data.categories]];
  data.series.forEach((s) => {
    rows.push([s.name, ...data.categories.map((_, i) => String(s.values[i] ?? 0))]);
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
  return { type: "barColumn", categories, series };
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
let active = { r: 1, c: 1 };
let anchor = { r: 1, c: 1 };
let editing = false;
let dragging = false;

export function mountGrid(container: HTMLElement, changed?: () => void): void {
  host = container;
  onChange = changed;
  document.addEventListener("mouseup", () => (dragging = false));
}

export function setGridData(data: ChartData): void {
  cells = cellsFromData(data);
  seriesColors = data.series.map((s) => s.color);
  clampActive();
  render();
}

export function getGridData(): ChartData {
  return dataFromCells(cells, seriesColors);
}

export function setSeriesColor(index: number, color: string): void {
  seriesColors[index] = color;
}

// ---- rendering -----------------------------------------------------------
function render(): void {
  if (!host) return;
  host.innerHTML = "";
  const table = document.createElement("table");
  table.className = "xgrid";
  cells.forEach((row, r) => {
    const tr = document.createElement("tr");
    row.forEach((val, c) => {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.className = "xcell" + (r === 0 || c === 0 ? " xhead" : "");
      inp.value = val;
      inp.readOnly = true;
      inp.dataset.r = String(r);
      inp.dataset.c = String(c);
      inp.addEventListener("input", () => (cells[r][c] = inp.value));
      inp.addEventListener("mousedown", (e) => onMouseDown(r, c, e));
      inp.addEventListener("mouseover", () => onMouseOver(r, c));
      inp.addEventListener("dblclick", () => beginEdit(r, c));
      inp.addEventListener("keydown", (e) => onKey(e, r, c));
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
  const R = cells.length;
  const C = cells[0].length;
  active = { r: Math.min(Math.max(0, r), R - 1), c: Math.min(Math.max(0, c), C - 1) };
  if (!extend) anchor = { ...active };
  editing = false;
  const el = inputAt(active.r, active.c);
  if (el) {
    el.readOnly = true;
    el.focus();
  }
  paint();
}

function beginEdit(r: number, c: number, replaceWith?: string): void {
  active = { r, c };
  anchor = { r, c };
  editing = true;
  const el = inputAt(r, c);
  if (!el) return;
  el.readOnly = false;
  if (replaceWith !== undefined) {
    el.value = replaceWith;
    cells[r][c] = replaceWith;
  }
  el.focus();
  const end = el.value.length;
  el.setSelectionRange(end, end);
  paint();
}

// ---- mouse ---------------------------------------------------------------
function onMouseDown(r: number, c: number, e: MouseEvent): void {
  if (editing && r === active.r && c === active.c) return; // clicking inside the cell being edited
  dragging = true;
  if (e.shiftKey) {
    active = { r, c }; // extend from existing anchor
  } else {
    anchor = { r, c };
    active = { r, c };
  }
  editing = false;
  const el = inputAt(r, c);
  if (el) el.readOnly = true;
  paint();
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
    anchor = { r: 0, c };
    active = { r: cells.length - 1, c };
    paint();
    return;
  }
  if (e.code === "Space" && e.shiftKey) {
    e.preventDefault();
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
