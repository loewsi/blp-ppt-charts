// Lightweight Excel-like data grid (no dependency, full control).
// Full-sheet layout: row 0 = category names, column 0 = series names, interior
// = values. Every cell is editable. Supports keyboard nav, Ctrl+Space /
// Shift+Space selection, Ctrl+"+"/Ctrl+"-" insert/remove, and Excel copy/paste.
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
let sel: { type: "cell" | "row" | "col"; r: number; c: number } = { type: "cell", r: 1, c: 1 };

export function mountGrid(container: HTMLElement, changed?: () => void): void {
  host = container;
  onChange = changed;
}

export function setGridData(data: ChartData): void {
  cells = cellsFromData(data);
  seriesColors = data.series.map((s) => s.color);
  render();
}

export function getGridData(): ChartData {
  syncFromDom();
  return dataFromCells(cells, seriesColors);
}

function syncFromDom(): void {
  if (!host) return;
  host.querySelectorAll<HTMLInputElement>("input.xcell").forEach((inp) => {
    const r = Number(inp.dataset.r);
    const c = Number(inp.dataset.c);
    if (cells[r]) cells[r][c] = inp.value;
  });
}

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
      inp.dataset.r = String(r);
      inp.dataset.c = String(c);
      if (r === 0 && c === 0) inp.disabled = true;
      inp.addEventListener("input", () => (cells[r][c] = inp.value));
      inp.addEventListener("focus", () => {
        active = { r, c };
        sel = { type: "cell", r, c };
        paint();
      });
      inp.addEventListener("keydown", onKey);
      inp.addEventListener("paste", onPaste);
      inp.addEventListener("copy", onCopy);
      td.appendChild(inp);
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  host.appendChild(table);
  paint();
}

function cellAt(r: number, c: number): HTMLInputElement | null {
  return host.querySelector<HTMLInputElement>(`input.xcell[data-r="${r}"][data-c="${c}"]`);
}

function focusCell(r: number, c: number): void {
  const el = cellAt(r, c);
  if (el) {
    el.focus();
    el.select();
  }
}

function paint(): void {
  host.querySelectorAll<HTMLInputElement>("input.xcell").forEach((inp) => {
    const r = Number(inp.dataset.r);
    const c = Number(inp.dataset.c);
    const selected =
      (sel.type === "col" && c === sel.c) ||
      (sel.type === "row" && r === sel.r) ||
      (sel.type === "cell" && r === active.r && c === active.c);
    inp.classList.toggle("sel", selected);
  });
}

function onKey(e: KeyboardEvent): void {
  const r = active.r;
  const c = active.c;
  const R = cells.length;
  const C = cells[0].length;

  // Selection: Ctrl+Space = column, Shift+Space = row (Excel).
  if (e.code === "Space" && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    sel = { type: "col", r, c };
    paint();
    return;
  }
  if (e.code === "Space" && e.shiftKey) {
    e.preventDefault();
    sel = { type: "row", r, c };
    paint();
    return;
  }

  // Insert / remove: Ctrl+"+" and Ctrl+"-" act on the current selection.
  if ((e.ctrlKey || e.metaKey) && (e.key === "+" || e.key === "=")) {
    e.preventDefault();
    if (sel.type === "col") cells = insertColumn(cells, Math.max(1, c));
    else cells = insertRow(cells, Math.max(1, r));
    seriesColors = syncColors();
    render();
    emit();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && (e.key === "-" || e.key === "_")) {
    e.preventDefault();
    if (sel.type === "col") cells = removeColumn(cells, c > 0 ? c : 1);
    else cells = removeRow(cells, r > 0 ? r : 1);
    seriesColors = syncColors();
    render();
    emit();
    return;
  }

  // Navigation.
  let nr = r;
  let nc = c;
  if (e.key === "Enter" || e.key === "ArrowDown") nr = Math.min(R - 1, r + 1);
  else if (e.key === "ArrowUp") nr = Math.max(0, r - 1);
  else if (e.key === "Tab" && !e.shiftKey) nc = c + 1 >= C ? (nr === R - 1 ? c : ((nr = r + 1), 1)) : c + 1;
  else if (e.key === "Tab" && e.shiftKey) nc = Math.max(0, c - 1);
  else return; // let other keys type normally (ArrowLeft/Right move the caret)

  e.preventDefault();
  if (nr === 0 && nc === 0) nc = 1; // skip the disabled corner
  active = { r: nr, c: nc };
  sel = { type: "cell", r: nr, c: nc };
  focusCell(nr, nc);
}

function onPaste(e: ClipboardEvent): void {
  const text = e.clipboardData?.getData("text") ?? "";
  if (!text.includes("\t") && !text.includes("\n")) return; // single value: let default paste
  e.preventDefault();
  syncFromDom();
  cells = pasteInto(cells, active.r, active.c, text);
  seriesColors = syncColors();
  render();
  focusCell(active.r, active.c);
  emit();
}

function onCopy(e: ClipboardEvent): void {
  syncFromDom();
  let text = "";
  if (sel.type === "col") text = cells.map((row) => row[sel.c] ?? "").join("\n");
  else if (sel.type === "row") text = (cells[sel.r] ?? []).join("\t");
  else return; // single cell: default copy
  e.preventDefault();
  e.clipboardData?.setData("text/plain", text);
}

/** Keep the seriesColors array length in sync with the number of series rows. */
function syncColors(): string[] {
  const nSeries = Math.max(0, cells.length - 1);
  const out: string[] = [];
  for (let i = 0; i < nSeries; i++) out.push(seriesColors[i] || PALETTE[i % PALETTE.length]);
  return out;
}

export function setSeriesColor(index: number, color: string): void {
  seriesColors[index] = color;
}

function emit(): void {
  if (onChange) onChange();
}
