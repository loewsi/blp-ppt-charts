// Excel-like data grid backed by RevoGrid (MIT). Full-sheet layout:
//   row 0            -> category names (cols 1..n)
//   column 0         -> series names (rows 1..m)
//   interior cells   -> numeric values
// Every meaningful label is an editable cell, so it feels like an Excel range,
// with native keyboard navigation and paste. Series colors live outside the grid
// (the chart style controls own them) and are preserved across edits here.
import { defineCustomElements } from "@revolist/revogrid/loader";
import type { ChartData, Series } from "../model/chartModel";
import { PALETTE } from "../model/chartModel";

defineCustomElements();

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = Record<string, any>;

let gridEl: HTMLElement & { columns?: unknown; source?: Row[]; getSource?: () => Row[] };
let source: Row[] = [];
let colsCount = 0; // number of category columns
let seriesColors: string[] = [];

function buildColumns(nCats: number): unknown[] {
  const cols: unknown[] = [{ prop: "0", name: "", size: 130 }];
  for (let c = 1; c <= nCats; c++) cols.push({ prop: String(c), name: "", size: 78 });
  return cols;
}

function buildSource(data: ChartData): Row[] {
  const header: Row = { "0": "" };
  data.categories.forEach((cat, i) => (header[String(i + 1)] = cat));
  const rows: Row[] = [header];
  data.series.forEach((s) => {
    const row: Row = { "0": s.name };
    data.categories.forEach((_, i) => (row[String(i + 1)] = s.values[i] ?? 0));
    rows.push(row);
  });
  return rows;
}

/** Create the grid element inside `container`. Call once. */
export function mountGrid(container: HTMLElement): void {
  gridEl = document.createElement("revo-grid") as HTMLElement & { source?: Row[] };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const g = gridEl as any;
  g.theme = "compact";
  g.resize = true;
  g.rowHeaders = false; // row 0 already holds category names; avoid a confusing gutter
  g.range = true; // range selection + fill + clipboard
  container.appendChild(gridEl);

  // Keep our local `source` in sync with edits (single-cell and pasted ranges).
  gridEl.addEventListener("afteredit", (e: Event) => {
    const detail = (e as CustomEvent).detail as
      | { rowIndex: number; prop: string; val: unknown }
      | { data: Record<number, Row> };
    if ("data" in detail && detail.data) {
      for (const [ri, changes] of Object.entries(detail.data)) {
        source[Number(ri)] = { ...source[Number(ri)], ...changes };
      }
    } else if ("rowIndex" in detail) {
      source[detail.rowIndex] = { ...source[detail.rowIndex], [detail.prop]: detail.val };
    }
  });
}

/** Load a ChartData into the grid. */
export function setGridData(data: ChartData): void {
  colsCount = data.categories.length;
  seriesColors = data.series.map((s) => s.color);
  source = buildSource(data);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const g = gridEl as any;
  g.columns = buildColumns(colsCount);
  g.source = source;
}

/** Read the current grid contents back into a ChartData. */
export function getGridData(): ChartData {
  // Prefer the grid's own source if it exposes it (covers edits we didn't catch).
  const live = (gridEl.getSource && gridEl.getSource()) || gridEl.source || source;
  const header = live[0] || {};
  const categories: string[] = [];
  for (let c = 1; c <= colsCount; c++) {
    const v = header[String(c)];
    categories.push(v != null && String(v).trim() ? String(v).trim() : `Cat ${c}`);
  }
  const series: Series[] = live.slice(1).map((row, ri) => ({
    name: row["0"] != null && String(row["0"]).trim() ? String(row["0"]).trim() : `Series ${ri + 1}`,
    color: seriesColors[ri] || PALETTE[ri % PALETTE.length],
    values: categories.map((_, i) => {
      const n = Number(String(row[String(i + 1)] ?? "").replace(/[,\s%]/g, ""));
      return isFinite(n) ? n : 0;
    }),
  }));
  return { type: "barColumn", categories, series };
}
