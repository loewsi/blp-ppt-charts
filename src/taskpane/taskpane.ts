import type { ChartData, ChartModel, ChartBox, Series } from "../model/chartModel";
import { DEFAULT_BOX, PALETTE, defaultData } from "../model/chartModel";
import { drawChart } from "../engine/render";
import { deleteChart, getSlideCharts } from "../engine/persistence";
import { newId } from "../util/id";

// ---- state ---------------------------------------------------------------
let currentData: ChartData = defaultData();
let currentBox: ChartBox = { ...DEFAULT_BOX };
let currentId: string | null = null; // null => "insert new" mode
let cachedModels: ChartModel[] = [];

// ---- boot ----------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    show("unsupported");
    return;
  }
  show("app");
  wire();
  renderGrid();
  setMode();
});

function show(which: "app" | "unsupported"): void {
  byId(which).hidden = false;
}

function wire(): void {
  byId("insertBtn").addEventListener("click", () => guard(insertChart));
  byId("updateBtn").addEventListener("click", () => guard(updateChart));
  byId("newBtn").addEventListener("click", () => resetToNew());
  byId("refreshBtn").addEventListener("click", () => guard(refreshList));
  byId("loadBtn").addEventListener("click", () => guard(loadSelected));
  byId("addSeriesBtn").addEventListener("click", () => addSeries());
  byId("addCatBtn").addEventListener("click", () => addCategory());
}

// ---- grid rendering ------------------------------------------------------
function renderGrid(): void {
  const table = byId("grid");
  table.innerHTML = "";

  // Header: "Series" + one column per category + remove column.
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  hr.appendChild(th("Series"));
  currentData.categories.forEach((cat, ci) => {
    const cell = document.createElement("th");
    cell.appendChild(input("cat", cat, "text"));
    hr.appendChild(cell);
    cell.querySelector("input")!.dataset.ci = String(ci);
  });
  hr.appendChild(th(""));
  thead.appendChild(hr);
  table.appendChild(thead);

  // Body: one row per series.
  const tbody = document.createElement("tbody");
  currentData.series.forEach((s, si) => {
    const tr = document.createElement("tr");

    const head = document.createElement("td");
    const wrap = document.createElement("div");
    wrap.className = "series-head";
    wrap.appendChild(input("sname", s.name, "text"));
    wrap.appendChild(input("scolor", s.color, "color"));
    head.appendChild(wrap);
    tr.appendChild(head);

    currentData.categories.forEach((_, ci) => {
      const td = document.createElement("td");
      const inp = input("val", String(s.values[ci] ?? 0), "number");
      inp.dataset.ci = String(ci);
      td.appendChild(inp);
      tr.appendChild(td);
    });

    const rm = document.createElement("td");
    rm.appendChild(removeBtn(() => removeSeries(si)));
    tr.appendChild(rm);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

/** Read the current DOM state back into a ChartData. */
function readGrid(): ChartData {
  const table = byId("grid");
  const categories = [...table.querySelectorAll<HTMLInputElement>(".cat")].map(
    (i, idx) => i.value.trim() || `Cat ${idx + 1}`
  );

  const series: Series[] = [...table.querySelectorAll<HTMLTableRowElement>("tbody tr")].map(
    (tr, idx) => {
      const name = tr.querySelector<HTMLInputElement>(".sname")!.value.trim() || `Series ${idx + 1}`;
      const color = tr.querySelector<HTMLInputElement>(".scolor")!.value || PALETTE[idx % PALETTE.length];
      const values = [...tr.querySelectorAll<HTMLInputElement>(".val")].map((v) => {
        const n = Number(v.value);
        return isFinite(n) ? n : 0;
      });
      return { name, color, values };
    }
  );

  return { type: "stackedColumn", categories, series };
}

// ---- structural edits ----------------------------------------------------
function addSeries(): void {
  currentData = readGrid();
  const idx = currentData.series.length;
  currentData.series.push({
    name: `Series ${idx + 1}`,
    color: PALETTE[idx % PALETTE.length],
    values: currentData.categories.map(() => 0),
  });
  renderGrid();
}

function addCategory(): void {
  currentData = readGrid();
  currentData.categories.push(`Cat ${currentData.categories.length + 1}`);
  currentData.series.forEach((s) => s.values.push(0));
  renderGrid();
}

function removeSeries(index: number): void {
  currentData = readGrid();
  if (currentData.series.length <= 1) return;
  currentData.series.splice(index, 1);
  renderGrid();
}

// ---- PowerPoint operations ----------------------------------------------
async function insertChart(): Promise<void> {
  const data = readGrid();
  if (!validate(data)) return;
  const box: ChartBox = { ...DEFAULT_BOX };
  await withSlide(async (context, slide) => {
    const model: ChartModel = { id: newId(), version: 1, data, box };
    await drawChart(context, slide, model);
    currentId = model.id;
  });
  currentData = data;
  currentBox = box;
  setMode();
  status("Chart inserted.");
}

async function updateChart(): Promise<void> {
  if (!currentId) return;
  const data = readGrid();
  if (!validate(data)) return;
  await withSlide(async (context, slide) => {
    await deleteChart(context, slide, currentId!);
    const model: ChartModel = { id: currentId!, version: 1, data, box: currentBox };
    await drawChart(context, slide, model);
  });
  currentData = data;
  status("Chart updated.");
}

async function refreshList(): Promise<void> {
  cachedModels = await withSlide((context, slide) => getSlideCharts(context, slide));
  const sel = byId("chartSelect") as HTMLSelectElement;
  sel.innerHTML = "";
  if (cachedModels.length === 0) {
    sel.appendChild(new Option("(no charts on this slide)", ""));
  } else {
    cachedModels.forEach((m, i) => {
      const label = `Chart ${i + 1}: ${m.data.series.map((s) => s.name).join(", ")}`;
      sel.appendChild(new Option(label, m.id));
    });
  }
  status(`Found ${cachedModels.length} chart(s) on this slide.`);
}

function loadSelected(): void {
  const sel = byId("chartSelect") as HTMLSelectElement;
  const id = sel.value;
  const model = cachedModels.find((m) => m.id === id);
  if (!model) {
    status("Select a chart first (press ↻ to rescan).", true);
    return;
  }
  currentData = model.data;
  currentBox = model.box;
  currentId = model.id;
  renderGrid();
  setMode();
  status("Chart loaded. Edit values, then Update.");
}

function resetToNew(): void {
  currentData = defaultData();
  currentBox = { ...DEFAULT_BOX };
  currentId = null;
  renderGrid();
  setMode();
  status("New chart. Edit values, then Insert.");
}

/** Run a callback against the selected slide (falls back to the first slide). */
async function withSlide<T>(
  cb: (context: PowerPoint.RequestContext, slide: PowerPoint.Slide) => Promise<T>
): Promise<T> {
  return await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide | undefined;
    try {
      const sel = context.presentation.getSelectedSlides();
      sel.load("items");
      await context.sync();
      slide = sel.items[0];
    } catch {
      // Older host without getSelectedSlides - fall back below.
    }
    if (!slide) slide = context.presentation.slides.getItemAt(0);
    return await cb(context, slide);
  });
}

// ---- helpers -------------------------------------------------------------
function validate(data: ChartData): boolean {
  if (data.categories.length === 0) {
    status("Add at least one category.", true);
    return false;
  }
  if (data.series.length === 0) {
    status("Add at least one series.", true);
    return false;
  }
  return true;
}

function setMode(): void {
  (byId("updateBtn") as HTMLButtonElement).disabled = currentId === null;
}

async function guard(fn: () => Promise<void> | void): Promise<void> {
  try {
    await fn();
  } catch (e) {
    status("Error: " + (e instanceof Error ? e.message : String(e)), true);
    // eslint-disable-next-line no-console
    console.error(e);
  }
}

function status(msg: string, isError = false): void {
  const el = byId("status");
  el.textContent = msg;
  el.classList.toggle("err", isError);
}

function byId<T extends HTMLElement = HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el as T;
}

function th(text: string): HTMLTableCellElement {
  const el = document.createElement("th");
  el.textContent = text;
  return el;
}

function input(cls: string, value: string, type: string): HTMLInputElement {
  const el = document.createElement("input");
  el.className = cls;
  el.type = type;
  el.value = value;
  return el;
}

function removeBtn(onClick: () => void): HTMLButtonElement {
  const b = document.createElement("button");
  b.className = "rm";
  b.title = "Remove series";
  b.textContent = "×";
  b.addEventListener("click", onClick);
  return b;
}
