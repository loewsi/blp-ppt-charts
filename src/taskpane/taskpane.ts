import type { ChartData, ChartModel, ChartBox, Series } from "../model/chartModel";
import { DEFAULT_BOX, PALETTE, defaultData } from "../model/chartModel";
import { drawChart } from "../engine/render";
import {
  deleteChart,
  getSlideCharts,
  getSelectedChartId,
  getChartBox,
} from "../engine/persistence";
import { newId } from "../util/id";

// ---- state ---------------------------------------------------------------
let currentData: ChartData = defaultData();
let currentBox: ChartBox = { ...DEFAULT_BOX };
let currentId: string | null = null; // null => "insert new" mode
let currentName = "";
let cachedModels: ChartModel[] = [];
let busy = false; // guards against the selection handler firing during our own edits

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

  // Auto-load whichever chart the user selects on the slide.
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );
});

/** Fires when the slide selection changes — pull the selected chart into the editor. */
async function onSelectionChanged(): Promise<void> {
  if (busy) return;
  try {
    await withSlide(async (context, slide) => {
      const id = await getSelectedChartId(context);
      if (!id || id === currentId) return;
      const models = await getSlideCharts(context, slide);
      const model = models.find((m) => m.id === id);
      if (model) {
        applyModel(model);
        status(`Editing ${model.name || "chart"} (selected on slide).`);
      }
    });
  } catch {
    // Ignore transient selection-read errors.
  }
}

function show(which: "app" | "unsupported"): void {
  byId(which).hidden = false;
}

function wire(): void {
  byId("insertBtn").addEventListener("click", () => guard(insertChart));
  byId("updateBtn").addEventListener("click", () => guard(updateChart));
  byId("newBtn").addEventListener("click", () => resetToNew());
  byId("editSelectedBtn").addEventListener("click", () => guard(editSelected));
  byId("refreshBtn").addEventListener("click", () => guard(refreshList));
  byId("loadBtn").addEventListener("click", () => guard(loadSelected));
  byId("addSeriesBtn").addEventListener("click", () => addSeries());
  byId("addCatBtn").addEventListener("click", () => addCategory());
  byId("pasteLoadBtn").addEventListener("click", () => loadPasted());
}

// ---- grid rendering ------------------------------------------------------
function renderGrid(): void {
  const table = byId("grid");
  table.innerHTML = "";

  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  hr.appendChild(th("Series"));
  currentData.categories.forEach((cat, ci) => {
    const cell = document.createElement("th");
    cell.appendChild(input("cat", cat, "text"));
    cell.appendChild(removeBtn(() => removeCategory(ci), "Remove category"));
    hr.appendChild(cell);
  });
  hr.appendChild(th(""));
  thead.appendChild(hr);
  table.appendChild(thead);

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
    rm.appendChild(removeBtn(() => removeSeries(si), "Remove series"));
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

function removeCategory(index: number): void {
  currentData = readGrid();
  if (currentData.categories.length <= 1) return;
  currentData.categories.splice(index, 1);
  currentData.series.forEach((s) => s.values.splice(index, 1));
  renderGrid();
}

/** Load a range pasted from Excel: top row = categories, first column = series names. */
function loadPasted(): void {
  const ta = byId("pasteArea") as HTMLTextAreaElement;
  const parsed = parsePasted(ta.value);
  if (!parsed) {
    status("Couldn't read that. Paste a grid with a header row and a header column.", true);
    return;
  }
  currentData = parsed;
  renderGrid();
  ta.value = "";
  const details = byId("pasteArea").closest("details");
  if (details) (details as HTMLDetailsElement).open = false;
  status(
    `Loaded ${parsed.categories.length} categories × ${parsed.series.length} series. Review, then Insert or Update.`
  );
}

function parsePasted(text: string): ChartData | null {
  const rows = text
    .replace(/\r/g, "")
    .split("\n")
    .filter((r) => r.trim().length > 0)
    .map((r) => r.split("\t"));
  if (rows.length < 2 || rows[0].length < 2) return null;

  const categories = rows[0].slice(1).map((c, i) => c.trim() || `Cat ${i + 1}`);
  const series: Series[] = rows.slice(1).map((r, ri) => ({
    name: (r[0] || "").trim() || `Series ${ri + 1}`,
    color: PALETTE[ri % PALETTE.length],
    values: categories.map((_, ci) => {
      const n = Number((r[ci + 1] || "").replace(/[,\s%]/g, ""));
      return isFinite(n) ? n : 0;
    }),
  }));
  return { type: "stackedColumn", categories, series };
}

// ---- PowerPoint operations ----------------------------------------------
async function insertChart(): Promise<void> {
  const data = readGrid();
  if (!validate(data)) return;
  const box: ChartBox = { ...DEFAULT_BOX };
  busy = true;
  try {
    await withSlide(async (context, slide) => {
      const existing = await getSlideCharts(context, slide);
      const name = `Chart ${existing.length + 1}`;
      const model: ChartModel = { id: newId(), version: 1, name, data, box };
      await drawChart(context, slide, model);
      currentId = model.id;
      currentName = name;
    });
  } finally {
    busy = false;
  }
  currentData = data;
  currentBox = box;
  setMode();
  status(`Inserted ${currentName}.`);
}

async function updateChart(): Promise<void> {
  if (!currentId) return;
  const data = readGrid();
  if (!validate(data)) return;
  busy = true;
  try {
    await withSlide(async (context, slide) => {
      // Redraw where the chart currently sits (respects moves/resizes).
      const liveBox = await getChartBox(context, slide, currentId!);
      const box = liveBox ?? currentBox;
      await deleteChart(context, slide, currentId!);
      const model: ChartModel = {
        id: currentId!,
        version: 1,
        name: currentName || "Chart",
        data,
        box,
      };
      await drawChart(context, slide, model);
      currentBox = box;
    });
  } finally {
    busy = false;
  }
  currentData = data;
  status(`Updated ${currentName || "chart"}.`);
}

/** Load whatever chart is currently selected on the slide. */
async function editSelected(): Promise<void> {
  await withSlide(async (context, slide) => {
    const id = await getSelectedChartId(context);
    if (!id) {
      status("Click a chart on the slide first, then press this.", true);
      return;
    }
    const models = await getSlideCharts(context, slide);
    const model = models.find((m) => m.id === id);
    if (!model) {
      status("That selection isn't a SlideChart.", true);
      return;
    }
    applyModel(model);
    status(`Editing ${model.name || "chart"} (selected on slide).`);
  });
}

async function refreshList(): Promise<void> {
  cachedModels = await withSlide((context, slide) => getSlideCharts(context, slide));
  const sel = byId("chartSelect") as HTMLSelectElement;
  sel.innerHTML = "";
  if (cachedModels.length === 0) {
    sel.appendChild(new Option("(no charts on this slide)", ""));
  } else {
    cachedModels.forEach((m) => {
      const label = `${m.name || "Chart"} — ${m.data.series.map((s) => s.name).join(", ")}`;
      sel.appendChild(new Option(label, m.id));
    });
  }
  status(`Found ${cachedModels.length} chart(s) on this slide.`);
}

function loadSelected(): void {
  const sel = byId("chartSelect") as HTMLSelectElement;
  const model = cachedModels.find((m) => m.id === sel.value);
  if (!model) {
    status("Pick a chart from the list first (press ↻ to rescan).", true);
    return;
  }
  applyModel(model);
  status(`Editing ${model.name || "chart"}.`);
}

function resetToNew(): void {
  currentData = defaultData();
  currentBox = { ...DEFAULT_BOX };
  currentId = null;
  currentName = "";
  renderGrid();
  setMode();
  status("New chart. Edit values, then Insert.");
}

/** Point the editor at an existing chart model. */
function applyModel(model: ChartModel): void {
  currentData = model.data;
  currentBox = model.box;
  currentId = model.id;
  currentName = model.name || "chart";
  renderGrid();
  setMode();
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
  const badge = byId("editing");
  if (currentId) {
    badge.textContent = `Editing: ${currentName || "chart"}`;
    badge.className = "editing on";
  } else {
    badge.textContent = "New chart — not inserted yet";
    badge.className = "editing";
  }
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

function removeBtn(onClick: () => void, title = "Remove"): HTMLButtonElement {
  const b = document.createElement("button");
  b.className = "rm";
  b.title = title;
  b.textContent = "×";
  b.addEventListener("click", onClick);
  return b;
}
