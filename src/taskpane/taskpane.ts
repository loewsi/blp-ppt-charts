import type {
  ChartData,
  ChartModel,
  ChartBox,
  Series,
  ChartOptions,
  Orientation,
  Grouping,
} from "../model/chartModel";
import {
  DEFAULT_BOX,
  PALETTE,
  PALETTES,
  CURRENT_SCHEMA_VERSION,
  defaultData,
  defaultOptions,
} from "../model/chartModel";
import { detectCapabilities } from "../office/capabilities";
import { loadMasterAccents } from "../office/theme";
import { drawChart } from "../engine/render";
import {
  deleteChart,
  getSlideCharts,
  getSelectedChartId,
  getChartBox,
  getPartBox,
  translatePart,
} from "../engine/persistence";
import { newId } from "../util/id";
import { mountGrid, setGridData, getGridData, setSeriesColor } from "./grid";

// ---- state ---------------------------------------------------------------
let currentData: ChartData = defaultData();
let currentBox: ChartBox = { ...DEFAULT_BOX };
let currentId: string | null = null; // null => "insert new" mode
let currentName = "";
let cachedModels: ChartModel[] = [];
let busy = false; // guards against the selection handler firing during our own edits
let applyTimer: ReturnType<typeof setTimeout> | undefined;

/** Debounced live apply: any edit re-renders the current chart automatically. */
function scheduleApply(): void {
  if (!currentId) return;
  if (applyTimer) clearTimeout(applyTimer);
  applyTimer = setTimeout(() => void guard(updateChart), 400);
}

// ---- boot ----------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    show("unsupported");
    return;
  }
  show("app");
  wire();
  mountGrid(byId("grid"), () => {
    currentData = getGridData();
    renderSeriesColors();
    scheduleApply();
  });
  setOptionsUI(defaultOptions());
  renderGrid();
  setMode();

  const caps = detectCapabilities();
  if (!caps.supportsGrouping) {
    status("Heads up: this PowerPoint build can't group shapes — charts insert as separate shapes.");
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );

  // Opt-in: no shape-resize event exists in Office.js, so we poll the chart's
  // size and re-lay-out once it settles (no live preview, so nothing jumps).
  setInterval(() => void pollResize(), 1000);

  // Ribbon insert commands run in this same (shared) runtime.
  registerRibbonActions();
  // Keep the runtime alive so selection show/hide works from document open.
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (Office as any).addin?.setStartupBehavior?.((Office as any).StartupBehavior.load);
  } catch {
    // Not a shared-runtime host; falls back to a manually opened pane.
  }
});

/* eslint-disable @typescript-eslint/no-explicit-any */
function registerRibbonActions(): void {
  const actions = (Office as any).actions;
  if (!actions?.associate) return;
  const mk = (over: Partial<ChartOptions>) => (event: { completed?: () => void }) => {
    void (async () => {
      try {
        await doInsert(over);
        showPane();
      } finally {
        event.completed?.();
      }
    })();
  };
  actions.associate("insertStackedColumn", mk({ orientation: "column", grouping: "stacked" }));
  actions.associate("insertClusteredColumn", mk({ orientation: "column", grouping: "clustered" }));
  actions.associate("insertStacked100", mk({ orientation: "column", grouping: "stacked100" }));
  actions.associate("insertBar", mk({ orientation: "bar", grouping: "stacked" }));
}

function showPane(): void {
  try {
    (Office as any).addin?.showAsTaskpane?.();
  } catch {
    /* not shared runtime */
  }
}

function hidePane(): void {
  try {
    (Office as any).addin?.hide?.();
  } catch {
    /* not shared runtime */
  }
}
/* eslint-enable @typescript-eslint/no-explicit-any */

let lastPolled: ChartBox | null = null;
let lastRenderedBox: ChartBox | null = null; // the chart's box right after our last draw
let lastResizeApplyMs = 0; // safety cap so resize can't rapid-fire

async function pollResize(): Promise<void> {
  if (busy || !currentId) return;
  let box: ChartBox | null = null;
  try {
    box = await withSlide((ctx, slide) => getChartBox(ctx, slide, currentId!));
  } catch {
    return;
  }
  if (!box) return;
  if (!lastRenderedBox) {
    lastRenderedBox = box;
    return;
  }
  // Only a SIZE change is a resize. A move (same size, different position) must
  // NOT re-render — just shift the model box by the same delta so later edits
  // redraw in place.
  if (sameSize(box, lastRenderedBox)) {
    currentBox = {
      left: currentBox.left + (box.left - lastRenderedBox.left),
      top: currentBox.top + (box.top - lastRenderedBox.top),
      width: currentBox.width,
      height: currentBox.height,
    };
    lastRenderedBox = box;
    return;
  }
  if (!lastPolled || !sameSize(box, lastPolled)) {
    lastPolled = box; // still dragging — wait for the size to settle
    return;
  }
  if (performance.now() - lastResizeApplyMs < 2500) return; // safety: at most once / 2.5s
  lastResizeApplyMs = performance.now();
  lastPolled = box;
  currentBox = box; // adopt the size the user dragged to
  await updateChart(); // redraws at currentBox and refreshes lastRenderedBox
}

function sameSize(a: ChartBox, b: ChartBox): boolean {
  // 3pt tolerance so sub-point jitter never reads as a resize.
  return Math.abs(a.width - b.width) < 3 && Math.abs(a.height - b.height) < 3;
}

async function onSelectionChanged(): Promise<void> {
  if (busy) return;
  try {
    await withSlide(async (context, slide) => {
      const id = await getSelectedChartId(context);
      if (!id) {
        hidePane(); // clicked away from any chart
        return;
      }
      showPane(); // a chart is selected → open the editor
      if (id === currentId) return;
      const models = await getSlideCharts(context, slide);
      const model = models.find((m) => m.id === id);
      if (model) {
        applyModel(model);
        status(`Editing ${model.name || "chart"}.`);
      }
    });
  } catch {
    // Ignore transient selection-read errors.
  }
}

function show(which: "app" | "unsupported"): void {
  byId(which).hidden = false;
}

const OPTION_IDS = [
  "optOrientation", "optGrouping", "optGap", "optTotals", "optLabels", "optReverse",
  "optLegend", "legendPosition", "optGridlines", "optAxis", "labelOverflow",
  "fontFamily", "segFontSize", "totFontSize",
  "nfDecimals", "nfScale", "nfPrefix", "nfSuffix", "nfHideZero",
];

function wire(): void {
  byId("insertBtn").addEventListener("click", () => guard(insertChart));
  byId("editSelectedBtn").addEventListener("click", () => guard(editSelected));
  byId("refreshBtn").addEventListener("click", () => guard(refreshList));
  byId("loadBtn").addEventListener("click", () => guard(loadSelected));
  byId("addSeriesBtn").addEventListener("click", () => addSeries());
  byId("removeSeriesBtn").addEventListener("click", () => removeSeries());
  byId("addCatBtn").addEventListener("click", () => addCategory());
  byId("removeCatBtn").addEventListener("click", () => removeCategory());
  byId("transposeBtn").addEventListener("click", () => transpose());
  byId("applyColorsBtn").addEventListener("click", () => guard(applyColors));
  byId("pasteLoadBtn").addEventListener("click", () => loadPasted());
  // Live apply: changing any style/format control re-renders the current chart.
  OPTION_IDS.forEach((id) => byId(id).addEventListener("change", () => scheduleApply()));
}

// ---- grid bridge ---------------------------------------------------------
/** Push the current model into the grid (and refresh the series color swatches). */
function renderGrid(): void {
  setGridData(currentData);
  renderSeriesColors();
}

/** Real-color swatches, one per series; editing one recolors that whole series. */
function renderSeriesColors(): void {
  const box = byId("seriesColors");
  box.innerHTML = "";
  currentData.series.forEach((s, i) => {
    const row = document.createElement("label");
    row.className = "swatch";
    const sw = document.createElement("input");
    sw.type = "color";
    sw.value = s.color;
    sw.addEventListener("change", () => {
      setSeriesColor(i, sw.value);
      currentData = readGrid();
      renderSeriesColors();
      scheduleApply();
      status("Color changed.");
    });
    const name = document.createElement("span");
    name.textContent = s.name;
    row.appendChild(sw);
    row.appendChild(name);
    box.appendChild(row);
  });
}

/** Read the grid back into a ChartData. */
function readGrid(): ChartData {
  return getGridData();
}

// ---- structural edits ----------------------------------------------------
function addSeries(): void {
  currentData = readGrid();
  const i = currentData.series.length;
  currentData.series.push({
    name: `Series ${i + 1}`,
    color: PALETTE[i % PALETTE.length],
    values: currentData.categories.map(() => 0),
  });
  renderGrid();
  scheduleApply();
}

function removeSeries(): void {
  currentData = readGrid();
  if (currentData.series.length <= 1) return;
  currentData.series.pop();
  renderGrid();
  scheduleApply();
}

function addCategory(): void {
  currentData = readGrid();
  currentData.categories.push(`Cat ${currentData.categories.length + 1}`);
  currentData.series.forEach((s) => s.values.push(0));
  renderGrid();
  scheduleApply();
}

function removeCategory(): void {
  currentData = readGrid();
  if (currentData.categories.length <= 1) return;
  currentData.categories.pop();
  currentData.series.forEach((s) => s.values.pop());
  renderGrid();
  scheduleApply();
}

function transpose(): void {
  const d = readGrid();
  const newCategories = d.series.map((s) => s.name);
  const newSeries: Series[] = d.categories.map((cat, ci) => ({
    name: cat,
    color: PALETTE[ci % PALETTE.length],
    values: d.series.map((s) => s.values[ci] ?? 0),
  }));
  currentData = { type: "barColumn", categories: newCategories, series: newSeries };
  renderGrid();
  scheduleApply();
}

async function applyColors(): Promise<void> {
  const scheme = (byId("colorScheme") as HTMLSelectElement).value;
  let palette: string[];
  if (scheme === "master") {
    palette = await withSlide((ctx, slide) => loadMasterAccents(ctx, slide));
    if (palette.length === 0) {
      status("Couldn't read master colors on this host — using BLP instead.", true);
      palette = PALETTES.blp;
    }
  } else {
    palette = PALETTES[scheme] ?? PALETTES.blp;
  }
  currentData = readGrid();
  currentData.series.forEach((s, i) => {
    s.color = palette[i % palette.length];
  });
  renderGrid();
  scheduleApply();
  status(`Applied ${scheme} colors.`);
}

// ---- paste fallback (RevoGrid also pastes natively into cells) -----------
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
  scheduleApply();
  status(
    `Loaded ${parsed.categories.length} categories × ${parsed.series.length} series.`
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
  return { type: "barColumn", categories, series };
}

// ---- PowerPoint operations ----------------------------------------------
async function insertChart(): Promise<void> {
  await doInsert({});
}

/** Insert a fresh chart of a given type; edits then apply live. Duplicate an
 *  existing chart by copy-pasting it on the slide. */
async function doInsert(over: Partial<ChartOptions>): Promise<void> {
  currentData = defaultData();
  currentName = "";
  currentId = null;
  setOptionsUI({ ...defaultOptions(), ...over });
  renderGrid();
  const data = readGrid();
  const box: ChartBox = { ...DEFAULT_BOX };
  busy = true;
  try {
    await withSlide(async (context, slide) => {
      const existing = await getSlideCharts(context, slide);
      const name = `Chart ${existing.length + 1}`;
      const model: ChartModel = {
        id: newId(),
        schemaVersion: CURRENT_SCHEMA_VERSION,
        name,
        data,
        box,
        options: readOptions(),
      };
      await drawChart(context, slide, model);
      currentId = model.id;
      currentName = name;
      lastRenderedBox = await getChartBox(context, slide, model.id); // baseline for resize detection
    });
  } finally {
    busy = false;
  }
  currentData = data;
  currentBox = box;
  setMode();
  status(`Inserted ${currentName}. Edit anything — it updates live.`);
}

async function updateChart(): Promise<void> {
  if (!currentId) return;
  const data = readGrid();
  if (!validate(data)) return;
  busy = true;
  try {
    await withSlide(async (context, slide) => {
      const opts = readOptions();
      const box = currentBox; // redraw at the intended box; the resize poll updates currentBox
      // Remember where the legend sits, so a redraw doesn't snap it back to default.
      const savedLegend = opts.showLegend ? await getPartBox(context, slide, currentId!, "legend") : null;
      await deleteChart(context, slide, currentId!);
      const model: ChartModel = {
        id: currentId!,
        schemaVersion: CURRENT_SCHEMA_VERSION,
        name: currentName || "Chart",
        data,
        box,
        options: opts,
      };
      await drawChart(context, slide, model);
      if (savedLegend) {
        const nowLegend = await getPartBox(context, slide, currentId!, "legend");
        if (nowLegend) {
          await translatePart(context, slide, currentId!, "legend", savedLegend.left - nowLegend.left, savedLegend.top - nowLegend.top);
        }
      }
      lastRenderedBox = await getChartBox(context, slide, currentId!); // baseline for resize detection
    });
  } finally {
    busy = false;
  }
  currentData = data;
}

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

function applyModel(model: ChartModel): void {
  currentData = model.data;
  currentBox = model.box;
  currentId = model.id;
  currentName = model.name || "chart";
  setOptionsUI(model.options);
  renderGrid();
  setMode();
}

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
  const badge = byId("editing");
  if (currentId) {
    badge.textContent = `Editing: ${currentName || "chart"} — changes apply live`;
    badge.className = "editing on";
  } else {
    badge.textContent = "No chart selected — click a chart, or Insert one";
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

// ---- chart-style options <-> UI -----------------------------------------
function readOptions(): ChartOptions {
  const v = (id: string) => (byId(id) as HTMLInputElement | HTMLSelectElement).value;
  const c = (id: string) => (byId(id) as HTMLInputElement).checked;
  const gapPct = Number(v("optGap"));
  return {
    orientation: v("optOrientation") as Orientation,
    grouping: v("optGrouping") as Grouping,
    gap: Math.min(0.9, Math.max(0, (isFinite(gapPct) ? gapPct : 35) / 100)),
    showTotals: c("optTotals"),
    showValueLabels: c("optLabels"),
    reverseCategories: c("optReverse"),
    showLegend: c("optLegend"),
    legendPosition: v("legendPosition") as "top" | "bottom" | "left" | "right",
    showGridlines: c("optGridlines"),
    showValueAxis: c("optAxis"),
    labelOverflow: v("labelOverflow") as "inside" | "outside",
    fontFamily: v("fontFamily"),
    segmentFontSize: clampInt(Number(v("segFontSize")), 6, 24),
    totalFontSize: clampInt(Number(v("totFontSize")), 6, 24),
    numberFormat: {
      decimals: clampInt(Number(v("nfDecimals")), 0, 3),
      scale: v("nfScale") as "none" | "k" | "M",
      prefix: v("nfPrefix"),
      suffix: v("nfSuffix"),
      hideZero: c("nfHideZero"),
    },
  };
}

function setOptionsUI(o: ChartOptions): void {
  (byId("optOrientation") as HTMLSelectElement).value = o.orientation;
  (byId("optGrouping") as HTMLSelectElement).value = o.grouping;
  (byId("optGap") as HTMLInputElement).value = String(Math.round(o.gap * 100));
  (byId("optTotals") as HTMLInputElement).checked = o.showTotals;
  (byId("optLabels") as HTMLInputElement).checked = o.showValueLabels;
  (byId("optReverse") as HTMLInputElement).checked = o.reverseCategories;
  (byId("optLegend") as HTMLInputElement).checked = o.showLegend;
  (byId("legendPosition") as HTMLSelectElement).value = o.legendPosition;
  (byId("optGridlines") as HTMLInputElement).checked = o.showGridlines;
  (byId("optAxis") as HTMLInputElement).checked = o.showValueAxis;
  (byId("labelOverflow") as HTMLSelectElement).value = o.labelOverflow;
  (byId("fontFamily") as HTMLSelectElement).value = o.fontFamily;
  (byId("segFontSize") as HTMLInputElement).value = String(o.segmentFontSize);
  (byId("totFontSize") as HTMLInputElement).value = String(o.totalFontSize);
  (byId("nfDecimals") as HTMLInputElement).value = String(o.numberFormat.decimals);
  (byId("nfScale") as HTMLSelectElement).value = o.numberFormat.scale;
  (byId("nfPrefix") as HTMLInputElement).value = o.numberFormat.prefix;
  (byId("nfSuffix") as HTMLInputElement).value = o.numberFormat.suffix;
  (byId("nfHideZero") as HTMLInputElement).checked = o.numberFormat.hideZero;
}

function clampInt(n: number, lo: number, hi: number): number {
  if (!isFinite(n)) return lo;
  return Math.max(lo, Math.min(hi, Math.round(n)));
}
