import type {
  ChartData,
  ChartModel,
  ChartBox,
  Series,
  ChartOptions,
  ChartType,
  Orientation,
  Grouping,
} from "../model/chartModel";
import {
  DEFAULT_BOX,
  PALETTE,
  PALETTES,
  CHART_TYPES,
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
  repairDuplicateChartIds,
} from "../engine/persistence";
import { newId } from "../util/id";
import { createOrUpdateAgenda } from "../office/agenda";
import { shadesFrom } from "../util/color";
import { mountGrid, setGridData, getGridData, setSeriesColor, setSeriesKind, getActive, getSelectionRange } from "./grid";

// ---- state ---------------------------------------------------------------
let currentData: ChartData = defaultData();
let currentBox: ChartBox = { ...DEFAULT_BOX };
let currentId: string | null = null; // null => "insert new" mode
let currentName = "";
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
  refreshVisibility();

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
let lastLegendPosition: string | null = null; // to detect a deliberate position change

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
      // A same-slide copy-paste clones the chart id; split the copy off first so
      // it becomes an independent chart before we try to load the selection.
      await repairDuplicateChartIds(context, slide);
      const id = await getSelectedChartId(context);
      if (!id) {
        // Clicked away from any chart → drop back to "no chart selected".
        if (currentId !== null) {
          currentId = null;
          currentName = "";
          lastRenderedBox = null;
          setMode();
          status("No chart selected — click a chart, or Insert one.");
        }
        hidePane();
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
  "chartType", "optOrientation", "optGrouping", "optGap", "optRefValue", "optAxisMin", "optAxisMax",
  "optTotals", "optLabels", "optReverse",
  "optLegend", "legendPosition", "optGridlines", "optAxis", "optAxisLine", "optConnectors",
  "optLineSecondaryAxis", "optLineAxisMin", "optLineAxisMax", "optReverseSeries", "optRefColor",
  "optPieHole", "optScatterQuadrant", "optScatterAxes",
  "optDiffArrow", "optDiffPercent", "optDiffFrom", "optDiffTo", "optDiffSeries", "optDiffPos",
  "optCagrArrow", "optCagrFrom", "optCagrTo", "optCagrSeries", "optCagrPeriods",
  "labelOverflow", "labelMode",
  "fontFamily", "segFontSize", "totFontSize",
  "nfDecimals", "nfScale", "nfPrefix", "nfSuffix", "nfHideZero",
  "nfThousands", "nfSep", "nfParens", "nfPlus",
];

function wire(): void {
  byId("insertBtn").addEventListener("click", () => guard(insertChart));
  byId("addSeriesBtn").addEventListener("click", () => addSeries());
  byId("removeSeriesBtn").addEventListener("click", () => removeSeries());
  byId("addCatBtn").addEventListener("click", () => addCategory());
  byId("removeCatBtn").addEventListener("click", () => removeCategory());
  byId("transposeBtn").addEventListener("click", () => transpose());
  byId("applyColorsBtn").addEventListener("click", () => guard(applyColors));
  byId("agendaBtn").addEventListener("click", () => guard(doAgenda));
  byId("shadeBtn").addEventListener("click", () => applyShades());
  // Live apply: changing any style/format control re-renders the current chart.
  OPTION_IDS.forEach((id) =>
    byId(id).addEventListener("change", () => {
      refreshVisibility();
      scheduleApply();
    })
  );
  byId("chartType").addEventListener("change", () => maybeNameScatterRows());
}

/** When switching to scatter, name the first rows X / Y / Size so the mapping is clear
 *  (only if they still have default names). */
function maybeNameScatterRows(): void {
  if (readChartType() !== "scatter") return;
  currentData = readGrid();
  const names = ["X", "Y", "Size"];
  let changed = false;
  currentData.series.forEach((s, i) => {
    if (i < 3 && /^(Product |Series |S\d|Cat )/i.test(s.name)) {
      s.name = names[i];
      changed = true;
    }
  });
  if (changed) {
    renderGrid();
    scheduleApply();
  }
}

/** Hide the option controls that don't apply to the current chart type / toggles,
 *  so the panel only shows what's relevant (Silvan: "only see the relevant options"). */
function refreshVisibility(): void {
  const type = readChartType();
  const barFamily = type === "barColumn" || type === "combination" || type === "line";
  const isPie = type === "pie";

  const showLabel = (inputId: string, show: boolean) => {
    const lbl = byId(inputId).closest("label") as HTMLElement | null;
    if (lbl) lbl.style.display = show ? "" : "none";
  };
  const showEl = (id: string, show: boolean) => {
    const e = document.getElementById(id);
    if (e) e.style.display = show ? "" : "none";
  };

  // Bar/line-family controls (irrelevant to pie).
  ["optOrientation", "optGrouping", "optGap", "optConnectors", "optReverse", "optReverseSeries",
   "optGridlines", "optAxis", "optAxisLine", "optLegend", "legendPosition", "optRefValue",
   "optAxisMin", "optAxisMax"].forEach((id) => showLabel(id, barFamily));

  // Combination only: the 2nd-axis toggle; its min/max only when the toggle is on.
  const isCombination = type === "combination";
  showLabel("optLineSecondaryAxis", isCombination);
  const secOn = (byId("optLineSecondaryAxis") as HTMLInputElement).checked;
  ["optLineAxisMin", "optLineAxisMax"].forEach((id) => showLabel(id, isCombination && secOn));
  // Per-series "line" checkboxes only for combination.
  document.querySelectorAll<HTMLElement>(".line-toggle").forEach((e) => (e.style.display = isCombination ? "" : "none"));

  // Pie-only / scatter-only.
  showLabel("optPieHole", isPie);
  showLabel("optScatterQuadrant", type === "scatter");
  showLabel("optScatterAxes", type === "scatter");
  // Scatter uses value labels + gridlines/axis; hide the rest of the bar controls there.
  if (type === "scatter") {
    ["optGridlines", "optAxis"].forEach((id) => showLabel(id, true));
  }

  // Reference-line color only once a reference value is set.
  const refSet = (byId("optRefValue") as HTMLInputElement).value.trim() !== "";
  showLabel("optRefColor", barFamily && refSet);

  // Arrows: whole section only for the column family; sub-fields only when the arrow is on.
  showEl("arrowsHeading", barFamily);
  showEl("arrowsOpts", barFamily);
  const diffOn = (byId("optDiffArrow") as HTMLSelectElement).value !== "off";
  ["optDiffPercent", "optDiffFrom", "optDiffTo", "optDiffSeries", "optDiffPos"].forEach((id) =>
    showLabel(id, diffOn)
  );
  const cagrOn = (byId("optCagrArrow") as HTMLSelectElement).value !== "off";
  ["optCagrFrom", "optCagrTo", "optCagrSeries", "optCagrPeriods"].forEach((id) => showLabel(id, cagrOn));

  // Data-entry hint per chart type.
  const hints: Partial<Record<typeof type, string>> = {
    waterfall: "Tip: type “e” in a value cell to make it a computed total / subtotal column.",
    scatter: "Rows: 1st = X, 2nd = Y, 3rd = bubble size (optional). Each column is a point.",
    pie: "The first series becomes the slices; each category is one slice.",
    mekko: "Column width = the category total; each column is 100%-stacked by series.",
  };
  const hint = byId("typeHint");
  const text = hints[type] ?? "";
  hint.textContent = text;
  hint.hidden = text === "";
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
    // Per-series line toggle (combination charts): draw this series as a line.
    const lineWrap = document.createElement("label");
    lineWrap.className = "chk mini line-toggle"; // shown only for combination (refreshVisibility)
    const lineChk = document.createElement("input");
    lineChk.type = "checkbox";
    lineChk.checked = s.kind === "line";
    lineChk.addEventListener("change", () => {
      setSeriesKind(i, lineChk.checked ? "line" : "bar");
      currentData = readGrid();
      scheduleApply();
      status(lineChk.checked ? `${s.name} shown as a line.` : `${s.name} shown as bars.`);
    });
    lineWrap.appendChild(lineChk);
    lineWrap.appendChild(document.createTextNode("line"));
    row.appendChild(sw);
    row.appendChild(name);
    row.appendChild(lineWrap);
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
  // Cell row = series index + 1; row 0 is the category header. Insert BELOW the
  // active series (splice index = active.r), or at the top when in the header row.
  const at = Math.min(currentData.series.length, getActive().r);
  currentData.series.splice(at, 0, {
    name: `Series ${i + 1}`,
    color: PALETTE[i % PALETTE.length],
    values: currentData.categories.map(() => 0),
  });
  renderGrid();
  scheduleApply();
}

function removeSeries(): void {
  currentData = readGrid();
  const { r0, r1 } = getSelectionRange();
  // Selection rows → series indices (row 0 is the header). Delete all selected, keep ≥1.
  const from = Math.max(0, r0 - 1);
  const to = Math.max(0, r1 - 1);
  const count = Math.min(to - from + 1, currentData.series.length - 1);
  if (count <= 0) return;
  currentData.series.splice(from, count);
  renderGrid();
  scheduleApply();
}

function addCategory(): void {
  currentData = readGrid();
  // Insert to the LEFT of the active category (col 0 is the series-name column).
  const at = Math.max(0, getActive().c - 1);
  currentData.categories.splice(at, 0, `Cat ${currentData.categories.length + 1}`);
  currentData.series.forEach((s) => s.values.splice(at, 0, 0));
  renderGrid();
  scheduleApply();
}

function removeCategory(): void {
  currentData = readGrid();
  const { c0, c1 } = getSelectionRange();
  const from = Math.max(0, c0 - 1);
  const to = Math.max(0, c1 - 1);
  const count = Math.min(to - from + 1, currentData.categories.length - 1);
  if (count <= 0) return;
  currentData.categories.splice(from, count);
  currentData.series.forEach((s) => s.values.splice(from, count));
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

/** Recolor every series as auto-generated shades of the picked base color. */
function applyShades(): void {
  const base = (byId("shadeBase") as HTMLInputElement).value;
  currentData = readGrid();
  const shades = shadesFrom(base, currentData.series.length);
  currentData.series.forEach((s, i) => (s.color = shades[i]));
  renderGrid();
  scheduleApply();
  status("Applied shades of the base color.");
}

async function doAgenda(): Promise<void> {
  const raw = (byId("agendaList") as HTMLTextAreaElement).value;
  const chapters = raw.split("\n").map((s) => s.trim()).filter(Boolean);
  if (chapters.length === 0) {
    status("Add at least one chapter (one per line).", true);
    return;
  }
  const result = await PowerPoint.run((context) => createOrUpdateAgenda(context, chapters));
  status(`Agenda ${result}.`);
}

async function applyColors(): Promise<void> {
  const scheme = (byId("colorScheme") as HTMLSelectElement).value;
  let palette: string[];
  if (scheme === "master") {
    palette = await withSlide((ctx, slide) => loadMasterAccents(ctx, slide));
    if (palette.length === 0) {
      status("Couldn't read master colors on this host — using BLP instead.", true);
      palette = PALETTES.blue;
    }
  } else {
    palette = PALETTES[scheme] ?? PALETTES.blue;
  }
  currentData = readGrid();
  currentData.series.forEach((s, i) => {
    s.color = palette[i % palette.length];
  });
  renderGrid();
  scheduleApply();
  status(`Applied ${scheme} colors.`);
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
  refreshVisibility();
  renderGrid();
  const data = readGrid();
  data.type = readChartType();
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
      lastLegendPosition = model.options.legendPosition;
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
  data.type = readChartType();
  if (!validate(data)) return;
  busy = true;
  try {
    await withSlide(async (context, slide) => {
      const opts = readOptions();
      const box = currentBox; // redraw at the intended box; the resize poll updates currentBox
      // Preserve a manually-moved legend — but if the user just picked a new default
      // position, let it snap there instead of translating it back.
      const positionChanged = lastLegendPosition !== opts.legendPosition;
      const savedLegend =
        opts.showLegend && !positionChanged ? await getPartBox(context, slide, currentId!, "legend") : null;
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
      lastLegendPosition = opts.legendPosition;
    });
  } finally {
    busy = false;
  }
  currentData = data;
}


function applyModel(model: ChartModel): void {
  currentData = model.data;
  currentBox = model.box;
  currentId = model.id;
  currentName = model.name || "chart";
  lastLegendPosition = model.options.legendPosition;
  setOptionsUI(model.options);
  (byId("chartType") as HTMLSelectElement).value = model.data.type;
  renderGrid();
  setMode();
  refreshVisibility();
}

function readChartType(): ChartType {
  const val = (byId("chartType") as HTMLSelectElement).value as ChartType;
  return CHART_TYPES.includes(val) ? val : "barColumn";
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
  const numOrNull = (id: string): number | null => {
    const raw = v(id).trim();
    const n = Number(raw);
    return raw !== "" && isFinite(n) ? n : null;
  };
  const idx1 = (id: string) => Math.max(0, (Number(v(id)) || 1) - 1); // 1-based UI → 0-based model
  return {
    orientation: v("optOrientation") as Orientation,
    grouping: v("optGrouping") as Grouping,
    gap: Math.min(0.9, Math.max(0, (isFinite(gapPct) ? gapPct : 35) / 100)),
    referenceValue: numOrNull("optRefValue"),
    axisMin: numOrNull("optAxisMin"),
    axisMax: numOrNull("optAxisMax"),
    showTotals: c("optTotals"),
    showValueLabels: c("optLabels"),
    reverseCategories: c("optReverse"),
    showLegend: c("optLegend"),
    legendPosition: v("legendPosition") as "top" | "bottom" | "left" | "right",
    showGridlines: c("optGridlines"),
    showValueAxis: c("optAxis"),
    showAxisLine: c("optAxisLine"),
    lineSecondaryAxis: c("optLineSecondaryAxis"),
    lineAxisMin: numOrNull("optLineAxisMin"),
    lineAxisMax: numOrNull("optLineAxisMax"),
    pieHole: Math.min(0.9, Math.max(0, (Number(v("optPieHole")) || 0) / 100)),
    scatterQuadrant: c("optScatterQuadrant"),
    scatterAxes: c("optScatterAxes"),
    showConnectors: c("optConnectors"),
    reverseSeries: c("optReverseSeries"),
    referenceColor: v("optRefColor") || "#E8412C",
    diffArrow: v("optDiffArrow") as "off" | "total" | "series",
    diffFrom: idx1("optDiffFrom"),
    diffTo: idx1("optDiffTo"),
    diffSeries: idx1("optDiffSeries"),
    diffPercent: c("optDiffPercent"),
    diffPos: v("optDiffPos").trim() === "" ? -1 : Math.max(0, Number(v("optDiffPos")) || 0),
    cagrArrow: v("optCagrArrow") as "off" | "total" | "series",
    cagrFrom: idx1("optCagrFrom"),
    cagrTo: idx1("optCagrTo"),
    cagrSeries: idx1("optCagrSeries"),
    cagrPeriods: Math.max(0, Number(v("optCagrPeriods")) || 0),
    labelOverflow: v("labelOverflow") as "inside" | "outside",
    labelMode: v("labelMode") as "value" | "percent" | "valuePercent" | "percentValue",
    fontFamily: v("fontFamily"),
    segmentFontSize: clampInt(Number(v("segFontSize")), 6, 24),
    totalFontSize: clampInt(Number(v("totFontSize")), 6, 24),
    numberFormat: {
      decimals: clampInt(Number(v("nfDecimals")), 0, 3),
      scaleExp: Number(v("nfScale")) || 0,
      prefix: v("nfPrefix"),
      suffix: v("nfSuffix"),
      hideZero: c("nfHideZero"),
      thousandsSep: c("nfThousands"),
      sep: v("nfSep") as "locale" | "comma" | "dot" | "apos" | "space",
      negParens: c("nfParens"),
      plusSign: c("nfPlus"),
    },
  };
}

function setOptionsUI(o: ChartOptions): void {
  (byId("optOrientation") as HTMLSelectElement).value = o.orientation;
  (byId("optGrouping") as HTMLSelectElement).value = o.grouping;
  (byId("optGap") as HTMLInputElement).value = String(Math.round(o.gap * 100));
  (byId("optRefValue") as HTMLInputElement).value = o.referenceValue == null ? "" : String(o.referenceValue);
  (byId("optAxisMin") as HTMLInputElement).value = o.axisMin == null ? "" : String(o.axisMin);
  (byId("optAxisMax") as HTMLInputElement).value = o.axisMax == null ? "" : String(o.axisMax);
  (byId("optTotals") as HTMLInputElement).checked = o.showTotals;
  (byId("optLabels") as HTMLInputElement).checked = o.showValueLabels;
  (byId("optReverse") as HTMLInputElement).checked = o.reverseCategories;
  (byId("optLegend") as HTMLInputElement).checked = o.showLegend;
  (byId("legendPosition") as HTMLSelectElement).value = o.legendPosition;
  (byId("optGridlines") as HTMLInputElement).checked = o.showGridlines;
  (byId("optAxis") as HTMLInputElement).checked = o.showValueAxis;
  (byId("optAxisLine") as HTMLInputElement).checked = o.showAxisLine;
  (byId("optLineSecondaryAxis") as HTMLInputElement).checked = o.lineSecondaryAxis;
  (byId("optLineAxisMin") as HTMLInputElement).value = o.lineAxisMin == null ? "" : String(o.lineAxisMin);
  (byId("optLineAxisMax") as HTMLInputElement).value = o.lineAxisMax == null ? "" : String(o.lineAxisMax);
  (byId("optConnectors") as HTMLInputElement).checked = o.showConnectors;
  (byId("optReverseSeries") as HTMLInputElement).checked = o.reverseSeries;
  (byId("optRefColor") as HTMLInputElement).value = o.referenceColor || "#E8412C";
  (byId("optPieHole") as HTMLInputElement).value = String(Math.round((o.pieHole || 0) * 100));
  (byId("optScatterQuadrant") as HTMLInputElement).checked = o.scatterQuadrant;
  (byId("optScatterAxes") as HTMLInputElement).checked = o.scatterAxes;
  (byId("optDiffArrow") as HTMLSelectElement).value = o.diffArrow;
  (byId("optDiffPercent") as HTMLInputElement).checked = o.diffPercent;
  (byId("optDiffFrom") as HTMLInputElement).value = String(o.diffFrom + 1);
  (byId("optDiffTo") as HTMLInputElement).value = String(o.diffTo + 1);
  (byId("optDiffSeries") as HTMLInputElement).value = String(o.diffSeries + 1);
  (byId("optDiffPos") as HTMLInputElement).value = o.diffPos < 0 ? "" : String(o.diffPos);
  (byId("optCagrArrow") as HTMLSelectElement).value = o.cagrArrow;
  (byId("optCagrFrom") as HTMLInputElement).value = String(o.cagrFrom + 1);
  (byId("optCagrTo") as HTMLInputElement).value = String(o.cagrTo + 1);
  (byId("optCagrSeries") as HTMLInputElement).value = String(o.cagrSeries + 1);
  (byId("optCagrPeriods") as HTMLInputElement).value = String(o.cagrPeriods);
  (byId("labelOverflow") as HTMLSelectElement).value = o.labelOverflow;
  (byId("labelMode") as HTMLSelectElement).value = o.labelMode;
  (byId("fontFamily") as HTMLSelectElement).value = o.fontFamily;
  (byId("segFontSize") as HTMLInputElement).value = String(o.segmentFontSize);
  (byId("totFontSize") as HTMLInputElement).value = String(o.totalFontSize);
  (byId("nfDecimals") as HTMLInputElement).value = String(o.numberFormat.decimals);
  (byId("nfScale") as HTMLSelectElement).value = String(o.numberFormat.scaleExp);
  (byId("nfPrefix") as HTMLInputElement).value = o.numberFormat.prefix;
  (byId("nfSuffix") as HTMLInputElement).value = o.numberFormat.suffix;
  (byId("nfHideZero") as HTMLInputElement).checked = o.numberFormat.hideZero;
  (byId("nfThousands") as HTMLInputElement).checked = o.numberFormat.thousandsSep;
  (byId("nfSep") as HTMLSelectElement).value = o.numberFormat.sep;
  (byId("nfParens") as HTMLInputElement).checked = o.numberFormat.negParens;
  (byId("nfPlus") as HTMLInputElement).checked = o.numberFormat.plusSign;
}

function clampInt(n: number, lo: number, hi: number): number {
  if (!isFinite(n)) return lo;
  return Math.max(lo, Math.min(hi, Math.round(n)));
}
