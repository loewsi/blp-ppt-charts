import type { ChartModel } from "../model/chartModel";
import { DEFAULT_OPTIONS } from "../model/chartModel";
import type { Primitive, ShapeMeta } from "./primitives";
import { formatNumber, formatPercent } from "./format";

const AXIS_COLOR = "#001C54";
const GRID_COLOR = "#D7E2F4";
const LABEL_LIGHT = "#FFFFFF";
const LABEL_DARK = "#001C54";

const PAD_MAIN = 22; // value-axis end: room for totals
const PAD_CROSS = 26; // category-axis end: room for category labels
const PAD_MINOR = 6;
const MIN_SEG_FOR_LABEL = 12; // pt along the value axis

/**
 * One engine for the whole column/bar family. Orientation (column|bar) and
 * grouping (stacked|clustered|stacked100) come from model.options. Emits generic
 * primitives, so render.ts and every annotation built later works unchanged.
 */
export function layoutBarColumn(model: ChartModel): Primitive[] {
  const opt = model.options ?? DEFAULT_OPTIONS;
  const nf = opt.numberFormat;
  const { data, box } = model;
  const prims: Primitive[] = [];
  const nCats = data.categories.length;
  const nSer = data.series.length;
  if (nCats === 0 || nSer === 0) return prims;

  const isColumn = opt.orientation === "column";
  const stacked = opt.grouping === "stacked" || opt.grouping === "stacked100";
  const norm100 = opt.grouping === "stacked100";

  const order = data.categories.map((_, i) => i);
  if (opt.reverseCategories) order.reverse();

  // Extra reserved bands (only when the feature is on, so defaults are unchanged).
  const legendBand = opt.showLegend ? 24 : 0;
  const axisBand = opt.showValueAxis ? 28 : 0;

  // Plot rectangle (points).
  let plotLeft: number, plotTop: number, plotW: number, plotH: number;
  if (isColumn) {
    plotLeft = box.left + PAD_MINOR + axisBand;
    plotTop = box.top + PAD_MAIN;
    plotW = box.width - PAD_MINOR * 2 - axisBand;
    plotH = box.height - PAD_MAIN - PAD_CROSS - legendBand;
  } else {
    plotLeft = box.left + PAD_CROSS;
    plotTop = box.top + PAD_MINOR;
    plotW = box.width - PAD_CROSS - PAD_MAIN;
    plotH = box.height - PAD_MINOR - legendBand - axisBand;
  }

  const catExtent = isColumn ? plotW : plotH;
  const valExtent = isColumn ? plotH : plotW;
  const slot = catExtent / nCats;
  const catThick = slot * (1 - opt.gap);

  const totals = data.categories.map((_, ci) =>
    data.series.reduce((s, se) => s + safe(se.values[ci]), 0)
  );

  let maxVal: number;
  if (norm100) maxVal = 1;
  else if (stacked) maxVal = Math.max(1, ...totals);
  else {
    let m = 1;
    for (const se of data.series) for (const v of se.values) m = Math.max(m, safe(v));
    maxVal = m;
  }
  const valScale = valExtent / maxVal;

  // Emit a rectangle for a segment/bar; returns its center for label placement.
  function emitRect(catStart: number, thick: number, v0: number, v1: number, fill: string, meta: ShapeMeta) {
    if (isColumn) {
      const x = plotLeft + catStart;
      const baseline = plotTop + plotH;
      const y = baseline - v1 * valScale;
      const h = (v1 - v0) * valScale;
      prims.push({ kind: "rect", x, y, w: thick, h, fill, meta });
      return { cx: x + thick / 2, cy: y + h / 2 };
    }
    const y = plotTop + catStart;
    const x = plotLeft + v0 * valScale;
    const w = (v1 - v0) * valScale;
    prims.push({ kind: "rect", x, y, w, h: thick, fill, meta });
    return { cx: x + w / 2, cy: y + thick / 2 };
  }

  function pushCenteredLabel(cx: number, cy: number, thick: number, text: string, size: number, meta: ShapeMeta) {
    const w = isColumn ? thick : 64;
    prims.push({ kind: "text", x: cx - w / 2, y: cy - 7, w, h: 14, text, color: LABEL_LIGHT, size, bold: false, align: "center", meta });
  }

  function pushTotal(catStart: number, thick: number, stackPx: number, text: string, meta: ShapeMeta) {
    if (isColumn) {
      const y = plotTop + plotH - stackPx - 18;
      prims.push({ kind: "text", x: plotLeft + catStart - 8, y, w: thick + 16, h: 16, text, color: LABEL_DARK, size: 10, bold: true, align: "center", meta });
    } else {
      prims.push({ kind: "text", x: plotLeft + stackPx + 4, y: plotTop + catStart + thick / 2 - 8, w: PAD_MAIN + 26, h: 16, text, color: LABEL_DARK, size: 10, bold: true, align: "left", meta });
    }
  }

  function pushCategoryLabel(catStart: number, thick: number, text: string, meta: ShapeMeta) {
    if (isColumn) {
      prims.push({ kind: "text", x: plotLeft + catStart - (slot - thick) / 2, y: plotTop + plotH + 3, w: slot, h: 16, text, color: LABEL_DARK, size: 9, bold: false, align: "center", meta });
    } else {
      prims.push({ kind: "text", x: box.left, y: plotTop + catStart + thick / 2 - 8, w: PAD_CROSS - 3, h: 16, text, color: LABEL_DARK, size: 9, bold: false, align: "right", meta });
    }
  }

  function axisText(t: number): string {
    if (norm100) return `${Math.round(t * 100)}%`;
    return formatNumber(t, { ...nf, hideZero: false });
  }

  function drawGridAxis() {
    for (const t of niceTicks(maxVal, 4)) {
      if (isColumn) {
        const y = plotTop + plotH - t * valScale;
        if (opt.showGridlines && t > 0) {
          prims.push({ kind: "line", x1: plotLeft, y1: y, x2: plotLeft + plotW, y2: y, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
        }
        if (opt.showValueAxis) {
          prims.push({ kind: "text", x: box.left, y: y - 7, w: axisBand - 3, h: 14, text: axisText(t), color: LABEL_DARK, size: 8, bold: false, align: "right", meta: { objectType: "valueAxis" } });
        }
      } else {
        const x = plotLeft + t * valScale;
        if (opt.showGridlines && t > 0) {
          prims.push({ kind: "line", x1: x, y1: plotTop, x2: x, y2: plotTop + plotH, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
        }
        if (opt.showValueAxis) {
          prims.push({ kind: "text", x: x - 22, y: plotTop + plotH + 2, w: 44, h: 14, text: axisText(t), color: LABEL_DARK, size: 8, bold: false, align: "center", meta: { objectType: "valueAxis" } });
        }
      }
    }
  }

  function drawLegend() {
    const sw = 10;
    const gapx = 6;
    const itemGap = 14;
    const widths = data.series.map((s) => sw + gapx + Math.max(20, s.name.length * 5.5));
    const total = widths.reduce((a, b) => a + b, 0) + itemGap * Math.max(0, data.series.length - 1);
    let x = box.left + Math.max(0, (box.width - total) / 2);
    const y = box.top + box.height - legendBand + 6;
    data.series.forEach((s, i) => {
      prims.push({ kind: "rect", x, y: y + 1, w: sw, h: sw, fill: s.color, meta: { objectType: "legendEntry", seriesIndex: i } });
      prims.push({ kind: "text", x: x + sw + gapx, y: y - 3, w: widths[i] - sw - gapx, h: 16, text: s.name, color: LABEL_DARK, size: 9, bold: false, align: "left", meta: { objectType: "legend", seriesIndex: i } });
      x += widths[i] + itemGap;
    });
  }

  // Gridlines/axis first, so gridlines sit behind the bars.
  if (opt.showGridlines || opt.showValueAxis) drawGridAxis();

  for (let k = 0; k < nCats; k++) {
    const ci = order[k];
    const slotStart = k * slot + (slot - catThick) / 2;
    const total = totals[ci] || 0;

    if (stacked) {
      let cum = 0;
      for (let si = 0; si < nSer; si++) {
        const raw = safe(data.series[si].values[ci]);
        const v = norm100 ? (total === 0 ? 0 : raw / total) : raw;
        if (v <= 0) continue;
        const r = emitRect(slotStart, catThick, cum, cum + v, data.series[si].color, {
          objectType: "segment",
          seriesIndex: si,
          categoryIndex: ci,
        });
        if (opt.showValueLabels && v * valScale >= MIN_SEG_FOR_LABEL) {
          const text = norm100 ? formatPercent(raw, total, nf.decimals) : formatNumber(raw, nf);
          if (text) pushCenteredLabel(r.cx, r.cy, catThick, text, 9, { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci });
        }
        cum += v;
      }
      if (opt.showTotals && !norm100) {
        const text = formatNumber(total, nf);
        if (text) pushTotal(slotStart, catThick, total * valScale, text, { objectType: "totalLabel", categoryIndex: ci });
      }
    } else {
      const laneThick = catThick / nSer;
      for (let si = 0; si < nSer; si++) {
        const raw = safe(data.series[si].values[ci]);
        if (raw <= 0) continue;
        const r = emitRect(slotStart + si * laneThick, laneThick, 0, raw, data.series[si].color, {
          objectType: "segment",
          seriesIndex: si,
          categoryIndex: ci,
        });
        if (opt.showValueLabels && raw * valScale >= MIN_SEG_FOR_LABEL) {
          const text = formatNumber(raw, nf);
          if (text) pushCenteredLabel(r.cx, r.cy, laneThick, text, 8, { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci });
        }
      }
    }

    pushCategoryLabel(slotStart, catThick, data.categories[ci], { objectType: "categoryLabel", categoryIndex: ci });
  }

  // Baseline along the category axis.
  const baselineMeta: ShapeMeta = { objectType: "baseline" };
  if (isColumn) {
    const baseline = plotTop + plotH;
    prims.push({ kind: "line", x1: plotLeft, y1: baseline, x2: plotLeft + plotW, y2: baseline, color: AXIS_COLOR, weight: 1, meta: baselineMeta });
  } else {
    prims.push({ kind: "line", x1: plotLeft, y1: plotTop, x2: plotLeft, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: baselineMeta });
  }

  if (opt.showLegend) drawLegend();

  return prims;
}

/** "Nice" axis tick values from 0 up to ~max (1/2/5 × 10^n steps). */
export function niceTicks(max: number, target = 4): number[] {
  if (!(max > 0)) return [0];
  const raw = max / target;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
  const ticks: number[] = [];
  for (let t = 0; t <= max + step * 1e-9; t += step) ticks.push(Number(t.toFixed(10)));
  return ticks;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}
