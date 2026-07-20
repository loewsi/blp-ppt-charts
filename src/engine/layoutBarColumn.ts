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

  // Reserved bands (only when a feature is on, so defaults are unchanged).
  const axisBand = opt.showValueAxis ? 28 : 0;
  const legendH = 24; // horizontal legend band (top/bottom)
  const legendW = 92; // vertical legend band (left/right)
  const legendPos = opt.legendPosition;

  // Accumulate insets on each side.
  let insTop = isColumn ? PAD_MAIN : PAD_MINOR;
  let insBottom = isColumn ? PAD_CROSS : PAD_MINOR;
  let insLeft = isColumn ? PAD_MINOR : PAD_CROSS;
  let insRight = isColumn ? PAD_MINOR : PAD_MAIN;
  if (opt.showValueAxis) {
    if (isColumn) insLeft += axisBand;
    else insBottom += axisBand;
  }
  if (opt.showLegend) {
    if (legendPos === "top") insTop += legendH;
    else if (legendPos === "bottom") insBottom += legendH;
    else if (legendPos === "left") insLeft += legendW;
    else insRight += legendW;
  }

  const plotLeft = box.left + insLeft;
  const plotTop = box.top + insTop;
  const plotW = box.width - insLeft - insRight;
  const plotH = box.height - insTop - insBottom;

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

  const fam = opt.fontFamily;
  const labelH = opt.segmentFontSize + 5; // approx label box height for collision checks

  function pushCenteredLabel(cx: number, cy: number, _thick: number, text: string, meta: ShapeMeta) {
    const w = estTextW(text, opt.segmentFontSize); // box only as wide as the text, centered on (cx, cy)
    prims.push({ kind: "text", x: cx - w / 2, y: cy - 7, w, h: 14, text, color: LABEL_LIGHT, size: opt.segmentFontSize, bold: false, align: "center", family: fam, meta });
  }

  // Small segment, kept inside: a chip filled with the segment color so the
  // white value stays readable even when nudged off-center to dodge a neighbour.
  function pushChipLabel(cx: number, cy: number, _thick: number, text: string, fill: string, meta: ShapeMeta) {
    const w = estTextW(text, opt.segmentFontSize);
    prims.push({ kind: "text", x: cx - w / 2, y: cy - 7, w, h: 14, text, color: LABEL_LIGHT, size: opt.segmentFontSize, bold: false, align: "center", family: fam, bg: fill, meta });
  }

  // Small segment: put the value just outside the column/bar instead of hiding it.
  function pushOutsideLabel(cx: number, cy: number, halfThick: number, text: string, meta: ShapeMeta) {
    const w = estTextW(text, opt.segmentFontSize);
    if (isColumn) {
      prims.push({ kind: "text", x: cx + halfThick + 3, y: cy - 7, w, h: 14, text, color: LABEL_DARK, size: opt.segmentFontSize, bold: false, align: "left", family: fam, meta });
    } else {
      prims.push({ kind: "text", x: cx - w / 2, y: cy - halfThick - 14, w, h: 14, text, color: LABEL_DARK, size: opt.segmentFontSize, bold: false, align: "center", family: fam, meta });
    }
  }

  function pushTotal(catStart: number, thick: number, stackPx: number, text: string, meta: ShapeMeta) {
    if (isColumn) {
      const y = plotTop + plotH - stackPx - 18;
      prims.push({ kind: "text", x: plotLeft + catStart - 8, y, w: thick + 16, h: 16, text, color: LABEL_DARK, size: opt.totalFontSize, bold: true, align: "center", family: fam, meta });
    } else {
      prims.push({ kind: "text", x: plotLeft + stackPx + 4, y: plotTop + catStart + thick / 2 - 8, w: PAD_MAIN + 26, h: 16, text, color: LABEL_DARK, size: opt.totalFontSize, bold: true, align: "left", family: fam, meta });
    }
  }

  function pushCategoryLabel(catStart: number, thick: number, text: string, meta: ShapeMeta) {
    if (isColumn) {
      prims.push({ kind: "text", x: plotLeft + catStart - (slot - thick) / 2, y: plotTop + plotH + 3, w: slot, h: 16, text, color: LABEL_DARK, size: 9, bold: false, align: "center", family: fam, meta });
    } else {
      prims.push({ kind: "text", x: box.left, y: plotTop + catStart + thick / 2 - 8, w: PAD_CROSS - 3, h: 16, text, color: LABEL_DARK, size: 9, bold: false, align: "right", family: fam, meta });
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
          prims.push({ kind: "text", x: box.left, y: y - 7, w: axisBand - 3, h: 14, text: axisText(t), color: LABEL_DARK, size: 8, bold: false, align: "right", family: fam, meta: { objectType: "valueAxis" } });
        }
      } else {
        const x = plotLeft + t * valScale;
        if (opt.showGridlines && t > 0) {
          prims.push({ kind: "line", x1: x, y1: plotTop, x2: x, y2: plotTop + plotH, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
        }
        if (opt.showValueAxis) {
          prims.push({ kind: "text", x: x - 22, y: plotTop + plotH + 2, w: 44, h: 14, text: axisText(t), color: LABEL_DARK, size: 8, bold: false, align: "center", family: fam, meta: { objectType: "valueAxis" } });
        }
      }
    }
  }

  function drawLegend() {
    const sw = 10;
    const gapx = 6;
    const textWs = data.series.map((s) => estTextW(s.name, 9)); // legend text sized to content
    if (legendPos === "top" || legendPos === "bottom") {
      const itemGap = 14;
      const widths = textWs.map((tw) => sw + gapx + tw);
      const total = widths.reduce((a, b) => a + b, 0) + itemGap * Math.max(0, data.series.length - 1);
      let x = box.left + Math.max(0, (box.width - total) / 2);
      const y = legendPos === "top" ? box.top + 4 : box.top + box.height - legendH + 6;
      data.series.forEach((s, i) => {
        prims.push({ kind: "rect", x, y: y + 8 - sw / 2, w: sw, h: sw, fill: s.color, meta: { objectType: "legendEntry", seriesIndex: i } });
        prims.push({ kind: "text", x: x + sw + gapx, y, w: textWs[i], h: 16, text: s.name, color: LABEL_DARK, size: 9, bold: false, align: "left", family: fam, meta: { objectType: "legend", seriesIndex: i } });
        x += widths[i] + itemGap;
      });
    } else {
      const x = legendPos === "left" ? box.left + 4 : box.left + box.width - legendW + 4;
      let y = plotTop;
      data.series.forEach((s, i) => {
        prims.push({ kind: "rect", x, y: y + 8 - sw / 2, w: sw, h: sw, fill: s.color, meta: { objectType: "legendEntry", seriesIndex: i } });
        prims.push({ kind: "text", x: x + sw + gapx, y: y + 1, w: Math.min(legendW - sw - gapx - 4, textWs[i]), h: 14, text: s.name, color: LABEL_DARK, size: 9, bold: false, align: "left", family: fam, meta: { objectType: "legend", seriesIndex: i } });
        y += 16;
      });
    }
  }

  // Gridlines/axis first, so gridlines sit behind the bars.
  if (opt.showGridlines || opt.showValueAxis) drawGridAxis();

  for (let k = 0; k < nCats; k++) {
    const ci = order[k];
    const slotStart = k * slot + (slot - catThick) / 2;
    const total = totals[ci] || 0;
    let prevLabelC = -Infinity; // center of the last placed small label (along value axis)
    let side = 1; // alternates the collision offset left/right (or up/down)

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
        if (opt.showValueLabels) {
          const text = norm100 ? formatPercent(raw, total, nf.decimals) : formatNumber(raw, nf);
          const m: ShapeMeta = { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci };
          if (text) {
            if (v * valScale >= MIN_SEG_FOR_LABEL) {
              pushCenteredLabel(r.cx, r.cy, catThick, text, m);
              prevLabelC = isColumn ? r.cy : r.cx;
            } else if (opt.labelOverflow === "outside") {
              pushOutsideLabel(r.cx, r.cy, catThick / 2, text, m);
            } else {
              // keep inside: chip + nudge off-center when it collides with the last one
              let cx = r.cx;
              let cy = r.cy;
              const cur = isColumn ? cy : cx;
              if (Math.abs(cur - prevLabelC) < labelH) {
                // Nudge just enough to clear the neighbour: half a label width.
                const off = estTextW(text, opt.segmentFontSize) / 2 + 3;
                if (isColumn) cx += side * off;
                else cy += side * (labelH * 0.6);
                side = -side;
              }
              pushChipLabel(cx, cy, catThick, text, data.series[si].color, m);
              prevLabelC = cur;
            }
          }
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
        if (opt.showValueLabels) {
          const text = formatNumber(raw, nf);
          const m: ShapeMeta = { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci };
          if (text) {
            if (raw * valScale >= MIN_SEG_FOR_LABEL) pushCenteredLabel(r.cx, r.cy, laneThick, text, m);
            else if (opt.labelOverflow === "outside") pushOutsideLabel(r.cx, r.cy, laneThick / 2, text, m);
            else pushChipLabel(r.cx, r.cy, laneThick, text, data.series[si].color, m);
          }
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

/** Rough width (pt) of a text box sized to its content, with a little padding. */
function estTextW(text: string, size: number): number {
  return Math.max(10, Math.round(text.length * size * 0.62) + 6);
}
