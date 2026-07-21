import type { ChartModel } from "../model/chartModel";
import { DEFAULT_OPTIONS } from "../model/chartModel";
import type { Primitive, ShapeMeta } from "./primitives";
import { formatNumber, formatPercent } from "./format";

const AXIS_COLOR = "#001C54";
const GRID_COLOR = "#D7E2F4";
const CONNECTOR_COLOR = "#9AA6BF";
const REF_COLOR = "#E8412C";
const ARROW_COLOR = "#001C54"; // difference / CAGR arrows
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
  const serOrder = data.series.map((_, i) => i); // stacking/cluster order (bottom→top)
  if (opt.reverseSeries) serOrder.reverse();

  // A series can be drawn as a line instead of a bar (combination charts). Lines
  // only make sense on a value axis, so we keep them to column orientation and skip
  // them for 100% stacked; elsewhere a "line" series just renders as a bar.
  const isLine = (si: number) => isColumn && !norm100 && data.series[si].kind === "line";
  const barOrder = serOrder.filter((si) => !isLine(si)); // series drawn as bars, in stack/cluster order
  const lineSeries = data.series.map((_, i) => i).filter(isLine); // series drawn as lines
  const nBar = barOrder.length;
  const useSecondary = opt.lineSecondaryAxis && lineSeries.length > 0; // line on its own right axis

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
    if (useSecondary && isColumn) insRight += axisBand; // right-hand secondary axis labels
  }
  if (opt.showLegend) {
    if (legendPos === "top") insTop += legendH;
    else if (legendPos === "bottom") insBottom += legendH;
    else if (legendPos === "left") insLeft += legendW;
    else insRight += legendW;
  }
  const cagrBand = opt.cagrArrow !== "off" && isColumn ? 20 : 0; // room for the CAGR arrow above the plot
  insTop += cagrBand;

  const plotLeft = box.left + insLeft;
  const plotTop = box.top + insTop;
  // Clamp so a very small box still yields positive geometry (chart stays visible).
  const plotW = Math.max(24, box.width - insLeft - insRight);
  const plotH = Math.max(24, box.height - insTop - insBottom);

  const catExtent = isColumn ? plotW : plotH;
  const valExtent = isColumn ? plotH : plotW;
  const slot = catExtent / nCats;
  const catThick = slot * (1 - opt.gap);

  // Totals cover the bar (stacked) series only; line series sit on their own.
  const totals = data.categories.map((_, ci) =>
    barOrder.reduce((s, si) => s + safe(data.series[si].values[ci]), 0)
  );

  // Value-axis range, including negatives; zero may sit inside the plot.
  let vMax: number;
  let vMin = 0;
  if (norm100) {
    vMax = 1;
  } else if (stacked) {
    let pMax = 0;
    let nMin = 0;
    for (let ci = 0; ci < nCats; ci++) {
      let p = 0;
      let ng = 0;
      for (const si of barOrder) {
        const v = safe(data.series[si].values[ci]);
        if (v > 0) p += v;
        else ng += v;
      }
      pMax = Math.max(pMax, p);
      nMin = Math.min(nMin, ng);
    }
    vMax = Math.max(1, pMax);
    vMin = Math.min(0, nMin);
  } else {
    let mx = 1;
    let mn = 0;
    for (const si of barOrder)
      for (const v of data.series[si].values) {
        mx = Math.max(mx, safe(v));
        mn = Math.min(mn, safe(v));
      }
    vMax = mx;
    vMin = Math.min(0, mn);
  }
  // Fold line-series values into the primary range so the line fits — unless the
  // line rides its own secondary axis.
  if (!useSecondary)
    for (const si of lineSeries)
      for (const v of data.series[si].values) {
        vMax = Math.max(vMax, safe(v));
        vMin = Math.min(0, vMin, safe(v));
      }

  // Secondary axis range (line series only).
  let lMin = 0;
  let lMax = 1;
  if (useSecondary) {
    lMin = Infinity;
    lMax = -Infinity;
    for (const si of lineSeries)
      for (const v of data.series[si].values) {
        const x = safe(v);
        lMin = Math.min(lMin, x);
        lMax = Math.max(lMax, x);
      }
    if (!isFinite(lMin)) {
      lMin = 0;
      lMax = 1;
    }
    lMin = Math.min(0, lMin);
    if (lMax <= lMin) lMax = lMin + 1;
  }
  // Manual axis overrides (auto scale unless the user fixes an end). Not for 100%.
  if (!norm100) {
    if (opt.axisMax != null && isFinite(opt.axisMax)) vMax = opt.axisMax;
    if (opt.axisMin != null && isFinite(opt.axisMin)) vMin = opt.axisMin;
    if (vMax <= vMin) vMax = vMin + 1; // keep a positive range
  }
  const vRange = vMax - vMin || 1;
  const vScale = valExtent / vRange;
  const yv = (v: number) => plotTop + (vMax - v) * vScale; // column: value -> y
  const xv = (v: number) => plotLeft + (v - vMin) * vScale; // bar: value -> x
  // Line series map through the secondary axis when enabled, else the primary one.
  const lScale = plotH / (lMax - lMin || 1);
  const lineY = (v: number) => (useSecondary ? plotTop + (lMax - v) * lScale : yv(v));

  // Emit a rectangle for a segment/bar; returns its center for label placement.
  function emitRect(catStart: number, thick: number, v0: number, v1: number, fill: string, meta: ShapeMeta) {
    const lo = Math.min(v0, v1);
    const hi = Math.max(v0, v1);
    if (isColumn) {
      const x = plotLeft + catStart;
      const yTop = yv(hi);
      const h = Math.max(0, yv(lo) - yTop);
      prims.push({ kind: "rect", x, y: yTop, w: thick, h, fill, meta });
      return { cx: x + thick / 2, cy: yTop + h / 2 };
    }
    const y = plotTop + catStart;
    const xL = xv(lo);
    const w = Math.max(0, xv(hi) - xL);
    prims.push({ kind: "rect", x: xL, y, w, h: thick, fill, meta });
    return { cx: xL + w / 2, cy: y + thick / 2 };
  }

  const fam = opt.fontFamily;
  const labelH = opt.segmentFontSize + 5; // approx label box height for collision checks

  // Does a centered value label fit inside the bar (so white-on-color stays readable)?
  // If not, the caller falls back to a chip so the text keeps a colored backing.
  function labelFits(text: string, thick: number): boolean {
    if (isColumn) return estTextW(text, opt.segmentFontSize) <= thick + 2; // text must fit the bar width
    return thick >= opt.segmentFontSize + 4; // bar must be tall enough for the text row
  }

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

  type Chip = { c: number; cx: number; cy: number; text: string; fill: string; meta: ShapeMeta };
  // Place a category's small "inside" chips: any run of overlapping labels is
  // spread to alternating sides (the one nearest the axis moves too); isolated
  // labels stay centered.
  function resolveChips(list: Chip[]) {
    list.sort((a, b) => a.c - b.c);
    let i = 0;
    while (i < list.length) {
      let j = i;
      while (j + 1 < list.length && Math.abs(list[j + 1].c - list[j].c) < labelH) j++;
      for (let k = i; k <= j; k++) {
        const it = list[k];
        let cx = it.cx;
        let cy = it.cy;
        if (j > i) {
          const dir = (k - i) % 2 === 0 ? 1 : -1;
          if (isColumn) cx += dir * (estTextW(it.text, opt.segmentFontSize) / 2 + 3);
          else cy += dir * (labelH * 0.7);
        }
        pushChipLabel(cx, cy, catThick, it.text, it.fill, it.meta);
      }
      i = j + 1;
    }
  }

  function pushTotal(catStart: number, thick: number, topValue: number, text: string, meta: ShapeMeta) {
    if (isColumn) {
      const y = yv(topValue) - 18;
      prims.push({ kind: "text", x: plotLeft + catStart - 8, y, w: thick + 16, h: 16, text, color: LABEL_DARK, size: opt.totalFontSize, bold: true, align: "center", family: fam, meta });
    } else {
      prims.push({ kind: "text", x: xv(topValue) + 4, y: plotTop + catStart + thick / 2 - 8, w: PAD_MAIN + 26, h: 16, text, color: LABEL_DARK, size: opt.totalFontSize, bold: true, align: "left", family: fam, meta });
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
    for (const t of axisTicks(vMin, vMax, 4)) {
      const isZero = Math.abs(t) < 1e-9;
      if (isColumn) {
        const y = yv(t);
        if (opt.showGridlines && !isZero) {
          prims.push({ kind: "line", x1: plotLeft, y1: y, x2: plotLeft + plotW, y2: y, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
        }
        if (opt.showValueAxis) {
          prims.push({ kind: "text", x: box.left, y: y - 7, w: axisBand - 3, h: 14, text: axisText(t), color: LABEL_DARK, size: 8, bold: false, align: "right", family: fam, meta: { objectType: "valueAxis" } });
        }
      } else {
        const x = xv(t);
        if (opt.showGridlines && !isZero) {
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
        prims.push({ kind: "text", x: x + sw + gapx, y, w: textWs[i], h: 16, text: s.name, color: LABEL_DARK, size: 9, bold: false, align: "left", family: fam, autofit: true, meta: { objectType: "legend", seriesIndex: i } });
        x += widths[i] + itemGap;
      });
    } else {
      const x = legendPos === "left" ? box.left + 4 : box.left + box.width - legendW + 4;
      let y = plotTop;
      data.series.forEach((s, i) => {
        prims.push({ kind: "rect", x, y: y + 8 - sw / 2, w: sw, h: sw, fill: s.color, meta: { objectType: "legendEntry", seriesIndex: i } });
        prims.push({ kind: "text", x: x + sw + gapx, y: y + 1, w: Math.min(legendW - sw - gapx - 4, textWs[i]), h: 14, text: s.name, color: LABEL_DARK, size: 9, bold: false, align: "left", family: fam, autofit: true, meta: { objectType: "legend", seriesIndex: i } });
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
    const chips: Chip[] = []; // small "inside" labels, positioned after the stack is known

    if (stacked) {
      let cumPos = 0;
      let cumNeg = 0;
      for (let sidx = 0; sidx < nBar; sidx++) {
        const si = barOrder[sidx];
        const color = data.series[si].color;
        const raw = safe(data.series[si].values[ci]);
        const v = norm100 ? (total === 0 ? 0 : raw / total) : raw;
        const m: ShapeMeta = { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci };
        if (v === 0) {
          // No bar to draw, but surface a "0" at the running stack position (in order),
          // not pinned to the axis, when zeros aren't hidden.
          if (!norm100 && opt.showValueLabels) {
            const t = formatNumber(raw, nf);
            if (t) {
              const cx = isColumn ? plotLeft + slotStart + catThick / 2 : xv(cumPos);
              const cy = isColumn ? yv(cumPos) : plotTop + slotStart + catThick / 2;
              chips.push({ c: isColumn ? cy : cx, cx, cy, text: t, fill: color, meta: m });
            }
          }
          continue;
        }
        let seg0: number;
        let seg1: number;
        if (v > 0) {
          seg0 = cumPos;
          seg1 = cumPos + v;
          cumPos += v;
        } else {
          seg0 = cumNeg + v; // extends below the running negative total
          seg1 = cumNeg;
          cumNeg += v;
        }
        const r = emitRect(slotStart, catThick, seg0, seg1, color, { objectType: "segment", seriesIndex: si, categoryIndex: ci });
        if (opt.showValueLabels) {
          const text = norm100 ? formatPercent(raw, total, nf.decimals, nf.sep) : formatNumber(raw, nf);
          const segPx = Math.abs(v) * vScale;
          if (text) {
            if (opt.labelOverflow === "outside") {
              pushOutsideLabel(r.cx, r.cy, catThick / 2, text, m); // outside = ALL labels, not just small
            } else if (segPx >= MIN_SEG_FOR_LABEL && labelFits(text, catThick)) {
              pushCenteredLabel(r.cx, r.cy, catThick, text, m);
            } else if (segPx >= MIN_SEG_FOR_LABEL) {
              pushChipLabel(r.cx, r.cy, catThick, text, color, m); // tall enough, but text wider than bar → chip
            } else {
              chips.push({ c: isColumn ? r.cy : r.cx, cx: r.cx, cy: r.cy, text, fill: color, meta: m });
            }
          }
        }
      }
      if (opt.showTotals) {
        const text = formatNumber(total, nf); // absolute total, shown even on 100% stacked
        const topPos = norm100 ? (total === 0 ? 0 : 1) : Math.max(cumPos, 0);
        if (text) pushTotal(slotStart, catThick, topPos, text, { objectType: "totalLabel", categoryIndex: ci });
      }
    } else {
      const laneThick = catThick / Math.max(1, nBar);
      for (let sidx = 0; sidx < nBar; sidx++) {
        const si = barOrder[sidx];
        const color = data.series[si].color;
        const raw = safe(data.series[si].values[ci]);
        if (raw === 0) continue;
        const r = emitRect(slotStart + sidx * laneThick, laneThick, 0, raw, color, { objectType: "segment", seriesIndex: si, categoryIndex: ci });
        if (opt.showValueLabels) {
          const text = formatNumber(raw, nf);
          const m: ShapeMeta = { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci };
          if (text) {
            const big = Math.abs(raw) * vScale >= MIN_SEG_FOR_LABEL;
            if (opt.labelOverflow === "outside") pushOutsideLabel(r.cx, r.cy, laneThick / 2, text, m);
            else if (big && labelFits(text, laneThick)) pushCenteredLabel(r.cx, r.cy, laneThick, text, m);
            else if (big) pushChipLabel(r.cx, r.cy, laneThick, text, color, m);
            else pushChipLabel(r.cx, r.cy, laneThick, text, color, m);
          }
        }
      }
    }

    resolveChips(chips);
    pushCategoryLabel(slotStart, catThick, data.categories[ci], { objectType: "categoryLabel", categoryIndex: ci });
  }

  // Baseline along the category axis.
  // Connectors: link the segment boundaries of adjacent stacked columns.
  if (opt.showConnectors && stacked && isColumn && nBar > 1) {
    for (let k = 0; k < nCats - 1; k++) {
      const ciA = order[k];
      const ciB = order[k + 1];
      const totA = totals[ciA] || 0;
      const totB = totals[ciB] || 0;
      const aRight = plotLeft + k * slot + (slot - catThick) / 2 + catThick;
      const bLeft = plotLeft + (k + 1) * slot + (slot - catThick) / 2;
      let cumA = 0;
      let cumB = 0;
      for (let sidx = 0; sidx < nBar - 1; sidx++) {
        const si = barOrder[sidx];
        const rawA = safe(data.series[si].values[ciA]);
        const rawB = safe(data.series[si].values[ciB]);
        cumA += norm100 ? (totA === 0 ? 0 : rawA / totA) : rawA;
        cumB += norm100 ? (totB === 0 ? 0 : rawB / totB) : rawB;
        const ya = yv(cumA);
        const yb = yv(cumB);
        prims.push({ kind: "line", x1: aRight, y1: ya, x2: bLeft, y2: yb, color: CONNECTOR_COLOR, weight: 0.75, meta: { objectType: "connector", categoryIndex: k } });
      }
    }
  }

  const baselineMeta: ShapeMeta = { objectType: "baseline" };
  if (isColumn) {
    const y0 = yv(0);
    prims.push({ kind: "line", x1: plotLeft, y1: y0, x2: plotLeft + plotW, y2: y0, color: AXIS_COLOR, weight: 1, meta: baselineMeta });
  } else {
    const x0 = xv(0);
    prims.push({ kind: "line", x1: x0, y1: plotTop, x2: x0, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: baselineMeta });
  }

  // Value-axis line (the "y-axis" for columns; the horizontal scale line for bars).
  if (opt.showAxisLine) {
    if (isColumn) {
      prims.push({ kind: "line", x1: plotLeft, y1: plotTop, x2: plotLeft, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: { objectType: "valueAxis" } });
    } else {
      prims.push({ kind: "line", x1: plotLeft, y1: plotTop + plotH, x2: plotLeft + plotW, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: { objectType: "valueAxis" } });
    }
  }

  // Reference/target line: a horizontal (column) or vertical (bar) marker at a
  // fixed value, drawn on top of the bars with its value labelled at the end.
  const rv = opt.referenceValue;
  const refColor = opt.referenceColor || REF_COLOR;
  if (rv != null && isFinite(rv) && !norm100 && rv >= vMin && rv <= vMax) {
    const text = formatNumber(rv, { ...nf, hideZero: false });
    const tw = estTextW(text, 9);
    if (isColumn) {
      const y = yv(rv);
      prims.push({ kind: "line", x1: plotLeft, y1: y, x2: plotLeft + plotW, y2: y, color: refColor, weight: 1.25, meta: { objectType: "valueLine" } });
      prims.push({ kind: "text", x: plotLeft + plotW - tw, y: y - 15, w: tw, h: 14, text, color: refColor, size: 9, bold: true, align: "right", family: fam, meta: { objectType: "valueLine" } });
    } else {
      const x = xv(rv);
      prims.push({ kind: "line", x1: x, y1: plotTop, x2: x, y2: plotTop + plotH, color: refColor, weight: 1.25, meta: { objectType: "valueLine" } });
      prims.push({ kind: "text", x: x - tw / 2, y: plotTop - 15, w: tw, h: 14, text, color: refColor, size: 9, bold: true, align: "center", family: fam, meta: { objectType: "valueLine" } });
    }
  }

  // Secondary axis tick labels on the right (line series scale).
  if (useSecondary && opt.showValueAxis && isColumn) {
    for (const t of axisTicks(lMin, lMax, 4)) {
      const y = plotTop + (lMax - t) * lScale;
      prims.push({ kind: "text", x: plotLeft + plotW + 3, y: y - 7, w: axisBand - 3, h: 14, text: formatNumber(t, { ...nf, hideZero: false }), color: LABEL_DARK, size: 8, bold: false, align: "left", family: fam, meta: { objectType: "valueAxis" } });
    }
  }

  // Line series (combination charts): a polyline across category centers with
  // markers + value labels, drawn on top of the bars. Column orientation only.
  for (const si of lineSeries) {
    const s = data.series[si];
    const pts = order.map((ci, k) => ({
      x: plotLeft + k * slot + slot / 2,
      y: lineY(safe(s.values[ci])),
      raw: safe(s.values[ci]),
      ci,
    }));
    for (let i = 0; i < pts.length - 1; i++) {
      prims.push({ kind: "line", x1: pts[i].x, y1: pts[i].y, x2: pts[i + 1].x, y2: pts[i + 1].y, color: s.color, weight: 1.75, meta: { objectType: "lineSeries", seriesIndex: si } });
    }
    const ms = 5;
    for (const p of pts) {
      prims.push({ kind: "rect", x: p.x - ms / 2, y: p.y - ms / 2, w: ms, h: ms, fill: s.color, meta: { objectType: "lineMarker", seriesIndex: si, categoryIndex: p.ci } });
      if (opt.showValueLabels) {
        const text = formatNumber(p.raw, nf);
        if (text) {
          const w = estTextW(text, opt.segmentFontSize);
          prims.push({ kind: "text", x: p.x - w / 2, y: p.y - 17, w, h: 14, text, color: LABEL_DARK, size: opt.segmentFontSize, bold: false, align: "center", family: fam, meta: { objectType: "segmentLabel", seriesIndex: si, categoryIndex: p.ci } });
        }
      }
    }
  }

  // Difference / CAGR arrows: vertical double-headed arrow between two category
  // levels, with dashed guides to each column top and a labelled delta.
  const clampIdx = (i: number, n: number) => Math.min(n - 1, Math.max(0, Math.round(i)));
  const catValue = (ci: number, mode: "total" | "series", seriesIdx: number) =>
    mode === "total" ? totals[ci] || 0 : safe(data.series[clampIdx(seriesIdx, nSer)].values[ci]);

  // Difference arrow: vertical double-headed arrow between two category levels,
  // with dashed guides to each column top and a labelled delta. The arrow's
  // horizontal position follows diffPos (a slot boundary), or the midpoint (auto).
  function drawSpanArrow(fromCi: number, toCi: number, y1: number, y2: number, label: string) {
    const kFrom = order.indexOf(fromCi);
    const kTo = order.indexOf(toCi);
    if (kFrom < 0 || kTo < 0) return;
    const xF = plotLeft + kFrom * slot + slot / 2;
    const xT = plotLeft + kTo * slot + slot / 2;
    const xA =
      opt.diffPos >= 0
        ? plotLeft + Math.min(nCats, Math.max(0, Math.round(opt.diffPos))) * slot
        : (xF + xT) / 2;
    prims.push({ kind: "line", x1: xF, y1, x2: xA, y2: y1, color: ARROW_COLOR, weight: 0.75, dashed: true, meta: { objectType: "differenceArrow" } });
    prims.push({ kind: "line", x1: xT, y1: y2, x2: xA, y2, color: ARROW_COLOR, weight: 0.75, dashed: true, meta: { objectType: "differenceArrow" } });
    prims.push({ kind: "arrow", x1: xA, y1, x2: xA, y2, color: ARROW_COLOR, weight: 1.25, doubleHeaded: true, meta: { objectType: "differenceArrow" } });
    const w = estTextW(label, 10);
    prims.push({ kind: "text", x: xA + 4, y: (y1 + y2) / 2 - 8, w, h: 16, text: label, color: ARROW_COLOR, size: 10, bold: true, align: "left", family: fam, bg: "#FFFFFF", meta: { objectType: "differenceArrow" } });
  }

  // CAGR arrow: a horizontal arrow above the plot from the "from" to the "to"
  // category, with a white % bubble in the middle (think-cell style).
  function drawCagrArrow(fromCi: number, toCi: number, label: string) {
    const kFrom = order.indexOf(fromCi);
    const kTo = order.indexOf(toCi);
    if (kFrom < 0 || kTo < 0) return;
    const xF = plotLeft + kFrom * slot + slot / 2;
    const xT = plotLeft + kTo * slot + slot / 2;
    const y = plotTop - cagrBand / 2 - 2; // in the reserved band above the plot
    prims.push({ kind: "arrow", x1: xF, y1: y, x2: xT, y2: y, color: ARROW_COLOR, weight: 1.25, meta: { objectType: "cagrArrow" } });
    const w = estTextW(label, 10) + 6;
    prims.push({ kind: "text", x: (xF + xT) / 2 - w / 2, y: y - 8, w, h: 16, text: label, color: ARROW_COLOR, size: 10, bold: true, align: "center", family: fam, bg: "#FFFFFF", meta: { objectType: "cagrArrow" } });
  }

  if (opt.diffArrow !== "off" && isColumn && nCats >= 2) {
    const fi = clampIdx(opt.diffFrom, nCats);
    const ti = clampIdx(opt.diffTo, nCats);
    if (fi !== ti) {
      const v1 = catValue(fi, opt.diffArrow, opt.diffSeries);
      const v2 = catValue(ti, opt.diffArrow, opt.diffSeries);
      const delta = v2 - v1;
      let label = formatNumber(delta, { ...nf, plusSign: true, hideZero: false });
      if (opt.diffPercent && v1 !== 0) {
        const pct = (delta / Math.abs(v1)) * 100;
        label += ` (${pct >= 0 ? "+" : "−"}${Math.abs(Math.round(pct))}%)`;
      }
      drawSpanArrow(fi, ti, yv(v1), yv(v2), label);
    }
  }

  if (opt.cagrArrow !== "off" && isColumn && nCats >= 2) {
    const fi = clampIdx(opt.cagrFrom, nCats);
    const ti = clampIdx(opt.cagrTo, nCats);
    if (fi !== ti) {
      const v1 = catValue(fi, opt.cagrArrow, opt.cagrSeries);
      const v2 = catValue(ti, opt.cagrArrow, opt.cagrSeries);
      const periods = opt.cagrPeriods > 0 ? opt.cagrPeriods : Math.abs(order.indexOf(ti) - order.indexOf(fi));
      drawCagrArrow(fi, ti, cagrLabel(v1, v2, periods));
    }
  }

  if (opt.showLegend) drawLegend();

  return prims;
}

/** CAGR label like "CAGR +12.5%". Returns "n/a" when it can't be computed. */
export function cagrLabel(from: number, to: number, periods: number): string {
  if (periods <= 0 || from <= 0 || to <= 0) return "CAGR n/a";
  const r = (Math.pow(to / from, 1 / periods) - 1) * 100;
  const rounded = Math.round(r * 10) / 10;
  return `CAGR ${rounded >= 0 ? "+" : "−"}${Math.abs(rounded)}%`;
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

/** Nice ticks covering [min, max], always including 0 when the range spans it. */
export function axisTicks(min: number, max: number, target = 4): number[] {
  const span = max - min || 1;
  const raw = span / target;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
  const ticks: number[] = [];
  const start = Math.ceil(min / step - 1e-9) * step;
  for (let t = start; t <= max + step * 1e-9; t += step) ticks.push(Number(t.toFixed(10)));
  return ticks;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

/** Rough width (pt) of a text box sized to its content, with a little padding. */
function estTextW(text: string, size: number): number {
  return Math.max(10, Math.round(text.length * size * 0.62) + 6);
}
