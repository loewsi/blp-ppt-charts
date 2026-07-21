import type { ChartModel } from "../model/chartModel";
import { DEFAULT_OPTIONS, PALETTES } from "../model/chartModel";
import type { Primitive } from "./primitives";
import { formatNumber } from "./format";
import { axisTicks } from "./layoutBarColumn";

const AXIS_COLOR = "#001C54";
const GRID_COLOR = "#D7E2F4";
const LABEL_DARK = "#001C54";
const PAD = 8;
const Y_BAND = 34; // room for y-axis tick labels (left)
const X_BAND = 18; // room for x-axis tick labels (bottom)

/**
 * Scatter / bubble (think-cell style). Grid convention: each ROW is a point; the
 * columns are X, Y, Size (bubble, optional), Group (color, optional). A point's
 * row name is its label; its color comes from its Group (if given) or its swatch.
 */
export function layoutScatter(model: ChartModel): Primitive[] {
  const opt = model.options ?? DEFAULT_OPTIONS;
  const nf = opt.numberFormat;
  const { data, box } = model;
  const prims: Primitive[] = [];
  const points = data.series; // each row/series is one point
  if (points.length === 0 || data.categories.length < 2) return prims; // need at least X + Y columns

  const hasSize = data.categories.length >= 3;
  const hasGroup = data.categories.length >= 4;
  const xs = points.map((p) => safe(p.values[0]));
  const ys = points.map((p) => safe(p.values[1]));
  const sizes = hasSize ? points.map((p) => Math.max(0, safe(p.values[2]))) : null;
  const fam = opt.fontFamily;

  // Color: by Group when a Group column is present (points sharing a group share a
  // color), else each point's own swatch color.
  const palette = PALETTES.blue;
  const groupColorFor = (i: number): string => {
    if (!hasGroup) return points[i].color;
    const g = safe(points[i].values[3]);
    return palette[(Math.round(g) % palette.length + palette.length) % palette.length];
  };

  const plotLeft = box.left + PAD + Y_BAND;
  const plotTop = box.top + PAD;
  const plotW = Math.max(24, box.width - PAD * 2 - Y_BAND);
  const plotH = Math.max(24, box.height - PAD * 2 - X_BAND);

  const range = (arr: number[]) => {
    let lo = Math.min(...arr);
    let hi = Math.max(...arr);
    if (!isFinite(lo) || !isFinite(hi)) {
      lo = 0;
      hi = 1;
    }
    if (hi === lo) {
      hi += 1;
      lo -= 1;
    }
    const pad = (hi - lo) * 0.08;
    return { lo: lo - pad, hi: hi + pad };
  };
  const xr = range(xs);
  const yr = range(ys);
  const xPix = (v: number) => plotLeft + ((v - xr.lo) / (xr.hi - xr.lo)) * plotW;
  const yPix = (v: number) => plotTop + ((yr.hi - v) / (yr.hi - yr.lo)) * plotH;

  // Gridlines + axis tick labels.
  for (const t of axisTicks(xr.lo, xr.hi, 4)) {
    const x = xPix(t);
    if (x < plotLeft - 0.5 || x > plotLeft + plotW + 0.5) continue;
    if (opt.showGridlines) prims.push({ kind: "line", x1: x, y1: plotTop, x2: x, y2: plotTop + plotH, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
    if (opt.showValueAxis) prims.push({ kind: "text", x: x - 22, y: plotTop + plotH + 2, w: 44, h: 14, text: formatNumber(t, { ...nf, hideZero: false }), color: LABEL_DARK, size: 8, bold: false, align: "center", family: fam, meta: { objectType: "valueAxis" } });
  }
  for (const t of axisTicks(yr.lo, yr.hi, 4)) {
    const y = yPix(t);
    if (y < plotTop - 0.5 || y > plotTop + plotH + 0.5) continue;
    if (opt.showGridlines) prims.push({ kind: "line", x1: plotLeft, y1: y, x2: plotLeft + plotW, y2: y, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
    if (opt.showValueAxis) prims.push({ kind: "text", x: box.left, y: y - 7, w: Y_BAND - 3, h: 14, text: formatNumber(t, { ...nf, hideZero: false }), color: LABEL_DARK, size: 8, bold: false, align: "right", family: fam, meta: { objectType: "valueAxis" } });
  }

  // Quadrant divider lines (at zero if in range, else the midpoint).
  if (opt.scatterQuadrant) {
    const xDiv = xr.lo <= 0 && xr.hi >= 0 ? 0 : (xr.lo + xr.hi) / 2;
    const yDiv = yr.lo <= 0 && yr.hi >= 0 ? 0 : (yr.lo + yr.hi) / 2;
    prims.push({ kind: "line", x1: xPix(xDiv), y1: plotTop, x2: xPix(xDiv), y2: plotTop + plotH, color: AXIS_COLOR, weight: 0.75, dashed: true, meta: { objectType: "quadrant" } });
    prims.push({ kind: "line", x1: plotLeft, y1: yPix(yDiv), x2: plotLeft + plotW, y2: yPix(yDiv), color: AXIS_COLOR, weight: 0.75, dashed: true, meta: { objectType: "quadrant" } });
  }

  // Axes (left + bottom) — optional.
  if (opt.scatterAxes) {
    prims.push({ kind: "line", x1: plotLeft, y1: plotTop, x2: plotLeft, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: { objectType: "valueAxis" } });
    prims.push({ kind: "line", x1: plotLeft, y1: plotTop + plotH, x2: plotLeft + plotW, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: { objectType: "baseline" } });
  }

  // Points (bubbles sized by area ∝ size).
  const maxSize = sizes ? Math.max(1, ...sizes) : 1;
  points.forEach((p, i) => {
    const cx = xPix(xs[i]);
    const cy = yPix(ys[i]);
    const r = sizes ? 5 + 16 * Math.sqrt(sizes[i] / maxSize) : 4.5;
    prims.push({ kind: "ellipse", x: cx - r, y: cy - r, w: r * 2, h: r * 2, fill: groupColorFor(i), meta: { objectType: "point", seriesIndex: i } });
    if (opt.showValueLabels && p.name) {
      const w = estTextW(p.name, opt.segmentFontSize);
      prims.push({ kind: "text", x: cx + r + 2, y: cy - 7, w, h: 14, text: p.name, color: LABEL_DARK, size: opt.segmentFontSize, bold: false, align: "left", family: fam, meta: { objectType: "point", seriesIndex: i } });
    }
  });

  return prims;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

function estTextW(text: string, size: number): number {
  return Math.max(10, Math.round(text.length * size * 0.62) + 6);
}
