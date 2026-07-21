import type { ChartModel } from "../model/chartModel";
import { DEFAULT_OPTIONS } from "../model/chartModel";
import type { Primitive, ShapeMeta } from "./primitives";
import { formatNumber } from "./format";

const AXIS_COLOR = "#001C54";
const CONNECTOR_COLOR = "#9AA6BF";
const LABEL_LIGHT = "#FFFFFF";
const LABEL_DARK = "#001C54";
const DEFAULT_FALL = "#E8412C";
const TOTAL_COLOR = "#001C54"; // computed "e" total/subtotal columns

const PAD_TOP = 22;
const PAD_BOTTOM = 26;
const PAD_SIDE = 8;
const MIN_BAR_FOR_LABEL = 12;

/**
 * Waterfall: one series of deltas rendered as floating rise/fall bars with a
 * running total, connectors between steps, and a zero baseline. Reuses the
 * shared number-format and font options. Column orientation (v1).
 */
export function layoutWaterfall(model: ChartModel): Primitive[] {
  const opt = model.options ?? DEFAULT_OPTIONS;
  const nf = opt.numberFormat;
  const fam = opt.fontFamily;
  const { data, box } = model;
  const prims: Primitive[] = [];
  const n = data.categories.length;
  const series0 = data.series[0];
  if (n === 0 || !series0) return prims;

  const nSer = data.series.length;
  const single = nSer === 1;
  const isTotal = (i: number) => data.totalFlags?.[i] ?? false;
  const riseColor = series0.color; // single-series rise/fall coloring
  const fallColor = data.series[1]?.color ?? DEFAULT_FALL;

  // Each step's delta is the SUM of its series; within the step the series stack
  // as sub-segments. subPos[i] holds the running sub-positions [before … after].
  // An "e" total column shows the running sum from the baseline and doesn't flow.
  const before: number[] = [];
  const after: number[] = [];
  const subPos: number[][] = [];
  let cum = 0;
  for (let i = 0; i < n; i++) {
    before.push(cum);
    if (isTotal(i)) {
      subPos.push([]);
      after.push(cum);
    } else {
      const ps = [cum];
      let c = cum;
      for (let s = 0; s < nSer; s++) {
        c += safe(data.series[s].values[i]);
        ps.push(c);
      }
      cum = c;
      subPos.push(ps);
      after.push(cum);
    }
  }

  const plotLeft = box.left + PAD_SIDE;
  const plotTop = box.top + PAD_TOP;
  const plotW = Math.max(24, box.width - PAD_SIDE * 2);
  const plotH = Math.max(24, box.height - PAD_TOP - PAD_BOTTOM);

  // Range spans every sub-position (so mixed-sign steps that overshoot still fit).
  const allV = [0, ...before, ...after, ...subPos.flat()];
  const hi = Math.max(...allV);
  const lo = Math.min(...allV);
  const range = hi - lo || 1;
  const scale = plotH / range;
  const yFor = (v: number) => plotTop + (hi - v) * scale;

  const slot = plotW / n;
  const gap = opt.gap;
  const barW = slot * (1 - gap);

  // Draw one sub-segment rect + its label.
  function drawSeg(x: number, a: number, b: number, fill: string, text: string, si: number, ci: number) {
    const yTop = yFor(Math.max(a, b));
    const h = Math.max(0, yFor(Math.min(a, b)) - yTop);
    prims.push({ kind: "rect", x, y: yTop, w: barW, h, fill, stroke: "#FFFFFF", meta: { objectType: "segment", seriesIndex: si, categoryIndex: ci } });
    if (opt.showValueLabels && text) {
      const cx = x + barW / 2;
      const lw = estTextW(text, opt.segmentFontSize);
      const m: ShapeMeta = { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci };
      if (h >= MIN_BAR_FOR_LABEL) {
        prims.push({ kind: "text", x: cx - lw / 2, y: yTop + h / 2 - 7, w: lw, h: 14, text, color: LABEL_LIGHT, size: opt.segmentFontSize, bold: false, align: "center", family: fam, meta: m });
      } else {
        prims.push({ kind: "text", x: cx - lw / 2, y: yTop - 16, w: lw, h: 14, text, color: LABEL_DARK, size: opt.segmentFontSize, bold: false, align: "center", family: fam, meta: m });
      }
    }
  }

  for (let i = 0; i < n; i++) {
    const x = plotLeft + i * slot + (slot - barW) / 2;
    if (isTotal(i)) {
      // "e" column: a single bar from the baseline to the running sum.
      drawSeg(x, 0, after[i], TOTAL_COLOR, formatNumber(after[i], { ...nf, hideZero: false }), 0, i);
    } else {
      const ps = subPos[i];
      for (let s = 0; s < nSer; s++) {
        const val = safe(data.series[s].values[i]);
        if (val === 0) continue;
        const fill = single ? (val >= 0 ? riseColor : fallColor) : data.series[s].color;
        // Respect the number-format flags: a leading "+" only appears when plusSign is on.
        drawSeg(x, ps[s], ps[s + 1], fill, formatNumber(val, nf), s, i);
      }
    }

    // Category label under the baseline.
    prims.push({ kind: "text", x: plotLeft + i * slot, y: plotTop + plotH + 3, w: slot, h: 16, text: data.categories[i], color: LABEL_DARK, size: 9, bold: false, align: "center", family: fam, meta: { objectType: "categoryLabel", categoryIndex: i } });

    // Connector from this step's end level to the next step's start.
    if (i < n - 1) {
      const y = yFor(after[i]);
      const xEnd = x + barW;
      const xNext = plotLeft + (i + 1) * slot + (slot - barW) / 2;
      prims.push({ kind: "line", x1: xEnd, y1: y, x2: xNext, y2: y, color: CONNECTOR_COLOR, weight: 1, meta: { objectType: "connector", categoryIndex: i } });
    }
  }

  // Zero baseline.
  const y0 = yFor(0);
  prims.push({ kind: "line", x1: plotLeft, y1: y0, x2: plotLeft + plotW, y2: y0, color: AXIS_COLOR, weight: 1, meta: { objectType: "baseline" } });

  return prims;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

function estTextW(text: string, size: number): number {
  return Math.max(10, Math.round(text.length * size * 0.62) + 6);
}
