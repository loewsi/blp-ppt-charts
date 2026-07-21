import type { ChartModel } from "../model/chartModel";
import { DEFAULT_OPTIONS } from "../model/chartModel";
import type { Primitive, ShapeMeta } from "./primitives";
import { formatNumber, segmentLabel } from "./format";

const AXIS_COLOR = "#001C54";
const GRID_COLOR = "#D7E2F4";
const LABEL_LIGHT = "#FFFFFF";
const LABEL_DARK = "#001C54";
const COL_GAP = 0; // mekko columns are contiguous (no gap)
const MIN_SEG = 12; // min segment height (pt) to carry a % label

/**
 * Mekko / Marimekko: columns whose WIDTH is proportional to the category total,
 * each column 100%-stacked by its series. Column total sits above; category name
 * below; segment share (%) inside.
 */
export function layoutMekko(model: ChartModel): Primitive[] {
  const opt = model.options ?? DEFAULT_OPTIONS;
  const nf = opt.numberFormat;
  const { data, box } = model;
  const prims: Primitive[] = [];
  const nCats = data.categories.length;
  const nSer = data.series.length;
  if (nCats === 0 || nSer === 0) return prims;

  const order = data.categories.map((_, i) => i);
  if (opt.reverseCategories) order.reverse();
  const serOrder = data.series.map((_, i) => i);
  if (opt.reverseSeries) serOrder.reverse();

  const totals = data.categories.map((_, ci) =>
    data.series.reduce((s, se) => s + Math.max(0, safe(se.values[ci])), 0)
  );
  const grand = totals.reduce((a, b) => a + b, 0);
  if (grand <= 0) return prims;

  const fam = opt.fontFamily;
  const axisBand = opt.showValueAxis ? 30 : 0;
  const insLeft = 6 + axisBand;
  const insTop = 20; // column totals
  const insBottom = 20; // category labels
  const insRight = 6;
  const plotLeft = box.left + insLeft;
  const plotTop = box.top + insTop;
  const plotW = Math.max(24, box.width - insLeft - insRight);
  const plotH = Math.max(24, box.height - insTop - insBottom);

  const yShare = (share: number) => plotTop + (1 - share) * plotH; // share 0..1 from the bottom

  // Horizontal % gridlines / axis labels.
  if (opt.showGridlines || opt.showValueAxis) {
    for (const pct of [0, 0.25, 0.5, 0.75, 1]) {
      const y = yShare(pct);
      if (opt.showGridlines && pct > 0 && pct < 1) {
        prims.push({ kind: "line", x1: plotLeft, y1: y, x2: plotLeft + plotW, y2: y, color: GRID_COLOR, weight: 0.75, meta: { objectType: "gridline" } });
      }
      if (opt.showValueAxis) {
        prims.push({ kind: "text", x: box.left, y: y - 7, w: axisBand - 3, h: 14, text: `${pct * 100}%`, color: LABEL_DARK, size: 8, bold: false, align: "right", family: fam, meta: { objectType: "valueAxis" } });
      }
    }
  }

  const availW = plotW - COL_GAP * Math.max(0, nCats - 1);
  let x = plotLeft;
  for (let k = 0; k < nCats; k++) {
    const ci = order[k];
    const total = totals[ci];
    const colW = availW * (total / grand);
    if (total > 0 && colW > 0.5) {
      let cum = 0; // cumulative share from the bottom
      for (let sidx = 0; sidx < nSer; sidx++) {
        const si = serOrder[sidx];
        const val = Math.max(0, safe(data.series[si].values[ci]));
        if (val === 0) continue;
        const share = val / total;
        const yTop = yShare(cum + share);
        const h = yShare(cum) - yTop;
        cum += share;
        const meta: ShapeMeta = { objectType: "segment", seriesIndex: si, categoryIndex: ci };
        prims.push({ kind: "rect", x, y: yTop, w: colW, h, fill: data.series[si].color, stroke: "#FFFFFF", meta });
        if (opt.showValueLabels && h >= MIN_SEG && colW >= 16) {
          const text = segmentLabel(val, total, opt.labelMode, nf);
          const w = estTextW(text, opt.segmentFontSize);
          prims.push({ kind: "text", x: x + colW / 2 - w / 2, y: yTop + h / 2 - 7, w, h: 14, text, color: LABEL_LIGHT, size: opt.segmentFontSize, bold: false, align: "center", family: fam, meta: { objectType: "segmentLabel", seriesIndex: si, categoryIndex: ci } });
        }
      }
      // Column total above, category name below.
      if (opt.showTotals) {
        const t = formatNumber(total, nf);
        if (t) prims.push({ kind: "text", x: x - 4, y: plotTop - 17, w: colW + 8, h: 16, text: t, color: LABEL_DARK, size: opt.totalFontSize, bold: true, align: "center", family: fam, meta: { objectType: "totalLabel", categoryIndex: ci } });
      }
      prims.push({ kind: "text", x: x - 4, y: plotTop + plotH + 3, w: colW + 8, h: 16, text: data.categories[ci], color: LABEL_DARK, size: 9, bold: false, align: "center", family: fam, meta: { objectType: "categoryLabel", categoryIndex: ci } });
    }
    x += colW + COL_GAP;
  }

  // Baseline.
  prims.push({ kind: "line", x1: plotLeft, y1: plotTop + plotH, x2: plotLeft + plotW, y2: plotTop + plotH, color: AXIS_COLOR, weight: 1, meta: { objectType: "baseline" } });

  return prims;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

function estTextW(text: string, size: number): number {
  return Math.max(10, Math.round(text.length * size * 0.62) + 6);
}
