import type { ChartModel } from "../model/chartModel";
import type { Primitive } from "./primitives";

const AXIS_COLOR = "#001C54";
const LABEL_COLOR = "#FFFFFF";
const TOTAL_COLOR = "#001C54";
const CAT_COLOR = "#001C54";

// Insets inside the chart box, in points.
const PAD_TOP = 22; // room for total labels above the tallest column
const PAD_BOTTOM = 20; // room for category labels below the baseline
const PAD_SIDE = 4;

const GAP_FRACTION = 0.35; // portion of each column slot left empty as gap
const MIN_SEGMENT_FOR_LABEL = 12; // hide labels on segments shorter than this (pt)

/** Pure geometry: ChartModel -> primitives. No Office.js here (easy to unit test). */
export function layoutStackedColumn(model: ChartModel): Primitive[] {
  const { box, data } = model;
  const prims: Primitive[] = [];
  const n = data.categories.length;
  if (n === 0 || data.series.length === 0) return prims;

  const plotLeft = box.left + PAD_SIDE;
  const plotTop = box.top + PAD_TOP;
  const plotWidth = box.width - PAD_SIDE * 2;
  const plotHeight = box.height - PAD_TOP - PAD_BOTTOM;
  const baseline = plotTop + plotHeight;

  const totals = data.categories.map((_, ci) =>
    data.series.reduce((sum, s) => sum + safe(s.values[ci]), 0)
  );
  const maxTotal = Math.max(1, ...totals);
  const yScale = plotHeight / maxTotal;

  const slot = plotWidth / n;
  const barW = slot * (1 - GAP_FRACTION);

  for (let ci = 0; ci < n; ci++) {
    const slotX = plotLeft + ci * slot;
    const barX = slotX + (slot - barW) / 2;
    let yCursor = baseline; // stack upward from the baseline

    for (const s of data.series) {
      const v = safe(s.values[ci]);
      const h = v * yScale;
      if (h <= 0) continue;
      const y = yCursor - h;
      prims.push({ kind: "rect", x: barX, y, w: barW, h, fill: s.color });
      if (h >= MIN_SEGMENT_FOR_LABEL) {
        prims.push({
          kind: "text",
          x: barX,
          y: y + h / 2 - 7,
          w: barW,
          h: 14,
          text: fmt(v),
          color: LABEL_COLOR,
          size: 9,
          bold: false,
          align: "center",
        });
      }
      yCursor = y;
    }

    // Total above the column.
    prims.push({
      kind: "text",
      x: barX - 8,
      y: baseline - totals[ci] * yScale - 18,
      w: barW + 16,
      h: 16,
      text: fmt(totals[ci]),
      color: TOTAL_COLOR,
      size: 10,
      bold: true,
      align: "center",
    });

    // Category label below the baseline.
    prims.push({
      kind: "text",
      x: slotX,
      y: baseline + 3,
      w: slot,
      h: 16,
      text: data.categories[ci],
      color: CAT_COLOR,
      size: 9,
      bold: false,
      align: "center",
    });
  }

  // Baseline axis.
  prims.push({
    kind: "line",
    x1: plotLeft,
    y1: baseline,
    x2: plotLeft + plotWidth,
    y2: baseline,
    color: AXIS_COLOR,
    weight: 1,
  });

  return prims;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

function fmt(v: number): string {
  if (Math.abs(v) >= 1000) {
    return (v / 1000).toLocaleString(undefined, { maximumFractionDigits: 1 }) + "k";
  }
  return v.toLocaleString(undefined, { maximumFractionDigits: 1 });
}
