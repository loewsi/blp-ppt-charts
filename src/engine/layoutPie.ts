import type { ChartModel } from "../model/chartModel";
import { DEFAULT_OPTIONS } from "../model/chartModel";
import type { Primitive, ShapeMeta } from "./primitives";
import { formatNumber, formatPercent } from "./format";
import { shadesFrom } from "../util/color";

const LABEL_DARK = "#001C54";
const MAX_FACET = 0.06; // ≈3.4° per facet → a ~105-gon, visually smooth
const LABEL_PAD = 34; // room for the category labels around the circle

/**
 * Pie / doughnut. Office.js has no arc geometry, so each slice is drawn as a fan
 * of thin isosceles-triangle facets whose outer corners sit on the circle — the
 * union reads as a round slice. The first series supplies the slice values; each
 * category is a slice. A doughnut is a white centre ellipse over the fan.
 */
export function layoutPie(model: ChartModel): Primitive[] {
  const opt = model.options ?? DEFAULT_OPTIONS;
  const nf = opt.numberFormat;
  const { data, box } = model;
  const prims: Primitive[] = [];
  const series = data.series[0];
  if (!series) return prims;

  const values = data.categories.map((_, i) => Math.max(0, safe(series.values[i])));
  const total = values.reduce((a, b) => a + b, 0);
  if (total <= 0) return prims;

  // Slice colors are shades of the (single) series color, so changing that swatch
  // — or applying a scheme — recolors the whole pie.
  const palette = shadesFrom(series.color, data.categories.length);
  const cx = box.left + box.width / 2;
  const cy = box.top + box.height / 2;
  const R = Math.max(10, Math.min(box.width, box.height) / 2 - LABEL_PAD);
  const fam = opt.fontFamily;

  const pt = (ang: number, r: number) => ({ x: cx + r * Math.cos(ang), y: cy + r * Math.sin(ang) });

  let a = -Math.PI / 2; // start at the top, sweep clockwise
  data.categories.forEach((cat, ci) => {
    const v = values[ci];
    if (v <= 0) return;
    const sweep = (v / total) * Math.PI * 2;
    const a0 = a;
    const a1 = a + sweep;
    const color = palette[ci % palette.length];

    // Fan of facets across the slice.
    const facets = Math.max(1, Math.ceil(sweep / MAX_FACET));
    const step = sweep / facets;
    const OVERLAP = 0.012; // rad — tiny overlap between facets of THIS slice closes anti-alias hairlines
    for (let f = 0; f < facets; f++) {
      const f0 = a0 + f * step;
      const f1 = f0 + step + (f < facets - 1 ? OVERLAP : 0); // no bleed past the slice's end
      const half = (f1 - f0) / 2;
      const mid = (f0 + f1) / 2;
      const w = 2 * R * Math.sin(half);
      const h = R * Math.cos(half);
      const scx = cx + (h / 2) * Math.cos(mid);
      const scy = cy + (h / 2) * Math.sin(mid);
      const meta: ShapeMeta = { objectType: "slice", categoryIndex: ci };
      prims.push({ kind: "triangle", x: scx - w / 2, y: scy - h / 2, w, h, rotation: (mid * 180) / Math.PI - 90, fill: color, meta });
    }
    a = a1;

    // Percentage inside; category name outside.
    const mid = (a0 + a1) / 2;
    if (opt.showValueLabels) {
      const inside = pt(mid, R * 0.62);
      const pct = formatPercent(v, total, nf.decimals, nf.sep);
      const w = estTextW(pct, opt.segmentFontSize);
      prims.push({ kind: "text", x: inside.x - w / 2, y: inside.y - 7, w, h: 14, text: pct, color: "#FFFFFF", size: opt.segmentFontSize, bold: true, align: "center", family: fam, meta: { objectType: "sliceLabel", categoryIndex: ci } });
    }
    const outside = pt(mid, R + 4);
    const label = `${cat}`;
    const lw = estTextW(label, 9);
    const rightHalf = Math.cos(mid) >= 0;
    prims.push({ kind: "text", x: rightHalf ? outside.x : outside.x - lw, y: outside.y - 7, w: lw, h: 14, text: label, color: LABEL_DARK, size: 9, bold: false, align: rightHalf ? "left" : "right", family: fam, meta: { objectType: "categoryLabel", categoryIndex: ci } });
  });

  // Doughnut hole (drawn over the fan; text still renders last so labels survive).
  if (opt.pieHole > 0) {
    const rHole = Math.min(0.9, opt.pieHole) * R;
    prims.push({ kind: "ellipse", x: cx - rHole, y: cy - rHole, w: rHole * 2, h: rHole * 2, fill: "#FFFFFF", meta: { objectType: "plotArea" } });
    // Total in the middle of a doughnut.
    if (opt.showTotals) {
      const t = formatNumber(total, nf);
      if (t) {
        const w = estTextW(t, opt.totalFontSize);
        prims.push({ kind: "text", x: cx - w / 2, y: cy - 8, w, h: 16, text: t, color: LABEL_DARK, size: opt.totalFontSize, bold: true, align: "center", family: fam, meta: { objectType: "totalLabel" } });
      }
    }
  }

  return prims;
}

function safe(v: number | undefined): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

function estTextW(text: string, size: number): number {
  return Math.max(10, Math.round(text.length * size * 0.62) + 6);
}
