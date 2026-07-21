import { describe, it, expect } from "vitest";
import { layoutPie } from "./layoutPie";
import type { ChartModel, ChartOptions } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";

function model(over: Partial<ChartOptions> = {}, values = [25, 25, 25, 25]): ChartModel {
  return {
    id: "t",
    schemaVersion: 1,
    name: "T",
    box: { left: 0, top: 0, width: 300, height: 300 },
    options: { ...defaultOptions(), ...over },
    data: {
      type: "pie",
      categories: values.map((_, i) => `C${i + 1}`),
      series: [{ name: "S", color: "#2E75FF", values }],
    },
  };
}

const facets = (p: ReturnType<typeof layoutPie>) => p.filter((s) => s.kind === "triangle" && s.meta?.objectType === "slice");

describe("layoutPie", () => {
  it("emits triangle facets and one category label per slice", () => {
    const prims = layoutPie(model());
    expect(facets(prims).length).toBeGreaterThanOrEqual(4 * 6); // each 90° slice → ≥6 facets at 15°
    expect(prims.filter((s) => s.meta?.objectType === "categoryLabel").length).toBe(4);
  });

  it("has every facet's apex converge on the pie centre (roundness invariant)", () => {
    // Pie centre for a 300×300 box.
    const cx = 150;
    const cy = 150;
    for (const f of facets(layoutPie(model()))) {
      if (f.kind !== "triangle") continue;
      // Apex = top-centre of the box, rotated about the box centre (clockwise, y-down).
      const ox = f.x + f.w / 2;
      const oy = f.y + f.h / 2;
      const th = (f.rotation * Math.PI) / 180;
      const s = Math.sin(th);
      const c = Math.cos(th);
      const apexX = ox + (s * f.h) / 2;
      const apexY = oy - (c * f.h) / 2;
      expect(Math.hypot(apexX - cx, apexY - cy)).toBeLessThan(1); // all facets meet at the centre
    }
  });

  it("keeps every facet corner within the bounding box", () => {
    const prims = layoutPie(model());
    for (const f of facets(prims)) {
      if (f.kind !== "triangle") continue;
      expect(f.x).toBeGreaterThanOrEqual(-1);
      expect(f.y).toBeGreaterThanOrEqual(-1);
      expect(f.x + f.w).toBeLessThanOrEqual(301);
      expect(f.y + f.h).toBeLessThanOrEqual(301);
    }
  });

  it("labels slices with their percentage share", () => {
    const labels = layoutPie(model()).filter((s) => s.kind === "text" && s.meta?.objectType === "sliceLabel");
    expect(labels.length).toBe(4);
    expect(labels.every((l) => l.kind === "text" && l.text === "25%")).toBe(true);
  });

  it("adds a doughnut hole + centre total when pieHole > 0", () => {
    const prims = layoutPie(model({ pieHole: 0.5, showTotals: true }));
    expect(prims.some((s) => s.kind === "ellipse")).toBe(true);
    expect(prims.some((s) => s.kind === "text" && s.meta?.objectType === "totalLabel" && s.text === "100")).toBe(true);
  });

  it("returns nothing for an empty / all-zero series", () => {
    expect(layoutPie(model({}, [0, 0, 0, 0]))).toEqual([]);
  });
});
