import { describe, it, expect } from "vitest";
import { layoutScatter } from "./layoutScatter";
import type { ChartModel, ChartOptions } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";

function model(over: Partial<ChartOptions> = {}, withSize = false): ChartModel {
  const series = [
    { name: "X", color: "#2E75FF", values: [1, 2, 3] },
    { name: "Y", color: "#001C54", values: [10, 20, 30] },
  ];
  if (withSize) series.push({ name: "Size", color: "#999", values: [5, 10, 40] });
  return {
    id: "t",
    schemaVersion: 1,
    name: "T",
    box: { left: 0, top: 0, width: 400, height: 300 },
    options: { ...defaultOptions(), ...over },
    data: { type: "scatter", categories: ["A", "B", "C"], series },
  };
}

const points = (p: ReturnType<typeof layoutScatter>) => p.filter((s) => s.kind === "ellipse" && s.meta?.objectType === "point");

describe("layoutScatter", () => {
  it("draws one point per category", () => {
    expect(points(layoutScatter(model())).length).toBe(3);
  });

  it("positions higher X further right and higher Y higher up", () => {
    const pts = points(layoutScatter(model())) as Array<{ x: number; y: number; w: number; meta?: { categoryIndex?: number } }>;
    const byCat = (i: number) => pts.find((p) => p.meta?.categoryIndex === i)!;
    expect(byCat(2).x).toBeGreaterThan(byCat(0).x); // X: 3 > 1
    expect(byCat(2).y).toBeLessThan(byCat(0).y); // Y: 30 higher on screen (smaller y)
  });

  it("bubble sizes scale with the third series", () => {
    const pts = points(layoutScatter(model({}, true))) as Array<{ w: number; meta?: { categoryIndex?: number } }>;
    const r = (i: number) => pts.find((p) => p.meta?.categoryIndex === i)!.w;
    expect(r(2)).toBeGreaterThan(r(0)); // size 40 > 5
  });

  it("adds quadrant lines only when enabled", () => {
    expect(layoutScatter(model({ scatterQuadrant: true })).some((s) => s.meta?.objectType === "quadrant")).toBe(true);
    expect(layoutScatter(model({})).some((s) => s.meta?.objectType === "quadrant")).toBe(false);
  });

  it("returns nothing without both X and Y series", () => {
    const m = model();
    m.data.series = [m.data.series[0]];
    expect(layoutScatter(m)).toEqual([]);
  });
});
