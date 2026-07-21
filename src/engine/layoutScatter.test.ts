import { describe, it, expect } from "vitest";
import { layoutScatter } from "./layoutScatter";
import type { ChartModel, ChartOptions } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";

// Rows = points; columns = X, Y, [Size], [Group].
function model(over: Partial<ChartOptions> = {}, cols: string[] = ["X", "Y"]): ChartModel {
  return {
    id: "t",
    schemaVersion: 1,
    name: "T",
    box: { left: 0, top: 0, width: 400, height: 300 },
    options: { ...defaultOptions(), ...over },
    data: {
      type: "scatter",
      categories: cols,
      series: [
        { name: "P1", color: "#111111", values: [1, 10, 5, 0] },
        { name: "P2", color: "#222222", values: [2, 20, 10, 1] },
        { name: "P3", color: "#333333", values: [3, 30, 40, 0] },
      ],
    },
  };
}

const points = (p: ReturnType<typeof layoutScatter>) => p.filter((s) => s.kind === "ellipse" && s.meta?.objectType === "point");

describe("layoutScatter (rows = points)", () => {
  it("draws one point per row", () => {
    expect(points(layoutScatter(model())).length).toBe(3);
  });

  it("positions higher X further right and higher Y higher up", () => {
    const pts = points(layoutScatter(model())) as Array<{ x: number; y: number; meta?: { seriesIndex?: number } }>;
    const byPt = (i: number) => pts.find((p) => p.meta?.seriesIndex === i)!;
    expect(byPt(2).x).toBeGreaterThan(byPt(0).x); // X: 3 > 1
    expect(byPt(2).y).toBeLessThan(byPt(0).y); // Y: 30 higher on screen
  });

  it("bubble sizes scale with the Size column when present", () => {
    const pts = points(layoutScatter(model({}, ["X", "Y", "Size"]))) as Array<{ w: number; meta?: { seriesIndex?: number } }>;
    const r = (i: number) => pts.find((p) => p.meta?.seriesIndex === i)!.w;
    expect(r(2)).toBeGreaterThan(r(0)); // size 40 > 5
  });

  it("colors points by Group when a Group column is present", () => {
    const pts = points(layoutScatter(model({}, ["X", "Y", "Size", "Group"]))) as Array<{ fill: string; meta?: { seriesIndex?: number } }>;
    const fill = (i: number) => pts.find((p) => p.meta?.seriesIndex === i)!.fill;
    expect(fill(0)).toBe(fill(2)); // groups 0 and 0 → same color
    expect(fill(0)).not.toBe(fill(1)); // group 1 differs
  });

  it("adds quadrant lines only when enabled", () => {
    expect(layoutScatter(model({ scatterQuadrant: true })).some((s) => s.meta?.objectType === "quadrant")).toBe(true);
    expect(layoutScatter(model({})).some((s) => s.meta?.objectType === "quadrant")).toBe(false);
  });

  it("returns nothing without at least X and Y columns", () => {
    expect(layoutScatter(model({}, ["X"]))).toEqual([]);
  });
});
