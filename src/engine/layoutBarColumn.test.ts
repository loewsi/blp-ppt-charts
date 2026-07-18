import { describe, it, expect } from "vitest";
import { layoutBarColumn } from "./layoutBarColumn";
import type { RectPrimitive } from "./primitives";
import type { ChartModel, ChartOptions } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";

function model(over: Partial<ChartOptions>, data?: ChartModel["data"]): ChartModel {
  return {
    id: "t",
    schemaVersion: 1,
    name: "T",
    box: { left: 0, top: 0, width: 400, height: 300 },
    options: { ...defaultOptions(), ...over },
    data:
      data ??
      {
        type: "barColumn",
        categories: ["A", "B"],
        series: [
          { name: "S1", color: "#111111", values: [10, 20] },
          { name: "S2", color: "#222222", values: [30, 40] },
        ],
      },
  };
}

const rects = (p: ReturnType<typeof layoutBarColumn>) =>
  p.filter((s): s is RectPrimitive => s.kind === "rect");
const segRects = (p: ReturnType<typeof layoutBarColumn>) =>
  rects(p).filter((r) => r.meta?.objectType === "segment");

describe("layoutBarColumn — stacked column", () => {
  const prims = layoutBarColumn(model({ orientation: "column", grouping: "stacked" }));

  it("emits one segment per positive value", () => {
    expect(segRects(prims).length).toBe(4);
  });

  it("tags every segment with series/category indices", () => {
    for (const r of segRects(prims)) {
      expect(typeof r.meta?.seriesIndex).toBe("number");
      expect(typeof r.meta?.categoryIndex).toBe("number");
    }
  });

  it("draws a baseline", () => {
    expect(prims.some((s) => s.kind === "line" && s.meta?.objectType === "baseline")).toBe(true);
  });

  it("adds one total label per category", () => {
    expect(prims.filter((s) => s.meta?.objectType === "totalLabel").length).toBe(2);
  });

  it("adds one category label per category", () => {
    expect(prims.filter((s) => s.meta?.objectType === "categoryLabel").length).toBe(2);
  });

  it("scales segment height with value", () => {
    const cat0 = segRects(prims).filter((r) => r.meta?.categoryIndex === 0);
    const s1 = cat0.find((r) => r.meta?.seriesIndex === 0)!; // value 10
    const s2 = cat0.find((r) => r.meta?.seriesIndex === 1)!; // value 30
    expect(s2.h).toBeGreaterThan(s1.h);
  });
});

describe("layoutBarColumn — 100% stacked", () => {
  it("normalizes every category to the same stack height", () => {
    const prims = layoutBarColumn(model({ grouping: "stacked100" }));
    const heightFor = (ci: number) =>
      segRects(prims)
        .filter((r) => r.meta?.categoryIndex === ci)
        .reduce((sum, r) => sum + r.h, 0);
    expect(Math.abs(heightFor(0) - heightFor(1))).toBeLessThan(0.5);
  });
});

describe("layoutBarColumn — clustered", () => {
  it("has no total labels", () => {
    const prims = layoutBarColumn(model({ grouping: "clustered" }));
    expect(prims.filter((s) => s.meta?.objectType === "totalLabel").length).toBe(0);
  });
});

describe("layoutBarColumn — zero values", () => {
  it("produces no segment for a zero value", () => {
    const prims = layoutBarColumn(
      model(
        { grouping: "stacked" },
        {
          type: "barColumn",
          categories: ["A"],
          series: [
            { name: "S1", color: "#111", values: [0] },
            { name: "S2", color: "#222", values: [5] },
          ],
        }
      )
    );
    expect(segRects(prims).length).toBe(1);
  });
});

describe("layoutBarColumn — bar orientation", () => {
  it("draws a vertical baseline", () => {
    const prims = layoutBarColumn(model({ orientation: "bar", grouping: "stacked" }));
    const baseline = prims.find((s) => s.kind === "line" && s.meta?.objectType === "baseline");
    expect(baseline && baseline.kind === "line" && Math.abs(baseline.x1 - baseline.x2) < 0.5).toBe(true);
  });
});
