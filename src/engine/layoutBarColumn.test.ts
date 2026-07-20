import { describe, it, expect } from "vitest";
import { layoutBarColumn, niceTicks } from "./layoutBarColumn";
import type { RectPrimitive, TextPrimitive } from "./primitives";
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

describe("layoutBarColumn — furniture (opt-in)", () => {
  it("defaults draw no legend/gridlines/axis", () => {
    const prims = layoutBarColumn(model({}));
    expect(prims.some((s) => s.meta?.objectType === "legend")).toBe(false);
    expect(prims.some((s) => s.meta?.objectType === "gridline")).toBe(false);
    expect(prims.some((s) => s.meta?.objectType === "valueAxis")).toBe(false);
  });

  it("legend adds one entry + label per series", () => {
    const prims = layoutBarColumn(model({ showLegend: true }));
    expect(prims.filter((s) => s.meta?.objectType === "legendEntry").length).toBe(2);
    expect(prims.filter((s) => s.meta?.objectType === "legend").length).toBe(2);
  });

  it("gridlines and axis labels appear when enabled", () => {
    const prims = layoutBarColumn(model({ showGridlines: true, showValueAxis: true }));
    expect(prims.some((s) => s.meta?.objectType === "gridline")).toBe(true);
    expect(prims.some((s) => s.meta?.objectType === "valueAxis")).toBe(true);
  });

  it("draws connectors between stacked columns only when enabled", () => {
    const on = layoutBarColumn(model({ showConnectors: true }));
    const off = layoutBarColumn(model({ showConnectors: false }));
    expect(on.some((s) => s.kind === "line" && s.meta?.objectType === "connector")).toBe(true);
    expect(off.some((s) => s.kind === "line" && s.meta?.objectType === "connector")).toBe(false);
  });
});

describe("layoutBarColumn — small labels & fonts", () => {
  const small = {
    type: "barColumn" as const,
    categories: ["A"],
    series: [
      { name: "S1", color: "#111", values: [98] },
      { name: "S2", color: "#222", values: [2] }, // tiny segment
    ],
  };
  const segLabels = (p: ReturnType<typeof layoutBarColumn>) =>
    p.filter((s): s is TextPrimitive => s.kind === "text" && s.meta?.objectType === "segmentLabel");

  it("keeps a small label inside as a segment-colored chip", () => {
    const labels = segLabels(layoutBarColumn(model({ labelOverflow: "inside" }, small)));
    expect(labels.length).toBe(2);
    expect(labels.find((l) => l.bg)?.bg).toBe("#222"); // small S2 gets a chip in its color
  });

  it("moves a small label outside (no chip) when labelOverflow=outside", () => {
    const labels = segLabels(layoutBarColumn(model({ labelOverflow: "outside" }, small)));
    expect(labels.length).toBe(2);
    expect(labels.some((l) => l.bg)).toBe(false);
  });

  it("applies font family and per-type sizes", () => {
    const prims = layoutBarColumn(model({ fontFamily: "Arial", segmentFontSize: 11, totalFontSize: 14 }));
    const seg = prims.find((s): s is TextPrimitive => s.kind === "text" && s.meta?.objectType === "segmentLabel");
    const tot = prims.find((s): s is TextPrimitive => s.kind === "text" && s.meta?.objectType === "totalLabel");
    expect(seg?.family).toBe("Arial");
    expect(seg?.size).toBe(11);
    expect(tot?.size).toBe(14);
  });
});

describe("niceTicks", () => {
  it("returns [0] for non-positive max", () => {
    expect(niceTicks(0)).toEqual([0]);
    expect(niceTicks(-5)).toEqual([0]);
  });

  it("starts at 0 and covers the max", () => {
    const ticks = niceTicks(40, 4);
    expect(ticks[0]).toBe(0);
    expect(ticks[ticks.length - 1]).toBeGreaterThanOrEqual(40);
  });

  it("is strictly increasing", () => {
    const ticks = niceTicks(37, 4);
    for (let i = 1; i < ticks.length; i++) expect(ticks[i]).toBeGreaterThan(ticks[i - 1]);
  });
});

describe("layoutBarColumn — bar orientation", () => {
  it("draws a vertical baseline", () => {
    const prims = layoutBarColumn(model({ orientation: "bar", grouping: "stacked" }));
    const baseline = prims.find((s) => s.kind === "line" && s.meta?.objectType === "baseline");
    expect(baseline && baseline.kind === "line" && Math.abs(baseline.x1 - baseline.x2) < 0.5).toBe(true);
  });
});
