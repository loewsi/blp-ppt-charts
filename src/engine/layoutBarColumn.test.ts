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

  it("outside moves ALL labels out, including big segments (not just small ones)", () => {
    // Default model: every segment is large; with outside none should be a centered chip.
    const labels = segLabels(layoutBarColumn(model({ labelOverflow: "outside" })));
    expect(labels.length).toBe(4);
    expect(labels.every((l) => !l.bg)).toBe(true); // no chips → all placed outside
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

describe("layoutBarColumn — negative values", () => {
  const withNeg = {
    type: "barColumn" as const,
    categories: ["A", "B"],
    series: [{ name: "S", color: "#111", values: [10, -6] }],
  };

  it("renders a bar for a negative value (not dropped)", () => {
    expect(segRects(layoutBarColumn(model({ grouping: "clustered" }, withNeg))).length).toBe(2);
  });

  it("places the negative bar below the zero baseline", () => {
    const prims = layoutBarColumn(model({ grouping: "clustered" }, withNeg));
    const baseline = prims.find((s) => s.kind === "line" && s.meta?.objectType === "baseline");
    const baseY = baseline && baseline.kind === "line" ? baseline.y1 : 0;
    const negBar = segRects(prims).find((r) => r.meta?.categoryIndex === 1)!;
    expect(negBar.y).toBeGreaterThanOrEqual(baseY - 1);
  });
});

describe("layoutBarColumn — reverse series / totals / axis line / zeros", () => {
  const bottomOf = (rs: RectPrimitive[]) => rs.reduce((a, b) => (a.y + a.h > b.y + b.h ? a : b));

  it("reverseSeries flips which series sits on the baseline", () => {
    const normal = segRects(layoutBarColumn(model({}))).filter((r) => r.meta?.categoryIndex === 0);
    const rev = segRects(layoutBarColumn(model({ reverseSeries: true }))).filter((r) => r.meta?.categoryIndex === 0);
    expect(bottomOf(normal).meta?.seriesIndex).toBe(0);
    expect(bottomOf(rev).meta?.seriesIndex).toBe(1);
  });

  it("shows the absolute total on 100% stacked", () => {
    const prims = layoutBarColumn(model({ grouping: "stacked100" }));
    const totals = prims.filter((s): s is TextPrimitive => s.kind === "text" && s.meta?.objectType === "totalLabel");
    expect(totals.length).toBe(2);
    expect(totals.map((t) => t.text)).toContain("40"); // cat A: 10 + 30, absolute (not 100%)
  });

  it("draws a value-axis line only when enabled", () => {
    const line = (p: ReturnType<typeof layoutBarColumn>) =>
      p.some((s) => s.kind === "line" && s.meta?.objectType === "valueAxis");
    expect(line(layoutBarColumn(model({ showAxisLine: true })))).toBe(true);
    expect(line(layoutBarColumn(model({})))).toBe(false);
  });

  it("labels a zero segment at the baseline only when zeros aren't hidden", () => {
    const data = {
      type: "barColumn" as const,
      categories: ["A"],
      series: [
        { name: "S1", color: "#111", values: [0] },
        { name: "S2", color: "#222", values: [5] },
      ],
    };
    const labelsWith = (hideZero: boolean) =>
      layoutBarColumn(model({ numberFormat: { ...defaultOptions().numberFormat, hideZero } }, data)).filter(
        (s): s is TextPrimitive => s.kind === "text" && s.meta?.objectType === "segmentLabel"
      );
    expect(labelsWith(false).some((l) => l.text === "0")).toBe(true);
    expect(labelsWith(true).some((l) => l.text === "0")).toBe(false);
  });
});

describe("layoutBarColumn — reference line", () => {
  const isRef = (s: { meta?: { objectType?: string } }) => s.meta?.objectType === "valueLine";

  it("draws no reference line by default", () => {
    expect(layoutBarColumn(model({})).some(isRef)).toBe(false);
  });

  it("draws a horizontal line + label at the value for a column chart", () => {
    const prims = layoutBarColumn(model({ referenceValue: 25, orientation: "column" }));
    const line = prims.find((s) => s.kind === "line" && isRef(s));
    expect(line && line.kind === "line" && Math.abs(line.y1 - line.y2) < 0.5).toBe(true); // horizontal
    expect(prims.some((s) => s.kind === "text" && isRef(s))).toBe(true);
  });

  it("draws a vertical line for a bar chart", () => {
    const prims = layoutBarColumn(model({ referenceValue: 25, orientation: "bar" }));
    const line = prims.find((s) => s.kind === "line" && isRef(s));
    expect(line && line.kind === "line" && Math.abs(line.x1 - line.x2) < 0.5).toBe(true); // vertical
  });

  it("skips a reference value outside the axis range", () => {
    expect(layoutBarColumn(model({ referenceValue: 9999 })).some(isRef)).toBe(false);
  });

  it("skips reference lines on 100% stacked charts", () => {
    expect(layoutBarColumn(model({ referenceValue: 0.5, grouping: "stacked100" })).some(isRef)).toBe(false);
  });
});

describe("layoutBarColumn — manual axis min/max", () => {
  // Column chart: a fixed, larger axisMax makes every segment shorter than with auto scale.
  it("fixing a larger axisMax shrinks the bars", () => {
    const auto = segRects(layoutBarColumn(model({ grouping: "clustered" })));
    const fixed = segRects(layoutBarColumn(model({ grouping: "clustered", axisMax: 200 })));
    const h = (p: typeof auto, si: number, ci: number) =>
      p.find((r) => r.meta?.seriesIndex === si && r.meta?.categoryIndex === ci)!.h;
    expect(h(fixed, 1, 1)).toBeLessThan(h(auto, 1, 1)); // value 40 on a 0..200 axis vs auto
  });

  it("axisMin below zero opens space under the baseline", () => {
    const prims = layoutBarColumn(model({ grouping: "clustered", axisMin: -50 }));
    const baseline = prims.find((s) => s.kind === "line" && s.meta?.objectType === "baseline");
    const baseY = baseline && baseline.kind === "line" ? baseline.y1 : 0;
    // With min −50 and a 300pt-tall plot, the zero baseline sits well below the plot top.
    expect(baseY).toBeGreaterThan(0);
  });

  it("guards against an inverted/empty range (max <= min)", () => {
    expect(() => layoutBarColumn(model({ axisMin: 100, axisMax: 100 }))).not.toThrow();
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
