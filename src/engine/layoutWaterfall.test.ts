import { describe, it, expect } from "vitest";
import { layoutWaterfall } from "./layoutWaterfall";
import type { ChartModel } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";
import type { RectPrimitive } from "./primitives";

function model(): ChartModel {
  return {
    id: "w",
    schemaVersion: 1,
    name: "W",
    box: { left: 0, top: 0, width: 400, height: 300 },
    options: defaultOptions(),
    data: {
      type: "waterfall",
      categories: ["Start", "A", "B"],
      series: [{ name: "D", color: "#2E75FF", values: [10, 5, -3] }],
    },
  };
}

describe("layoutWaterfall", () => {
  const p = layoutWaterfall(model());
  const bars = p.filter((s): s is RectPrimitive => s.kind === "rect" && s.meta?.objectType === "segment");

  it("emits one bar per category", () => expect(bars.length).toBe(3));

  it("draws a zero baseline", () =>
    expect(p.some((s) => s.kind === "line" && s.meta?.objectType === "baseline")).toBe(true));

  it("connects consecutive steps", () =>
    expect(p.filter((s) => s.kind === "line" && s.meta?.objectType === "connector").length).toBe(2));

  it("colours falls differently from rises", () => {
    const rise = bars.find((r) => r.meta?.categoryIndex === 0); // +10
    const fall = bars.find((r) => r.meta?.categoryIndex === 2); // -3
    expect(rise?.fill).toBe("#2E75FF");
    expect(fall?.fill).not.toBe("#2E75FF");
  });

  it("labels each category", () =>
    expect(p.filter((s) => s.meta?.objectType === "categoryLabel").length).toBe(3));
});

describe("layoutWaterfall — multi-series steps", () => {
  function multi(): ChartModel {
    const m = model();
    m.data.categories = ["Q1", "Q2"];
    m.data.series = [
      { name: "A", color: "#111111", values: [6, 4] },
      { name: "B", color: "#222222", values: [4, -1] },
    ];
    return m;
  }
  const p = layoutWaterfall(multi());
  const bars = p.filter((s): s is RectPrimitive => s.kind === "rect" && s.meta?.objectType === "segment");

  it("stacks each series as its own sub-segment within a step", () => {
    expect(bars.length).toBe(4); // 2 steps × 2 series
  });

  it("colours sub-segments by series (not rise/fall)", () => {
    expect(bars.some((r) => r.fill === "#111111")).toBe(true);
    expect(bars.some((r) => r.fill === "#222222")).toBe(true);
  });

  it("still connects the two steps once", () => {
    expect(p.filter((s) => s.kind === "line" && s.meta?.objectType === "connector").length).toBe(1);
  });
});

describe("layoutWaterfall — 'e' total column", () => {
  function withTotal(): ChartModel {
    const m = model();
    // Start +10, A +5, then a total column (e) → should show 15 from the baseline.
    m.data.categories = ["Start", "A", "Total"];
    m.data.series = [{ name: "D", color: "#2E75FF", values: [10, 5, 0] }];
    m.data.totalFlags = [false, false, true];
    return m;
  }

  const p = layoutWaterfall(withTotal());
  const bars = p.filter((s): s is RectPrimitive => s.kind === "rect" && s.meta?.objectType === "segment");
  const baseline = p.find((s) => s.kind === "line" && s.meta?.objectType === "baseline");
  const baseY = baseline && baseline.kind === "line" ? baseline.y1 : 0;

  it("the total column reaches the baseline (bar from 0 to the running sum)", () => {
    const totalBar = bars.find((r) => r.meta?.categoryIndex === 2)!;
    expect(Math.abs(totalBar.y + totalBar.h - baseY)).toBeLessThan(1); // bottom sits on the baseline
  });

  it("the total column is taller than either delta (spans the whole sum)", () => {
    const h = (i: number) => bars.find((r) => r.meta?.categoryIndex === i)!.h;
    expect(h(2)).toBeGreaterThan(h(0)); // total 15 vs first delta 10
  });

  it("labels the total with the running sum", () => {
    expect(p.some((s) => s.kind === "text" && s.meta?.objectType === "segmentLabel" && s.text === "15")).toBe(true);
  });
});
