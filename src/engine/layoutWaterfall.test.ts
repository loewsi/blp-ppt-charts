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
