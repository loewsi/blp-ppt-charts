import { describe, it, expect } from "vitest";
import { layoutLine } from "./layoutLine";
import type { ChartModel } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";

function model(): ChartModel {
  return {
    id: "t",
    schemaVersion: 1,
    name: "T",
    box: { left: 0, top: 0, width: 400, height: 300 },
    options: defaultOptions(),
    data: {
      type: "line",
      categories: ["A", "B", "C"],
      series: [
        { name: "S1", color: "#111111", values: [10, 20, 15] },
        { name: "S2", color: "#222222", values: [5, 8, 12] },
      ],
    },
  };
}

describe("layoutLine", () => {
  const prims = layoutLine(model());

  it("draws no bar segments (everything is a line)", () => {
    expect(prims.some((s) => s.kind === "rect" && s.meta?.objectType === "segment")).toBe(false);
  });

  it("draws a polyline for each series (n-1 segments) plus markers", () => {
    const lines = prims.filter((s) => s.kind === "line" && s.meta?.objectType === "lineSeries");
    const markers = prims.filter((s) => s.kind === "rect" && s.meta?.objectType === "lineMarker");
    expect(lines.length).toBe(4); // 2 series × (3 points − 1)
    expect(markers.length).toBe(6); // 2 series × 3 points
  });

  it("still draws category labels", () => {
    expect(prims.filter((s) => s.meta?.objectType === "categoryLabel").length).toBe(3);
  });
});
