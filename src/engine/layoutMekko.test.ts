import { describe, it, expect } from "vitest";
import { layoutMekko } from "./layoutMekko";
import type { ChartModel, ChartOptions } from "../model/chartModel";
import { defaultOptions } from "../model/chartModel";
import type { RectPrimitive as Rect } from "./primitives";

function model(over: Partial<ChartOptions> = {}): ChartModel {
  return {
    id: "t",
    schemaVersion: 1,
    name: "T",
    box: { left: 0, top: 0, width: 400, height: 300 },
    options: { ...defaultOptions(), ...over },
    data: {
      type: "mekko",
      categories: ["Big", "Small"],
      series: [
        { name: "S1", color: "#111", values: [80, 10] },
        { name: "S2", color: "#222", values: [20, 10] },
      ],
    },
  };
}

const segs = (p: ReturnType<typeof layoutMekko>) =>
  p.filter((s): s is Rect => s.kind === "rect" && s.meta?.objectType === "segment");

describe("layoutMekko", () => {
  it("emits one segment per non-zero value", () => {
    expect(segs(layoutMekko(model())).length).toBe(4);
  });

  it("makes the higher-total column wider", () => {
    const s = segs(layoutMekko(model()));
    const widthOfCat = (ci: number) => s.find((r) => r.meta?.categoryIndex === ci)!.w;
    expect(widthOfCat(0)).toBeGreaterThan(widthOfCat(1)); // total 100 vs 20
  });

  it("normalizes each column to full plot height (100% stacked)", () => {
    const s = segs(layoutMekko(model()));
    const heightFor = (ci: number) => s.filter((r) => r.meta?.categoryIndex === ci).reduce((a, r) => a + r.h, 0);
    expect(Math.abs(heightFor(0) - heightFor(1))).toBeLessThan(0.5);
  });

  it("returns nothing for an all-zero dataset", () => {
    const m = model();
    m.data.series.forEach((se) => (se.values = [0, 0]));
    expect(layoutMekko(m)).toEqual([]);
  });
});
