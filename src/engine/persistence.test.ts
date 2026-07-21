import { describe, it, expect } from "vitest";
import { planDuplicateRepair, planOrphanLegends, type RepairEntry } from "./persistence";

const chart = (id: string, cx: number, cy: number): RepairEntry => ({ id, hasModel: true, part: "chart", cx, cy });
const legend = (id: string, cx: number, cy: number): RepairEntry => ({ id, hasModel: false, part: "legend", cx, cy });

describe("planDuplicateRepair", () => {
  it("does nothing when every id is unique", () => {
    expect(planDuplicateRepair([chart("A", 0, 0), chart("B", 100, 0)])).toEqual([]);
  });

  it("reassigns the copy when an id appears twice (no legends)", () => {
    const plan = planDuplicateRepair([chart("A", 0, 0), chart("A", 20, 20)]);
    expect(plan).toEqual([{ anchorIndex: 1, legendIndex: null }]);
  });

  it("pairs each duplicated chart with its nearest legend", () => {
    // z-order: original chart+legend, then pasted chart+legend (offset by +20,+20)
    const plan = planDuplicateRepair([
      chart("A", 0, 0), // 0 original chart
      legend("A", 0, 50), // 1 original legend
      chart("A", 20, 20), // 2 pasted chart
      legend("A", 20, 70), // 3 pasted legend
    ]);
    // Original (index 0) keeps its id + legend 1; the copy (index 2) gets a new id + legend 3.
    expect(plan).toEqual([{ anchorIndex: 2, legendIndex: 3 }]);
  });

  it("handles a chart pasted more than once", () => {
    const plan = planDuplicateRepair([chart("A", 0, 0), chart("A", 20, 20), chart("A", 40, 40)]);
    expect(plan.map((p) => p.anchorIndex)).toEqual([1, 2]);
  });

  it("leaves other ids untouched while repairing one", () => {
    const plan = planDuplicateRepair([chart("A", 0, 0), chart("A", 20, 20), chart("B", 200, 0)]);
    expect(plan).toEqual([{ anchorIndex: 1, legendIndex: null }]);
  });
});

describe("planOrphanLegends", () => {
  it("flags a legend whose chart anchor is gone", () => {
    // chart A intact; chart B was deleted but its legend (index 2) remains
    const entries = [chart("A", 0, 0), legend("A", 0, 50), legend("B", 200, 50)];
    expect(planOrphanLegends(entries)).toEqual([2]);
  });

  it("keeps legends that still have their chart", () => {
    expect(planOrphanLegends([chart("A", 0, 0), legend("A", 0, 50)])).toEqual([]);
  });
});
