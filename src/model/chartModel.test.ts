import { describe, it, expect } from "vitest";
import {
  normalizeModel,
  defaultData,
  defaultOptions,
  CURRENT_SCHEMA_VERSION,
  type ChartModel,
} from "./chartModel";

describe("normalizeModel", () => {
  it("migrates a legacy model (version, stackedColumn type, no options)", () => {
    const legacy = {
      id: "c1",
      version: 1,
      name: "Old",
      data: { type: "stackedColumn", categories: ["A"], series: [{ name: "S", color: "#111", values: [1] }] },
      box: { left: 0, top: 0, width: 100, height: 100 },
    } as unknown as ChartModel;

    const m = normalizeModel(legacy);
    expect(m.schemaVersion).toBe(1);
    expect(m.data.type).toBe("barColumn");
    expect(m.options).toBeDefined();
    expect(m.options.orientation).toBe("column");
    expect(m.options.numberFormat).toBeDefined();
  });

  it("fills defaults for partial options", () => {
    const partial = {
      id: "c2",
      schemaVersion: 1,
      name: "P",
      data: defaultData(),
      box: { left: 0, top: 0, width: 100, height: 100 },
      options: { grouping: "clustered" },
    } as unknown as ChartModel;

    const m = normalizeModel(partial);
    expect(m.options.grouping).toBe("clustered"); // preserved
    expect(m.options.orientation).toBe("column"); // defaulted
    expect(m.options.numberFormat.hideZero).toBe(true); // defaulted
  });

  it("preserves a current model unchanged in meaning", () => {
    const current: ChartModel = {
      id: "c3",
      schemaVersion: CURRENT_SCHEMA_VERSION,
      name: "N",
      data: defaultData(),
      box: { left: 1, top: 2, width: 3, height: 4 },
      options: defaultOptions(),
    };
    const m = normalizeModel(current);
    expect(m.id).toBe("c3");
    expect(m.box).toEqual({ left: 1, top: 2, width: 3, height: 4 });
    expect(m.data.series.length).toBe(current.data.series.length);
  });
});
