import { describe, it, expect } from "vitest";
import {
  cellsFromData,
  dataFromCells,
  pasteInto,
  insertColumn,
  removeColumn,
  insertRow,
  removeRow,
} from "./grid";
import type { ChartData } from "../model/chartModel";

const data: ChartData = {
  type: "barColumn",
  categories: ["Q1", "Q2"],
  series: [
    { name: "A", color: "#111", values: [1, 2] },
    { name: "B", color: "#222", values: [3, 4] },
  ],
};

describe("cells <-> data", () => {
  it("round-trips", () => {
    const cells = cellsFromData(data);
    expect(cells[0]).toEqual(["", "Q1", "Q2"]);
    expect(cells[1]).toEqual(["A", "1", "2"]);
    const back = dataFromCells(cells, ["#111", "#222"]);
    expect(back.categories).toEqual(["Q1", "Q2"]);
    expect(back.series[0]).toEqual({ name: "A", color: "#111", values: [1, 2] });
  });

  it("defaults blank names and parses messy numbers", () => {
    const back = dataFromCells([["", ""], ["", "1,000"]], []);
    expect(back.categories[0]).toBe("Cat 1");
    expect(back.series[0].name).toBe("Series 1");
    expect(back.series[0].values).toEqual([1000]);
  });
});

describe("pasteInto", () => {
  it("places a block at the target", () => {
    const out = pasteInto(cellsFromData(data), 1, 1, "9\t8\n7\t6");
    expect([out[1][1], out[1][2], out[2][1], out[2][2]]).toEqual(["9", "8", "7", "6"]);
  });

  it("grows rows and columns when the block overflows", () => {
    const out = pasteInto([["", "Q1"], ["A", "1"]], 1, 1, "1\t2\t3\n4\t5\t6");
    expect(out.length).toBe(3);
    expect(out[0].length).toBe(4);
    expect(out[2][3]).toBe("6");
  });
});

describe("insert / remove", () => {
  it("insertColumn adds a blank column", () => {
    const out = insertColumn(cellsFromData(data), 1);
    expect(out[0].length).toBe(4);
    expect(out[0][1]).toBe("");
  });

  it("removeColumn keeps at least one category", () => {
    let c = removeColumn(cellsFromData(data), 1); // 3 -> 2 cols
    expect(c[0].length).toBe(2);
    c = removeColumn(c, 1); // blocked at the last category
    expect(c[0].length).toBe(2);
  });

  it("insertRow / removeRow adjust series", () => {
    let c = insertRow(cellsFromData(data), 1);
    expect(c.length).toBe(4);
    c = removeRow(c, 1);
    expect(c.length).toBe(3);
  });
});
