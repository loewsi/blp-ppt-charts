import { describe, it, expect } from "vitest";
import { evaluateGrid, colToIndex, type Grid } from "./formula";

// A1-style over the whole grid: B2 = row 1, col 1 (first data value).
function grid(): Grid {
  return [
    ["", "Q1", "Q2", "Q3"], // row 0 (A1..D1)
    ["Rev", "10", "20", "30"], // row 1 (A2..D2)
    ["Cost", "4", "6", "8"], // row 2 (A3..D3)
    ["Profit", "", "", ""], // row 3 (A4..D4)
  ];
}

describe("colToIndex", () => {
  it("maps letters to indices", () => {
    expect(colToIndex("A")).toBe(0);
    expect(colToIndex("B")).toBe(1);
    expect(colToIndex("Z")).toBe(25);
    expect(colToIndex("AA")).toBe(26);
  });
});

describe("evaluateGrid", () => {
  it("passes literals through", () => {
    const { values, display } = evaluateGrid(grid());
    expect(values[1][1]).toBe(10);
    expect(display[1][1]).toBe("10");
  });

  it("evaluates arithmetic with cell refs", () => {
    const g = grid();
    g[3][1] = "=B2-B3"; // 10 - 4 = 6
    const { values, display } = evaluateGrid(g);
    expect(values[3][1]).toBe(6);
    expect(display[3][1]).toBe("6");
  });

  it("respects operator precedence + parentheses", () => {
    const g = grid();
    g[3][1] = "=B2+B3*2"; // 10 + 8 = 18
    g[3][2] = "=(C2+C3)*2"; // (20+6)*2 = 52
    const { values } = evaluateGrid(g);
    expect(values[3][1]).toBe(18);
    expect(values[3][2]).toBe(52);
  });

  it("supports SUM over a range", () => {
    const g = grid();
    g[3][1] = "=SUM(B2:D2)"; // 10+20+30
    const { values } = evaluateGrid(g);
    expect(values[3][1]).toBe(60);
  });

  it("supports AVERAGE/MIN/MAX/COUNT", () => {
    const g = grid();
    g[3][1] = "=AVERAGE(B2:D2)";
    g[3][2] = "=MIN(B2:D2)";
    g[3][3] = "=MAX(B2:D2)";
    g[0][0] = "=COUNT(B2:D2)";
    const { values } = evaluateGrid(g);
    expect(values[3][1]).toBe(20);
    expect(values[3][2]).toBe(10);
    expect(values[3][3]).toBe(30);
    expect(values[0][0]).toBe(3);
  });

  it("chains formulas that reference other formulas", () => {
    const g = grid();
    g[3][1] = "=B2-B3"; // 6
    g[3][2] = "=B4*10"; // references the formula in B4 → 60
    const { values } = evaluateGrid(g);
    expect(values[3][2]).toBe(60);
  });

  it("ignores text cells in a range", () => {
    const g = grid();
    g[3][1] = "=SUM(A2:D2)"; // A2 = "Rev" (text) is skipped
    const { values } = evaluateGrid(g);
    expect(values[3][1]).toBe(60);
  });

  it("flags a cyclic reference", () => {
    const g = grid();
    g[3][1] = "=B4"; // B4 refers to itself
    const { display } = evaluateGrid(g);
    expect(display[3][1]).toBe("#CYCLE");
  });

  it("flags a malformed formula", () => {
    const g = grid();
    g[3][1] = "=B2+";
    const { display } = evaluateGrid(g);
    expect(display[3][1]).toBe("#ERR");
  });
});
