import { describe, it, expect } from "vitest";
import { mixColor, shadesFrom } from "./color";

describe("mixColor", () => {
  it("returns the base unchanged at t=0", () => {
    expect(mixColor("#2E75FF", 0)).toBe("#2e75ff");
  });
  it("mixes toward white for t>0", () => {
    expect(mixColor("#000000", 0.5)).toBe("#808080");
  });
  it("mixes toward black for t<0", () => {
    expect(mixColor("#ffffff", -0.5)).toBe("#808080");
  });
});

describe("shadesFrom", () => {
  it("returns n shades", () => {
    expect(shadesFrom("#2E75FF", 5)).toHaveLength(5);
  });
  it("goes light → dark (first lighter than last)", () => {
    const [first, , last] = shadesFrom("#2E75FF", 3);
    const lum = (h: string) => parseInt(h.slice(1, 3), 16) + parseInt(h.slice(3, 5), 16) + parseInt(h.slice(5, 7), 16);
    expect(lum(first)).toBeGreaterThan(lum(last));
  });
  it("handles a single series", () => {
    expect(shadesFrom("#2E75FF", 1)).toHaveLength(1);
  });
});
