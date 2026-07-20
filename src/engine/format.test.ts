import { describe, it, expect } from "vitest";
import { formatNumber, formatPercent } from "./format";
import { DEFAULT_NUMBER_FORMAT } from "../model/chartModel";

const base = DEFAULT_NUMBER_FORMAT;

describe("formatNumber", () => {
  // Assertions avoid locale-sensitive grouping/decimal separators.
  it("formats a plain integer", () => {
    expect(formatNumber(5, { ...base, decimals: 0 })).toBe("5");
  });

  it("hides zero when hideZero is set", () => {
    expect(formatNumber(0, { ...base, hideZero: true })).toBe("");
  });

  it("shows zero when hideZero is off", () => {
    expect(formatNumber(0, { ...base, hideZero: false })).toBe("0");
  });

  it("applies prefix and suffix", () => {
    expect(formatNumber(5, { ...base, prefix: "$", suffix: "x" })).toBe("$5x");
  });

  it("scales thousands", () => {
    expect(formatNumber(1000, { ...base, scale: "k", decimals: 0 })).toBe("1k");
  });

  it("scales millions", () => {
    expect(formatNumber(3_000_000, { ...base, scale: "M", decimals: 0 })).toBe("3M");
  });

  it("keeps decimals (separator may be locale-specific)", () => {
    expect(formatNumber(3.14159, { ...base, decimals: 2 })).toMatch(/^3[.,]14$/);
  });

  it("groups thousands when enabled", () => {
    // Separator is locale-specific (comma, dot, space, or apostrophe in de-CH).
    expect(formatNumber(1234567, { ...base, thousandsSep: true })).toMatch(/^1\D234\D567$/);
  });

  it("shows negatives with a minus by default", () => {
    expect(formatNumber(-5, base)).toBe("−5");
  });

  it("shows negatives in parentheses when enabled", () => {
    expect(formatNumber(-5, { ...base, negParens: true })).toBe("(5)");
  });

  it("adds a plus sign to positives when enabled", () => {
    expect(formatNumber(5, { ...base, plusSign: true })).toBe("+5");
  });
});

describe("formatPercent", () => {
  it("computes a whole-number share", () => {
    expect(formatPercent(25, 100, 0)).toBe("25%");
  });

  it("returns empty for a zero total", () => {
    expect(formatPercent(5, 0)).toBe("");
  });

  it("supports decimals", () => {
    expect(formatPercent(1, 3, 1)).toMatch(/^33[.,]3%$/);
  });
});
