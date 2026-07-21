import type { NumberFormat, SeparatorStyle } from "../model/chartModel";

// One number formatter, reused by segment labels, totals, axis ticks, and (later)
// arrows and value lines — so every number in a chart formats consistently.

const SEPARATORS: Record<Exclude<SeparatorStyle, "locale">, { group: string; decimal: string }> = {
  comma: { group: ",", decimal: "." }, // 1,234.56
  dot: { group: ".", decimal: "," }, // 1.234,56
  apos: { group: "'", decimal: "." }, // 1'234.56  (Swiss)
  space: { group: " ", decimal: "," }, // 1 234,56  (thin space)
};

/** Format an absolute number's digits with the chosen separators. */
function formatBody(abs: number, decimals: number, grouping: boolean, sep: SeparatorStyle): string {
  if (sep === "locale" || !SEPARATORS[sep]) {
    return abs.toLocaleString(undefined, {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
      useGrouping: grouping,
    });
  }
  const { group, decimal } = SEPARATORS[sep];
  const fixed = abs.toFixed(decimals); // "1234.56"
  const [intPart, fracPart] = fixed.split(".");
  const grouped = grouping ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, group) : intPart;
  return fracPart ? `${grouped}${decimal}${fracPart}` : grouped;
}

export function formatNumber(value: number, f: NumberFormat): string {
  // displayed = value × 10^scaleExp (e.g. scaleExp −3 shows a value in thousands).
  const x = f.scaleExp ? value * Math.pow(10, f.scaleExp) : value;

  // Hide only when the DISPLAYED value (after scale + rounding) is zero.
  const shown = Number(x.toFixed(f.decimals));
  if (f.hideZero && shown === 0) return "";

  const negative = shown < 0;
  const body = formatBody(Math.abs(x), f.decimals, !!f.thousandsSep, f.sep);
  const core = `${f.prefix}${body}${f.suffix}`;
  if (negative) return f.negParens ? `(${core})` : `−${core}`;
  return f.plusSign ? `+${core}` : core;
}

/** A segment label showing value, percent, or a combination of both. */
export function segmentLabel(
  value: number,
  total: number,
  mode: "value" | "percent" | "valuePercent" | "percentValue",
  f: NumberFormat
): string {
  const val = formatNumber(value, f);
  const pct = formatPercent(value, total, f.decimals, f.sep);
  switch (mode) {
    case "percent":
      return pct;
    case "valuePercent":
      return val && pct ? `${val} (${pct})` : val || pct;
    case "percentValue":
      return pct && val ? `${pct} (${val})` : pct || val;
    default:
      return val;
  }
}

/** Percent of a total, e.g. for 100% stacked labels: always "NN%". */
export function formatPercent(value: number, total: number, decimals = 0, sep: SeparatorStyle = "comma"): string {
  if (total === 0) return "";
  const pct = (value / total) * 100;
  const body = formatBody(Math.abs(pct), decimals, false, sep);
  return `${pct < 0 ? "−" : ""}${body}%`;
}
