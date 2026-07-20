import type { NumberFormat } from "../model/chartModel";

// One number formatter, reused by segment labels, totals, axis ticks, and (later)
// arrows and value lines — so every number in a chart formats consistently.
export function formatNumber(value: number, f: NumberFormat): string {
  if (f.hideZero && value === 0) return "";

  let x = value;
  let scaleSuffix = "";
  if (f.scale === "k") {
    x = x / 1000;
    scaleSuffix = "k";
  } else if (f.scale === "M") {
    x = x / 1_000_000;
    scaleSuffix = "M";
  }

  const negative = x < 0;
  const body = Math.abs(x).toLocaleString(undefined, {
    minimumFractionDigits: f.decimals,
    maximumFractionDigits: f.decimals,
    useGrouping: !!f.thousandsSep,
  });

  const core = `${f.prefix}${body}${scaleSuffix}${f.suffix}`;
  if (negative) return f.negParens ? `(${core})` : `−${core}`;
  return f.plusSign ? `+${core}` : core;
}

/** Percent of a total, e.g. for 100% stacked labels: always "NN%". */
export function formatPercent(value: number, total: number, decimals = 0): string {
  if (total === 0) return "";
  const pct = (value / total) * 100;
  return `${pct.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals })}%`;
}
