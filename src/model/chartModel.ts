// The data model stored inside the .pptx (as a shape tag) so a chart stays
// "live-linked": we read it back and re-render when numbers or options change.

// The bar/column family. Variants are driven by ChartOptions, not separate types,
// so one layout engine covers column/bar x stacked/clustered/100%. Waterfall etc.
// will be added as new families later.
export type ChartType = "barColumn" | "waterfall";

export type Orientation = "column" | "bar";
export type Grouping = "stacked" | "clustered" | "stacked100";

export interface Series {
  name: string;
  color: string; // "#RRGGBB"
  values: number[]; // one value per category
}

export interface ChartData {
  type: ChartType;
  categories: string[];
  series: Series[];
}

/** How numbers render in labels/totals/axis. Reused everywhere (see format.ts). */
export interface NumberFormat {
  decimals: number; // 0..3
  scale: "none" | "k" | "M"; // divide by 1 / 1e3 / 1e6
  prefix: string; // e.g. "$"
  suffix: string; // e.g. " kg"
  hideZero: boolean; // blank instead of 0
  thousandsSep: boolean; // group thousands (1,234)
  negParens: boolean; // show negatives as (123) instead of −123
  plusSign: boolean; // show a leading + on positives
}

export interface ChartOptions {
  orientation: Orientation;
  grouping: Grouping;
  gap: number; // 0..0.9 gap between categories (fraction of a slot)
  showTotals: boolean; // totals at the end of each stack
  showValueLabels: boolean; // value inside each segment/bar
  reverseCategories: boolean;
  showLegend: boolean; // series legend
  legendPosition: "top" | "bottom" | "left" | "right";
  showGridlines: boolean; // value-axis gridlines across the plot
  showValueAxis: boolean; // value-axis tick labels
  showConnectors: boolean; // connector lines between stacked segments across categories
  referenceValue: number | null; // horizontal reference/target line at this value (null = off)
  axisMin: number | null; // fix the value-axis minimum (null = auto)
  axisMax: number | null; // fix the value-axis maximum (null = auto)
  labelOverflow: "inside" | "outside"; // small-segment labels: keep inside (chip + offset) or place outside
  fontFamily: string; // font for all labels
  segmentFontSize: number; // pt, segment/value labels
  totalFontSize: number; // pt, total labels
  numberFormat: NumberFormat;
}

/** Position + size of the chart on the slide, in points (1in = 72pt). */
export interface ChartBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ChartModel {
  id: string;
  schemaVersion: number;
  name: string;
  data: ChartData;
  box: ChartBox;
  options: ChartOptions;
}

/** Bump when the stored model shape changes; normalizeModel migrates older ones. */
export const CURRENT_SCHEMA_VERSION = 1;

export const DEFAULT_BOX: ChartBox = { left: 120, top: 150, width: 380, height: 230 };

/** BLP brand palette, cycled across series. */
export const PALETTE = ["#2E75FF", "#001C54", "#9CC0FF", "#5A93FF", "#C7DEFF", "#3F5B8C"];

// Named presets applied to all series at once. "master" is handled separately
// (read live from the theme). BLP is mainly-blue: the first several entries are
// blue shades, so typical charts stay blue and only pull in other accents when
// there are many series.
export const PALETTES: Record<string, string[]> = {
  blp: [
    "#2E75FF", // BLP blue
    "#001C54", // dark navy
    "#5A93FF",
    "#9CC0FF",
    "#12377E",
    "#7FB0FF",
    "#E8412C", // accents only kick in past ~6 series
    "#F5A623",
    "#2FA84F",
    "#7B4FE0",
  ],
  grayscale: ["#2B2B2B", "#555555", "#808080", "#A9A9A9", "#C7C7C7", "#E2E2E2"],
};

export const DEFAULT_NUMBER_FORMAT: NumberFormat = {
  decimals: 0,
  scale: "none",
  prefix: "",
  suffix: "",
  hideZero: true,
  thousandsSep: false,
  negParens: false,
  plusSign: false,
};

export const DEFAULT_OPTIONS: ChartOptions = {
  orientation: "column",
  grouping: "stacked",
  gap: 0.35,
  showTotals: true,
  showValueLabels: true,
  reverseCategories: false,
  showLegend: false,
  legendPosition: "bottom",
  showGridlines: false,
  showValueAxis: false,
  showConnectors: false,
  referenceValue: null,
  axisMin: null,
  axisMax: null,
  labelOverflow: "inside",
  fontFamily: "Roboto",
  segmentFontSize: 9,
  totalFontSize: 10,
  numberFormat: { ...DEFAULT_NUMBER_FORMAT },
};

export function defaultOptions(): ChartOptions {
  return { ...DEFAULT_OPTIONS, numberFormat: { ...DEFAULT_NUMBER_FORMAT } };
}

export function defaultData(): ChartData {
  return {
    type: "barColumn",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "Product A", color: PALETTE[0], values: [10, 12, 9, 14] },
      { name: "Product B", color: PALETTE[1], values: [6, 8, 7, 9] },
    ],
  };
}

/** Fill in schemaVersion/options/type for models saved by older builds. */
export function normalizeModel(m: ChartModel): ChartModel {
  const legacyVersion = (m as { version?: number }).version;
  return {
    ...m,
    schemaVersion: m.schemaVersion ?? legacyVersion ?? CURRENT_SCHEMA_VERSION,
    data: { ...m.data, type: m.data?.type === "waterfall" ? "waterfall" : "barColumn" },
    options: m.options
      ? { ...DEFAULT_OPTIONS, ...m.options, numberFormat: { ...DEFAULT_NUMBER_FORMAT, ...m.options.numberFormat } }
      : defaultOptions(),
  };
}
