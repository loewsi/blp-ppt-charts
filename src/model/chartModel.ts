// The data model stored inside the .pptx (as a shape tag) so a chart stays
// "live-linked": we read it back and re-render when numbers or options change.

// The bar/column family. Variants are driven by ChartOptions, not separate types,
// so one layout engine covers column/bar x stacked/clustered/100%. Waterfall etc.
// will be added as new families later.
export type ChartType = "barColumn";

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
}

export interface ChartOptions {
  orientation: Orientation;
  grouping: Grouping;
  gap: number; // 0..0.9 gap between categories (fraction of a slot)
  showTotals: boolean; // totals at the end of each stack
  showValueLabels: boolean; // value inside each segment/bar
  reverseCategories: boolean;
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
  version: 1;
  name: string;
  data: ChartData;
  box: ChartBox;
  options: ChartOptions;
}

export const DEFAULT_BOX: ChartBox = { left: 120, top: 150, width: 380, height: 230 };

/** BLP brand palette, cycled across series. */
export const PALETTE = ["#2E75FF", "#001C54", "#9CC0FF", "#5A93FF", "#C7DEFF", "#3F5B8C"];

export const DEFAULT_NUMBER_FORMAT: NumberFormat = {
  decimals: 0,
  scale: "none",
  prefix: "",
  suffix: "",
  hideZero: true,
};

export const DEFAULT_OPTIONS: ChartOptions = {
  orientation: "column",
  grouping: "stacked",
  gap: 0.35,
  showTotals: true,
  showValueLabels: true,
  reverseCategories: false,
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

/** Fill in options/type for models saved by older builds. */
export function normalizeModel(m: ChartModel): ChartModel {
  return {
    ...m,
    data: { ...m.data, type: "barColumn" },
    options: m.options
      ? { ...DEFAULT_OPTIONS, ...m.options, numberFormat: { ...DEFAULT_NUMBER_FORMAT, ...m.options.numberFormat } }
      : defaultOptions(),
  };
}
