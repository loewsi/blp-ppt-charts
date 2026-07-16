// The data model that is stored inside the .pptx (as a shape tag) so a chart
// stays "live-linked": we can read it back and re-render when numbers change.

// v1 ships "stackedColumn". Waterfall/clustered slot in here later.
export type ChartType = "stackedColumn";

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
  data: ChartData;
  box: ChartBox;
}

export const DEFAULT_BOX: ChartBox = { left: 180, top: 120, width: 600, height: 320 };

/** BLP brand palette, cycled across series. */
export const PALETTE = ["#2E75FF", "#001C54", "#9CC0FF", "#5A93FF", "#C7DEFF", "#3F5B8C"];

export function defaultData(): ChartData {
  return {
    type: "stackedColumn",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: [
      { name: "Product A", color: PALETTE[0], values: [10, 12, 9, 14] },
      { name: "Product B", color: PALETTE[1], values: [6, 8, 7, 9] },
    ],
  };
}
