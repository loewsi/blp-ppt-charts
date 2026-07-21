import type { ChartModel, ChartType } from "../model/chartModel";
import type { Primitive } from "./primitives";
import { layoutBarColumn } from "./layoutBarColumn";
import { layoutLine } from "./layoutLine";
import { layoutMekko } from "./layoutMekko";
import { layoutPie } from "./layoutPie";
import { layoutScatter } from "./layoutScatter";
import { layoutWaterfall } from "./layoutWaterfall";

type LayoutFn = (model: ChartModel) => Primitive[];

const REGISTRY: Record<ChartType, LayoutFn> = {
  barColumn: layoutBarColumn,
  line: layoutLine,
  combination: layoutBarColumn, // combination = bars + per-series line kinds (D8)
  waterfall: layoutWaterfall,
  pie: layoutPie,
  scatter: layoutScatter,
  mekko: layoutMekko,
};

export function computeLayout(model: ChartModel): Primitive[] {
  // Fall back to the bar/column engine for older models (e.g. "stackedColumn").
  const fn = REGISTRY[model.data.type] ?? layoutBarColumn;
  return fn(model);
}
