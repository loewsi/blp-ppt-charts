import type { ChartModel, ChartType } from "../model/chartModel";
import type { Primitive } from "./primitives";
import { layoutBarColumn } from "./layoutBarColumn";
import { layoutWaterfall } from "./layoutWaterfall";

type LayoutFn = (model: ChartModel) => Primitive[];

const REGISTRY: Record<ChartType, LayoutFn> = {
  barColumn: layoutBarColumn,
  waterfall: layoutWaterfall,
};

export function computeLayout(model: ChartModel): Primitive[] {
  // Fall back to the bar/column engine for older models (e.g. "stackedColumn").
  const fn = REGISTRY[model.data.type] ?? layoutBarColumn;
  return fn(model);
}
