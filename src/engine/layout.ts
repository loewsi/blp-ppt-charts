import type { ChartModel, ChartType } from "../model/chartModel";
import type { Primitive } from "./primitives";
import { layoutBarColumn } from "./layoutBarColumn";

type LayoutFn = (model: ChartModel) => Primitive[];

// Register future chart families here (e.g. waterfall) - nothing else changes.
const REGISTRY: Record<ChartType, LayoutFn> = {
  barColumn: layoutBarColumn,
};

export function computeLayout(model: ChartModel): Primitive[] {
  // Fall back to the bar/column engine for older models (e.g. "stackedColumn").
  const fn = REGISTRY[model.data.type] ?? layoutBarColumn;
  return fn(model);
}
