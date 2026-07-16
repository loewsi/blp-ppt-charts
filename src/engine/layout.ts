import type { ChartModel, ChartType } from "../model/chartModel";
import type { Primitive } from "./primitives";
import { layoutStackedColumn } from "./layoutStackedColumn";

type LayoutFn = (model: ChartModel) => Primitive[];

// Register future chart types here (e.g. waterfall) - nothing else changes.
const REGISTRY: Record<ChartType, LayoutFn> = {
  stackedColumn: layoutStackedColumn,
};

export function computeLayout(model: ChartModel): Primitive[] {
  const fn = REGISTRY[model.data.type];
  if (!fn) throw new Error(`No layout registered for chart type "${model.data.type}"`);
  return fn(model);
}
