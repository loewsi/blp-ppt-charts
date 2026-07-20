import type { ChartModel } from "../model/chartModel";
import type { Primitive } from "./primitives";
import { layoutBarColumn } from "./layoutBarColumn";

/**
 * Line chart = the column engine with every series drawn as a line. Reuses the
 * shared axis/labels/legend/annotation machinery (markers, value labels, CAGR &
 * difference arrows all work), so a line chart is just "all series are lines".
 */
export function layoutLine(model: ChartModel): Primitive[] {
  const m: ChartModel = {
    ...model,
    options: { ...model.options, orientation: "column" }, // lines need a value axis
    data: {
      ...model.data,
      series: model.data.series.map((s) => ({ ...s, kind: "line" as const })),
    },
  };
  return layoutBarColumn(m);
}
