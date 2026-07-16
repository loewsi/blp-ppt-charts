import type { ChartModel } from "../model/chartModel";
import { TAG_ID, TAG_MODEL } from "./tags";

/** Read every BLP chart on a slide by parsing the model stamped on anchor shapes. */
export async function getSlideCharts(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide
): Promise<ChartModel[]> {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();

  const probes = shapes.items.map((s) => {
    const tag = s.tags.getItemOrNullObject(TAG_MODEL);
    tag.load("value, isNullObject");
    return tag;
  });
  await context.sync();

  const models: ChartModel[] = [];
  for (const tag of probes) {
    if (tag.isNullObject) continue;
    try {
      models.push(JSON.parse(tag.value) as ChartModel);
    } catch {
      // Ignore shapes whose tag isn't valid JSON.
    }
  }
  return models;
}

/** Delete all shapes belonging to a chart id. Returns how many were removed. */
export async function deleteChart(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  id: string
): Promise<number> {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();

  const probes = shapes.items.map((s) => {
    const tag = s.tags.getItemOrNullObject(TAG_ID);
    tag.load("value, isNullObject");
    return { shape: s, tag };
  });
  await context.sync();

  let removed = 0;
  for (const { shape, tag } of probes) {
    if (!tag.isNullObject && tag.value === id) {
      shape.delete();
      removed++;
    }
  }
  await context.sync();
  return removed;
}
