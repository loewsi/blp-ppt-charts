import type { ChartModel, ChartBox } from "../model/chartModel";
import { normalizeModel } from "../model/chartModel";
import { TAG_ID, TAG_MODEL, TAG_PART } from "./tags";

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
      models.push(normalizeModel(JSON.parse(tag.value) as ChartModel));
    } catch {
      // Ignore shapes whose tag isn't valid JSON.
    }
  }
  return models;
}

/** Read a chart's current bounding box on the slide (so Update redraws in place). */
export async function getChartBox(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  id: string
): Promise<ChartBox | null> {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();

  const probes = shapes.items.map((s) => {
    const tag = s.tags.getItemOrNullObject(TAG_ID);
    const part = s.tags.getItemOrNullObject(TAG_PART);
    tag.load("value, isNullObject");
    part.load("value, isNullObject");
    s.load("left, top, width, height");
    return { s, tag, part };
  });
  await context.sync();

  const mine = probes.filter((p) => !p.tag.isNullObject && p.tag.value === id);
  // Prefer the chart part (exclude the separate legend group) for the box.
  const chartOnly = mine.filter((p) => !p.part.isNullObject && p.part.value === "chart");
  const matching = (chartOnly.length > 0 ? chartOnly : mine).map((p) => p.s);
  if (matching.length === 0) return null;

  // Union of all matching shapes (a single group, or many shapes on old hosts).
  let left = Infinity,
    top = Infinity,
    right = -Infinity,
    bottom = -Infinity;
  for (const s of matching) {
    left = Math.min(left, s.left);
    top = Math.min(top, s.top);
    right = Math.max(right, s.left + s.width);
    bottom = Math.max(bottom, s.top + s.height);
  }
  return { left, top, width: right - left, height: bottom - top };
}

/** Read the chart id of whatever is selected on the slide (group or shape), or null. */
export async function getSelectedChartId(
  context: PowerPoint.RequestContext
): Promise<string | null> {
  const sel = context.presentation.getSelectedShapes();
  sel.load("items");
  await context.sync();

  const probes = sel.items.map((s) => {
    const tag = s.tags.getItemOrNullObject(TAG_ID);
    tag.load("value, isNullObject");
    return tag;
  });
  await context.sync();

  for (const tag of probes) {
    if (!tag.isNullObject) return tag.value;
  }
  return null;
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
