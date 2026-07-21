import type { ChartModel, ChartBox } from "../model/chartModel";
import { normalizeModel } from "../model/chartModel";
import { newId } from "../util/id";
import { TAG_ID, TAG_MODEL, TAG_PART } from "./tags";

// ---- duplicate-id repair (same-slide copy-paste) -------------------------
// A copy-paste clones a chart's group(s) with the SAME id, so editing one would
// delete both. We detect ids that carry more than one "anchor" (a group holding
// the model) and split the extras onto fresh ids, pairing each with its nearest
// legend group so a pasted chart becomes independently editable.

export interface RepairEntry {
  id: string;
  hasModel: boolean; // true for the chart's anchor group (carries TAG_MODEL)
  part: string | null; // "chart" | "legend" | null
  cx: number; // shape center, for pairing a legend to its chart
  cy: number;
}

export interface RepairAssignment {
  anchorIndex: number; // entry to receive a new id
  legendIndex: number | null; // its paired legend entry, if any
}

/** Pure: decide which duplicated anchors need a fresh id and which legend goes with each. */
export function planDuplicateRepair(entries: RepairEntry[]): RepairAssignment[] {
  const byId = new Map<string, number[]>();
  entries.forEach((e, i) => {
    const arr = byId.get(e.id) ?? [];
    arr.push(i);
    byId.set(e.id, arr);
  });

  const out: RepairAssignment[] = [];
  for (const idxs of byId.values()) {
    const anchors = idxs.filter((i) => entries[i].hasModel);
    if (anchors.length <= 1) continue; // no duplication for this id
    const legends = idxs.filter((i) => entries[i].part === "legend");

    // Greedily reserve each anchor's nearest legend (anchor order = z-order, so the
    // original keeps its legend first); only the extra anchors get reassigned.
    const used = new Set<number>();
    const legendFor = new Map<number, number | null>();
    for (const ai of anchors) {
      let best = -1;
      let bestD = Infinity;
      for (const li of legends) {
        if (used.has(li)) continue;
        const d = (entries[li].cx - entries[ai].cx) ** 2 + (entries[li].cy - entries[ai].cy) ** 2;
        if (d < bestD) {
          bestD = d;
          best = li;
        }
      }
      if (best >= 0) used.add(best);
      legendFor.set(ai, best >= 0 ? best : null);
    }
    for (let a = 1; a < anchors.length; a++) {
      out.push({ anchorIndex: anchors[a], legendIndex: legendFor.get(anchors[a]) ?? null });
    }
  }
  return out;
}

/** Pure: legend entries whose id has no chart anchor (chart was deleted → orphan legend). */
export function planOrphanLegends(entries: RepairEntry[]): number[] {
  const withAnchor = new Set(entries.filter((e) => e.hasModel).map((e) => e.id));
  const out: number[] = [];
  entries.forEach((e, i) => {
    if (e.part === "legend" && !withAnchor.has(e.id)) out.push(i);
  });
  return out;
}

/** Find duplicated chart ids on the slide and reassign fresh ids to the copies. */
export async function repairDuplicateChartIds(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide
): Promise<number> {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();

  const probes = shapes.items.map((s) => {
    const id = s.tags.getItemOrNullObject(TAG_ID);
    const model = s.tags.getItemOrNullObject(TAG_MODEL);
    const part = s.tags.getItemOrNullObject(TAG_PART);
    id.load("value, isNullObject");
    model.load("value, isNullObject");
    part.load("value, isNullObject");
    s.load("left, top, width, height");
    return { s, id, model, part };
  });
  await context.sync();

  const refs = probes.filter((p) => !p.id.isNullObject);
  const entries: RepairEntry[] = refs.map((p) => ({
    id: p.id.value,
    hasModel: !p.model.isNullObject,
    part: p.part.isNullObject ? null : p.part.value,
    cx: p.s.left + p.s.width / 2,
    cy: p.s.top + p.s.height / 2,
  }));

  const plan = planDuplicateRepair(entries);
  const orphans = planOrphanLegends(entries);
  if (plan.length === 0 && orphans.length === 0) return 0;

  // Next free "Chart N" number, so a pasted copy doesn't reuse the original's name.
  let maxNum = 0;
  for (const r of refs) {
    if (r.model.isNullObject) continue;
    try {
      const nm = (JSON.parse(r.model.value) as ChartModel).name ?? "";
      const match = /(\d+)\s*$/.exec(nm);
      if (match) maxNum = Math.max(maxNum, Number(match[1]));
    } catch {
      /* ignore */
    }
  }

  for (const asg of plan) {
    const fresh = newId();
    const anchor = refs[asg.anchorIndex];
    anchor.s.tags.add(TAG_ID, fresh); // PowerPoint tag add() upserts by key
    try {
      const m = JSON.parse(anchor.model.value) as ChartModel;
      m.id = fresh;
      m.name = `Chart ${++maxNum}`; // the pasted copy gets the next free number
      anchor.s.tags.add(TAG_MODEL, JSON.stringify(m));
    } catch {
      // leave a malformed model alone; the id split still de-conflicts editing
    }
    if (asg.legendIndex != null) refs[asg.legendIndex].s.tags.add(TAG_ID, fresh);
  }

  // Delete legend groups left behind when their chart was deleted in PowerPoint.
  for (const i of orphans) refs[i].s.delete();

  await context.sync();
  return plan.length + orphans.length;
}

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

/** Bounding box of the shapes of a chart belonging to a given part ("chart"|"legend"). */
export async function getPartBox(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  id: string,
  part: string
): Promise<ChartBox | null> {
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  const probes = shapes.items.map((s) => {
    const tagId = s.tags.getItemOrNullObject(TAG_ID);
    const tagPart = s.tags.getItemOrNullObject(TAG_PART);
    tagId.load("value, isNullObject");
    tagPart.load("value, isNullObject");
    s.load("left, top, width, height");
    return { s, tagId, tagPart };
  });
  await context.sync();
  const mine = probes.filter(
    (p) => !p.tagId.isNullObject && p.tagId.value === id && !p.tagPart.isNullObject && p.tagPart.value === part
  );
  if (mine.length === 0) return null;
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
  for (const { s } of mine) {
    left = Math.min(left, s.left);
    top = Math.min(top, s.top);
    right = Math.max(right, s.left + s.width);
    bottom = Math.max(bottom, s.top + s.height);
  }
  return { left, top, width: right - left, height: bottom - top };
}

/** Shift all shapes of a chart part by (dx, dy) — used to restore a moved legend. */
export async function translatePart(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  id: string,
  part: string,
  dx: number,
  dy: number
): Promise<void> {
  if (Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5) return;
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  const probes = shapes.items.map((s) => {
    const tagId = s.tags.getItemOrNullObject(TAG_ID);
    const tagPart = s.tags.getItemOrNullObject(TAG_PART);
    tagId.load("value, isNullObject");
    tagPart.load("value, isNullObject");
    s.load("left, top");
    return { s, tagId, tagPart };
  });
  await context.sync();
  for (const { s, tagId, tagPart } of probes) {
    if (!tagId.isNullObject && tagId.value === id && !tagPart.isNullObject && tagPart.value === part) {
      s.left += dx;
      s.top += dy;
    }
  }
  await context.sync();
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
