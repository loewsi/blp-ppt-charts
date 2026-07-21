import type { ChartModel } from "../model/chartModel";
import type { Primitive, Align } from "./primitives";
import { computeLayout } from "./layout";
import { TAG_ID, TAG_ANCHOR, TAG_MODEL, TAG_PART } from "./tags";

/** Draw a chart's primitives as native PowerPoint shapes and tag them so the
 *  chart can be found, re-read and re-drawn later. */
export async function drawChart(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  model: ChartModel
): Promise<void> {
  const raw = computeLayout(model);
  // Create fills/lines first and text last, so every label sits in FRONT of the
  // bars (z-order follows creation order within the group).
  const prims = [...raw.filter((p) => p.kind !== "text"), ...raw.filter((p) => p.kind === "text")];
  const shapes = slide.shapes;
  // A primitive can map to several native shapes (an arrow = line + head), so keep
  // each primitive's shapes together to preserve the legend/chart split below.
  const createdPer = prims.map((p) => makeShapes(shapes, p));
  await context.sync(); // shapes must exist before they can be grouped

  // The legend is its own group so it can be moved independently and doesn't
  // stretch when the chart is resized.
  const isLegend = (i: number) => {
    const t = prims[i].meta?.objectType;
    return t === "legend" || t === "legendEntry";
  };
  const chartShapes = createdPer.filter((_, i) => !isLegend(i)).flat();
  const legendShapes = createdPer.filter((_, i) => isLegend(i)).flat();

  const canGroup = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");

  if (canGroup && chartShapes.length > 1) {
    const chartGroup = shapes.addGroup(chartShapes);
    tagChart(chartGroup, model);
    if (legendShapes.length > 1) {
      const legendGroup = shapes.addGroup(legendShapes);
      legendGroup.tags.add(TAG_ID, model.id);
      legendGroup.tags.add(TAG_PART, "legend");
    } else if (legendShapes.length === 1) {
      legendShapes[0].tags.add(TAG_ID, model.id);
      legendShapes[0].tags.add(TAG_PART, "legend");
    }
  } else {
    // Fallback for hosts without grouping (PowerPointApi < 1.8): tag each shape.
    // Iterate per-primitive so the legend/chart split and the single anchor stay correct.
    let anchorDone = false;
    createdPer.forEach((arr, i) => {
      const part = isLegend(i) ? "legend" : "chart";
      for (const shape of arr) {
        shape.tags.add(TAG_ID, model.id);
        shape.tags.add(TAG_PART, part);
        if (!anchorDone && part === "chart") {
          shape.tags.add(TAG_ANCHOR, "1");
          shape.tags.add(TAG_MODEL, JSON.stringify(model));
          anchorDone = true;
        }
      }
    });
  }
  await context.sync();
}

function tagChart(shape: PowerPoint.Shape, model: ChartModel): void {
  shape.tags.add(TAG_ID, model.id);
  shape.tags.add(TAG_ANCHOR, "1");
  shape.tags.add(TAG_PART, "chart");
  shape.tags.add(TAG_MODEL, JSON.stringify(model));
}

/** A line segment as a thin rectangle (axis-aligned) or a rotated rectangle
 *  (diagonal) — deterministic on Windows + Mac, unlike PowerPoint connectors. */
function makeLine(
  shapes: PowerPoint.ShapeCollection,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  color: string,
  weight: number
): PowerPoint.Shape {
  const adx = Math.abs(x2 - x1);
  const ady = Math.abs(y2 - y1);
  if (adx < 0.5 || ady < 0.5) {
    const horizontal = ady < 0.5;
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: Math.min(x1, x2) - (horizontal ? 0 : weight / 2),
      top: Math.min(y1, y2) - (horizontal ? weight / 2 : 0),
      width: horizontal ? adx : weight,
      height: horizontal ? weight : ady,
    });
    s.fill.setSolidColor(color);
    s.lineFormat.visible = false;
    return s;
  }
  const len = Math.hypot(x2 - x1, y2 - y1);
  const angle = (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI;
  const midX = (x1 + x2) / 2;
  const midY = (y1 + y2) / 2;
  const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: midX - len / 2,
    top: midY - weight / 2,
    width: len,
    height: weight,
  });
  s.fill.setSolidColor(color);
  s.lineFormat.visible = false;
  s.rotation = angle;
  return s;
}

/** A filled triangular arrowhead whose apex sits at (tipX,tipY), pointing along dirDeg. */
function makeHead(
  shapes: PowerPoint.ShapeCollection,
  tipX: number,
  tipY: number,
  dirDeg: number,
  size: number,
  color: string
): PowerPoint.Shape {
  const w = size * 1.3;
  const h = size;
  const rad = (dirDeg * Math.PI) / 180;
  // The triangle's apex points "up"; its center sits half a height back from the tip.
  const cx = tipX - (h / 2) * Math.cos(rad);
  const cy = tipY - (h / 2) * Math.sin(rad);
  const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
    left: cx - w / 2,
    top: cy - h / 2,
    width: w,
    height: h,
  });
  s.fill.setSolidColor(color);
  s.lineFormat.visible = false;
  s.rotation = dirDeg + 90; // rotate "up" apex to point along the direction
  return s;
}

function makeShapes(shapes: PowerPoint.ShapeCollection, p: Primitive): PowerPoint.Shape[] {
  if (p.kind === "rect") {
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: p.x,
      top: p.y,
      width: p.w,
      height: p.h,
    });
    s.fill.setSolidColor(p.fill);
    s.lineFormat.visible = false;
    return [s];
  }

  if (p.kind === "line") {
    return [makeLine(shapes, p.x1, p.y1, p.x2, p.y2, p.color, p.weight)];
  }

  if (p.kind === "arrow") {
    const out = [makeLine(shapes, p.x1, p.y1, p.x2, p.y2, p.color, p.weight)];
    const angle = (Math.atan2(p.y2 - p.y1, p.x2 - p.x1) * 180) / Math.PI;
    const size = p.headSize ?? 7;
    out.push(makeHead(shapes, p.x2, p.y2, angle, size, p.color));
    if (p.doubleHeaded) out.push(makeHead(shapes, p.x1, p.y1, angle + 180, size, p.color));
    return out;
  }

  if (p.kind === "triangle") {
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
      left: p.x,
      top: p.y,
      width: p.w,
      height: p.h,
    });
    s.fill.setSolidColor(p.fill);
    s.lineFormat.visible = false;
    if (p.rotation) s.rotation = p.rotation;
    return [s];
  }

  if (p.kind === "ellipse") {
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, {
      left: p.x,
      top: p.y,
      width: p.w,
      height: p.h,
    });
    s.fill.setSolidColor(p.fill);
    s.lineFormat.visible = false;
    return [s];
  }

  // text
  const s = shapes.addTextBox(p.text, { left: p.x, top: p.y, width: p.w, height: p.h });
  if (p.bg) s.fill.setSolidColor(p.bg);
  else s.fill.clear();
  s.lineFormat.visible = false;
  const tf = s.textFrame;
  tf.topMargin = 0;
  tf.bottomMargin = 0;
  tf.leftMargin = 0;
  tf.rightMargin = 0;
  tf.verticalAlignment = "Middle"; // vertically center text in its box
  const tr = tf.textRange;
  tr.font.size = p.size;
  tr.font.bold = p.bold;
  tr.font.color = p.color;
  tr.font.name = p.family ?? "Roboto";
  tr.paragraphFormat.horizontalAlignment = hAlign(p.align);
  if (p.autofit) tf.autoSizeSetting = "AutoSizeShapeToFitText"; // box hugs the text exactly
  return [s];
}

function hAlign(a: Align): PowerPoint.ParagraphHorizontalAlignment {
  switch (a) {
    case "left":
      return PowerPoint.ParagraphHorizontalAlignment.left;
    case "right":
      return PowerPoint.ParagraphHorizontalAlignment.right;
    default:
      return PowerPoint.ParagraphHorizontalAlignment.center;
  }
}
