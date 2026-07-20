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
  const created = prims.map((p) => makeShape(shapes, p));
  await context.sync(); // shapes must exist before they can be grouped

  // The legend is its own group so it can be moved independently and doesn't
  // stretch when the chart is resized.
  const isLegend = (i: number) => {
    const t = prims[i].meta?.objectType;
    return t === "legend" || t === "legendEntry";
  };
  const chartShapes = created.filter((_, i) => !isLegend(i));
  const legendShapes = created.filter((_, i) => isLegend(i));

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
    created.forEach((shape, i) => {
      shape.tags.add(TAG_ID, model.id);
      shape.tags.add(TAG_PART, isLegend(i) ? "legend" : "chart");
      if (i === 0) {
        shape.tags.add(TAG_ANCHOR, "1");
        shape.tags.add(TAG_MODEL, JSON.stringify(model));
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

function makeShape(shapes: PowerPoint.ShapeCollection, p: Primitive): PowerPoint.Shape {
  if (p.kind === "rect") {
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: p.x,
      top: p.y,
      width: p.w,
      height: p.h,
    });
    s.fill.setSolidColor(p.fill);
    s.lineFormat.visible = false;
    return s;
  }

  if (p.kind === "line") {
    // All lines render as thin rectangles (deterministic on Windows + Mac).
    // Horizontal/vertical are axis-aligned; diagonals are a rotated rectangle,
    // which avoids PowerPoint connectors misdrawing the slant direction.
    const adx = Math.abs(p.x2 - p.x1);
    const ady = Math.abs(p.y2 - p.y1);
    if (adx < 0.5 || ady < 0.5) {
      const horizontal = ady < 0.5;
      const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: Math.min(p.x1, p.x2) - (horizontal ? 0 : p.weight / 2),
        top: Math.min(p.y1, p.y2) - (horizontal ? p.weight / 2 : 0),
        width: horizontal ? adx : p.weight,
        height: horizontal ? p.weight : ady,
      });
      s.fill.setSolidColor(p.color);
      s.lineFormat.visible = false;
      return s;
    }
    // Diagonal: a thin rectangle of length |P1P2| rotated to the segment angle.
    const len = Math.hypot(p.x2 - p.x1, p.y2 - p.y1);
    const angle = (Math.atan2(p.y2 - p.y1, p.x2 - p.x1) * 180) / Math.PI;
    const midX = (p.x1 + p.x2) / 2;
    const midY = (p.y1 + p.y2) / 2;
    const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: midX - len / 2,
      top: midY - p.weight / 2,
      width: len,
      height: p.weight,
    });
    s.fill.setSolidColor(p.color);
    s.lineFormat.visible = false;
    s.rotation = angle;
    return s;
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
  return s;
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
