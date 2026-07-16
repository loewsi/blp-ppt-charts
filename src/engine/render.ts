import type { ChartModel } from "../model/chartModel";
import type { Primitive, Align } from "./primitives";
import { computeLayout } from "./layout";
import { TAG_ID, TAG_ANCHOR, TAG_MODEL } from "./tags";

/** Draw a chart's primitives as native PowerPoint shapes and tag them so the
 *  chart can be found, re-read and re-drawn later. */
export async function drawChart(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  model: ChartModel
): Promise<void> {
  const prims = computeLayout(model);
  const shapes = slide.shapes;
  const created = prims.map((p) => makeShape(shapes, p));
  await context.sync(); // shapes must exist before they can be grouped

  const canGroup =
    created.length > 1 &&
    Office.context.requirements.isSetSupported("PowerPointApi", "1.8");

  if (canGroup) {
    // One selectable/movable object; tag the group with the model.
    const group = shapes.addGroup(created);
    tagChart(group, model);
  } else {
    // Fallback for hosts without grouping (PowerPointApi < 1.8): tag each shape.
    created.forEach((shape, i) => {
      shape.tags.add(TAG_ID, model.id);
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
    // Horizontal/vertical lines render as thin rectangles — deterministic and
    // crisp on Windows + Mac (PowerPoint connectors misdraw zero-height boxes
    // as diagonals). True diagonals fall back to a real connector.
    const dx = Math.abs(p.x2 - p.x1);
    const dy = Math.abs(p.y2 - p.y1);
    const isHorizontal = dy < 0.5;
    const isVertical = dx < 0.5;
    if (isHorizontal || isVertical) {
      const w = isHorizontal ? dx : p.weight;
      const h = isHorizontal ? p.weight : dy;
      const s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: Math.min(p.x1, p.x2) - (isVertical ? p.weight / 2 : 0),
        top: Math.min(p.y1, p.y2) - (isHorizontal ? p.weight / 2 : 0),
        width: w,
        height: h,
      });
      s.fill.setSolidColor(p.color);
      s.lineFormat.visible = false;
      return s;
    }
    const s = shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: Math.min(p.x1, p.x2),
      top: Math.min(p.y1, p.y2),
      width: dx,
      height: dy,
    });
    s.lineFormat.color = p.color;
    s.lineFormat.weight = p.weight;
    s.lineFormat.visible = true;
    return s;
  }

  // text
  const s = shapes.addTextBox(p.text, { left: p.x, top: p.y, width: p.w, height: p.h });
  s.fill.clear();
  s.lineFormat.visible = false;
  const tf = s.textFrame;
  tf.topMargin = 0;
  tf.bottomMargin = 0;
  tf.leftMargin = 0;
  tf.rightMargin = 0;
  const tr = tf.textRange;
  tr.font.size = p.size;
  tr.font.bold = p.bold;
  tr.font.color = p.color;
  tr.font.name = "Roboto";
  tr.paragraphFormat.horizontalAlignment = hAlign(p.align);
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
