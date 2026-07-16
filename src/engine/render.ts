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
  const created: PowerPoint.Shape[] = [];

  for (const p of prims) {
    created.push(makeShape(shapes, p));
  }

  // Tag every shape with the chart id; stamp the model onto the first one.
  created.forEach((shape, i) => {
    shape.tags.add(TAG_ID, model.id);
    if (i === 0) {
      shape.tags.add(TAG_ANCHOR, "1");
      shape.tags.add(TAG_MODEL, JSON.stringify(model));
    }
  });

  await context.sync();
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
    const s = shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: Math.min(p.x1, p.x2),
      top: Math.min(p.y1, p.y2),
      width: Math.abs(p.x2 - p.x1),
      height: Math.abs(p.y2 - p.y1),
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
