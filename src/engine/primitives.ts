// A chart-type-agnostic drawing language. Every layout function emits these,
// and render.ts knows how to turn each one into a native PowerPoint shape.
// Adding waterfall later = a new layout that emits the same primitives
// (rects for bars, lines for connectors, text for labels) - render.ts is untouched.

export type Align = "left" | "center" | "right";

// Semantic identity of a rendered shape, so each primitive maps back to its
// chart-model role (enables sub-object selection and update-in-place later).
export type ObjectType =
  | "chartContainer"
  | "plotArea"
  | "segment"
  | "segmentLabel"
  | "totalLabel"
  | "categoryLabel"
  | "legend"
  | "legendEntry"
  | "valueAxis"
  | "gridline"
  | "baseline"
  | "connector"
  | "cagrArrow"
  | "differenceArrow"
  | "valueLine"
  | "lineSeries"
  | "lineMarker"
  | "slice"
  | "sliceLabel"
  | "point"
  | "quadrant";

export interface ShapeMeta {
  objectType: ObjectType;
  seriesIndex?: number;
  categoryIndex?: number;
}

export interface RectPrimitive {
  kind: "rect";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string; // "#RRGGBB"
  rounded?: boolean; // render as a rounded rectangle (e.g. a label bubble)
  meta?: ShapeMeta;
}

export interface TextPrimitive {
  kind: "text";
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  color: string;
  size: number; // points
  bold: boolean;
  align: Align;
  family?: string; // font family; defaults to Roboto in the adapter
  bg?: string; // optional background fill (e.g. segment color behind a small label)
  autofit?: boolean; // let PowerPoint size the box to the text (used for legend)
  meta?: ShapeMeta;
}

export interface LinePrimitive {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  weight: number; // points
  dashed?: boolean; // rendered as a dashed line (used for leaders / CAGR spans)
  meta?: ShapeMeta;
}

/** A line from (x1,y1)→(x2,y2) with a filled triangular head at the "to" end
 *  (and optionally the "from" end). Rendered as a line body + rotated triangle(s). */
export interface ArrowPrimitive {
  kind: "arrow";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  weight: number; // points
  headSize?: number; // triangle height in points (default 7)
  doubleHeaded?: boolean; // head at both ends
  meta?: ShapeMeta;
}

/** An isosceles triangle in its bounding box, rotated about the box center.
 *  Used to build pie/doughnut slices as a fan of thin facets. */
export interface TrianglePrimitive {
  kind: "triangle";
  x: number; // bounding-box top-left
  y: number;
  w: number;
  h: number;
  rotation: number; // degrees, clockwise
  fill: string;
  meta?: ShapeMeta;
}

/** An ellipse/circle in its bounding box (scatter points, doughnut hole). */
export interface EllipsePrimitive {
  kind: "ellipse";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  meta?: ShapeMeta;
}

export type Primitive =
  | RectPrimitive
  | TextPrimitive
  | LinePrimitive
  | ArrowPrimitive
  | TrianglePrimitive
  | EllipsePrimitive;
