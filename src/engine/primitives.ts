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
  | "cagrArrow"
  | "differenceArrow"
  | "valueLine";

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
  meta?: ShapeMeta;
}

export type Primitive = RectPrimitive | TextPrimitive | LinePrimitive;
