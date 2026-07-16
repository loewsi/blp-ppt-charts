// A chart-type-agnostic drawing language. Every layout function emits these,
// and render.ts knows how to turn each one into a native PowerPoint shape.
// Adding waterfall later = a new layout that emits the same primitives
// (rects for bars, lines for connectors, text for labels) - render.ts is untouched.

export type Align = "left" | "center" | "right";

export interface RectPrimitive {
  kind: "rect";
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string; // "#RRGGBB"
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
}

export interface LinePrimitive {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  weight: number; // points
}

export type Primitive = RectPrimitive | TextPrimitive | LinePrimitive;
