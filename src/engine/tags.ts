// PowerPoint shape tag keys. Keys are stored upper-cased by PowerPoint, so we
// declare them upper-case to match on read-back.
export const TAG_ID = "BLPCHARTID"; // present on every shape belonging to a chart
export const TAG_ANCHOR = "BLPCHARTANCHOR"; // marks the shape that carries the model
export const TAG_MODEL = "BLPCHARTMODEL"; // JSON-serialized ChartModel (on the anchor)
