// Deck-level agenda / table-of-contents. A pane list of chapters generates (and
// re-generates in place) an "Agenda" slide with an auto-numbered list. The two
// text boxes are tagged so a later run updates them instead of duplicating.
const AGENDA_TAG = "BLPAGENDA";
const DARK = "#001C54";

/** Build the numbered list text. */
export function numberedList(chapters: string[]): string {
  return chapters.map((c, i) => `${i + 1}.  ${c}`).join("\n");
}

/** Create the agenda slide, or update the existing one's title + list in place. */
export async function createOrUpdateAgenda(
  context: PowerPoint.RequestContext,
  chapters: string[]
): Promise<"created" | "updated"> {
  const clean = chapters.map((c) => c.trim()).filter(Boolean);
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();

  // Load every slide's shapes, then probe each shape for the agenda tag.
  const shapeLists = slides.items.map((sl) => {
    sl.shapes.load("items");
    return sl.shapes;
  });
  await context.sync();

  const probes: { shape: PowerPoint.Shape; tag: PowerPoint.Tag }[] = [];
  for (const shapes of shapeLists) {
    for (const shape of shapes.items) {
      const tag = shape.tags.getItemOrNullObject(AGENDA_TAG);
      tag.load("value, isNullObject");
      probes.push({ shape, tag });
    }
  }
  await context.sync();

  const listShape = probes.find((p) => !p.tag.isNullObject && p.tag.value === "list")?.shape;
  const titleShape = probes.find((p) => !p.tag.isNullObject && p.tag.value === "title")?.shape;

  if (listShape) {
    listShape.textFrame.textRange.text = numberedList(clean);
    if (titleShape) titleShape.textFrame.textRange.text = "Agenda";
    await context.sync();
    return "updated";
  }

  // No agenda yet — add a slide at the end and lay out the title + list.
  slides.add();
  await context.sync();
  slides.load("items");
  await context.sync();
  const slide = slides.items[slides.items.length - 1];

  const title = slide.shapes.addTextBox("Agenda", { left: 40, top: 30, width: 620, height: 50 });
  styleText(title, 28, true);
  title.tags.add(AGENDA_TAG, "title");

  const list = slide.shapes.addTextBox(numberedList(clean), { left: 50, top: 104, width: 600, height: 380 });
  styleText(list, 18, false);
  list.tags.add(AGENDA_TAG, "list");

  await context.sync();
  return "created";
}

function styleText(shape: PowerPoint.Shape, size: number, bold: boolean): void {
  shape.fill.clear();
  shape.lineFormat.visible = false;
  const tr = shape.textFrame.textRange;
  tr.font.size = size;
  tr.font.bold = bold;
  tr.font.color = DARK;
  tr.font.name = "Roboto";
}
