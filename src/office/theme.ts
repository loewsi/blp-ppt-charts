// Reads the presentation master's accent colors live (PowerPointApi 1.10).
// Falls back to an empty array on older hosts so callers can use a preset.
export async function loadMasterAccents(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide
): Promise<string[]> {
  try {
    const scheme = slide.themeColorScheme;
    const keys = [
      PowerPoint.ThemeColor.accent1,
      PowerPoint.ThemeColor.accent2,
      PowerPoint.ThemeColor.accent3,
      PowerPoint.ThemeColor.accent4,
      PowerPoint.ThemeColor.accent5,
      PowerPoint.ThemeColor.accent6,
    ];
    const results = keys.map((k) => scheme.getThemeColor(k));
    await context.sync();
    return results.map((r) => normalizeHex(r.value)).filter(Boolean);
  } catch {
    return [];
  }
}

function normalizeHex(v: string): string {
  if (!v) return "";
  const s = v.trim();
  return s.startsWith("#") ? s.toUpperCase() : "#" + s.toUpperCase();
}
