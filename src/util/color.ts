// Small color helpers for generating a shade ramp from one base color.

function parse(hex: string): { r: number; g: number; b: number } {
  const h = (hex || "").replace("#", "");
  return {
    r: parseInt(h.slice(0, 2), 16) || 0,
    g: parseInt(h.slice(2, 4), 16) || 0,
    b: parseInt(h.slice(4, 6), 16) || 0,
  };
}

function toHex(r: number, g: number, b: number): string {
  const c = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");
  return `#${c(r)}${c(g)}${c(b)}`;
}

/** Mix a color toward white (t > 0) or black (t < 0). t in [-1, 1]; 0 = unchanged. */
export function mixColor(hex: string, t: number): string {
  const { r, g, b } = parse(hex);
  if (t >= 0) return toHex(r + (255 - r) * t, g + (255 - g) * t, b + (255 - b) * t);
  const k = 1 + t;
  return toHex(r * k, g * k, b * k);
}

/** N distinct shades of a base color, spread from light to dark. */
export function shadesFrom(hex: string, n: number): string[] {
  if (n <= 1) return [mixColor(hex, 0)];
  const out: string[] = [];
  for (let i = 0; i < n; i++) out.push(mixColor(hex, 0.55 - i * (0.9 / (n - 1))));
  return out;
}
