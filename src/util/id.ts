export function newId(): string {
  const c = globalThis.crypto;
  if (c && typeof c.randomUUID === "function") return "c_" + c.randomUUID();
  return "c_" + Math.random().toString(36).slice(2) + Date.now().toString(36);
}
