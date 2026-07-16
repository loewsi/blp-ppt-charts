// Generates manifest.prod.xml from manifest.xml by repointing the localhost dev
// URLs at a public host (e.g. GitHub Pages).
//
// Usage:
//   PUBLIC_URL=https://<user>.github.io/<repo> npm run gen-manifest
//   npm run gen-manifest -- https://<user>.github.io/<repo>
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const publicUrl = (process.env.PUBLIC_URL || process.argv[2] || "").replace(/\/+$/, "");

if (!publicUrl) {
  console.error("Missing PUBLIC_URL. Example:");
  console.error("  PUBLIC_URL=https://loewsi.github.io/blp-ppt-charts npm run gen-manifest");
  process.exit(1);
}

const origin = new URL(publicUrl).origin; // AppDomain must be the origin, not a path
const src = readFileSync(join(root, "manifest.xml"), "utf8");

let out = src.replace(
  "<AppDomain>https://localhost:3000</AppDomain>",
  `<AppDomain>${origin}</AppDomain>`
);
out = out.replaceAll("https://localhost:3000", publicUrl);

writeFileSync(join(root, "manifest.prod.xml"), out);
console.log(`Wrote manifest.prod.xml -> ${publicUrl}`);
