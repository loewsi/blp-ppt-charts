import { defineConfig } from "vite";
import * as devCerts from "office-addin-dev-certs";

// Office requires the add-in to be served over HTTPS, even in development.
// `npm run certs` installs a locally-trusted certificate that this reads.
export default defineConfig(async () => {
  let https: { key: Buffer; cert: Buffer } | undefined;
  try {
    const opts = await devCerts.getHttpsServerOptions();
    https = { key: opts.key, cert: opts.cert };
  } catch {
    console.warn("[vite] No dev certificate found. Run `npm run certs` first.");
  }

  return {
    // GitHub Pages serves a project site under /<repo>/, so the build needs
    // that as its base. Dev/localhost stays at "/". CI sets BASE_PATH.
    base: process.env.BASE_PATH ?? "/",
    root: ".",
    server: { host: "localhost", port: 3000, strictPort: true, https },
    preview: { host: "localhost", port: 3000, strictPort: true, https },
    build: {
      outDir: "dist",
      emptyOutDir: true,
      rollupOptions: {
        input: {
          taskpane: "taskpane.html",
          commands: "commands.html",
        },
      },
    },
  };
});
