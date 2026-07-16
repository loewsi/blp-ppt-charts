# SlideChart — PowerPoint add-in

Live-linked charts for PowerPoint, **cross-platform (Windows + Mac + web)**, built on
Office.js. This is a separate project from the VBA **QuickTools** add-in (which is
Windows-only). v1 ships **stacked column**; the engine is built so **waterfall** drops in
as one more layout file.

## Quick start (development)

```bash
npm install
npm run certs      # one-time: trusts https://localhost (accept the prompt)
npm run gen-icons  # regenerate placeholder icons if needed
npm run dev        # serves https://localhost:3000
```

Leave `npm run dev` running, then sideload `manifest.xml` (below). Open PowerPoint →
**Home** tab → **SlideChart** group → **Chart editor** to open the task pane.

Other scripts: `npm run build` (type-check + production bundle to `dist/`),
`npm run validate` (validate the manifest).

## Sideloading (dev)

The **manifest** points at `https://localhost:3000`, so the dev server must be running on
the same machine you test from.

**Windows — shared-folder catalog**
1. Put `manifest.xml` in a folder you share (right-click → Properties → Sharing).
2. PowerPoint → File → Options → Trust Center → Trust Center Settings → **Trusted Add-in
   Catalogs** → paste the `\\PC\Share` URL → Add → tick **Show in Menu** → OK → restart PowerPoint.
3. Insert → **My Add-ins** → **Shared Folder** → **SlideChart**.

**Mac — wef folder**
1. Copy `manifest.xml` to
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
   (create `wef` if missing).
2. Restart PowerPoint → Insert → **My Add-ins** → **SlideChart**.

## Deployment — GitHub Pages (public, free)

An add-in is a small web app: PowerPoint loads its UI from a URL, so the files must live on
a public HTTPS host. We use **GitHub Pages**. Hosting it publicly (vs internal SharePoint)
is what lets you use it personally *and* share it with people outside BLP.

**One-time setup**
1. Create a GitHub repo named `blp-ppt-charts` (public).
2. `git remote add origin https://github.com/<user>/blp-ppt-charts.git`
3. `git push -u origin main`
4. On GitHub: **Settings → Pages → Build and deployment → Source = GitHub Actions**.

The workflow in `.github/workflows/deploy.yml` then builds and publishes on every push to
`main`. It derives all URLs automatically from the repo, so nothing is hard-coded. After the
first run your add-in is live at:

```
https://<user>.github.io/blp-ppt-charts/taskpane.html
manifest:  https://<user>.github.io/blp-ppt-charts/manifest.xml
```

**Installing it (you or anyone you share with)** — download that `manifest.xml` and sideload
it (see *Sideloading* above; the wef-folder method on Mac, trusted-catalog on Windows). To
share externally, just send that manifest link. For one-click install for anyone, publish to
**AppSource** later (bigger process; not needed for personal/small-scale use).

**Manual build without CI (optional)**
```bash
BASE_PATH=/blp-ppt-charts/ npm run build
PUBLIC_URL=https://<user>.github.io/blp-ppt-charts npm run gen-manifest
```

## How it works

```
data grid (task pane)  →  ChartModel  →  layout*.ts → Primitive[]  →  render.ts → native PPT shapes
                                             ↑                              ↓
                                      persistence.ts  ←── model stored in a shape TAG (saved in the .pptx)
```

- `src/model/chartModel.ts` — the data model (persisted inside the presentation).
- `src/engine/primitives.ts` — chart-type-agnostic drawing language (rect / text / line).
- `src/engine/layoutStackedColumn.ts` — pure geometry: model → primitives.
- `src/engine/layout.ts` — registry mapping chart type → layout function (**add waterfall here**).
- `src/engine/render.ts` — primitives → native PowerPoint shapes + tags.
- `src/engine/persistence.ts` — find / read / delete charts via shape tags.
- `src/taskpane/taskpane.ts` — the UI controller (grid, insert / update / load).

**Live-linked** = each chart's model is JSON-stamped onto its first shape's tag, which is
saved in the `.pptx` and travels to Mac. *Update* reads the model back, deletes the old
shapes, and redraws — that's the "re-layout when numbers change" behaviour.

## Units

All geometry is in **points** (1 in = 72 pt). A 16:9 slide is 960 × 540 pt.
