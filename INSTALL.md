# Installing SlideChart

There are two very different situations:

- **You're an end user** who just wants to use the add-in → once it's published to
  **AppSource**, it's a one-click install (see *A*). Until then, use one of the
  sideload methods below.
- **You're testing/developing** or handing it to a few people before it's in the store →
  sideload it (*B* web, *C* Windows desktop, *D* Mac).

Distribution to many people (inside or outside BLP) is *E* (AppSource) or, for BLP staff
only, *F* (M365 admin). **Sideloading is not how you distribute to everyone** — it's for
testing or a handful of technical users.

---

## A. End user — from the store (once published)

1. In PowerPoint: **Insert** tab → **Add-ins** → **Store** (or **Get Add-ins**).
2. Search **"SlideChart"**.
3. Click **Add**. Done — no admin, no files, works on Windows/Mac/web.

---

## B. PowerPoint on the web (easiest to test — no admin)

1. Go to **office.com** → sign in (BLP *or* personal Microsoft account) → open **PowerPoint** → new presentation.
2. **Insert** tab → **Add-ins**.
3. In the dialog, click **Upload My Add-in** (top-right).
4. **Browse** → choose the `manifest.xml` file → **Upload**.
5. The **SlideChart** button appears on the **Home** tab.

---

## C. Windows desktop (trusted catalog)

Windows desktop has no "Upload" button, so the manifest must live in a **shared folder**
registered as a trusted catalog.

**Step 1 — Put the manifest in a shared folder** (one-time, needs admin once)
1. Create a folder, e.g. `C:\PPTAddins`, and copy `manifest.xml` into it.
2. Right-click the folder → **Properties** → **Sharing** → **Share…** → add **Everyone** → **Share**.
   (Or in an **Administrator** PowerShell:
   `New-SmbShare -Name ppt -Path C:\PPTAddins -FullAccess Everyone`)
3. Note the network path shown, e.g. `\\YOUR-PC\ppt`.

**Step 2 — Register the catalog in PowerPoint** (one-time)
1. PowerPoint → **File** → **Options** → **Trust Center** → **Trust Center Settings…**
2. **Trusted Add-in Catalogs**.
3. In **Catalog Url**, paste the network path (e.g. `\\YOUR-PC\ppt`) → **Add catalog**.
4. Tick the **Show in Menu** checkbox on the row that appears → **OK** → **OK**.
5. **Close and reopen PowerPoint** (required).

**Step 3 — Insert it**
1. **Insert** tab → **Add-ins** (or **My Add-ins** → **See All**).
2. Open the **SHARED FOLDER** tab → select **SlideChart** → **Add**.
3. The **SlideChart** button appears on the **Home** tab.

---

## D. Mac desktop

1. In Finder, press **⌘⇧G** and go to:
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
   (create the `wef` folder if it doesn't exist).
2. Copy `manifest.xml` into it.
3. Restart PowerPoint → **Insert** → **Add-ins** → **My Add-ins** → **SlideChart**.

---

## E. Distribute to everyone via AppSource (the real distribution channel)

Reaches anyone, inside **or** outside BLP, one-click, all platforms. High-level:
1. Create a free **Microsoft Partner Center** account.
2. Prepare required assets: privacy-policy URL, support URL, icons, screenshots, description.
3. Submit `manifest.xml` (pointing at the public GitHub Pages host) for validation.
4. Microsoft reviews (typically a few days to ~2 weeks). Once approved, it's in the store.

---

## F. Distribute to BLP staff only (M365 admin)

For internal-only rollout, a BLP Microsoft 365 admin can push it centrally:
1. **Microsoft 365 admin center** → **Settings** → **Integrated apps** → **Upload custom apps**.
2. Provide the `manifest.xml` (or its URL) and assign to users/groups.
3. It appears automatically for those users — no action on their part.
*(This covers BLP staff only; external people still need AppSource.)*

---

## Troubleshooting

- **Add-in doesn't appear (desktop):** confirm you restarted PowerPoint after adding the
  catalog, and that the network path is a real **share** (`\\PC\name`), not a local path.
- **Blank task pane / won't load:** the web host must be reachable. Open the manifest's
  `taskpane.html` URL in a browser — it should load.
- **"This add-in could not be started":** usually a manifest URL that isn't publicly
  reachable over HTTPS.

## Uninstall

- **Web / Mac:** remove via **My Add-ins** (web) or delete the file from `wef` (Mac).
- **Windows desktop:** remove the catalog under **Trust Center → Trusted Add-in Catalogs**,
  or clear the Office cache at
  `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
