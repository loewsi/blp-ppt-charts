# ADR-0001: Core architecture for SlideChart

Status: Accepted (2026-07-18) · Context: MVP, free/anonymous, one chart family (column/bar).

## Decisions and rationale

1. **Office.js add-in (not VSTO/COM).**
   Cross-platform (Windows + Mac + web) is a hard requirement and the user base includes
   individuals and unrelated organisations. VSTO/COM is Windows-only. Office.js is the only
   technology that meets the reach requirement with one codebase.

2. **Accept that cross-platform outweighs guaranteed cold-start offline editing.**
   Office.js UIs load from HTTPS, so editing after an *offline restart* cannot be reliably
   guaranteed. We prioritise reach and honesty over a false offline claim. Mitigations: all
   operations are local once loaded; charts remain fully usable as native shapes offline.

3. **All chart operations are local TypeScript.**
   No backend is required for normal use. This protects privacy, removes latency and outages
   from the core loop, and keeps the product free/anonymous. Remote services (release check,
   telemetry, AI, licensing) are optional and must never gate chart actions.

4. **Charts render as native PowerPoint shapes.**
   So they are visible/printable/exportable without the add-in and for recipients who don't
   have it, and remain basically editable in PowerPoint. Shapes are the *output*; the stored
   model is the *source of truth*.

5. **Chart model stored inside the presentation (shape tags).**
   Keeps presentation data in the user's file (privacy proposition) and lets a chart be
   reconstructed and copied with the file. Custom XML parts are unavailable in PowerPoint
   Office.js; document settings don't travel per-chart; tags do. Large models will be chunked.

6. **Separate domain model → render plan → adapter.**
   The engine (model, layout, number format, render plan) is platform-neutral and has **no**
   dependency on React, the task-pane UI, Office.js, auth, or billing. Only `render.ts`
   (the adapter) touches Office.js. This enables unit testing and future reuse (waterfall etc.).

7. **No accounts, billing, orgs, or licensing in the MVP.**
   Explicitly out of scope. The architecture doesn't preclude them later (they'd be optional
   services behind interfaces), but building them now would blur scope and add a backend the
   core doesn't need.

8. **The VBA/RibbonX productivity tool stays a separate product.**
   Different execution model, deployment, and platform behaviour. Merging now would blur scope.
   The add-in is designed so general productivity commands *could* be added later.

## Known limitations / risks

- Cold-start offline editing not guaranteed (see decision 2).
- `addGroup` / `getSelectedShapes` / selection events / tag-size limits vary by host; detected
  at runtime with fallbacks; **Mac not yet runtime-verified**.
- `Update` currently delete-and-recreates shapes → suboptimal undo; in-place update planned.
- Copying a chart can duplicate `chartId`; detection/repair not yet built.
- No transactional render guard yet.

## Not chosen (and why)

- **VSTO/COM add-in** — Windows-only, fails reach requirement.
- **Native PowerPoint charts via API** — PowerPoint Office.js has no chart-creation API.
- **Custom XML parts for storage** — not exposed in PowerPoint Office.js.
- **Monorepo `packages/*` split now** — current `src/*` modules already isolate concerns
  (engine is Office-free); a formal package split adds tooling overhead without near-term
  benefit. Revisit if the engine is reused by a second app.
