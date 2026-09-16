# ppt-shortcuts

Personal Office.js add-in that binds keyboard shortcuts to shape-sizing commands in
PowerPoint for macOS. Modeled on Efficient Elements; built and owned locally.

**Status: Phase 1 feasibility spike.** Five commands, real keystrokes, end-to-end.
Stop and report before Phase 2 (see the build brief).

## Commands

| Key | Action ID | What it does |
|---|---|---|
| `⌃⇧⌥W` | `matchWidth` | target width := reference width |
| `⌃⇧⌥H` | `matchHeight` | target height := reference height |
| `⌃⇧⌥B` | `matchBoth` | both, non-proportional |
| `⌃⇧⌥I` | `fitInside` | proportional contain: `k = min(refW/w, refH/h)` |
| `⌃⇧⌥O` | `fillOutside` | proportional cover: `k = max(refW/w, refH/h)` |
| `⌃⇧⌥P` | `pickupReference` | store the selected shape as reference (only meaningful in `pickup` mode) |

`Ctrl` in the shortcuts file is the physical Control key on a Mac (`⌃`), and `Alt` is Option (`⌥`).

Behaviour switches live in `CONFIG` at the top of [docs/commands.js](docs/commands.js):

- `REFERENCE_MODE` — `"lastSelected"` (Efficient Elements model) or `"pickup"`. **Phase 0 decides this.**
- `RECENTER_ON_REF` (default `true`) — fit/fill land on the reference's centre.
- `KEEP_CENTER` (default `true`) — match commands grow around the target's own centre.

The task pane exposes runtime toggles for all three so you can experiment without redeploying.

## Layout

```
ppt-shortcuts/
├── docs/                  # served by GitHub Pages (https://cueland.github.io/ppt-shortcuts/)
│   ├── taskpane.html      # shared-runtime host page + Phase 0 probe UI
│   ├── commands.js        # Office.actions.associate handlers, geometry, resolveReference()
│   ├── shortcuts.json     # action IDs → key combos
│   └── assets/            # ribbon icons
├── manifest.xml           # sideloaded locally — NOT served
├── scripts/sideload.sh    # copy manifest into PowerPoint's wef dir and restart PowerPoint
└── scripts/clear-cache.sh # nuke the add-in cache when a change refuses to show up
```

The action IDs in `shortcuts.json` (`actions[].id` and `shortcuts[].action`) and the keys of the
`commands` object in `commands.js` must match exactly — `Office.onReady` associates every key of
`commands`, so adding a command means adding it in both places.

## One-time setup

1. **PowerPoint ≥ 16.105.2 (26012530).** PowerPoint → About PowerPoint. Below this build,
   shortcut registration silently does nothing.
2. **Hosting.** Push this repo to GitHub as a **public** repo named `ppt-shortcuts`, then
   Settings → Pages → Source: `main` branch, `/docs` folder. Confirm
   `https://cueland.github.io/ppt-shortcuts/shortcuts.json` loads in a browser.
   If the account or repo name differ, search-and-replace the base URL in `manifest.xml`.
3. **Sideload.**
   ```bash
   scripts/sideload.sh
   ```
   This copies `manifest.xml` to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
   (note the lowercase `p` in `Powerpoint`) and restarts PowerPoint.
4. Open any presentation. A **Shortcuts › ppt-shortcuts** button appears on the Home tab.
   The diagnostics line at the top of the pane reports `Office.context.diagnostics.version`
   and which requirement sets the host supports.

## Phase 0 — selection-order probe (do this first)

Everything Efficient-Elements-like ("the last shape you clicked is the master") depends on whether
`getSelectedShapes()` returns shapes in **selection order** or **z-order**. VBA returns z-order.

1. Open the task pane → **Create probe shapes**. Adds A, B, C and brings A to the front, so
   z-order ≠ any natural click order.
2. Click on the slide, then shift-click to select **C → A → B**.
3. **Dump selection.** The table shows array index, `name`, `id`, `zOrderPosition`.
   - Array order **C, A, B** → selection order is preserved → keep `REFERENCE_MODE = "lastSelected"`.
   - Array order tracks `zOrderPosition` → set `REFERENCE_MODE = "pickup"` in `CONFIG`.
     The `pickupReference` binding is already registered, so nothing else changes.

Repeat with a different order (e.g. B → C → A) to rule out a coincidence.

## Phase 1 — acceptance

Work through these with the task pane open so the log shows before/after geometry:

1. [ ] PowerPoint version ≥ 16.105.2 (diagnostics line).
2. [ ] Add-in loads; version reported.
3. [ ] Phase 0 answered; `REFERENCE_MODE` set accordingly.
4. [ ] All five shortcuts fire from the keyboard **with the task pane closed**.
5. [ ] Geometry: **Create acceptance shapes** adds a 100×200 target and a 300×150 reference.
   Select target then reference (or pick up the reference), then:
   - `⌃⇧⌥I` → target becomes **75×150**
   - `⌘Z`, then `⌃⇧⌥O` → target becomes **300×600**
   The log prints a ⚠ if the read-back geometry differs from the computed value — that would
   indicate PowerPoint is honouring "lock aspect ratio" on `width`/`height` writes, which we'd
   need to work around.
6. [ ] **Show registered shortcuts** reports no conflicts, and pressing each combo shows no
   chooser dialog.

## Iteration loop

edit → commit → push → wait ~1 min for Pages → `scripts/sideload.sh` (restarts PowerPoint).

**Office caches add-in files aggressively.** If a change doesn't appear:

```bash
scripts/clear-cache.sh && scripts/sideload.sh
```

When `shortcuts.json` changes, also bump the `?v=` on the `ExtendedOverrides` URL in `manifest.xml`.
Budget for this — it looks exactly like a code bug the first time.

## Debugging

- Command handlers log to the task pane's **Log** panel (last 200 lines, kept even while the
  pane is closed because the shared runtime stays alive) and to `console.log`.
- To attach Safari's Web Inspector to the add-in's webview:
  `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`, restart
  PowerPoint, then right-click inside the task pane → Inspect Element.

## Out of scope (no API surface)

Adjustment handles, rotation, screen colour picker, email/save selected slides, animations.
See the feasibility inventory alongside the build brief before scoping Phase 2.
