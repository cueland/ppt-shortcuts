# ChristiantialElements

Personal Office.js add-in that binds keyboard shortcuts to shape commands in PowerPoint
for macOS. Modeled on Efficient Elements; built and owned locally. Repo slug and Pages URL
stay `ppt-shortcuts`; the product name is ChristiantialElements.

**Status: Phase 1 passed** (five sizing commands fire from real keystrokes). Now adding
commands from the operation inventory in priority order, with in-pane key assignment.

## How to use it

Select the shapes you want to change, then shift-click the shape you want them to copy
**last**. Press a key. The last-selected shape is the reference (the Efficient Elements model).

Default keys (change them in the pane):

| Key | Command |
|---|---|
| `⌃⇧⌥W` | Match width |
| `⌃⇧⌥H` | Match height |
| `⌃⇧⌥B` | Match width and height |
| `⌃⇧⌥I` | Fit inside reference (proportional contain) |
| `⌃⇧⌥O` | Fill reference (proportional cover) |
| `⌃⇧⌥S` | Add sticky note (initials + timestamp, cursor on the next line) |
| `⌃⇧⌥K` | Show / hide the pane |

Behaviour switches live in `CONFIG` at the top of [docs/commands.js](docs/commands.js):

- `RECENTER_ON_REF` (default `true`) — fit/fill land on the reference's centre.
- `KEEP_CENTER` (default `true`) — match commands grow around the target's own centre.

## Sticky notes

`⌃⇧⌥S` drops a BCG-style reviewer note on the current slide: a bright text box with
`CU 15 Sep 26 - 8:28p:` as its first line and the cursor waiting on the second, so you just
type. Initials and the default colour are set in the pane (Sticky notes section) and remembered.
Each palette colour is also its own command (`Add sticky — Pink` etc.) if you want a key per
colour; the palette itself is `STICKY.COLORS` in `commands.js`. Additional stickies on the same
slide cascade down-left so they don't stack. Styling copies the sample deck (143pt wide,
top-right, thin-thick dark-blue outline, 12pt bold, auto-fit) minus the drop shadow, which the
JS API can't set.

## Assigning keys

Press `⌃⇧⌥K` (or Home › ChristiantialElements › Open) to open the pane. Click a command, then press the key
you want — or click a key in the map. Esc cancels; Delete removes the command's key. The change
is live immediately; nothing to redeploy.

Rules the pane enforces:

1. **Only keys in the bank can be bound.** The manifest registers a fixed bank — `⌃⇧⌥` + A–Z/0–9
   and `⌃⌥` + A–Z/0–9, 72 keys — and PowerPoint only ever sends us those. Anything else (say
   `⌘B`) is refused with the reason ("PowerPoint: Bold").
2. **Conflicts need a second press.** A bank key already bound to another command, or one a
   known tool uses (macOS defaults, Rectangle window manager), warns first; pressing it again
   binds anyway. Bank keys are dashed in the map when something else is known to use them.
3. One key per command.

Why a bank instead of rebinding live: the add-in-only manifest fixes shortcuts at load time,
and this PowerPoint build reports no `KeyboardShortcuts 1.1`, so `Office.actions.replaceShortcuts`
isn't available. Registering every plausible key up front and routing them through a keymap is
the only way to make assignment instant.

The keymap is stored in the add-in (`OfficeRuntime.storage`, with `localStorage` fallback).
**Backup / restore** in the pane shows it as JSON; paste that into `DEFAULT_KEYMAP` in
`commands.js` to make it the shipped default.

## Layout

```
ppt-shortcuts/
├── docs/                      # served by GitHub Pages (https://cueland.github.io/ppt-shortcuts/)
│   ├── taskpane.html          # shared-runtime host page + assignment UI
│   ├── commands.js            # COMMANDS registry, KEY BANK, geometry, keymap, recorder, pane
│   ├── shortcuts.json         # GENERATED from the bank in commands.js: one Office action per key
│   ├── native-shortcuts.js    # PowerPoint / macOS / Rectangle shortcuts for conflict warnings
│   └── assets/                # ribbon icons
├── manifest.xml               # sideloaded locally — NOT served
├── scripts/build-shortcuts.py # regenerates shortcuts.json from the bank in commands.js (--check verifies)
├── scripts/sideload.sh        # copy manifest into PowerPoint's wef dir and restart PowerPoint
└── scripts/clear-cache.sh     # nuke the add-in cache when a change refuses to show up
```

**Adding a command** = one entry in the `COMMANDS` array in `commands.js` (`id`, `group`,
`label`, `desc`, `run`). It appears in the pane immediately and can be bound to any bank key.
No manifest or JSON change.

**Changing the bank** (new modifier set or keys) = edit `KEY_BANK_MODIFIER_SETS` / `KEY_BANK_KEYS`
in `commands.js`, run `scripts/build-shortcuts.py`, bump `?v=` on `ExtendedOverrides` in
`manifest.xml`, push, re-sideload.

**Every push that changes `commands.js`**: bump `BUILD` in it and the `?b=` on both script tags in
`taskpane.html`. The pane's diagnostics line shows the loaded `Build`, so you can tell at a glance
whether PowerPoint is running fresh code or a cached copy.

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
4. Open any presentation. A **ChristiantialElements › Open** button appears on the Home tab.
   The diagnostics line at the top of the pane reports `Office.context.diagnostics.version`
   and which requirement sets the host supports.

## Phase 0 — selection-order probe: ANSWERED

Whether "last shape you clicked is the master" can work depends on whether
`getSelectedShapes()` returns shapes in **selection order** or **z-order** (VBA returns z-order).

**Result (2026-09-15, PowerPoint for Mac 16.112.3): selection order.** With A forced to the
front, selecting C → A → B dumped `[C, A, B]`. `resolveReference()` therefore takes the last
element, and no explicit pick-up-reference action is needed.

The probe is still in the task pane (**Create probe shapes** / **Dump selection**) in case a
future PowerPoint build changes the behaviour.

## Phase 1 — acceptance: PASSED (2026-09-15, PowerPoint 16.112.909.4)

1. [x] PowerPoint version ≥ 16.105.2.
2. [x] Add-in loads; version reported.
3. [x] Phase 0 answered: selection order preserved; last-selected = reference.
4. [x] All five shortcuts fire from the keyboard with the task pane closed.
5. [x] Geometry verified; read-back matched computed values exactly, so PowerPoint does **not**
   apply lock-aspect-ratio to `width`/`height` writes. (`fitInside` 364.87×207.13 → 237.01×134.55,
   ratio preserved to 4 s.f.)
6. [x] No chooser dialogs on `⌃⇧⌥` combos.

Not yet verified: whether the `⌃⌥` half of the bank fires (first `⌃⇧⌥` half does). Bind
something to a `⌃⌥` key and press it; if nothing happens, that set can be dropped from
`scripts/build-shortcuts.py`.

`KeyboardShortcuts 1.1` reports unsupported on this build. That set only gates the
customisation APIs (`getShortcuts`, `replaceShortcuts`, `areShortcutsInUse`); shortcuts
themselves work without it.

**Test helpers** in the pane still provide the probe shapes and the 100×200 / 300×150 pair.

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
