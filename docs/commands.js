/* global Office, OfficeRuntime, PowerPoint, NATIVE_SHORTCUTS */
/*
 * ChristiantialElements — keyboard-bound shape commands for PowerPoint on macOS.
 *
 * Runs in the add-in's shared runtime (lifetime="long"), so the same JS context serves
 * both the keyboard-shortcut actions and the task pane. That's why command handlers can
 * write straight into the task pane's #log element when the pane happens to be open.
 *
 * Key model: shortcuts.json registers a fixed BANK of key combos, each as a generic "slot"
 * action. A KEYMAP {combo → command id}, stored in the add-in, decides what each slot runs.
 * Assigning a key is therefore instant — no redeploy, no cache clearing.
 *
 * This file is deliberately self-contained (the bank is defined HERE and shortcuts.json is
 * generated from it by scripts/build-shortcuts.py). Office caches add-in files one by one,
 * so a page that depends on two scripts staying in step can end up with a mismatched pair
 * and die on load — taking every shortcut with it.
 *
 * Layout of this file:
 *   1. CONFIG + KEY BANK — behaviour switches; the registered key combos
 *   2. geometry          — pure functions, no Office API (unit-testable in any JS runtime)
 *   3. selection         — loadSelection / resolveReference (last selected = reference)
 *   4. COMMANDS          — the registry: one entry per command, the only place to add one
 *   5. keys              — combo parsing/display, keymap store, slot dispatcher
 *   6. recorder + pane   — assignment UI, key-map grid, probe, log
 *   7. Office.onReady    — Office.actions.associate wiring
 */
"use strict";

// Shown in the pane and the log so you can tell which build PowerPoint actually loaded.
const BUILD = "2026-09-15.11";

// ---------------------------------------------------------------------------
// 1. CONFIG + KEY BANK
// ---------------------------------------------------------------------------

// The key bank. scripts/build-shortcuts.py parses these two lines to generate
// docs/shortcuts.json — keep them single-line JSON arrays. After changing them: run the
// script, bump ?v= on ExtendedOverrides in manifest.xml, push, re-sideload.
const KEY_BANK_MODIFIER_SETS = ["Ctrl+Shift+Alt", "Ctrl+Alt"];
const KEY_BANK_KEYS = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9"];
const KEY_BANK = KEY_BANK_MODIFIER_SETS.flatMap((mods) => KEY_BANK_KEYS.map((k) => `${mods}+${k}`));
/** Office action id for a bank combo — must equal what build-shortcuts.py writes. */
function slotId(combo) { return "k_" + combo.replace(/\+/g, "_"); }

const CONFIG = {
  // fitInside / fillOutside: move the result onto the reference's centre (true) or
  // scale in place around the target's own centre (false).
  RECENTER_ON_REF: true,

  // matchWidth / matchHeight / matchBoth: grow/shrink around the target's own centre
  // (true) instead of PowerPoint's native top-left anchoring (false).
  KEEP_CENTER: true,
};

// Keys bound out of the box. Every combo must be in KEY_BANK; every value a COMMANDS id.
const DEFAULT_KEYMAP = {
  "Ctrl+Shift+Alt+W": "matchWidth",
  "Ctrl+Shift+Alt+H": "matchHeight",
  "Ctrl+Shift+Alt+B": "matchBoth",
  "Ctrl+Shift+Alt+I": "fitInside",
  "Ctrl+Shift+Alt+O": "fillOutside",
  "Ctrl+Shift+Alt+K": "togglePane",
  "Ctrl+Shift+Alt+S": "addSticky",
};

// ---------------------------------------------------------------------------
// 2. Geometry — pure functions. `target` and `ref` are {left, top, width, height}.
// ---------------------------------------------------------------------------
const geometry = {
  /** Match width and/or height of `ref`, non-proportionally. */
  match(target, ref, { width, height }, keepCenter = CONFIG.KEEP_CENTER) {
    const w = width ? ref.width : target.width;
    const h = height ? ref.height : target.height;
    let { left, top } = target;
    if (keepCenter) {
      left = target.left + (target.width - w) / 2;
      top = target.top + (target.height - h) / 2;
    }
    return { left, top, width: w, height: h };
  },

  /**
   * Proportional scale. mode "contain": k = min(refW/w, refH/h). mode "cover": k = max(...).
   * Recentres on `ref` when recenterOnRef, otherwise on the target's own centre.
   */
  scale(target, ref, mode, recenterOnRef = CONFIG.RECENTER_ON_REF) {
    if (target.width <= 0 || target.height <= 0) {
      throw new Error("Target has zero width or height; cannot scale proportionally.");
    }
    const kx = ref.width / target.width;
    const ky = ref.height / target.height;
    const k = mode === "contain" ? Math.min(kx, ky) : Math.max(kx, ky);
    const w = target.width * k;
    const h = target.height * k;
    const anchor = recenterOnRef ? ref : target;
    return {
      left: anchor.left + (anchor.width - w) / 2,
      top: anchor.top + (anchor.height - h) / 2,
      width: w,
      height: h,
    };
  },
};

// ---------------------------------------------------------------------------
// 3. Selection
// ---------------------------------------------------------------------------

/** Plain-object snapshot of a loaded shape proxy. */
function snapshot(shape) {
  return { id: shape.id, name: shape.name, left: shape.left, top: shape.top, width: shape.width, height: shape.height };
}

/** Load the current selection and return an array of plain snapshots plus the proxies. */
async function loadSelection(context) {
  const shapes = context.presentation.getSelectedShapes();
  shapes.load("items/id,items/name,items/left,items/top,items/width,items/height");
  await context.sync();
  return shapes.items.map((proxy) => ({ ...snapshot(proxy), proxy }));
}

/**
 * Given the selected shapes (in the order the API returned them), decide which one is the
 * reference and which are the targets.
 *
 * Phase 0 result (2026-09-15, PowerPoint for Mac 16.112.3): getSelectedShapes() returns
 * shapes in SELECTION order, not z-order — selecting C → A → B dumps [C, A, B] with A on
 * top. So the Efficient Elements model holds: the last shape you selected is the reference.
 *
 * Returns { reference: snapshot, targets: [{...snapshot, proxy}] }.
 */
function resolveReference(selected) {
  if (selected.length < 2) {
    throw new Error("Select at least two shapes: targets first, reference last.");
  }
  return { reference: selected[selected.length - 1], targets: selected.slice(0, -1) };
}

/**
 * Shared driver: load selection, resolve reference, compute new geometry for each target
 * with `compute(target, reference)`, write it back, then re-read and report.
 */
async function applyToTargets(actionId, compute) {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    const { reference, targets } = resolveReference(selected);

    log(`${actionId}: reference "${reference.name}" ${fmt(reference)} → ${targets.length} target(s)`);

    for (const target of targets) {
      const next = compute(target, reference);
      target.proxy.left = next.left;
      target.proxy.top = next.top;
      target.proxy.width = next.width;
      target.proxy.height = next.height;
      target.expected = next;
    }
    await context.sync();

    // Verify: re-read and compare, so any host-side adjustment is visible in the log.
    for (const target of targets) target.proxy.load("left,top,width,height");
    await context.sync();
    for (const target of targets) {
      const got = snapshot(target.proxy);
      const ok = approxEqual(got, target.expected);
      log(`  "${target.name}" ${fmt(target)} → ${fmt(got)}${ok ? "" : "  ⚠ expected " + fmt(target.expected)}`);
    }
  });
}

// ---------------------------------------------------------------------------
// 3b. Sticky notes — BCG-style reviewer notes: bright box, "initials date - time:" header,
//     cursor left on the empty next line so you just start typing.
//     Geometry and styling copied from the sample deck (sample stickie.pptx): 143pt wide,
//     top-right with 24pt/54pt margins, 2.25pt thin-thick outline in the theme's dark blue,
//     12pt bold. Height is the sample's 80pt × 1.25, fixed (no auto-fit).
//
//     No shadow: the sample has a real soft drop shadow, but PowerPoint's JS API exposes no
//     shape effects at all (shadow, glow, reflection, soft edges — verified against the preview
//     API and via the Dump Shape API probe, 2026-09-15). OOXML injection is Word-only.
// ---------------------------------------------------------------------------
const STICKY = {
  WIDTH: 143,
  HEIGHT: 100,
  MARGIN: { top: 24, right: 54 },
  CASCADE: 18,          // each additional sticky on a slide steps down-left by this much
  FONT_SIZE: 12,
  LINE: { color: "#0E2841", weight: 2.25, style: "ThinThick" },
  SLIDE: { width: 960, height: 540 }, // 16:9 default; the JS API exposes no slide size
  // Highlighter palette. Names become command ids (sticky_yellow …), so keep them stable.
  COLORS: [
    { name: "Yellow", hex: "#FFFF00" },
    { name: "Green",  hex: "#66FF33" },
    { name: "Pink",   hex: "#FF66CC" },
    { name: "Orange", hex: "#FFA500" },
    { name: "Blue",   hex: "#33CCFF" },
    { name: "Purple", hex: "#CC99FF" },
  ],
};

// User settings (initials, default colour) — persisted like the keymap.
const SETTINGS_STORAGE_KEY = "ppt-shortcuts.settings.v1";
const DEFAULT_SETTINGS = { initials: "CU", stickyColor: "Yellow" };
let settings = { ...DEFAULT_SETTINGS };
async function loadSettings() {
  try {
    const raw = await store.get(SETTINGS_STORAGE_KEY);
    if (raw) settings = { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch (err) { log("settings unreadable, using defaults: " + err.message); }
}
async function saveSettings() { await store.set(SETTINGS_STORAGE_KEY, JSON.stringify(settings)); }

/** "15 Sep 26 - 8:28p" — the sample's format. */
function stickyStamp(d = new Date()) {
  const mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getMonth()];
  const yy = String(d.getFullYear()).slice(-2);
  const h24 = d.getHours();
  const h = h24 % 12 || 12;
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${d.getDate()} ${mon} ${yy} - ${h}:${mm}${h24 >= 12 ? "p" : "a"}`;
}

function stickyColorHex(name) {
  const c = STICKY.COLORS.find((x) => x.name === name) || STICKY.COLORS[0];
  return c.hex;
}

/** Add a sticky to the current slide and leave the cursor on its empty second line. */
async function addSticky(colorName) {
  const hex = stickyColorHex(colorName || settings.stickyColor);
  const { WIDTH: w, HEIGHT: h } = STICKY;
  const header = `${settings.initials} ${stickyStamp()}:`;

  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    // Cascade below any stickies already on the slide so they don't stack exactly on top.
    const existing = shapes.items.filter((s) => /^Sticky\b/.test(s.name)).length;
    const n = existing + 1;
    const left = STICKY.SLIDE.width - STICKY.MARGIN.right - w - existing * STICKY.CASCADE;
    const top = STICKY.MARGIN.top + existing * STICKY.CASCADE;

    const box = shapes.addTextBox(header + "\n", { left, top, width: w, height: h });
    box.name = `Sticky ${n}`;
    box.fill.setSolidColor(hex);
    box.lineFormat.color = STICKY.LINE.color;
    box.lineFormat.weight = STICKY.LINE.weight;
    box.lineFormat.style = STICKY.LINE.style;
    const tf = box.textFrame;
    tf.wordWrap = true;
    tf.autoSizeSetting = "AutoSizeNone";
    tf.textRange.font.size = STICKY.FONT_SIZE;
    tf.textRange.font.bold = true;
    box.load("id");
    await context.sync();

    // Put the insertion point at the start of the (empty) second line.
    try {
      tf.textRange.load("text");
      await context.sync();
      tf.textRange.getSubstring(tf.textRange.text.length, 0).setSelected();
      await context.sync();
    } catch (err) {
      log("sticky: could not place the cursor (" + (err.message || err) + "); selecting the note instead");
      slide.setSelectedShapes([box.id]);
      await context.sync();
    }
    log(`sticky ${n}: "${header}" ${hex} ${w}×${h} at (${left}, ${top})`);
  });
}

// ---------------------------------------------------------------------------
// 4. COMMANDS — the registry. Adding a command = adding one entry here.
//    id: stable, used in the keymap. group/label/desc: shown in the pane.
// ---------------------------------------------------------------------------
const COMMANDS = [
  { id: "matchWidth", group: "Size", label: "Match width",
    desc: "Target width := reference width.",
    run: () => applyToTargets("matchWidth", (t, r) => geometry.match(t, r, { width: true })) },
  { id: "matchHeight", group: "Size", label: "Match height",
    desc: "Target height := reference height.",
    run: () => applyToTargets("matchHeight", (t, r) => geometry.match(t, r, { height: true })) },
  { id: "matchBoth", group: "Size", label: "Match width and height",
    desc: "Both dimensions, non-proportional.",
    run: () => applyToTargets("matchBoth", (t, r) => geometry.match(t, r, { width: true, height: true })) },
  { id: "fitInside", group: "Size", label: "Fit inside reference",
    desc: "Scale proportionally so the target fits within the reference.",
    run: () => applyToTargets("fitInside", (t, r) => geometry.scale(t, r, "contain")) },
  { id: "fillOutside", group: "Size", label: "Fill reference",
    desc: "Scale proportionally so the target covers the reference.",
    run: () => applyToTargets("fillOutside", (t, r) => geometry.scale(t, r, "cover")) },
  { id: "addSticky", group: "Sticky", label: "Add sticky (default colour)",
    desc: "Reviewer note with your initials and a timestamp; cursor lands on the next line.",
    run: () => addSticky() },
  ...STICKY.COLORS.map((c) => ({ id: "sticky_" + c.name.toLowerCase(), group: "Sticky", label: `Add sticky — ${c.name}`,
    desc: `Sticky in ${c.name} (${c.hex}).`, run: () => addSticky(c.name) })),
  { id: "togglePane", group: "Add-in", label: "Show / hide this pane",
    desc: "Open the pane to assign keys; press again to hide it.",
    run: () => togglePane() },
];
const COMMAND_BY_ID = Object.fromEntries(COMMANDS.map((c) => [c.id, c]));

/** Run a command by id, logging instead of throwing (errors would vanish inside the shortcut runtime). */
async function runCommand(id) {
  const cmd = COMMAND_BY_ID[id];
  if (!cmd) { log(`Unknown command "${id}".`); return; }
  const t0 = performance.now();
  try {
    await cmd.run();
    log(`${id} done in ${Math.round(performance.now() - t0)} ms`);
  } catch (err) {
    log(`${id} FAILED: ${err && err.message ? err.message : err}`);
    if (err && err.debugInfo) log("  debugInfo: " + JSON.stringify(err.debugInfo));
  }
}

// Tracked from Office.addin.onVisibilityModeChanged, with document.visibilityState as a
// cross-check (the shared-runtime page is what the pane displays, so when the pane is
// closed the page usually reports "hidden").
let paneVisible = null; // null = no event yet
function paneLooksVisible() {
  if (paneVisible !== null) return paneVisible;
  return document.visibilityState === "visible";
}
async function togglePane() {
  if (!Office.addin || !Office.addin.showAsTaskpane) throw new Error("Office.addin.showAsTaskpane unavailable (needs SharedRuntime 1.1).");
  const before = `event=${paneVisible} doc=${document.visibilityState}`;
  if (paneLooksVisible()) {
    log(`togglePane: hiding (${before})`);
    await Office.addin.hide();
  } else {
    log(`togglePane: showing (${before})`);
    await Office.addin.showAsTaskpane();
  }
}

// ---------------------------------------------------------------------------
// 5. Keys — canonical combos, keymap store, slot dispatcher
// ---------------------------------------------------------------------------

const MOD_ORDER = ["Cmd", "Ctrl", "Shift", "Alt"];
const MOD_ALIASES = { command: "Cmd", cmd: "Cmd", meta: "Cmd", control: "Ctrl", ctrl: "Ctrl", shift: "Shift", alt: "Alt", option: "Alt", opt: "Alt" };
const MOD_GLYPH = { Cmd: "⌘", Ctrl: "⌃", Shift: "⇧", Alt: "⌥" };

/** Parse "ctrl+alt+w" (any order/case) into canonical "Ctrl+Alt+W". */
function canonicalCombo(text) {
  const parts = text.split("+").map((p) => p.trim()).filter(Boolean);
  const mods = new Set();
  let key = null;
  for (const p of parts) {
    const m = MOD_ALIASES[p.toLowerCase()];
    if (m) mods.add(m);
    else key = p.length === 1 ? p.toUpperCase() : p;
  }
  if (!key) return null;
  return [...MOD_ORDER.filter((m) => mods.has(m)), key].join("+");
}

/** "Ctrl+Shift+Alt+W" → "⌃⇧⌥W" */
function displayCombo(combo) {
  if (!combo) return "—";
  const parts = combo.split("+");
  const key = parts.pop();
  return parts.map((m) => MOD_GLYPH[m] || m).join("") + key;
}

/** Turn a KeyboardEvent into a canonical combo, or null for a bare modifier press. */
function comboFromEvent(e) {
  const code = e.code || "";
  let key = null;
  if (/^Key[A-Z]$/.test(code)) key = code.slice(3);
  else if (/^Digit[0-9]$/.test(code)) key = code.slice(5);
  else if (/^F[0-9]{1,2}$/.test(code)) key = code;
  else key = { Minus: "-", Equal: "=", Comma: ",", Period: ".", Slash: "/", Semicolon: ";", Quote: "'",
    BracketLeft: "[", BracketRight: "]", Backslash: "\\", Backquote: "`", Space: "Space",
    ArrowLeft: "Left", ArrowRight: "Right", ArrowUp: "Up", ArrowDown: "Down",
    Enter: "Return", Backspace: "Backspace", Delete: "Delete", Escape: "Esc", Tab: "Tab" }[code] || null;
  if (!key) return null;
  const mods = [];
  if (e.metaKey) mods.push("Cmd");
  if (e.ctrlKey) mods.push("Ctrl");
  if (e.shiftKey) mods.push("Shift");
  if (e.altKey) mods.push("Alt");
  return [...mods, key].join("+");
}

const BANK_SET = new Set(KEY_BANK);
const NATIVE_BY_COMBO = (() => {
  const m = {};
  // native-shortcuts.js is optional: without it you just lose the conflict warnings.
  const list = typeof NATIVE_SHORTCUTS !== "undefined" ? NATIVE_SHORTCUTS : [];
  for (const [combo, what, source] of list) {
    const c = canonicalCombo(combo);
    (m[c] = m[c] || []).push({ what, source });
  }
  return m;
})();

// ---- keymap persistence: OfficeRuntime.storage (shared-runtime store) with localStorage fallback ----
const KEYMAP_STORAGE_KEY = "ppt-shortcuts.keymap.v1";
const store = {
  async get(key) {
    try {
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
        const v = await OfficeRuntime.storage.getItem(key);
        if (v != null) return v;
      }
    } catch (_) { /* fall through */ }
    try { return localStorage.getItem(key); } catch (_) { return null; }
  },
  async set(key, value) {
    try { if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) await OfficeRuntime.storage.setItem(key, value); } catch (_) { /* ignore */ }
    try { localStorage.setItem(key, value); } catch (_) { /* ignore */ }
  },
};

let keymap = { ...DEFAULT_KEYMAP };

/** Keep only entries whose combo is in the bank and whose command exists. Returns the dropped ones. */
function sanitizeKeymap(map) {
  const clean = {};
  const dropped = [];
  for (const [combo, id] of Object.entries(map || {})) {
    const c = canonicalCombo(combo);
    if (c && BANK_SET.has(c) && COMMAND_BY_ID[id] && !Object.values(clean).includes(id)) clean[c] = id;
    else dropped.push(`${combo} → ${id}`);
  }
  return { clean, dropped };
}

async function loadKeymap() {
  const raw = await store.get(KEYMAP_STORAGE_KEY);
  if (!raw) { keymap = { ...DEFAULT_KEYMAP }; return; }
  try {
    const { clean, dropped } = sanitizeKeymap(JSON.parse(raw));
    keymap = clean;
    if (dropped.length) log(`keymap: ignored ${dropped.length} stale entr${dropped.length === 1 ? "y" : "ies"}: ${dropped.join("; ")}`);
    // New shipped defaults (a command added since the map was saved) join the map when both
    // the key and the command are still free — never overriding a choice you made.
    const added = [];
    for (const [combo, id] of Object.entries(DEFAULT_KEYMAP)) {
      if (!keymap[combo] && !comboForCommand(id) && COMMAND_BY_ID[id]) { keymap[combo] = id; added.push(`${displayCombo(combo)} → ${COMMAND_BY_ID[id].label}`); }
    }
    if (added.length) { await saveKeymap(); log("keymap: added new default(s): " + added.join(", ")); }
  } catch (err) {
    log("keymap: stored value unreadable, using defaults (" + err.message + ")");
    keymap = { ...DEFAULT_KEYMAP };
  }
}

async function saveKeymap() {
  await store.set(KEYMAP_STORAGE_KEY, JSON.stringify(keymap));
}

function comboForCommand(id) {
  return Object.keys(keymap).find((c) => keymap[c] === id) || null;
}

/** Bind combo → command id, replacing whatever either side had. */
async function bind(combo, id) {
  const prev = comboForCommand(id);
  if (prev) delete keymap[prev];
  keymap[combo] = id;
  await saveKeymap();
  log(`bound ${displayCombo(combo)} → ${COMMAND_BY_ID[id].label}`);
}

async function unbind(id) {
  const prev = comboForCommand(id);
  if (!prev) return;
  delete keymap[prev];
  await saveKeymap();
  log(`removed ${displayCombo(prev)} from ${COMMAND_BY_ID[id].label}`);
}

/** Every registered slot lands here. */
async function onSlot(combo) {
  if (recorder.active) { recorder.capture(combo, "office"); return; }
  const id = keymap[combo];
  if (!id) { log(`${displayCombo(combo)} isn't assigned. Press ${displayCombo(comboForCommand("togglePane"))} to open the pane and assign it.`); return; }
  await runCommand(id);
}

// ---------------------------------------------------------------------------
// 6. Recorder + task pane
// ---------------------------------------------------------------------------

/**
 * Assignment flow: click a command (or press its row) → recorder.active with that command →
 * press a combo (DOM keydown) or click a cell in the grid or, if Office grabs the keystroke
 * first, the slot action itself reports it. Conflicts require the same combo a second time.
 */
const recorder = {
  active: false,
  commandId: null,
  pendingConfirm: null, // combo awaiting a second press
  lastCapture: 0,

  start(commandId) {
    this.active = true;
    this.commandId = commandId;
    this.pendingConfirm = null;
    renderPane();
    setStatus(`Press the key for “${COMMAND_BY_ID[commandId].label}” — or click a key below. Esc cancels, Delete removes.`, "info");
  },
  stop() {
    this.active = false;
    this.commandId = null;
    this.pendingConfirm = null;
    renderPane();
  },

  async capture(combo, source) {
    if (!this.active) return;
    // A key can arrive twice (DOM keydown + Office slot). Ignore the echo.
    const now = performance.now();
    if (now - this.lastCapture < 400 && combo === this._lastCombo) return;
    this.lastCapture = now;
    this._lastCombo = combo;

    const id = this.commandId;
    const cmd = COMMAND_BY_ID[id];

    if (combo === "Esc") { this.stop(); setStatus("Cancelled.", "muted"); return; }
    if (combo === "Backspace" || combo === "Delete") {
      await unbind(id); this.stop(); setStatus(`Removed the key from “${cmd.label}”.`, "ok"); return;
    }

    // 1. Only bank keys can be bound — the manifest fixes what PowerPoint will send us.
    if (!BANK_SET.has(combo)) {
      const native = NATIVE_BY_COMBO[combo];
      const why = native ? ` It's already ${native.map((n) => `${n.source}: ${n.what}`).join(", ")}.` : "";
      setStatus(`${displayCombo(combo)} can't be bound — only ${KEY_BANK_MODIFIER_SETS.map((m) => displayCombo(m + "+key")).join(" and ")} are registered.${why} Try another key.`, "warn");
      return;
    }

    // 2. Known conflicts need a second press to confirm.
    const problems = [];
    const holder = keymap[combo];
    if (holder && holder !== id) problems.push(`already runs “${COMMAND_BY_ID[holder].label}”`);
    for (const n of NATIVE_BY_COMBO[combo] || []) problems.push(`is ${n.source}'s “${n.what}”`);
    if (problems.length && this.pendingConfirm !== combo) {
      this.pendingConfirm = combo;
      setStatus(`${displayCombo(combo)} ${problems.join(" and ")}. Press it again to use it anyway, or pick another key.`, "warn");
      return;
    }

    await bind(combo, id);
    this.stop();
    setStatus(`${displayCombo(combo)} → ${cmd.label}${source === "office" ? " (confirmed live from PowerPoint)" : ""}`, "ok");
  },
};

function onKeyDown(e) {
  if (!recorder.active) return;
  // Let the export textarea keep its keys.
  if (e.target && e.target.tagName === "TEXTAREA") return;
  const combo = comboFromEvent(e);
  if (!combo) return;
  e.preventDefault();
  e.stopPropagation();
  recorder.capture(combo, "dom");
}

// ---- rendering ----

function el(id) { return document.getElementById(id); }

function setStatus(text, kind) {
  const s = el("recorder-status");
  if (!s) return;
  s.textContent = text;
  s.className = "status " + (kind || "");
}

function renderCommandList() {
  const root = el("cmd-list");
  if (!root) return;
  const esc = (v) => String(v).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const groups = [...new Set(COMMANDS.map((c) => c.group))];
  root.innerHTML = groups.map((g) =>
    `<div class="group">${esc(g)}</div>` +
    COMMANDS.filter((c) => c.group === g).map((c) => {
      const combo = comboForCommand(c.id);
      const active = recorder.active && recorder.commandId === c.id;
      return `<div class="cmd${active ? " active" : ""}" data-id="${c.id}" tabindex="0" title="${esc(c.desc)}">
        <span class="cmd-label">${esc(c.label)}</span>
        <span class="cmd-actions">
          <kbd class="${combo ? "" : "empty"}">${active ? "press a key…" : esc(displayCombo(combo))}</kbd>
          <button class="run" data-run="${c.id}" title="Run now">▶</button>
          ${combo ? `<button class="x" data-unbind="${c.id}" title="Remove key">×</button>` : ""}
        </span>
      </div>`;
    }).join("")
  ).join("");
}

// Physical layout for the key map: number row on top, then the three QWERTY rows.
// Any bank key not listed here is appended in a final row.
const KEY_ROWS = [
  ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"],
  ["Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P"],
  ["A", "S", "D", "F", "G", "H", "J", "K", "L"],
  ["Z", "X", "C", "V", "B", "N", "M"],
];

/** What a bank combo is used for: {kind: "assigned"|"conflict"|"free", label, detail}. */
function describeCombo(combo) {
  const id = keymap[combo];
  const native = NATIVE_BY_COMBO[combo] || [];
  if (id) return { kind: "assigned", label: COMMAND_BY_ID[id].label, detail: native.map((n) => `also ${n.source}: ${n.what}`) };
  if (native.length) return { kind: "conflict", label: native[0].source, detail: native.map((n) => `${n.source}: ${n.what}`) };
  return { kind: "free", label: "", detail: ["available"] };
}

function renderKeyGrid() {
  const root = el("keymap-grid");
  if (!root) return;
  const inRows = new Set(KEY_ROWS.flat());
  const extra = KEY_BANK_KEYS.filter((k) => !inRows.has(k));
  const rows = extra.length ? [...KEY_ROWS, extra] : KEY_ROWS;
  root.innerHTML = KEY_BANK_MODIFIER_SETS.map((mods) => {
    const html = rows.map((row, r) =>
      `<div class="krow r${r}">` + row.filter((k) => KEY_BANK_KEYS.includes(k)).map((k) => {
        const combo = `${mods}+${k}`;
        const d = describeCombo(combo);
        const active = recorder.active && keymap[combo] && recorder.commandId === keymap[combo];
        const cls = ["cell", d.kind, active ? "active" : ""].join(" ");
        return `<button class="${cls}" data-combo="${combo}" aria-label="${displayCombo(combo)}: ${d.detail.join("; ")}"><b>${k}</b><i>${d.label}</i></button>`;
      }).join("") + "</div>"
    ).join("");
    return `<div class="mods">${displayCombo(mods + "+")}</div><div class="keyboard">${html}</div>`;
  }).join("");
}

// Custom hover card (the webview's native title tooltips are slow and easy to miss).
function showTip(cell) {
  const tip = el("keytip");
  if (!tip || !cell) return;
  const combo = cell.dataset.combo;
  const d = describeCombo(combo);
  tip.innerHTML = `<b>${displayCombo(combo)}</b> ${d.kind === "assigned" ? "→ " + d.label : ""}<br>` +
    (d.kind === "free" ? "<span class='muted'>available</span>" : d.detail.map((t) => `<span>${t}</span>`).join("<br>"));
  tip.hidden = false;
  const r = cell.getBoundingClientRect();
  const w = tip.offsetWidth;
  const left = Math.max(8, Math.min(window.innerWidth - w - 8, r.left + r.width / 2 - w / 2));
  tip.style.left = left + window.scrollX + "px";
  tip.style.top = r.bottom + 6 + window.scrollY + "px";
}
function hideTip() { const tip = el("keytip"); if (tip) tip.hidden = true; }

function renderExport() {
  const ta = el("keymap-json");
  if (ta && document.activeElement !== ta) ta.value = JSON.stringify(keymap, null, 2);
}

function renderPane() {
  renderCommandList();
  renderKeyGrid();
  renderExport();
}

function renderStickySwatches() {
  const root = el("sticky-colors");
  if (!root) return;
  root.innerHTML = STICKY.COLORS.map((c) =>
    `<button class="swatch${c.name === settings.stickyColor ? " selected" : ""}" data-color="${c.name}" style="background:${c.hex}" title="${c.name} ${c.hex}"><span>${c.name}</span></button>`
  ).join("");
}

function renderDiagnostics() {
  const d = el("diagnostics");
  if (!d) return;
  const info = Office.context.diagnostics || {};
  const sets = ["PowerPointApi 1.4", "PowerPointApi 1.5", "PowerPointApi 1.8", "PowerPointApi 1.10", "SharedRuntime 1.1", "KeyboardShortcuts 1.1"]
    .map((s) => { const [name, ver] = s.split(" "); return `${s}: ${supports(name, ver) ? "yes" : "no"}`; });
  d.innerHTML =
    `<div><b>Host</b> ${info.host || "?"} · <b>Platform</b> ${info.platform || "?"} · <b>Version</b> ${info.version || "?"} · <b>Build</b> ${BUILD} · <b>Bank</b> ${KEY_BANK.length} keys</div>` +
    `<div class="muted">${sets.join(" · ")}</div>`;
}

function wireTaskPane() {
  const on = (id, evt, fn) => { const e = el(id); if (e) e.addEventListener(evt, fn); };

  // Command list: click row → record; ▶ → run; × → unbind.
  on("cmd-list", "click", async (e) => {
    const run = e.target.closest("[data-run]");
    if (run) { e.stopPropagation(); await runCommand(run.dataset.run); return; }
    const x = e.target.closest("[data-unbind]");
    if (x) { e.stopPropagation(); await unbind(x.dataset.unbind); renderPane(); setStatus("Removed.", "ok"); return; }
    const row = e.target.closest(".cmd");
    if (row) recorder.start(row.dataset.id);
  });
  on("cmd-list", "keydown", (e) => {
    const row = e.target.closest && e.target.closest(".cmd");
    if (row && !recorder.active && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); recorder.start(row.dataset.id); }
  });

  // Grid: click a key while recording → capture; otherwise clicking an assigned key selects that command.
  on("keymap-grid", "click", (e) => {
    const cell = e.target.closest("[data-combo]");
    if (!cell) return;
    const combo = cell.dataset.combo;
    if (recorder.active) { recorder.capture(combo, "grid"); return; }
    if (keymap[combo]) { recorder.start(keymap[combo]); return; }
    const d = describeCombo(combo);
    setStatus(`${displayCombo(combo)} — ${d.detail.join("; ")}. Pick a command above first, then click a key.`, "muted");
  });
  on("keymap-grid", "mouseover", (e) => showTip(e.target.closest("[data-combo]")));
  on("keymap-grid", "mouseout", (e) => { if (!e.relatedTarget || !e.relatedTarget.closest || !e.relatedTarget.closest("[data-combo]")) hideTip(); });
  on("keymap-grid", "focusin", (e) => showTip(e.target.closest("[data-combo]")));
  on("keymap-grid", "focusout", hideTip);

  document.addEventListener("keydown", onKeyDown, true);

  // Export / import / reset
  on("btn-apply-json", "click", async () => {
    try {
      const { clean, dropped } = sanitizeKeymap(JSON.parse(el("keymap-json").value));
      keymap = clean;
      await saveKeymap();
      renderPane();
      setStatus(dropped.length ? `Applied; ignored: ${dropped.join("; ")}` : "Applied.", dropped.length ? "warn" : "ok");
    } catch (err) { setStatus("Not valid JSON: " + err.message, "warn"); }
  });
  on("btn-copy-json", "click", async () => {
    try { await navigator.clipboard.writeText(JSON.stringify(keymap, null, 2)); setStatus("Copied.", "ok"); }
    catch (_) { el("keymap-json").select(); setStatus("Select-all + copy from the box (clipboard API unavailable).", "muted"); }
  });
  on("btn-reset-json", "click", async () => {
    keymap = { ...DEFAULT_KEYMAP };
    await saveKeymap();
    renderPane();
    setStatus("Reset to defaults.", "ok");
  });

  // Probe + log
  on("btn-dump", "click", () => dumpSelection().catch((e) => log("dumpSelection FAILED: " + e.message)));
  on("btn-probe-shapes", "click", () => createProbeShapes().catch((e) => log("createProbeShapes FAILED: " + e.message)));
  on("btn-acceptance-shapes", "click", () => createAcceptanceShapes().catch((e) => log("createAcceptanceShapes FAILED: " + e.message)));
  on("btn-shape-api", "click", () => dumpShapeApi().catch((e) => log("dumpShapeApi FAILED: " + e.message)));
  on("btn-clear-log", "click", () => { logBuffer.length = 0; log("log cleared"); });

  for (const key of ["RECENTER_ON_REF", "KEEP_CENTER"]) {
    const box = el("cfg-" + key);
    if (box) {
      box.checked = CONFIG[key];
      box.addEventListener("change", () => { CONFIG[key] = box.checked; log(`${key} = ${box.checked}`); });
    }
  }

  // Sticky settings
  const initials = el("sticky-initials");
  if (initials) {
    initials.value = settings.initials;
    initials.addEventListener("change", async () => {
      settings.initials = initials.value.trim().toUpperCase() || DEFAULT_SETTINGS.initials;
      initials.value = settings.initials;
      await saveSettings();
      log(`initials = ${settings.initials}`);
    });
  }
  on("sticky-colors", "click", async (e) => {
    const sw = e.target.closest("[data-color]");
    if (!sw) return;
    settings.stickyColor = sw.dataset.color;
    await saveSettings();
    renderStickySwatches();
    log(`sticky colour = ${settings.stickyColor}`);
  });
  on("btn-add-sticky", "click", () => runCommand("addSticky"));

  renderDiagnostics();
  renderStickySwatches();
  renderPane();
  const l = el("log");
  if (l) l.textContent = logBuffer.join("\n");
}

// ---- probe helpers (Phase 0 / Phase 1) ----

const supports = (set, ver) =>
  !!(Office.context && Office.context.requirements && Office.context.requirements.isSetSupported(set, ver));

/** Dump the selection as returned by getSelectedShapes(), with zOrderPosition when available. */
async function dumpSelection() {
  const hasZ = supports("PowerPointApi", "1.8");
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/name" + (hasZ ? ",items/zOrderPosition" : ""));
    await context.sync();
    const rows = shapes.items.map((s, i) => ({ index: i, name: s.name, id: s.id, zOrderPosition: hasZ ? s.zOrderPosition : "n/a" }));
    renderTable("probe-output", rows);
    log(`dumpSelection: array order [${rows.map((r) => r.name).join(" → ")}]`);
  });
}

/** Add three named rectangles with z-order deliberately different from a natural click order. */
async function createProbeShapes() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const specs = [
      { name: "A", left: 60, top: 100, width: 120, height: 80 },
      { name: "B", left: 240, top: 100, width: 120, height: 80 },
      { name: "C", left: 420, top: 100, width: 120, height: 80 },
    ];
    const made = specs.map((spec) => {
      const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, spec);
      shape.name = spec.name;
      shape.textFrame.textRange.text = spec.name;
      return shape;
    });
    await context.sync();
    if (supports("PowerPointApi", "1.8")) {
      made[0].setZOrder(PowerPoint.ShapeZOrder.bringToFront);
      await context.sync();
    }
    log("Created probe shapes A, B, C (A brought to front).");
  });
}

/**
 * Ask the host what a Shape really supports.
 *  1. Every member the runtime defines on the proxy's prototype chain (inspected via
 *     descriptors — touching an unloaded getter throws).
 *  2. Then just try it: shape.load(name) and shape.set({name: …}) for every plausible
 *     effect-property name, logging what PowerPoint answers. If any of them takes, it exists.
 */
async function dumpShapeApi() {
  const CANDIDATES = ["shadow", "shadowFormat", "effects", "effectFormat", "glow", "glowFormat", "reflection", "reflectionFormat", "softEdge", "softEdges", "softEdgeFormat", "threeDFormat", "style"];
  await PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id,items/name");
    await context.sync();
    if (!sel.items.length) throw new Error("select a shape first");
    const shape = sel.items[0];

    // 1. prototype members
    const members = new Map();
    for (let proto = Object.getPrototypeOf(shape); proto && proto !== Object.prototype; proto = Object.getPrototypeOf(proto)) {
      for (const k of Object.getOwnPropertyNames(proto)) {
        if (k.startsWith("_") || k === "constructor" || members.has(k)) continue;
        const d = Object.getOwnPropertyDescriptor(proto, k);
        members.set(k, typeof d.value === "function" ? "method" : "property");
      }
    }
    const list = [...members.keys()].sort();
    const effectish = list.filter((k) => /shadow|glow|reflect|soft|effect|bevel|3d|style/i.test(k));
    log(`Shape API on this host: ${list.length} members. Effect-related names: ${effectish.length ? effectish.join(", ") : "none"}.`);

    // 2. try each candidate for real
    const rows = [];
    for (const name of CANDIDATES) {
      const row = { name, onPrototype: members.has(name) ? "yes" : "no", load: "", set: "" };
      try { shape.load(name); await context.sync(); row.load = "ok: " + JSON.stringify(shape[name] && shape[name].toJSON ? shape[name].toJSON() : shape[name]).slice(0, 60); }
      catch (err) { row.load = "✗ " + (err.message || err).slice(0, 80); }
      try { shape.set({ [name]: { visible: true } }); await context.sync(); row.set = "ok (no error)"; }
      catch (err) { row.set = "✗ " + (err.message || err).slice(0, 80); }
      rows.push(row);
    }
    renderTable("probe-output", rows);
    for (const r of rows) log(`  ${r.name}: prototype=${r.onPrototype} · load → ${r.load} · set → ${r.set}`);
    log("Full member list: " + list.join(", "));
  });
}

/** Add the Phase 1 acceptance pair: 100×200 target and 300×150 reference. */
async function createAcceptanceShapes() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const target = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: 60, top: 80, width: 100, height: 200 });
    target.name = "Target 100x200";
    target.textFrame.textRange.text = "target";
    const ref = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle, { left: 300, top: 80, width: 300, height: 150 });
    ref.name = "Reference 300x150";
    ref.textFrame.textRange.text = "reference";
    await context.sync();
    log("Created acceptance shapes. Expect fitInside → 75×150, fillOutside → 300×600.");
  });
}

// ---- log / formatting ----

const logBuffer = [];
function log(message) {
  const line = `${new Date().toLocaleTimeString()}  ${message}`;
  logBuffer.push(line);
  if (logBuffer.length > 200) logBuffer.shift();
  console.log("[ChristiantialElements] " + message);
  const l = el("log");
  if (l) { l.textContent = logBuffer.join("\n"); l.scrollTop = l.scrollHeight; }
}

function fmt(g) {
  const n = (v) => (Math.round(v * 100) / 100).toString();
  return `${n(g.width)}×${n(g.height)} @ (${n(g.left)}, ${n(g.top)})`;
}

function approxEqual(a, b, eps = 0.05) {
  return ["left", "top", "width", "height"].every((k) => Math.abs(a[k] - b[k]) < eps);
}

function renderTable(elementId, rows) {
  const root = el(elementId);
  if (!root) return;
  if (rows.length === 0) { root.innerHTML = "<p class='muted'>Nothing selected.</p>"; return; }
  const cols = Object.keys(rows[0]);
  const esc = (v) => String(v).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
  root.innerHTML =
    "<table><thead><tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr></thead><tbody>" +
    rows.map((r) => "<tr>" + cols.map((c) => `<td>${esc(r[c])}</td>`).join("") + "</tr>").join("") +
    "</tbody></table>";
}

// ---------------------------------------------------------------------------
// 7. Wiring
// ---------------------------------------------------------------------------
Office.onReady(async (info) => {
  // Every bank combo is an Office action; the id must equal slotId(combo) as in shortcuts.json.
  if (Office.actions && typeof Office.actions.associate === "function") {
    for (const combo of KEY_BANK) Office.actions.associate(slotId(combo), () => onSlot(combo));
  } else {
    log("Office.actions.associate unavailable — not running inside an Office shared runtime.");
  }
  if (Office.addin && Office.addin.onVisibilityModeChanged) {
    try {
      await Office.addin.onVisibilityModeChanged((args) => {
        paneVisible = args.visibilityMode === "Taskpane";
        log(`pane visibility → ${args.visibilityMode}`);
      });
    } catch (err) { log("onVisibilityModeChanged unavailable: " + (err.message || err)); }
  }

  await loadKeymap();
  await loadSettings();
  log(`ready: build ${BUILD} · host=${info.host} platform=${info.platform} version=${(Office.context.diagnostics || {}).version} · ${KEY_BANK.length} slots · ${Object.keys(keymap).length} bound`);

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wireTaskPane);
  else wireTaskPane();
});
