/* global Office, OfficeRuntime, PowerPoint, KEY_BANK, KEY_BANK_MODIFIER_SETS, KEY_BANK_KEYS, slotId, NATIVE_SHORTCUTS */
/*
 * ChristiantialElements — keyboard-bound shape commands for PowerPoint on macOS.
 *
 * Runs in the add-in's shared runtime (lifetime="long"), so the same JS context serves
 * both the keyboard-shortcut actions and the task pane. That's why command handlers can
 * write straight into the task pane's #log element when the pane happens to be open.
 *
 * Key model: shortcuts.json registers a fixed BANK of key combos (see keybank.js), each as
 * a generic "slot" action. A KEYMAP {combo → command id}, stored in the add-in, decides what
 * each slot runs. Assigning a key is therefore instant — no redeploy, no cache clearing.
 *
 * Layout of this file:
 *   1. CONFIG            — behaviour switches from the build brief (§5)
 *   2. geometry          — pure functions, no Office API (unit-testable in any JS runtime)
 *   3. selection         — loadSelection / resolveReference (last selected = reference)
 *   4. COMMANDS          — the registry: one entry per command, the only place to add one
 *   5. keys              — combo parsing/display, keymap store, slot dispatcher
 *   6. recorder + pane   — assignment UI, key-map grid, probe, log
 *   7. Office.onReady    — Office.actions.associate wiring
 */
"use strict";

// ---------------------------------------------------------------------------
// 1. CONFIG
// ---------------------------------------------------------------------------
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
  for (const [combo, what, source] of NATIVE_SHORTCUTS) {
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

function renderKeyGrid() {
  const root = el("keymap-grid");
  if (!root) return;
  root.innerHTML = KEY_BANK_MODIFIER_SETS.map((mods) => {
    const cells = KEY_BANK_KEYS.map((k) => {
      const combo = `${mods}+${k}`;
      const id = keymap[combo];
      const native = NATIVE_BY_COMBO[combo];
      const cls = ["cell", id ? "assigned" : "free", native ? "native" : "", recorder.active && recorder.commandId === id ? "active" : ""].join(" ");
      const title = id ? COMMAND_BY_ID[id].label : native ? native.map((n) => `${n.source}: ${n.what}`).join("; ") : "available";
      return `<button class="${cls}" data-combo="${combo}" title="${displayCombo(combo)} — ${title}"><b>${k}</b>${id ? `<i>${COMMAND_BY_ID[id].label}</i>` : ""}</button>`;
    }).join("");
    return `<div class="mods">${displayCombo(mods + "+")}</div><div class="cells">${cells}</div>`;
  }).join("");
}

function renderExport() {
  const ta = el("keymap-json");
  if (ta && document.activeElement !== ta) ta.value = JSON.stringify(keymap, null, 2);
}

function renderPane() {
  renderCommandList();
  renderKeyGrid();
  renderExport();
}

function renderDiagnostics() {
  const d = el("diagnostics");
  if (!d) return;
  const info = Office.context.diagnostics || {};
  const sets = ["PowerPointApi 1.4", "PowerPointApi 1.5", "PowerPointApi 1.8", "PowerPointApi 1.10", "SharedRuntime 1.1", "KeyboardShortcuts 1.1"]
    .map((s) => { const [name, ver] = s.split(" "); return `${s}: ${supports(name, ver) ? "yes" : "no"}`; });
  d.innerHTML =
    `<div><b>Host</b> ${info.host || "?"} · <b>Platform</b> ${info.platform || "?"} · <b>Version</b> ${info.version || "?"} · <b>Bank</b> ${KEY_BANK.length} keys</div>` +
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
    if (recorder.active) recorder.capture(combo, "grid");
    else if (keymap[combo]) recorder.start(keymap[combo]);
    else setStatus("Pick a command above first, then click a key.", "muted");
  });

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
  on("btn-clear-log", "click", () => { logBuffer.length = 0; log("log cleared"); });

  for (const key of ["RECENTER_ON_REF", "KEEP_CENTER"]) {
    const box = el("cfg-" + key);
    if (box) {
      box.checked = CONFIG[key];
      box.addEventListener("change", () => { CONFIG[key] = box.checked; log(`${key} = ${box.checked}`); });
    }
  }

  renderDiagnostics();
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
  log(`ready: host=${info.host} platform=${info.platform} version=${(Office.context.diagnostics || {}).version} · ${KEY_BANK.length} slots · ${Object.keys(keymap).length} bound`);

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wireTaskPane);
  else wireTaskPane();
});
