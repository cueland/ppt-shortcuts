/* global Office, OfficeRuntime, PowerPoint, NATIVE_SHORTCUTS */
/*
 * Expropriated Elements — keyboard-driven shape tools for PowerPoint on macOS.
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
 *   2. helpers           — Office proxies, selection, rect maths, settings store
 *   3. operations        — one async function per feature, grouped like EE's task pane
 *   4. COMMANDS          — the registry: one entry per command, the only place to add one
 *   5. keys              — combo parsing/display, keymap store, slot dispatcher
 *   6. pane              — icon toolbar (run / assign modes), params, recorder, key map, log
 *   7. Office.onReady    — Office.actions.associate wiring
 */
"use strict";

// Shown in the pane and the log so you can tell which build PowerPoint actually loaded.
const BUILD = "2026-09-20.28";

// ---------------------------------------------------------------------------
// 1. CONFIG + KEY BANK
// ---------------------------------------------------------------------------

// The key bank. scripts/build-shortcuts.py parses these two lines to generate
// docs/shortcuts.json — keep them single-line JSON arrays. After changing them: run the
// script, bump ?v= on ExtendedOverrides in manifest.xml, push, re-sideload.
const KEY_BANK_MODIFIER_SETS = ["Ctrl+Shift+Alt", "Ctrl+Alt", "Ctrl+Shift", "Shift+Alt"];
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
  // Slide size in points. The JS API exposes no slide dimensions; 16:9 default.
  SLIDE: { width: 960, height: 540 },
  // Swap: also exchange sizes / layer order.
  SWAP_SIZE: false,
  SWAP_ZORDER: true,
  // Stack: gap between stacked shapes (pt).
  STACK_GAP: 0,
};

// Keys bound out of the box. Every combo must be in KEY_BANK; every value a COMMANDS id.
const DEFAULT_KEYMAP = {
  "Ctrl+Shift+Alt+W": "matchWidth",
  "Ctrl+Shift+Alt+H": "matchHeight",
  "Ctrl+Shift+Alt+B": "matchBoth",
  "Ctrl+Shift+Alt+I": "fitInside",
  "Ctrl+Shift+Alt+O": "fillOutside",
  "Ctrl+Shift+Alt+S": "addSticky",
  "Ctrl+Shift+Alt+D": "duplicateSlide",
  "Ctrl+Shift+Alt+K": "togglePane",
};

// Brand/colour palette for the fill / line / font colour commands. Editable in the pane.
const DEFAULT_PALETTE = ["#000000", "#404040", "#808080", "#BFBFBF", "#FFFFFF", "#0E2841", "#156082", "#00B050", "#E97132", "#FFC000"];

// ---------------------------------------------------------------------------
// 2. Helpers
// ---------------------------------------------------------------------------

// ---- persistence: OfficeRuntime.storage (shared-runtime store) with localStorage fallback ----
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

// ---- user settings (persisted) ----
const SETTINGS_STORAGE_KEY = "ppt-shortcuts.settings.v1";
const DEFAULT_SETTINGS = {
  initials: "CU",
  stickyColor: "Yellow",
  palette: DEFAULT_PALETTE,
  resizeFactor: 1.1,          // Magic Resizer: ×factor (1.1 = +10 %)
  resizeFont: true,
  resizeLine: true,
  slice: { rows: 2, cols: 2, gapX: 6, gapY: 6 },
  matrix: { rows: 2, cols: 3, gapX: 8, gapY: 8 },
  margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
  nudge: 1,
  fontSize: 12,
  gotoSlide: 1,
  myFormats: [],               // [{name, props}]
  agenda: "",                  // one item per line
  masterLabel: "Confidential",
  masterPos: { left: 700, top: 500, width: 220, height: 24 },
};
let settings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
async function loadSettings() {
  try {
    const raw = await store.get(SETTINGS_STORAGE_KEY);
    if (raw) settings = { ...JSON.parse(JSON.stringify(DEFAULT_SETTINGS)), ...JSON.parse(raw) };
  } catch (err) { log("settings unreadable, using defaults: " + err.message); }
}
async function saveSettings() { await store.set(SETTINGS_STORAGE_KEY, JSON.stringify(settings)); }

// ---- rect maths on plain objects {left, top, width, height} ----
const R = {
  right: (r) => r.left + r.width,
  bottom: (r) => r.top + r.height,
  cx: (r) => r.left + r.width / 2,
  cy: (r) => r.top + r.height / 2,
  slide: () => ({ left: 0, top: 0, width: CONFIG.SLIDE.width, height: CONFIG.SLIDE.height }),
};

/** Plain-object snapshot of a loaded shape proxy. */
function snapshot(shape) {
  return { id: shape.id, name: shape.name, left: shape.left, top: shape.top, width: shape.width, height: shape.height };
}

/** Load the current selection and return an array of plain snapshots plus the proxies (in selection order). */
async function loadSelection(context, extra = "") {
  const shapes = context.presentation.getSelectedShapes();
  shapes.load("items/id,items/name,items/left,items/top,items/width,items/height" + (extra ? "," + extra : ""));
  await context.sync();
  return shapes.items.map((proxy) => ({ ...snapshot(proxy), proxy }));
}

/**
 * Phase 0 result (2026-09-15, PowerPoint for Mac 16.112.3): getSelectedShapes() returns
 * shapes in SELECTION order, so the Efficient Elements model holds: the last shape you
 * selected is the reference ("Master"). With a single shape selected, the slide is the reference.
 */
function resolveReference(selected, { allowSlide = false } = {}) {
  if (selected.length === 0) throw new Error("Nothing selected.");
  if (selected.length === 1) {
    if (!allowSlide) throw new Error("Select at least two shapes: targets first, reference last.");
    return { reference: { ...R.slide(), name: "slide" }, targets: selected };
  }
  return { reference: selected[selected.length - 1], targets: selected.slice(0, -1) };
}

/** Write {left, top, width, height} onto proxies, sync, then re-read and log what stuck. */
async function writeGeometry(context, actionId, items) {
  for (const it of items) {
    const n = it.next;
    if (n.left != null) it.proxy.left = n.left;
    if (n.top != null) it.proxy.top = n.top;
    if (n.width != null) it.proxy.width = Math.max(0.5, n.width);
    if (n.height != null) it.proxy.height = Math.max(0.5, n.height);
  }
  await context.sync();
  for (const it of items) it.proxy.load("left,top,width,height");
  await context.sync();
  for (const it of items) log(`  "${it.name}" ${fmt(it)} → ${fmt(snapshot(it.proxy))}`);
}

/** Reference/targets driver: compute(target, reference) → partial geometry for each target. */
async function applyToTargets(actionId, compute, opts = {}) {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    const { reference, targets } = resolveReference(selected, opts);
    log(`${actionId}: reference "${reference.name}" ${fmt(reference)} → ${targets.length} target(s)`);
    for (const t of targets) t.next = compute(t, reference, targets);
    await writeGeometry(context, actionId, targets.filter((t) => t.next));
  });
}

/** All-selected driver: compute(items) sets item.next for every selected shape. */
async function applyToAll(actionId, compute, minCount = 1) {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (selected.length < minCount) throw new Error(`Select at least ${minCount} shape${minCount > 1 ? "s" : ""}.`);
    log(`${actionId}: ${selected.length} shape(s)`);
    compute(selected);
    await writeGeometry(context, actionId, selected.filter((t) => t.next));
  });
}

const supports = (set, ver) =>
  !!(Office.context && Office.context.requirements && Office.context.requirements.isSetSupported(set, ver));

// ---------------------------------------------------------------------------
// 3. Operations
// ---------------------------------------------------------------------------

// ---- Size (D) ----
const geometry = {
  match(target, ref, { width, height }, keepCenter = CONFIG.KEEP_CENTER) {
    const w = width ? ref.width : target.width;
    const h = height ? ref.height : target.height;
    let { left, top } = target;
    if (keepCenter) { left = target.left + (target.width - w) / 2; top = target.top + (target.height - h) / 2; }
    return { left, top, width: w, height: h };
  },
  scale(target, ref, mode, recenterOnRef = CONFIG.RECENTER_ON_REF) {
    if (target.width <= 0 || target.height <= 0) throw new Error("Target has zero width or height; cannot scale proportionally.");
    const k = mode === "contain" ? Math.min(ref.width / target.width, ref.height / target.height) : Math.max(ref.width / target.width, ref.height / target.height);
    const w = target.width * k, h = target.height * k;
    const a = recenterOnRef ? ref : target;
    return { left: a.left + (a.width - w) / 2, top: a.top + (a.height - h) / 2, width: w, height: h };
  },
};

/**
 * Move one edge of the target to a coordinate, keeping the opposite edge where it is. If the
 * moved edge would cross the opposite edge, the interval flips: the old opposite edge becomes
 * the moved edge and the target coordinate becomes the new opposite edge. Shrinks as readily
 * as it grows. `targetOf(ref)` gives the coordinate (an edge of the reference).
 */
function moveEdge(edge, targetOf) {
  return (t, r) => {
    const P = targetOf(r);
    switch (edge) {
      case "right": return { left: Math.min(P, t.left), width: Math.abs(P - t.left) };
      case "left": return { left: Math.min(P, R.right(t)), width: Math.abs(R.right(t) - P) };
      case "bottom": return { top: Math.min(P, t.top), height: Math.abs(P - t.top) };
      case "top": return { top: Math.min(P, R.bottom(t)), height: Math.abs(R.bottom(t) - P) };
    }
  };
}

/** Stretch: the target's edge goes to the SAME edge of the reference (its far edge on that side). */
function stretch(edge) {
  return moveEdge(edge, { right: R.right, left: (r) => r.left, bottom: R.bottom, top: (r) => r.top }[edge]);
}

/** Fill gap: the target's edge goes to the OPPOSITE edge of the reference (its near edge). */
function fillGap(edge) {
  return moveEdge(edge, { right: (r) => r.left, left: R.right, bottom: (r) => r.top, top: R.bottom }[edge]);
}

/** Magic Resizer: scale every selected shape by settings.resizeFactor around its centre, optionally font + line. */
async function magicResize(factorOverride) {
  const k = factorOverride || Number(settings.resizeFactor) || 1;
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (!selected.length) throw new Error("Nothing selected.");
    for (const s of selected) {
      s.next = { left: R.cx(s) - (s.width * k) / 2, top: R.cy(s) - (s.height * k) / 2, width: s.width * k, height: s.height * k };
      if (settings.resizeLine) { s.proxy.lineFormat.load("weight,visible"); }
      if (settings.resizeFont) { s.tf = s.proxy.getTextFrameOrNullObject ? s.proxy.getTextFrameOrNullObject() : null; if (s.tf) s.tf.load("isNullObject,textRange/font/size"); }
    }
    await context.sync();
    for (const s of selected) {
      try { if (settings.resizeLine && s.proxy.lineFormat.visible && s.proxy.lineFormat.weight > 0) s.proxy.lineFormat.weight = s.proxy.lineFormat.weight * k; } catch (_) { /* no line */ }
      try { if (settings.resizeFont && s.tf && !s.tf.isNullObject && s.tf.textRange.font.size) s.tf.textRange.font.size = Math.max(1, Math.round(s.tf.textRange.font.size * k * 2) / 2); } catch (_) { /* no text */ }
    }
    log(`magicResize ×${k} (${settings.resizeFont ? "font " : ""}${settings.resizeLine ? "line" : ""})`);
    await writeGeometry(context, "magicResize", selected);
  });
}

/** Slice / multiply: split the selected shape into rows × cols with gaps. Copies fill, line, text; geometry preset is not readable so new cells are rectangles. */
async function sliceShape() {
  const { rows, cols, gapX, gapY } = settings.slice;
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (selected.length !== 1) throw new Error("Select exactly one shape to slice.");
    const s = selected[0];
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    s.proxy.fill.load("foregroundColor,type,transparency");
    s.proxy.lineFormat.load("color,weight,visible,dashStyle");
    const tf = s.proxy.getTextFrameOrNullObject ? s.proxy.getTextFrameOrNullObject() : null;
    if (tf) tf.load("isNullObject,textRange/text,textRange/font/size,textRange/font/bold,textRange/font/color,textRange/font/name");
    await context.sync();
    const cw = (s.width - (cols - 1) * gapX) / cols, ch = (s.height - (rows - 1) * gapY) / rows;
    if (cw <= 0 || ch <= 0) throw new Error("Gaps too large for that shape.");
    const ids = [];
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const left = s.left + c * (cw + gapX), top = s.top + r * (ch + gapY);
      if (r === 0 && c === 0) { s.proxy.left = left; s.proxy.top = top; s.proxy.width = cw; s.proxy.height = ch; ids.push(s.id); continue; }
      const n = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left, top, width: cw, height: ch });
      n.name = `${s.name} ${r + 1}-${c + 1}`;
      try { if (s.proxy.fill.type === "Solid") n.fill.setSolidColor(s.proxy.fill.foregroundColor); } catch (_) { /* keep default */ }
      try {
        n.lineFormat.visible = s.proxy.lineFormat.visible;
        if (s.proxy.lineFormat.visible) { n.lineFormat.color = s.proxy.lineFormat.color; n.lineFormat.weight = s.proxy.lineFormat.weight; n.lineFormat.dashStyle = s.proxy.lineFormat.dashStyle; }
      } catch (_) { /* ignore */ }
      try {
        if (tf && !tf.isNullObject) {
          n.textFrame.textRange.text = tf.textRange.text;
          if (tf.textRange.font.size) n.textFrame.textRange.font.size = tf.textRange.font.size;
          if (tf.textRange.font.bold != null) n.textFrame.textRange.font.bold = tf.textRange.font.bold;
          if (tf.textRange.font.color) n.textFrame.textRange.font.color = tf.textRange.font.color;
          if (tf.textRange.font.name) n.textFrame.textRange.font.name = tf.textRange.font.name;
        }
      } catch (_) { /* ignore */ }
      n.load("id");
      ids.push(n);
    }
    await context.sync();
    log(`slice: "${s.name}" → ${rows}×${cols} cells of ${cw.toFixed(1)}×${ch.toFixed(1)}`);
    try { slide.setSelectedShapes(ids.map((x) => (typeof x === "string" ? x : x.id))); await context.sync(); } catch (_) { /* fine */ }
  });
}

// ---- Position (C) ----

function alignEdge(edge) {
  return (t, r) => {
    switch (edge) {
      case "left": return { left: r.left };
      case "right": return { left: R.right(r) - t.width };
      case "top": return { top: r.top };
      case "bottom": return { top: R.bottom(r) - t.height };
      case "center": return { left: R.cx(r) - t.width / 2 };
      case "middle": return { top: R.cy(r) - t.height / 2 };
    }
  };
}

/** Dock: move the target in `dir` until it touches the reference. */
function dock(dir) {
  return (t, r) => {
    switch (dir) {
      case "right": return { left: r.left - t.width };
      case "left": return { left: R.right(r) };
      case "down": return { top: r.top - t.height };
      case "up": return { top: R.bottom(r) };
    }
  };
}

/** Distribute: outer two stay, the gaps between all shapes are evened. Needs 3+. */
function distribute(axis) {
  return (items) => {
    const pos = axis === "h" ? "left" : "top", size = axis === "h" ? "width" : "height";
    const sorted = [...items].sort((a, b) => a[pos] - b[pos]);
    const first = sorted[0], last = sorted[sorted.length - 1];
    const inner = sorted.slice(1, -1);
    const innerSize = inner.reduce((s, x) => s + x[size], 0);
    const gap = (last[pos] - (first[pos] + first[size]) - innerSize) / (inner.length + 1);
    let cursor = first[pos] + first[size] + gap;
    for (const x of inner) { x.next = { [pos]: cursor }; cursor += x[size] + gap; }
  };
}

/** Stack: butt shapes together in selection order; align the cross-axis edge to the first shape. */
function stack(axis) {
  return (items) => {
    const gap = Number(CONFIG.STACK_GAP) || 0;
    const first = items[0];
    let cursor = axis === "h" ? R.right(first) : R.bottom(first);
    for (const x of items.slice(1)) {
      x.next = axis === "h" ? { left: cursor + gap, top: first.top } : { top: cursor + gap, left: first.left };
      cursor += gap + (axis === "h" ? x.width : x.height);
    }
  };
}

/** Golden canon: place targets inside the reference vertically so bottom margin = 2 × top margin. */
function goldenCanon(t, r) {
  const free = r.height - t.height;
  return { top: r.top + free / 3 };
}

/** Align in matrix: arrange all selected shapes into rows × cols starting at the first shape's position. */
function matrix(items) {
  const { rows, cols, gapX, gapY } = settings.matrix;
  const origin = items[0];
  const cw = Math.max(...items.map((x) => x.width)), ch = Math.max(...items.map((x) => x.height));
  items.forEach((x, i) => {
    if (i >= rows * cols) return;
    const r = Math.floor(i / cols), c = i % cols;
    x.next = { left: origin.left + c * (cw + gapX) + (cw - x.width) / 2, top: origin.top + r * (ch + gapY) + (ch - x.height) / 2 };
  });
}

/** Swap two shapes: positions by centre, optionally size and z-order. */
async function swapShapes() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context, supports("PowerPointApi", "1.8") ? "items/zOrderPosition" : "");
    if (selected.length !== 2) throw new Error("Select exactly two shapes to swap.");
    const [a, b] = selected;
    const ac = { x: R.cx(a), y: R.cy(a) }, bc = { x: R.cx(b), y: R.cy(b) };
    if (CONFIG.SWAP_SIZE) {
      a.next = { left: b.left, top: b.top, width: b.width, height: b.height };
      b.next = { left: a.left, top: a.top, width: a.width, height: a.height };
    } else {
      a.next = { left: bc.x - a.width / 2, top: bc.y - a.height / 2 };
      b.next = { left: ac.x - b.width / 2, top: ac.y - b.height / 2 };
    }
    await writeGeometry(context, "swap", [a, b]);
    if (CONFIG.SWAP_ZORDER && supports("PowerPointApi", "1.8")) {
      const za = a.proxy.zOrderPosition, zb = b.proxy.zOrderPosition;
      const [lo, hi] = za < zb ? [a, b] : [b, a];
      const steps = Math.abs(za - zb);
      for (let i = 0; i < steps; i++) lo.proxy.setZOrder(PowerPoint.ShapeZOrder.bringForward);
      for (let i = 0; i < steps - 1; i++) hi.proxy.setZOrder(PowerPoint.ShapeZOrder.sendBackward);
      await context.sync();
      log(`  z-order swapped (${za} ↔ ${zb})`);
    }
  });
}

/** Nudge all selected shapes by settings.nudge points. */
function nudge(dx, dy) {
  return (items) => { const n = Number(settings.nudge) || 1; for (const x of items) x.next = { left: x.left + dx * n, top: x.top + dy * n }; };
}

/** Align in table: snap each selected loose shape into the centre of the table cell it overlaps. Assumes uniform cells (the API exposes no row/column sizes). */
async function alignInTable() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context, "items/type");
    const table = selected.find((s) => s.proxy.type === "Table");
    if (!table) throw new Error("Include a table in the selection (select it last).");
    const t = table.proxy.getTable();
    t.load("rowCount,columnCount");
    await context.sync();
    const cw = table.width / t.columnCount, ch = table.height / t.rowCount;
    const loose = selected.filter((s) => s !== table);
    for (const s of loose) {
      const c = Math.min(t.columnCount - 1, Math.max(0, Math.floor((R.cx(s) - table.left) / cw)));
      const r = Math.min(t.rowCount - 1, Math.max(0, Math.floor((R.cy(s) - table.top) / ch)));
      s.next = { left: table.left + c * cw + (cw - s.width) / 2, top: table.top + r * ch + (ch - s.height) / 2 };
    }
    log(`alignInTable: ${loose.length} shape(s) into ${t.rowCount}×${t.columnCount} table (uniform cells assumed)`);
    await writeGeometry(context, "alignInTable", loose);
  });
}

// ---- Colour (F) ----
async function applyColor(kind, hex) {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (!selected.length) throw new Error("Nothing selected.");
    for (const s of selected) {
      try {
        if (kind === "fill") s.proxy.fill.setSolidColor(hex);
        else if (kind === "line") { s.proxy.lineFormat.visible = true; s.proxy.lineFormat.color = hex; }
        else if (kind === "font") s.proxy.textFrame.textRange.font.color = hex;
      } catch (err) { log(`  "${s.name}": ${err.message || err}`); }
    }
    await context.sync();
    log(`${kind} colour ${hex} → ${selected.length} shape(s)`);
  });
}

// ---- Text (G) ----
async function forEachTextFrame(actionId, fn) {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (!selected.length) throw new Error("Nothing selected.");
    const frames = selected.map((s) => ({ s, tf: s.proxy.getTextFrameOrNullObject ? s.proxy.getTextFrameOrNullObject() : s.proxy.textFrame }));
    for (const f of frames) f.tf.load("isNullObject,wordWrap,autoSizeSetting,leftMargin,rightMargin,topMargin,bottomMargin,textRange/text");
    await context.sync();
    const live = frames.filter((f) => !f.tf.isNullObject);
    // Toggles decide ONE target for the whole selection: on unless everything is already on.
    const all = live.map((f) => f.tf);
    let n = 0;
    for (const f of live) { await fn(f.tf, f.s, context, all); n++; }
    await context.sync();
    log(`${actionId}: ${n} text frame(s)`);
  });
}

const setMargins = () => forEachTextFrame("setMargins", (tf) => { const m = settings.margins; tf.leftMargin = +m.left; tf.rightMargin = +m.right; tf.topMargin = +m.top; tf.bottomMargin = +m.bottom; });
const marginsZero = () => forEachTextFrame("marginsZero", (tf) => { tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0; });
// Toggle between "shape fits text" and "no auto-fit". Never "text fits shape", and never
// touches word wrap — the two toggles are independent.
const fitFormToText = () => forEachTextFrame("fitShapeToggle", (tf, s, context, all) => {
  const allOn = all.every((t) => t.autoSizeSetting === "AutoSizeShapeToFitText");
  tf.autoSizeSetting = allOn ? "AutoSizeNone" : "AutoSizeShapeToFitText";
});
const wrapToggle = () => forEachTextFrame("wrapToggle", (tf, s, context, all) => { tf.wordWrap = !all.every((t) => t.wordWrap); });
const setFontSize = () => forEachTextFrame("setFontSize", (tf) => { tf.textRange.font.size = Number(settings.fontSize) || 12; });
const bulletsToggle = () => forEachTextFrame("bulletsToggle", async (tf, s, context) => {
  const bf = tf.textRange.paragraphFormat.bulletFormat; bf.load("visible"); await context.sync(); bf.visible = !bf.visible;
});

/** Split the text box being edited at the cursor: text before stays, text after goes into a new box below. */
async function splitTextBox() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (selected.length !== 1) throw new Error("Click into one text box first.");
    const s = selected[0];
    const cursor = context.presentation.getSelectedTextRange();
    cursor.load("start,length");
    const tf = s.proxy.textFrame;
    tf.load("textRange/text,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,autoSizeSetting,textRange/font/size,textRange/font/bold,textRange/font/color,textRange/font/name");
    s.proxy.fill.load("foregroundColor,type");
    s.proxy.lineFormat.load("color,weight,visible");
    await context.sync();
    const text = tf.textRange.text, at = cursor.start;
    if (at <= 0 || at >= text.length) throw new Error("Put the cursor somewhere in the middle of the text.");
    const before = text.slice(0, at).replace(/[\r\n]+$/, ""), after = text.slice(at).replace(/^[\r\n]+/, "");
    tf.textRange.text = before;
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const n = slide.shapes.addTextBox(after, { left: s.left, top: R.bottom(s) + 6, width: s.width, height: s.height });
    n.name = s.name + " (split)";
    try { if (s.proxy.fill.type === "Solid") n.fill.setSolidColor(s.proxy.fill.foregroundColor); } catch (_) { /* ignore */ }
    try { n.lineFormat.visible = s.proxy.lineFormat.visible; if (s.proxy.lineFormat.visible) { n.lineFormat.color = s.proxy.lineFormat.color; n.lineFormat.weight = s.proxy.lineFormat.weight; } } catch (_) { /* ignore */ }
    const nt = n.textFrame;
    nt.leftMargin = tf.leftMargin; nt.rightMargin = tf.rightMargin; nt.topMargin = tf.topMargin; nt.bottomMargin = tf.bottomMargin;
    nt.wordWrap = tf.wordWrap; nt.autoSizeSetting = tf.autoSizeSetting;
    if (tf.textRange.font.size) nt.textRange.font.size = tf.textRange.font.size;
    if (tf.textRange.font.bold != null) nt.textRange.font.bold = tf.textRange.font.bold;
    if (tf.textRange.font.color) nt.textRange.font.color = tf.textRange.font.color;
    if (tf.textRange.font.name) nt.textRange.font.name = tf.textRange.font.name;
    await context.sync();
    log(`splitTextBox: "${s.name}" at ${at} → "${before.slice(0, 20)}…" + "${after.slice(0, 20)}…"`);
  });
}

/** Merge selected text boxes into the first (selection order), paragraphs appended; the rest are deleted. */
async function mergeTextBoxes() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (selected.length < 2) throw new Error("Select two or more text boxes.");
    for (const s of selected) { s.tf = s.proxy.textFrame; s.tf.load("textRange/text"); }
    await context.sync();
    const first = selected[0];
    const merged = selected.map((s) => s.tf.textRange.text.replace(/[\r\n]+$/, "")).filter(Boolean).join("\n");
    first.tf.textRange.text = merged;
    for (const s of selected.slice(1)) s.proxy.delete();
    await context.sync();
    log(`mergeTextBoxes: ${selected.length} → "${first.name}"`);
  });
}

// ---- Format Wizard / My Formats / Painter (B7–B9) ----
const FORMAT_PROPS = "fill/foregroundColor,fill/type,fill/transparency,lineFormat/color,lineFormat/weight,lineFormat/visible,lineFormat/dashStyle";
const TEXT_PROPS = "isNullObject,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,autoSizeSetting,verticalAlignment,textRange/font/name,textRange/font/size,textRange/font/bold,textRange/font/italic,textRange/font/color,textRange/paragraphFormat/horizontalAlignment";
let formatClipboard = null;
let painterOn = false;

async function readFormat(context, proxy) {
  proxy.load(FORMAT_PROPS);
  const tf = proxy.getTextFrameOrNullObject ? proxy.getTextFrameOrNullObject() : proxy.textFrame;
  tf.load(TEXT_PROPS);
  await context.sync();
  const f = {
    fill: proxy.fill.type === "Solid" ? { color: proxy.fill.foregroundColor, transparency: proxy.fill.transparency } : proxy.fill.type === "NoFill" ? { none: true } : null,
    line: { visible: proxy.lineFormat.visible, color: proxy.lineFormat.color, weight: proxy.lineFormat.weight, dashStyle: proxy.lineFormat.dashStyle },
  };
  if (!tf.isNullObject) {
    f.text = { leftMargin: tf.leftMargin, rightMargin: tf.rightMargin, topMargin: tf.topMargin, bottomMargin: tf.bottomMargin, wordWrap: tf.wordWrap, autoSizeSetting: tf.autoSizeSetting, verticalAlignment: tf.verticalAlignment };
    const fo = tf.textRange.font;
    f.font = { name: fo.name, size: fo.size, bold: fo.bold, italic: fo.italic, color: fo.color };
    f.para = { horizontalAlignment: tf.textRange.paragraphFormat.horizontalAlignment };
  }
  return f;
}

function writeFormat(proxy, f) {
  try { if (f.fill && f.fill.none) proxy.fill.clear(); else if (f.fill && f.fill.color) { proxy.fill.setSolidColor(f.fill.color); if (f.fill.transparency != null) proxy.fill.transparency = f.fill.transparency; } } catch (_) { /* ignore */ }
  try { if (f.line) { proxy.lineFormat.visible = !!f.line.visible; if (f.line.visible) { if (f.line.color) proxy.lineFormat.color = f.line.color; if (f.line.weight) proxy.lineFormat.weight = f.line.weight; if (f.line.dashStyle) proxy.lineFormat.dashStyle = f.line.dashStyle; } } } catch (_) { /* ignore */ }
  if (f.text || f.font || f.para) {
    try {
      const tf = proxy.textFrame;
      if (f.text) for (const k of Object.keys(f.text)) if (f.text[k] != null) tf[k] = f.text[k];
      if (f.font) for (const k of Object.keys(f.font)) if (f.font[k] != null) tf.textRange.font[k] = f.font[k];
      if (f.para && f.para.horizontalAlignment) tf.textRange.paragraphFormat.horizontalAlignment = f.para.horizontalAlignment;
    } catch (_) { /* shape has no text */ }
  }
}

/** Pick up the reference's format. One shape selected → pick up only. Several → pick up from the last and apply to the rest. */
async function formatPickup() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (!selected.length) throw new Error("Nothing selected.");
    const ref = selected[selected.length - 1];
    formatClipboard = await readFormat(context, ref.proxy);
    log(`format picked up from "${ref.name}": ${JSON.stringify(formatClipboard).slice(0, 120)}…`);
    const targets = selected.slice(0, -1);
    for (const t of targets) writeFormat(t.proxy, formatClipboard);
    if (targets.length) { await context.sync(); log(`  applied to ${targets.length} shape(s)`); }
    renderFormats();
  });
}

async function formatApply(preset) {
  const f = preset || formatClipboard;
  if (!f) throw new Error("Nothing picked up yet (use Pick up format first).");
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (!selected.length) throw new Error("Nothing selected.");
    for (const t of selected) writeFormat(t.proxy, f);
    await context.sync();
    log(`format applied to ${selected.length} shape(s)`);
  });
}

async function formatPainterToggle() {
  painterOn = !painterOn;
  if (painterOn && !formatClipboard) { painterOn = false; throw new Error("Pick up a format first, then turn the painter on."); }
  log(`format painter ${painterOn ? "ON — every new selection gets the format until you toggle it off" : "off"}`);
  renderFormats();
}

async function saveMyFormat() {
  if (!formatClipboard) throw new Error("Pick up a format first.");
  const name = (el("myformat-name") && el("myformat-name").value.trim()) || `Format ${settings.myFormats.length + 1}`;
  settings.myFormats = settings.myFormats.filter((f) => f.name !== name).concat([{ name, props: formatClipboard }]);
  await saveSettings();
  renderFormats();
  log(`saved format "${name}"`);
}

// ---- Tools (I) ----
async function selectSimilar() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context, "items/type");
    if (selected.length !== 1) throw new Error("Select exactly one shape as the pattern.");
    const s = selected[0];
    s.proxy.fill.load("foregroundColor,type");
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const all = slide.shapes;
    all.load("items/id,items/type,items/fill/foregroundColor,items/fill/type");
    await context.sync();
    const ids = all.items.filter((x) => x.type === s.proxy.type && x.fill.type === s.proxy.fill.type && (x.fill.type !== "Solid" || x.fill.foregroundColor === s.proxy.fill.foregroundColor)).map((x) => x.id);
    slide.setSelectedShapes(ids);
    await context.sync();
    log(`selectSimilar: ${ids.length} shape(s) like "${s.name}" (${s.proxy.type}, ${s.proxy.fill.type} ${s.proxy.fill.foregroundColor || ""})`);
  });
}

/** Decompose a table into one text box per cell (uniform cell sizes; the API exposes no row/column sizes). */
async function decomposeTable() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context, "items/type");
    const table = selected.find((s) => s.proxy.type === "Table");
    if (!table) throw new Error("Select a table.");
    const t = table.proxy.getTable();
    t.load("rowCount,columnCount");
    await context.sync();
    const cells = [];
    for (let r = 0; r < t.rowCount; r++) for (let c = 0; c < t.columnCount; c++) {
      const cell = t.getCellOrNullObject(r, c);
      cell.load("isNullObject,text");
      cells.push({ r, c, cell });
    }
    await context.sync();
    const cw = table.width / t.columnCount, ch = table.height / t.rowCount;
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    let n = 0;
    for (const { r, c, cell } of cells) {
      if (cell.isNullObject) continue;
      const box = slide.shapes.addTextBox(cell.text || "", { left: table.left + c * cw, top: table.top + r * ch, width: cw, height: ch });
      box.name = `${table.name} r${r + 1}c${c + 1}`;
      box.lineFormat.visible = true; box.lineFormat.color = "#808080"; box.lineFormat.weight = 0.75;
      box.textFrame.autoSizeSetting = "AutoSizeNone";
      n++;
    }
    table.proxy.delete();
    await context.sync();
    log(`decomposeTable: ${n} cells → text boxes`);
  });
}

/** Insert the selected slides as pictures on a new slide, tiled to fit. */
async function slidesAsPictures() {
  return PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedSlides();
    sel.load("items/id");
    await context.sync();
    if (!sel.items.length) throw new Error("Select one or more slides (in the thumbnail pane).");
    const images = [];
    for (const s of sel.items) { const img = s.getImageAsBase64({ width: 640 }); images.push(img); }
    await context.sync();
    const newSlide = context.presentation.slides.add();
    context.presentation.slides.load("items/id");
    await context.sync();
    const target = context.presentation.slides.items[context.presentation.slides.items.length - 1];
    context.presentation.setSelectedSlides([target.id]);
    await context.sync();
    const n = images.length, cols = Math.ceil(Math.sqrt(n)), rows = Math.ceil(n / cols);
    const gap = 12, W = CONFIG.SLIDE.width - 2 * 36, H = CONFIG.SLIDE.height - 2 * 36;
    const w = (W - (cols - 1) * gap) / cols, h = Math.min(w * 9 / 16, (H - (rows - 1) * gap) / rows);
    for (let i = 0; i < n; i++) {
      const r = Math.floor(i / cols), c = i % cols;
      await new Promise((resolve, reject) => Office.context.document.setSelectedDataAsync(images[i].value, {
        coercionType: Office.CoercionType.Image, imageLeft: 36 + c * (w + gap), imageTop: 36 + r * (h + gap), imageWidth: w, imageHeight: h,
      }, (res) => (res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(res.error && res.error.message)))));
    }
    log(`slidesAsPictures: ${n} slide(s) tiled ${rows}×${cols} on a new slide (added at the end — the API has no insert-at)`);
    void newSlide;
  });
}

/** Read the selected shape's geometry into the Dimensions panel. */
async function refreshDimensions() {
  const box = el("dims");
  if (!box) return;
  try {
    await PowerPoint.run(async (context) => {
      const selected = await loadSelection(context);
      if (selected.length !== 1) { box.dataset.id = ""; ["left", "top", "width", "height"].forEach((k) => { el("dim-" + k).value = ""; }); el("dims-name").textContent = selected.length ? `${selected.length} shapes` : "nothing selected"; return; }
      const s = selected[0];
      box.dataset.id = s.id;
      el("dims-name").textContent = s.name;
      for (const k of ["left", "top", "width", "height"]) el("dim-" + k).value = (Math.round(s[k] * 100) / 100).toString();
    });
  } catch (_) { /* pane closed or no doc */ }
}

async function applyDimensions() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (selected.length !== 1) throw new Error("Select exactly one shape.");
    const s = selected[0];
    s.next = {};
    for (const k of ["left", "top", "width", "height"]) { const v = parseFloat(el("dim-" + k).value); if (!isNaN(v)) s.next[k] = v; }
    await writeGeometry(context, "dimensions", [s]);
  });
}

/** Render the selected shape as PNG into the pane (right-click → save / copy). */
async function exportImage() {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    if (selected.length !== 1) throw new Error("Select exactly one shape.");
    const img = selected[0].proxy.getImageAsBase64({ width: 1600 });
    await context.sync();
    const out = el("export-output");
    if (out) out.innerHTML = `<img src="data:image/png;base64,${img.value}" alt="${selected[0].name}"><p class="muted">Right-click the image → Copy Image / Save Image.</p>`;
    log(`exportImage: "${selected[0].name}" rendered (${Math.round(img.value.length * 0.75 / 1024)} KB)`);
  });
}

async function setVisible(visible) {
  return PowerPoint.run(async (context) => {
    if (!visible) {
      const selected = await loadSelection(context);
      if (!selected.length) throw new Error("Nothing selected.");
      for (const s of selected) s.proxy.visible = false;
      await context.sync();
      log(`hid ${selected.length} shape(s)`);
    } else {
      const shapes = context.presentation.getSelectedSlides().getItemAt(0).shapes;
      shapes.load("items/id,items/visible");
      await context.sync();
      const hidden = shapes.items.filter((s) => s.visible === false);
      for (const s of hidden) s.visible = true;
      await context.sync();
      log(`unhid ${hidden.length} shape(s)`);
    }
  });
}


// ---- Privileged notice: toggle a legal banner at the top of every slide ----
//   Copied from "Privileged notification.pptx": rounded rectangle 451.6×19.5pt at (266.3, 0),
//   accent fill #F15D22 with the theme's 15 % shade outline, white 14pt centred text.
const NOTICE = {
  name: "EE privileged notice",
  text: "Legally Privileged – Subject to MC Approval and Consultation",
  left: 266.3, top: 0, width: 451.6, height: 19.5,
  fill: "#F15D22", line: "#240E05", lineWeight: 1,
  fontSize: 14, fontColor: "#FFFFFF",
};

async function privilegedNoticeToggle() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const per = slides.items.map((sl) => { const sh = sl.shapes; sh.load("items/id,items/name"); return { sl, sh }; });
    await context.sync();
    // On every slide already → remove everywhere. Otherwise → add to the slides missing it.
    const has = (x) => x.sh.items.some((s) => s.name === NOTICE.name);
    const everywhere = per.length > 0 && per.every(has);
    let n = 0;
    if (everywhere) {
      for (const x of per) for (const s of x.sh.items) if (s.name === NOTICE.name) { s.delete(); n++; }
      await context.sync();
      log(`privileged notice: removed from ${n} slide(s)`);
    } else {
      for (const x of per) {
        if (has(x)) continue;
        const box = x.sh.addGeometricShape("RoundRectangle" /* enum key is roundRectangle; the literal avoids an undefined type → line */, { left: NOTICE.left, top: NOTICE.top, width: NOTICE.width, height: NOTICE.height });
        box.name = NOTICE.name;
        box.fill.setSolidColor(NOTICE.fill);
        box.lineFormat.color = NOTICE.line; box.lineFormat.weight = NOTICE.lineWeight; box.lineFormat.visible = true;
        const tf = box.textFrame;
        tf.textRange.text = NOTICE.text;
        tf.verticalAlignment = "Middle";
        tf.wordWrap = false; tf.autoSizeSetting = "AutoSizeNone";
        tf.textRange.font.size = NOTICE.fontSize; tf.textRange.font.color = NOTICE.fontColor; tf.textRange.font.bold = false;
        tf.textRange.paragraphFormat.horizontalAlignment = "Center";
        n++;
      }
      await context.sync();
      log(`privileged notice: added to ${n} slide(s) (now on all ${per.length})`);
    }
  });
}

/** Duplicate the current slide in place: export it as a .pptx in memory, insert it right after itself. Selection stays on the original. */
async function duplicateSlide() {
  return PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedSlides();
    sel.load("items/id");
    await context.sync();
    if (!sel.items.length) throw new Error("No slide selected.");
    const slide = sel.items[0];
    const exported = slide.exportAsBase64();
    await context.sync();
    context.presentation.insertSlidesFromBase64(exported.value, { formatting: "KeepSourceFormatting", targetSlideId: slide.id });
    await context.sync();
    log(`duplicateSlide: copy inserted after the current slide (${Math.round(exported.value.length * 0.75 / 1024)} KB)`);
  });
}


// ---- Export selected slides as a new presentation (package surgery on the real .pptx) ----
//   getFileAsync gives the whole current deck as bytes; we delete the unselected slide parts
//   (masters/layouts/theme/media untouched) and open the result with PowerPoint.createPresentation.
//   simplify=true also drops layouts and masters that no remaining slide uses.

function getDeckBytes() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(r.error && r.error.message || "getFileAsync failed"));
      const file = r.value, parts = [];
      let i = 0;
      const next = () => {
        if (i >= file.sliceCount) {
          file.closeAsync(() => {});
          const total = parts.reduce((n, p) => n + p.length, 0);
          const out = new Uint8Array(total); let off = 0;
          for (const p of parts) { out.set(p, off); off += p.length; }
          return resolve(out);
        }
        file.getSliceAsync(i, (sr) => {
          if (sr.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(() => {}); return reject(new Error(sr.error && sr.error.message || "getSliceAsync failed")); }
          const d = sr.value.data;
          parts.push(d instanceof Uint8Array ? d : Array.isArray(d) ? Uint8Array.from(d) : new Uint8Array(d));
          i++; next();
        });
      };
      next();
    });
  });
}

const xmlParse = (text) => new DOMParser().parseFromString(text, "application/xml");
const xmlOut = (doc) => new XMLSerializer().serializeToString(doc);
const relsPath = (partPath) => { const i = partPath.lastIndexOf("/"); return partPath.slice(0, i) + "/_rels/" + partPath.slice(i + 1) + ".rels"; };
const resolveTarget = (fromPart, target) => {
  if (target.startsWith("/")) return target.slice(1);
  const base = fromPart.slice(0, fromPart.lastIndexOf("/")).split("/");
  for (const seg of target.split("/")) { if (seg === "..") base.pop(); else if (seg !== ".") base.push(seg); }
  return base.join("/");
};

async function readRels(zip, partPath) {
  const f = zip.file(relsPath(partPath));
  if (!f) return { doc: null, rels: [] };
  const doc = xmlParse(await f.async("string"));
  const rels = [...doc.getElementsByTagName("Relationship")].map((el) => ({ el, id: el.getAttribute("Id"), type: el.getAttribute("Type"), target: resolveTarget(partPath, el.getAttribute("Target")), raw: el.getAttribute("Target"), mode: el.getAttribute("TargetMode") }));
  return { doc, rels };
}

function removeContentType(ctDoc, partPath) {
  for (const o of [...ctDoc.getElementsByTagName("Override")]) if (o.getAttribute("PartName") === "/" + partPath) o.parentNode.removeChild(o);
}

async function deletePart(zip, ctDoc, partPath) {
  zip.remove(partPath); zip.remove(relsPath(partPath)); removeContentType(ctDoc, partPath);
}

async function buildSubsetDeck(bytes, keepIndices, simplify) {
  if (typeof JSZip === "undefined") throw new Error("JSZip didn't load (offline?) — the export needs it.");
  const zip = await JSZip.loadAsync(bytes);
  const ctDoc = xmlParse(await zip.file("[Content_Types].xml").async("string"));
  const presPath = "ppt/presentation.xml";
  const presDoc = xmlParse(await zip.file(presPath).async("string"));
  const presRels = await readRels(zip, presPath);
  const sldIds = [...presDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sldId")];
  const keep = new Set(keepIndices);
  let removed = 0;
  for (let i = 0; i < sldIds.length; i++) {
    if (keep.has(i)) continue;
    const el = sldIds[i];
    const rid = el.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || el.getAttribute("r:id");
    const rel = presRels.rels.find((r) => r.id === rid);
    el.parentNode.removeChild(el);
    if (rel) {
      rel.el.parentNode.removeChild(rel.el);
      // notes slide + comments hanging off this slide go too
      const srels = await readRels(zip, rel.target);
      for (const r of srels.rels) if (/notesSlide|comments/.test(r.type) && r.mode !== "External") await deletePart(zip, ctDoc, r.target);
      await deletePart(zip, ctDoc, rel.target);
      removed++;
    }
  }
  let droppedLayouts = 0, droppedMasters = 0;
  if (simplify) {
    // layouts used by the remaining slides
    const usedLayouts = new Set();
    for (const rel of presRels.rels.filter((r) => /\/slide$/.test(r.type))) {
      const srels = await readRels(zip, rel.target);
      for (const r of srels.rels) if (/slideLayout$/.test(r.type)) usedLayouts.add(r.target);
    }
    const masterRels = presRels.rels.filter((r) => /slideMaster$/.test(r.type));
    const sldMasterIds = [...presDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sldMasterId")];
    for (const mrel of masterRels) {
      const mPath = mrel.target;
      const mDoc = xmlParse(await zip.file(mPath).async("string"));
      const mRels = await readRels(zip, mPath);
      const layoutIds = [...mDoc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sldLayoutId")];
      let kept = 0;
      for (const lid of layoutIds) {
        const rid = lid.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || lid.getAttribute("r:id");
        const lrel = mRels.rels.find((r) => r.id === rid);
        if (lrel && !usedLayouts.has(lrel.target)) {
          lid.parentNode.removeChild(lid); lrel.el.parentNode.removeChild(lrel.el);
          await deletePart(zip, ctDoc, lrel.target); droppedLayouts++;
        } else kept++;
      }
      if (kept === 0) {
        const mid = sldMasterIds.find((x) => (x.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || x.getAttribute("r:id")) === mrel.id);
        if (mid) mid.parentNode.removeChild(mid);
        mrel.el.parentNode.removeChild(mrel.el);
        await deletePart(zip, ctDoc, mPath); droppedMasters++;
      } else {
        zip.file(mPath, xmlOut(mDoc)); zip.file(relsPath(mPath), xmlOut(mRels.doc));
      }
    }
  }
  zip.file(presPath, xmlOut(presDoc));
  zip.file(relsPath(presPath), xmlOut(presRels.doc));
  zip.file("[Content_Types].xml", xmlOut(ctDoc));
  const base64 = await zip.generateAsync({ type: "base64", compression: "DEFLATE" });
  return { base64, removed, droppedLayouts, droppedMasters, total: sldIds.length };
}

/** Selected slides → a new presentation window, template intact (or simplified). */
async function exportSelectedSlides(simplify) {
  const keepIndices = await PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedSlides(); sel.load("items/id");
    const all = context.presentation.slides; all.load("items/id");
    await context.sync();
    if (!sel.items.length) throw new Error("Select one or more slides first (thumbnail pane).");
    const ids = new Set(sel.items.map((s) => s.id));
    return all.items.map((s, i) => (ids.has(s.id) ? i : -1)).filter((i) => i >= 0);
  });
  log(`export: reading the deck…`);
  const bytes = await getDeckBytes();
  log(`export: ${Math.round(bytes.length / 1024)} KB, keeping ${keepIndices.length} slide(s)${simplify ? ", simplifying template" : ""}`);
  const r = await buildSubsetDeck(bytes, keepIndices, simplify);
  await PowerPoint.createPresentation(r.base64);
  log(`export: opened a new presentation with ${keepIndices.length} of ${r.total} slides (removed ${r.removed}${simplify ? `, dropped ${r.droppedLayouts} layout(s) + ${r.droppedMasters} master(s)` : ""}). Save it with ⌘S.`);
}

async function goToSlide() {
  const n = Math.max(1, parseInt(settings.gotoSlide, 10) || 1);
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    if (n > slides.items.length) throw new Error(`Only ${slides.items.length} slides.`);
    context.presentation.setSelectedSlides([slides.items[n - 1].id]);
    await context.sync();
    log(`goToSlide ${n}`);
  });
}

// ---- Wizards (B1, B2) — v1 ----

/** Agenda v1: one agenda slide + one divider per item, appended at the end (the API has no insert-at). Items from settings.agenda, one per line. */
async function agendaWizard() {
  const items = String(settings.agenda || "").split("\n").map((l) => l.trim()).filter(Boolean);
  if (!items.length) throw new Error("Enter agenda items (one per line) in the Agenda box first.");
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const addSlide = async () => { slides.add(); slides.load("items/id"); await context.sync(); return slides.items[slides.items.length - 1]; };
    const W = CONFIG.SLIDE.width;
    const drawList = (slide, activeIndex) => {
      const title = slide.shapes.addTextBox("Agenda", { left: 60, top: 40, width: W - 120, height: 50 });
      title.name = "CE agenda title"; title.textFrame.textRange.font.size = 32; title.textFrame.textRange.font.bold = true;
      items.forEach((it, i) => {
        const box = slide.shapes.addTextBox(`${i + 1}  ${it}`, { left: 80, top: 120 + i * 36, width: W - 160, height: 32 });
        box.name = `CE agenda item ${i + 1}`;
        box.textFrame.textRange.font.size = 20;
        box.textFrame.textRange.font.bold = activeIndex === i;
        box.textFrame.textRange.font.color = activeIndex === null || activeIndex === i ? "#000000" : "#9A9A9A";
      });
    };
    // Where to put them: right after the current slide (moveTo, PowerPointApi 1.8).
    const cur = context.presentation.getSelectedSlides();
    cur.load("items/id");
    slides.load("items/id");
    await context.sync();
    let insertAt = cur.items.length ? slides.items.findIndex((x) => x.id === cur.items[0].id) + 1 : slides.items.length;
    const overview = await addSlide(); drawList(overview, null);
    const made = [overview];
    for (let i = 0; i < items.length; i++) { const s = await addSlide(); drawList(s, i); made.push(s); }
    await context.sync();
    if (supports("PowerPointApi", "1.8")) { for (const s of made) { s.moveTo(insertAt++); await context.sync(); } }
    log(`agenda: overview + ${items.length} divider slide(s) inserted after the current slide`);
  });
}

/** Master Wizard v1: add (or remove) a text label on every slide layout of every master. */
async function masterLabel(add) {
  return PowerPoint.run(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id");
    await context.sync();
    let n = 0;
    for (const m of masters.items) {
      const layouts = m.layouts;
      layouts.load("items/id");
      await context.sync();
      for (const l of layouts.items) {
        if (add) {
          const p = settings.masterPos;
          const box = l.shapes.addTextBox(settings.masterLabel, { left: +p.left, top: +p.top, width: +p.width, height: +p.height });
          box.name = "CE master label";
          box.textFrame.textRange.font.size = 9; box.textFrame.textRange.font.color = "#808080";
          box.textFrame.textRange.paragraphFormat.horizontalAlignment = "Right";
          n++;
        } else {
          l.shapes.load("items/id,items/name");
          await context.sync();
          for (const s of l.shapes.items) if (s.name === "CE master label") { s.delete(); n++; }
        }
      }
    }
    await context.sync();
    log(`master label ${add ? "added to" : "removed from"} ${n} layout(s)`);
  });
}

// ---- Sticky notes (J1) ----
//   Geometry and styling copied from the sample deck (sample stickie.pptx): 143pt wide,
//   top-right with 24pt/54pt margins, 2.25pt thin-thick outline in the theme's dark blue,
//   12pt bold. Height is the sample's 80pt × 1.25, fixed (no auto-fit).
//   No shadow: PowerPoint's JS API exposes no shape effects at all (shadow, glow, reflection,
//   soft edges — verified against the preview API and via the Dump Shape API probe).
const STICKY = {
  WIDTH: 143, HEIGHT: 100,
  MARGIN: { top: 24, right: 54 },
  CASCADE: 18,
  FONT_SIZE: 12,
  LINE: { color: "#0E2841", weight: 2.25, style: "ThinThick" },
  COLORS: [
    { name: "Yellow", hex: "#FFFF00" }, { name: "Green", hex: "#66FF33" }, { name: "Pink", hex: "#FF66CC" },
    { name: "Orange", hex: "#FFA500" }, { name: "Blue", hex: "#33CCFF" }, { name: "Purple", hex: "#CC99FF" },
  ],
};

/** "15 Sep 26 - 8:28p" — the sample's format. */
function stickyStamp(d = new Date()) {
  const mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getMonth()];
  const h24 = d.getHours(), h = h24 % 12 || 12;
  return `${d.getDate()} ${mon} ${String(d.getFullYear()).slice(-2)} - ${h}:${String(d.getMinutes()).padStart(2, "0")}${h24 >= 12 ? "p" : "a"}`;
}
function stickyColorHex(name) { return (STICKY.COLORS.find((x) => x.name === name) || STICKY.COLORS[0]).hex; }

async function addSticky(colorName) {
  const hex = stickyColorHex(colorName || settings.stickyColor);
  const { WIDTH: w, HEIGHT: h } = STICKY;
  const header = `${settings.initials} ${stickyStamp()}:`;
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();
    const existing = shapes.items.filter((s) => /^Sticky\b/.test(s.name)).length;
    const n = existing + 1;
    const left = CONFIG.SLIDE.width - STICKY.MARGIN.right - w - existing * STICKY.CASCADE;
    const top = STICKY.MARGIN.top + existing * STICKY.CASCADE;
    const box = shapes.addTextBox(header + "\n", { left, top, width: w, height: h });
    box.name = `Sticky ${n}`;
    box.fill.setSolidColor(hex);
    box.lineFormat.color = STICKY.LINE.color; box.lineFormat.weight = STICKY.LINE.weight; box.lineFormat.style = STICKY.LINE.style;
    const tf = box.textFrame;
    tf.wordWrap = true; tf.autoSizeSetting = "AutoSizeNone";
    tf.textRange.font.size = STICKY.FONT_SIZE; tf.textRange.font.bold = true;
    box.load("id");
    await context.sync();
    try {
      tf.textRange.load("text"); await context.sync();
      tf.textRange.getSubstring(tf.textRange.text.length, 0).setSelected(); await context.sync();
    } catch (err) {
      log("sticky: could not place the cursor (" + (err.message || err) + "); selecting the note instead");
      slide.setSelectedShapes([box.id]); await context.sync();
    }
    log(`sticky ${n}: "${header}" ${hex} ${w}×${h} at (${left}, ${top})`);
  });
}

// ---- pane visibility ----
let paneVisible = null;
function paneLooksVisible() { return paneVisible !== null ? paneVisible : document.visibilityState === "visible"; }
async function togglePane() {
  if (!Office.addin || !Office.addin.showAsTaskpane) throw new Error("Office.addin.showAsTaskpane unavailable (needs SharedRuntime 1.1).");
  const before = `event=${paneVisible} doc=${document.visibilityState}`;
  if (paneLooksVisible()) { log(`togglePane: hiding (${before})`); await Office.addin.hide(); }
  else { log(`togglePane: showing (${before})`); await Office.addin.showAsTaskpane(); }
}

// ---------------------------------------------------------------------------
// 4. COMMANDS — the registry. Adding a command = adding one entry here.
//    id: stable (used in the keymap). group: pane section. icon: key into ICONS.
// ---------------------------------------------------------------------------
const COMMANDS = [
  // Position
  { id: "alignLeft", group: "Position", label: "Align left", icon: "alignLeft", desc: "Left edges to the reference (single shape: to the slide).", run: () => applyToTargets("alignLeft", alignEdge("left"), { allowSlide: true }) },
  { id: "alignRight", group: "Position", label: "Align right", icon: "alignRight", desc: "Right edges to the reference.", run: () => applyToTargets("alignRight", alignEdge("right"), { allowSlide: true }) },
  { id: "alignTop", group: "Position", label: "Align top", icon: "alignTop", desc: "Top edges to the reference.", run: () => applyToTargets("alignTop", alignEdge("top"), { allowSlide: true }) },
  { id: "alignBottom", group: "Position", label: "Align bottom", icon: "alignBottom", desc: "Bottom edges to the reference.", run: () => applyToTargets("alignBottom", alignEdge("bottom"), { allowSlide: true }) },
  { id: "alignCenter", group: "Position", label: "Align centre", icon: "alignCenter", desc: "Horizontal centres to the reference.", run: () => applyToTargets("alignCenter", alignEdge("center"), { allowSlide: true }) },
  { id: "alignMiddle", group: "Position", label: "Align middle", icon: "alignMiddle", desc: "Vertical centres to the reference.", run: () => applyToTargets("alignMiddle", alignEdge("middle"), { allowSlide: true }) },
  { id: "distributeH", group: "Position", label: "Distribute horizontally", icon: "distributeH", desc: "Outer two stay; gaps evened. 3+ shapes.", run: () => applyToAll("distributeH", distribute("h"), 3) },
  { id: "distributeV", group: "Position", label: "Distribute vertically", icon: "distributeV", desc: "Outer two stay; gaps evened. 3+ shapes.", run: () => applyToAll("distributeV", distribute("v"), 3) },
  { id: "dockLeft", group: "Position", label: "Dock left", icon: "dockLeft", desc: "Move left until touching the reference.", run: () => applyToTargets("dockLeft", dock("left")) },
  { id: "dockRight", group: "Position", label: "Dock right", icon: "dockRight", desc: "Move right until touching the reference.", run: () => applyToTargets("dockRight", dock("right")) },
  { id: "dockUp", group: "Position", label: "Dock up", icon: "dockUp", desc: "Move up until touching the reference.", run: () => applyToTargets("dockUp", dock("up")) },
  { id: "dockDown", group: "Position", label: "Dock down", icon: "dockDown", desc: "Move down until touching the reference.", run: () => applyToTargets("dockDown", dock("down")) },
  { id: "stackH", group: "Position", label: "Stack horizontally", icon: "stackH", desc: "Butt shapes together left→right in selection order.", run: () => applyToAll("stackH", stack("h"), 2) },
  { id: "stackV", group: "Position", label: "Stack vertically", icon: "stackV", desc: "Butt shapes together top→bottom in selection order.", run: () => applyToAll("stackV", stack("v"), 2) },
  { id: "swap", group: "Position", label: "Swap", icon: "swap", desc: "Exchange two shapes' positions (and layer order).", run: () => swapShapes() },
  { id: "goldenCanon", group: "Position", label: "Golden canon", icon: "golden", desc: "Place inside the reference with bottom margin = 2× top.", run: () => applyToTargets("goldenCanon", goldenCanon) },
  { id: "matrix", group: "Position", label: "Align in matrix", icon: "matrix", desc: "Arrange into rows × cols (settings below).", run: () => applyToAll("matrix", matrix, 2) },
  { id: "alignInTable", group: "Position", label: "Align in table", icon: "table", desc: "Snap loose shapes into the table cells they overlap (table selected last).", run: () => alignInTable() },
  { id: "nudgeLeft", group: "Position", label: "Nudge left", icon: "nudgeLeft", desc: "Move by the nudge amount.", run: () => applyToAll("nudgeLeft", nudge(-1, 0)) },
  { id: "nudgeRight", group: "Position", label: "Nudge right", icon: "nudgeRight", desc: "Move by the nudge amount.", run: () => applyToAll("nudgeRight", nudge(1, 0)) },
  { id: "nudgeUp", group: "Position", label: "Nudge up", icon: "nudgeUp", desc: "Move by the nudge amount.", run: () => applyToAll("nudgeUp", nudge(0, -1)) },
  { id: "nudgeDown", group: "Position", label: "Nudge down", icon: "nudgeDown", desc: "Move by the nudge amount.", run: () => applyToAll("nudgeDown", nudge(0, 1)) },
  // Size
  { id: "matchWidth", group: "Size", label: "Match width", icon: "matchWidth", desc: "Target width := reference width.", run: () => applyToTargets("matchWidth", (t, r) => geometry.match(t, r, { width: true })) },
  { id: "matchHeight", group: "Size", label: "Match height", icon: "matchHeight", desc: "Target height := reference height.", run: () => applyToTargets("matchHeight", (t, r) => geometry.match(t, r, { height: true })) },
  { id: "matchBoth", group: "Size", label: "Match both", icon: "matchBoth", desc: "Both dimensions, non-proportional.", run: () => applyToTargets("matchBoth", (t, r) => geometry.match(t, r, { width: true, height: true })) },
  { id: "fitInside", group: "Size", label: "Fit inside", icon: "fitInside", desc: "Scale proportionally to fit within the reference.", run: () => applyToTargets("fitInside", (t, r) => geometry.scale(t, r, "contain")) },
  { id: "fillOutside", group: "Size", label: "Fill reference", icon: "fillOutside", desc: "Scale proportionally to cover the reference.", run: () => applyToTargets("fillOutside", (t, r) => geometry.scale(t, r, "cover")) },
  { id: "stretchLeft", group: "Size", label: "Stretch left", icon: "stretchLeft", desc: "Left edge → reference's left edge (shrinks or flips as needed).", run: () => applyToTargets("stretchLeft", stretch("left")) },
  { id: "stretchRight", group: "Size", label: "Stretch right", icon: "stretchRight", desc: "Right edge → reference's right edge (shrinks or flips as needed).", run: () => applyToTargets("stretchRight", stretch("right")) },
  { id: "stretchUp", group: "Size", label: "Stretch up", icon: "stretchUp", desc: "Top edge → reference's top edge (shrinks or flips as needed).", run: () => applyToTargets("stretchUp", stretch("top")) },
  { id: "stretchDown", group: "Size", label: "Stretch down", icon: "stretchDown", desc: "Bottom edge → reference's bottom edge (shrinks or flips as needed).", run: () => applyToTargets("stretchDown", stretch("bottom")) },
  { id: "fillLeft", group: "Size", label: "Fill gap left", icon: "fillLeft", desc: "Left edge → reference's right edge (shrinks or flips as needed).", run: () => applyToTargets("fillLeft", fillGap("left")) },
  { id: "fillRight", group: "Size", label: "Fill gap right", icon: "fillRight", desc: "Right edge → reference's left edge (shrinks or flips as needed).", run: () => applyToTargets("fillRight", fillGap("right")) },
  { id: "fillUp", group: "Size", label: "Fill gap up", icon: "fillUp", desc: "Top edge → reference's bottom edge (shrinks or flips as needed).", run: () => applyToTargets("fillUp", fillGap("top")) },
  { id: "fillDown", group: "Size", label: "Fill gap down", icon: "fillDown", desc: "Bottom edge → reference's top edge (shrinks or flips as needed).", run: () => applyToTargets("fillDown", fillGap("bottom")) },
  { id: "resizeUp", group: "Size", label: "Resize +", icon: "resizeUp", desc: "Magic Resizer: scale by the factor (settings below).", run: () => magicResize() },
  { id: "resizeDown", group: "Size", label: "Resize −", icon: "resizeDown", desc: "Magic Resizer: scale by 1 / factor.", run: () => magicResize(1 / (Number(settings.resizeFactor) || 1.1)) },
  { id: "slice", group: "Size", label: "Slice / multiply", icon: "slice", desc: "Split one shape into rows × cols (settings below).", run: () => sliceShape() },
  // Colour (generated per palette slot below)
  // Text
  { id: "setMargins", group: "Text", label: "Set margins", icon: "margins", desc: "Apply the margins below.", run: () => setMargins() },
  { id: "marginsZero", group: "Text", label: "Zero margins", icon: "marginsZero", desc: "All four text margins to 0.", run: () => marginsZero() },
  { id: "fitFormToText", group: "Text", label: "Fit shape ↔ off", icon: "fitText", desc: "Toggle: shape resizes to its text ↔ no auto-fit (never text-shrinks-to-shape).", run: () => fitFormToText() },
  { id: "wrapToggle", group: "Text", label: "Wrap on/off", icon: "wrap", desc: "Toggle word wrap (on unless every selected box is already on).", run: () => wrapToggle() },
  { id: "splitTextBox", group: "Text", label: "Split at cursor", icon: "split", desc: "Two boxes from one, at the cursor.", run: () => splitTextBox() },
  { id: "mergeTextBoxes", group: "Text", label: "Merge boxes", icon: "merge", desc: "Combine in selection order.", run: () => mergeTextBoxes() },
  { id: "bulletsToggle", group: "Text", label: "Bullets on/off", icon: "bullets", desc: "Toggle bullets.", run: () => bulletsToggle() },
  { id: "setFontSize", group: "Text", label: "Set font size", icon: "fontSize", desc: "Apply the font size below.", run: () => setFontSize() },
  // Format
  { id: "formatPickup", group: "Format", label: "Pick up format", icon: "pickup", desc: "From the last-selected shape; applies to the others if several are selected.", run: () => formatPickup() },
  { id: "formatApply", group: "Format", label: "Apply format", icon: "apply", desc: "Apply the picked-up format to the selection.", run: () => formatApply() },
  { id: "formatPainter", group: "Format", label: "Painter on/off", icon: "painter", desc: "Apply the picked-up format to every new selection until toggled off.", run: () => formatPainterToggle() },
  { id: "saveMyFormat", group: "Format", label: "Save as My Format", icon: "star", desc: "Store the picked-up format as a named preset.", run: () => saveMyFormat() },
  ...[1, 2, 3, 4, 5].map((i) => ({ id: "myFormat" + i, group: "Format", label: "My Format " + i, icon: "preset", badge: String(i), desc: "Apply saved format #" + i + ".", run: () => { const f = settings.myFormats[i - 1]; if (!f) throw new Error(`No saved format #${i}.`); return formatApply(f.props); } })),
  // Tools
  { id: "selectSimilar", group: "Tools", label: "Select similar", icon: "similar", desc: "Select shapes on the slide with the same type and fill.", run: () => selectSimilar() },
  { id: "decomposeTable", group: "Tools", label: "Decompose table", icon: "decompose", desc: "Table → one text box per cell.", run: () => decomposeTable() },
  { id: "slidesAsPictures", group: "Tools", label: "Slides as pictures", icon: "pictures", desc: "Selected slides tiled as images on a new slide.", run: () => slidesAsPictures() },
  { id: "exportImage", group: "Tools", label: "Export as image", icon: "export", desc: "Render the selected shape to PNG in the pane.", run: () => exportImage() },
  { id: "hide", group: "Tools", label: "Hide selected", icon: "hide", desc: "Hide (keeps position and layer).", run: () => setVisible(false) },
  { id: "unhide", group: "Tools", label: "Unhide all", icon: "unhide", desc: "Show every hidden shape on the slide.", run: () => setVisible(true) },
  { id: "exportSlides", group: "Tools", label: "Slides → new deck", icon: "newDeck", desc: "Open the selected slides as a new presentation, template kept exactly.", run: () => exportSelectedSlides(false) },
  { id: "exportSlidesSimplified", group: "Tools", label: "Slides → new deck (simplified)", icon: "newDeckLite", desc: "Same, but drop layouts and masters the selected slides don't use.", run: () => exportSelectedSlides(true) },
  { id: "privilegedNotice", group: "Tools", label: "Privileged notice", icon: "notice", desc: "Put the 'Legally Privileged – Subject to MC…' banner on every slide that lacks it; if every slide already has it, remove it from all.", run: () => privilegedNoticeToggle() },
  { id: "duplicateSlide", group: "Tools", label: "Duplicate slide", icon: "duplicate", desc: "Insert an exact copy of the current slide right after it (backup before editing).", run: () => duplicateSlide() },
  { id: "goToSlide", group: "Tools", label: "Go to slide", icon: "goto", desc: "Jump to the slide number below.", run: () => goToSlide() },
  { id: "agendaWizard", group: "Tools", label: "Agenda", icon: "agenda", desc: "Agenda + divider slides from the items below (v1, appended at the end).", run: () => agendaWizard() },
  { id: "masterLabelAdd", group: "Tools", label: "Master label +", icon: "master", desc: "Add the label below to every slide layout.", run: () => masterLabel(true) },
  { id: "masterLabelRemove", group: "Tools", label: "Master label −", icon: "masterOff", desc: "Remove it from every layout.", run: () => masterLabel(false) },
  // Sticky
  { id: "addSticky", group: "Sticky", label: "Sticky", icon: "sticky", desc: "Reviewer note with initials + timestamp; cursor on the next line.", run: () => addSticky() },
  ...STICKY.COLORS.map((c) => ({ id: "sticky_" + c.name.toLowerCase(), group: "Sticky", label: c.name, icon: "sticky", color: c.hex, desc: `Sticky in ${c.name}.`, run: () => addSticky(c.name) })),
  // Add-in
  { id: "togglePane", group: "Add-in", label: "Show / hide pane", icon: "pane", desc: "Open the pane; press again to hide it.", run: () => togglePane() },
];

// Colour commands: one per palette slot × fill/line/font. Colours come from settings.palette at run time.
for (const kind of ["fill", "line", "font"]) {
  for (let i = 1; i <= DEFAULT_PALETTE.length; i++) {
    COMMANDS.push({ id: `${kind}${i}`, group: "Colour", kind, slot: i, label: `${kind[0].toUpperCase() + kind.slice(1)} ${i}`, icon: kind, desc: `${kind} colour = palette slot ${i}.`,
      run: () => { const hex = (settings.palette || DEFAULT_PALETTE)[i - 1]; if (!hex) throw new Error(`Palette slot ${i} is empty.`); return applyColor(kind, hex); } });
  }
}
// Position/Size icon order mirrors the ribbon: L/centre/R, T/middle/B, dock-L/distribute/dock-R …
const ICON_ORDER = ["alignLeft", "alignCenter", "alignRight", "alignTop", "alignMiddle", "alignBottom", "dockLeft", "distributeH", "dockRight", "dockUp", "distributeV", "dockDown", "stackH", "stackV", "swap", "alignInTable", "matrix", "goldenCanon", "nudgeLeft", "nudgeRight", "nudgeUp", "nudgeDown", "stretchLeft", "matchWidth", "stretchRight", "stretchUp", "matchHeight", "stretchDown", "fillLeft", "matchBoth", "fillRight", "fillUp", "fitInside", "fillDown", "fillOutside", "resizeUp", "resizeDown", "slice"];
COMMANDS.sort((a, b) => { const ia = ICON_ORDER.indexOf(a.id), ib = ICON_ORDER.indexOf(b.id); return (ia < 0 || ib < 0) ? 0 : ia - ib; });
const COMMAND_BY_ID = Object.fromEntries(COMMANDS.map((c) => [c.id, c]));
const GROUPS = ["Position", "Size", "Colour", "Text", "Format", "Tools", "Sticky", "Add-in"];

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

// ---------------------------------------------------------------------------
// 5. Keys — canonical combos, keymap store, slot dispatcher
// ---------------------------------------------------------------------------

const MOD_ORDER = ["Cmd", "Ctrl", "Shift", "Alt"];
const MOD_ALIASES = { command: "Cmd", cmd: "Cmd", meta: "Cmd", control: "Ctrl", ctrl: "Ctrl", shift: "Shift", alt: "Alt", option: "Alt", opt: "Alt" };
const MOD_GLYPH = { Cmd: "⌘", Ctrl: "⌃", Shift: "⇧", Alt: "⌥" };

function canonicalCombo(text) {
  const parts = text.split("+").map((p) => p.trim()).filter(Boolean);
  const mods = new Set();
  let key = null;
  for (const p of parts) { const m = MOD_ALIASES[p.toLowerCase()]; if (m) mods.add(m); else key = p.length === 1 ? p.toUpperCase() : p; }
  if (!key) return null;
  return [...MOD_ORDER.filter((m) => mods.has(m)), key].join("+");
}
function displayCombo(combo) {
  if (!combo) return "—";
  const parts = combo.split("+"); const key = parts.pop();
  return parts.map((m) => MOD_GLYPH[m] || m).join("") + key;
}
function comboFromEvent(e) {
  const code = e.code || "";
  let key = null;
  if (/^Key[A-Z]$/.test(code)) key = code.slice(3);
  else if (/^Digit[0-9]$/.test(code)) key = code.slice(5);
  else if (/^F[0-9]{1,2}$/.test(code)) key = code;
  else key = { Minus: "-", Equal: "=", Comma: ",", Period: ".", Slash: "/", Semicolon: ";", Quote: "'", BracketLeft: "[", BracketRight: "]", Backslash: "\\", Backquote: "`", Space: "Space",
    ArrowLeft: "Left", ArrowRight: "Right", ArrowUp: "Up", ArrowDown: "Down", Enter: "Return", Backspace: "Backspace", Delete: "Delete", Escape: "Esc", Tab: "Tab" }[code] || null;
  if (!key) return null;
  const mods = [];
  if (e.metaKey) mods.push("Cmd"); if (e.ctrlKey) mods.push("Ctrl"); if (e.shiftKey) mods.push("Shift"); if (e.altKey) mods.push("Alt");
  return [...mods, key].join("+");
}

const BANK_SET = new Set(KEY_BANK);
const NATIVE_BY_COMBO = (() => {
  const m = {};
  const list = typeof NATIVE_SHORTCUTS !== "undefined" ? NATIVE_SHORTCUTS : [];
  for (const row of list) { if (!Array.isArray(row) || row.length < 3) continue; const c = canonicalCombo(row[0]); if (c) (m[c] = m[c] || []).push({ what: row[1], source: row[2] }); }
  return m;
})();

const KEYMAP_STORAGE_KEY = "ppt-shortcuts.keymap.v1";
let keymap = { ...DEFAULT_KEYMAP };

function sanitizeKeymap(map) {
  const clean = {}, dropped = [];
  for (const [combo, id] of Object.entries(map || {})) {
    const c = canonicalCombo(combo);
    if (c && BANK_SET.has(c) && COMMAND_BY_ID[id] && !Object.values(clean).includes(id)) clean[c] = id;
    else dropped.push(`${combo} → ${id}`);
  }
  return { clean, dropped };
}
function comboForCommand(id) { return Object.keys(keymap).find((c) => keymap[c] === id) || null; }

async function loadKeymap() {
  const raw = await store.get(KEYMAP_STORAGE_KEY);
  if (!raw) { keymap = { ...DEFAULT_KEYMAP }; return; }
  try {
    const { clean, dropped } = sanitizeKeymap(JSON.parse(raw));
    keymap = clean;
    if (dropped.length) log(`keymap: ignored ${dropped.length} stale entr${dropped.length === 1 ? "y" : "ies"}: ${dropped.join("; ")}`);
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
async function saveKeymap() { await store.set(KEYMAP_STORAGE_KEY, JSON.stringify(keymap)); }
async function bind(combo, id) {
  const prev = comboForCommand(id); if (prev) delete keymap[prev];
  keymap[combo] = id; await saveKeymap();
  log(`bound ${displayCombo(combo)} → ${COMMAND_BY_ID[id].label}`);
}
async function unbind(id) {
  const prev = comboForCommand(id); if (!prev) return;
  delete keymap[prev]; await saveKeymap();
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
// 6. Pane — icon toolbar, assign mode, recorder, key map, params, log
// ---------------------------------------------------------------------------

// Inline SVG pictograms. 24×24, stroke = currentColor. `t` rotates/flips a base icon.
const S = (inner, t) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"${t ? ` style="transform:${t}"` : ""}>${inner}</svg>`;
const ICON_BASE = {
  align: '<line x1="4" y1="3" x2="4" y2="21"/><rect x="6" y="6" width="11" height="4"/><rect x="6" y="14" width="7" height="4"/>',
  alignC: '<line x1="12" y1="3" x2="12" y2="21"/><rect x="6.5" y="6" width="11" height="4"/><rect x="8.5" y="14" width="7" height="4"/>',
  dist: '<rect x="3" y="7" width="4" height="10"/><rect x="10" y="7" width="4" height="10"/><rect x="17" y="7" width="4" height="10"/>',
  dock: '<line x1="20" y1="3" x2="20" y2="21"/><rect x="3" y="8" width="8" height="8"/><path d="M12 12h5m-2-2 2 2-2 2"/>',
  stackI: '<rect x="3" y="7" width="8" height="10"/><rect x="11" y="7" width="8" height="10"/><path d="M21 3v18" opacity=".35"/>',
  // EE-style: target on the left, reference on the right, an arrow from the target to the
  // reference's NEAR edge (gap) or through it to its FAR edge (stretch). Rotated for U/D/L.
  stretch: '<rect x="2" y="8" width="7" height="8"/><rect x="15" y="5" width="7" height="14" fill="currentColor" opacity=".25" stroke="none"/><rect x="15" y="5" width="7" height="14"/><path d="M9 12h12.5M19 9.5l2.5 2.5-2.5 2.5"/>',
  fillg: '<rect x="2" y="8" width="7" height="8"/><rect x="15" y="5" width="7" height="14" fill="currentColor" opacity=".25" stroke="none"/><rect x="15" y="5" width="7" height="14"/><path d="M9 12h5.5M12 9.5l2.5 2.5-2.5 2.5"/>',
  nudge: '<rect x="8" y="8" width="8" height="8"/><path d="M4 12h2m-1-2-1 2 1 2" /><path d="M3 12h3"/>',
};
const ICONS = {
  alignLeft: S(ICON_BASE.align), alignRight: S(ICON_BASE.align, "rotate(180deg)"), alignTop: S(ICON_BASE.align, "rotate(90deg)"), alignBottom: S(ICON_BASE.align, "rotate(-90deg)"),
  alignCenter: S(ICON_BASE.alignC), alignMiddle: S(ICON_BASE.alignC, "rotate(90deg)"),
  distributeH: S(ICON_BASE.dist), distributeV: S(ICON_BASE.dist, "rotate(90deg)"),
  dockRight: S(ICON_BASE.dock), dockLeft: S(ICON_BASE.dock, "rotate(180deg)"), dockDown: S(ICON_BASE.dock, "rotate(90deg)"), dockUp: S(ICON_BASE.dock, "rotate(-90deg)"),
  stackH: S(ICON_BASE.stackI), stackV: S(ICON_BASE.stackI, "rotate(90deg)"),
  swap: S('<path d="M4 8h13m-3-3 3 3-3 3M20 16H7m3-3-3 3 3 3"/>'),
  golden: S('<rect x="3" y="3" width="18" height="18"/><rect x="7" y="6" width="10" height="8"/>'),
  matrix: S('<circle cx="6" cy="6" r="1.6"/><circle cx="12" cy="6" r="1.6"/><circle cx="18" cy="6" r="1.6"/><circle cx="6" cy="12" r="1.6"/><circle cx="12" cy="12" r="1.6"/><circle cx="18" cy="12" r="1.6"/><circle cx="6" cy="18" r="1.6"/><circle cx="12" cy="18" r="1.6"/><circle cx="18" cy="18" r="1.6"/>'),
  table: S('<rect x="3" y="4" width="18" height="16"/><path d="M3 10h18M3 15h18M9 4v16M15 4v16"/><circle cx="12" cy="12.5" r="1.8" fill="currentColor" stroke="none"/>'),
  nudgeLeft: S(ICON_BASE.nudge), nudgeRight: S(ICON_BASE.nudge, "rotate(180deg)"), nudgeUp: S(ICON_BASE.nudge, "rotate(90deg)"), nudgeDown: S(ICON_BASE.nudge, "rotate(-90deg)"),
  matchWidth: S('<rect x="6" y="9" width="12" height="6"/><path d="M2 12h20M4 10l-2 2 2 2M20 10l2 2-2 2"/>'),
  matchHeight: S('<rect x="9" y="6" width="6" height="12"/><path d="M12 2v20M10 4l2-2 2 2M10 20l2 2 2-2"/>'),
  matchBoth: S('<rect x="7" y="7" width="10" height="10"/><path d="M3 3l4 4M21 21l-4-4M21 3l-4 4M3 21l4-4"/>'),
  fitInside: S('<rect x="3" y="3" width="18" height="18" stroke-dasharray="3 2"/><rect x="8" y="8" width="8" height="8" fill="currentColor" stroke="none" opacity=".8"/>'),
  fillOutside: S('<rect x="8" y="8" width="8" height="8" stroke-dasharray="3 2"/><rect x="3" y="3" width="18" height="18" fill="currentColor" stroke="none" opacity=".35"/>'),
  stretchRight: S(ICON_BASE.stretch), stretchLeft: S(ICON_BASE.stretch, "rotate(180deg)"), stretchDown: S(ICON_BASE.stretch, "rotate(90deg)"), stretchUp: S(ICON_BASE.stretch, "rotate(-90deg)"),
  fillRight: S(ICON_BASE.fillg), fillLeft: S(ICON_BASE.fillg, "rotate(180deg)"), fillDown: S(ICON_BASE.fillg, "rotate(90deg)"), fillUp: S(ICON_BASE.fillg, "rotate(-90deg)"),
  resizeUp: S('<rect x="4" y="4" width="16" height="16"/><path d="M12 8v8M8 12h8"/>'),
  resizeDown: S('<rect x="4" y="4" width="16" height="16"/><path d="M8 12h8"/>'),
  slice: S('<rect x="4" y="4" width="16" height="16"/><path d="M12 4v16M4 12h16"/>'),
  fill: S('<circle cx="12" cy="12" r="8" fill="var(--sw)" stroke="rgba(127,127,127,.6)"/>'),
  line: S('<circle cx="12" cy="12" r="7.5" stroke="var(--sw)" stroke-width="3.5"/>'),
  font: S('<text x="12" y="17.5" text-anchor="middle" font-size="17" font-weight="700" fill="var(--sw)" stroke="rgba(127,127,127,.5)" stroke-width=".6">A</text>'),
  margins: S('<rect x="3" y="3" width="18" height="18"/><rect x="7" y="7" width="10" height="10" stroke-dasharray="2 2"/>'),
  marginsZero: S('<rect x="3" y="3" width="18" height="18"/><path d="M6 8h12M6 12h12M6 16h8"/>'),
  fitText: S('<rect x="4" y="6" width="16" height="12"/><path d="M7 10h10M7 14h6"/><path d="M20 3l1 1-1 1M4 21l-1-1 1-1"/>'),
  wrap: S('<path d="M4 6h16M4 12h11a3 3 0 0 1 0 6h-3m0 0 2-2m-2 2 2 2M4 18h5"/>'),
  split: S('<rect x="3" y="5" width="8" height="14"/><rect x="13" y="5" width="8" height="14"/><path d="M12 2v20" stroke-dasharray="2 2"/>'),
  merge: S('<rect x="3" y="5" width="7" height="14"/><rect x="14" y="5" width="7" height="14"/><path d="M10 12h4"/>'),
  bullets: S('<circle cx="5" cy="7" r="1.5" fill="currentColor" stroke="none"/><circle cx="5" cy="12" r="1.5" fill="currentColor" stroke="none"/><circle cx="5" cy="17" r="1.5" fill="currentColor" stroke="none"/><path d="M9 7h11M9 12h11M9 17h8"/>'),
  fontSize: S('<text x="3" y="19" font-size="17" font-weight="700" fill="currentColor" stroke="none">A</text><text x="14" y="19" font-size="10" font-weight="700" fill="currentColor" stroke="none">A</text>'),
  pickup: S('<path d="M4 20l6-6M14 4l6 6-8 8-6-6z"/><path d="M12 10l2 2"/>'),
  apply: S('<path d="M4 20l6-6M14 4l6 6-8 8-6-6z"/><path d="M3 4h5M5 2v4"/>'),
  painter: S('<path d="M4 20l6-6M14 4l6 6-8 8-6-6z"/><path d="M2 12c2-3 4-3 6 0s4 3 6 0" opacity=".7"/>'),
  star: S('<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z"/>'),
  preset: S('<rect x="4" y="4" width="16" height="16" rx="3"/>'),
  similar: S('<rect x="3" y="3" width="8" height="8"/><rect x="13" y="13" width="8" height="8"/><path d="M16 3l1.3 2.7 2.7 1.3-2.7 1.3L16 11l-1.3-2.7L12 7l2.7-1.3z" fill="currentColor" stroke="none"/>'),
  decompose: S('<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>'),
  pictures: S('<rect x="3" y="5" width="8" height="6"/><rect x="13" y="5" width="8" height="6"/><rect x="3" y="13" width="8" height="6"/><rect x="13" y="13" width="8" height="6"/>'),
  export: S('<rect x="3" y="3" width="18" height="18"/><path d="M3 17l5-5 4 4 3-3 6 6"/><circle cx="16" cy="8" r="2"/>'),
  hide: S('<path d="M3 3l18 18M10 6.5A9.7 9.7 0 0 1 12 6c5 0 9 6 9 6a15 15 0 0 1-3.2 3.5M6.5 8A15 15 0 0 0 3 12s4 6 9 6a9 9 0 0 0 3-.5"/>'),
  unhide: S('<path d="M3 12s4-6 9-6 9 6 9 6-4 6-9 6-9-6-9-6z"/><circle cx="12" cy="12" r="2.5"/>'),
  newDeck: S('<rect x="3" y="5" width="12" height="9"/><path d="M15 9h6M18 6l3 3-3 3"/><rect x="3" y="17" width="12" height="3" opacity=".5"/>'),
  newDeckLite: S('<rect x="3" y="5" width="12" height="9"/><path d="M15 9h6M18 6l3 3-3 3"/><path d="M3 18.5h12" stroke-dasharray="2 2" opacity=".6"/>'),
  notice: S('<rect x="3" y="3" width="18" height="6" rx="3" fill="currentColor" opacity=".85" stroke="none"/><path d="M4 13h16M4 17h11" opacity=".5"/>'),
  duplicate: S('<rect x="3" y="7" width="13" height="10"/><path d="M8 7V4h13v10h-3"/><path d="M9.5 12h4M11.5 10v4"/>'),
  goto: S('<rect x="3" y="4" width="18" height="16"/><path d="M9 9l-1.5 6M15.5 9L14 15M7 11h10M6.5 13.5h10"/>'),
  agenda: S('<path d="M5 7h14M5 12h14M5 17h9"/><circle cx="3" cy="7" r=".8" fill="currentColor"/><circle cx="3" cy="12" r=".8" fill="currentColor"/><circle cx="3" cy="17" r=".8" fill="currentColor"/>'),
  master: S('<rect x="3" y="4" width="18" height="14"/><path d="M8 21h8M13 14h5" /><path d="M13 14h5" stroke-width="3" opacity=".4"/>'),
  masterOff: S('<rect x="3" y="4" width="18" height="14"/><path d="M8 21h8M14 12l4 4m0-4-4 4"/>'),
  sticky: S('<path d="M4 4h16v11l-5 5H4z" fill="var(--sw, #FFFF00)" stroke="rgba(0,0,0,.55)"/><path d="M15 20v-5h5" stroke="rgba(0,0,0,.55)"/>'),
  pane: S('<rect x="3" y="4" width="18" height="16"/><path d="M15 4v16"/>'),
};

function el(id) { return document.getElementById(id); }

// ---- recorder (assign mode) ----
const recorder = {
  active: false, commandId: null, pendingConfirm: null, lastCapture: 0, _lastCombo: null,
  start(commandId) {
    this.active = true; this.commandId = commandId; this.pendingConfirm = null;
    renderToolbar(); renderKeyGrid();
    setStatus(`Press the key for “${COMMAND_BY_ID[commandId].label}” — or click a key in the map. Esc cancels, Delete removes.`, "info");
  },
  stop() { this.active = false; this.commandId = null; this.pendingConfirm = null; renderToolbar(); renderKeyGrid(); },
  async capture(combo, source) {
    if (!this.active) return;
    const now = performance.now();
    if (now - this.lastCapture < 400 && combo === this._lastCombo) return;
    this.lastCapture = now; this._lastCombo = combo;
    const id = this.commandId, cmd = COMMAND_BY_ID[id];
    if (combo === "Esc") { this.stop(); setStatus("Cancelled.", "muted"); return; }
    if (combo === "Backspace" || combo === "Delete") { await unbind(id); this.stop(); setStatus(`Removed the key from “${cmd.label}”.`, "ok"); return; }
    if (!BANK_SET.has(combo)) {
      const native = NATIVE_BY_COMBO[combo];
      const why = native ? ` It's already ${native.map((n) => `${n.source}: ${n.what}`).join(", ")}.` : "";
      setStatus(`${displayCombo(combo)} can't be bound — only ${KEY_BANK_MODIFIER_SETS.map((m) => displayCombo(m + "+key")).join(" and ")} are registered.${why} Try another key.`, "warn");
      return;
    }
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
  if (e.target && (e.target.tagName === "TEXTAREA" || e.target.tagName === "INPUT")) return;
  const combo = comboFromEvent(e);
  if (!combo) return;
  e.preventDefault(); e.stopPropagation();
  recorder.capture(combo, "dom");
}

// ---- rendering ----
let assignMode = false;
let compact = false;
try { compact = localStorage.getItem("ce-compact") === "1"; } catch (_) { /* ignore */ }
function setCompact(on) {
  compact = on;
  document.body.classList.toggle("compact", on);
  const b = el("btn-compact"); if (b) b.classList.toggle("on", on);
  try { localStorage.setItem("ce-compact", on ? "1" : "0"); } catch (_) { /* ignore */ }
}

function setStatus(text, kind) { const s = el("recorder-status"); if (!s) return; s.textContent = text; s.className = "status " + (kind || ""); }
const esc = (v) => String(v).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function renderToolbar() {
  const root = el("toolbar");
  if (!root) return;
  const palette = settings.palette || DEFAULT_PALETTE;
  root.innerHTML = GROUPS.map((g) => {
    const cmds = COMMANDS.filter((c) => c.group === g);
    const tools = cmds.map((c) => {
      const combo = comboForCommand(c.id);
      const active = recorder.active && recorder.commandId === c.id;
      const sw = c.color || (c.kind ? palette[c.slot - 1] || "transparent" : null);
      const style = sw ? ` style="--sw:${sw}"` : "";
      return `<button class="tool${active ? " active" : ""}${combo ? " bound" : ""}" data-id="${c.id}" title="${esc(c.label)}${combo ? "  " + displayCombo(combo) : ""} — ${esc(c.desc)}"${style}>
        <span class="ic">${ICONS[c.icon] || ICONS.preset}${c.badge ? `<b class="badge">${c.badge}</b>` : ""}</span>
        <span class="lbl">${esc(c.label)}</span>
        <kbd class="${combo ? "" : "empty"}">${combo ? displayCombo(combo) : "·"}</kbd>
      </button>`;
    }).join("");
    return `<details class="grp" data-group="${g}" open><summary>${g}<span class="count">${cmds.length}</span></summary><div class="tools">${tools}</div>${PARAMS[g] ? `<div class="params">${PARAMS[g]()}</div>` : ""}</details>`;
  }).join("");
  // preserve open/closed state
  for (const d of root.querySelectorAll("details.grp")) { const k = "grp-open-" + d.dataset.group; try { if (localStorage.getItem(k) === "0") d.open = false; } catch (_) { /* ignore */ } d.addEventListener("toggle", () => { try { localStorage.setItem(k, d.open ? "1" : "0"); } catch (_) { /* ignore */ } }); }
  wireParams();
}

// Per-group parameter panels (settings that commands read at run time).
const num = (id, label, value, step = 1) => `<label>${label} <input type="number" id="${id}" value="${value}" step="${step}"></label>`;
const PARAMS = {
  Position: () => `${num("p-nudge", "Nudge pt", settings.nudge)} ${num("p-mrows", "Matrix rows", settings.matrix.rows)} ${num("p-mcols", "cols", settings.matrix.cols)} ${num("p-mgapx", "gap x", settings.matrix.gapX)} ${num("p-mgapy", "gap y", settings.matrix.gapY)}`,
  Size: () => `${num("p-factor", "Resize ×", settings.resizeFactor, 0.05)} <label><input type="checkbox" id="p-rfont" ${settings.resizeFont ? "checked" : ""}> font</label> <label><input type="checkbox" id="p-rline" ${settings.resizeLine ? "checked" : ""}> line</label>
    <span class="sep"></span> ${num("p-srows", "Slice rows", settings.slice.rows)} ${num("p-scols", "cols", settings.slice.cols)} ${num("p-sgapx", "gap x", settings.slice.gapX)} ${num("p-sgapy", "gap y", settings.slice.gapY)}`,
  Colour: () => `<label>Palette (hex, comma-separated) <input type="text" id="p-palette" class="wide" value="${esc((settings.palette || DEFAULT_PALETTE).join(", "))}"></label>`,
  Text: () => `${num("p-ml", "Margins L", settings.margins.left, 0.1)} ${num("p-mr", "R", settings.margins.right, 0.1)} ${num("p-mt", "T", settings.margins.top, 0.1)} ${num("p-mb", "B", settings.margins.bottom, 0.1)} <span class="sep"></span> ${num("p-fsize", "Font size", settings.fontSize, 0.5)}`,
  Format: () => `<div id="formats"></div>`,
  Tools: () => `${num("p-goto", "Slide #", settings.gotoSlide)} <span class="sep"></span>
    <label>Master label <input type="text" id="p-mlabel" value="${esc(settings.masterLabel)}"></label> ${num("p-mleft", "L", settings.masterPos.left)} ${num("p-mtop", "T", settings.masterPos.top)} ${num("p-mw", "W", settings.masterPos.width)} ${num("p-mh", "H", settings.masterPos.height)}
    <label class="block">Agenda items (one per line)<textarea id="p-agenda" rows="4">${esc(settings.agenda || "")}</textarea></label>
    <div id="dims" class="dims"><b>Dimensions</b> <span id="dims-name" class="muted"></span> ${num("dim-left", "L", "", 0.1)} ${num("dim-top", "T", "", 0.1)} ${num("dim-width", "W", "", 0.1)} ${num("dim-height", "H", "", 0.1)} <button id="btn-dims-apply">Apply</button></div>
    <div id="export-output"></div>`,
  Sticky: () => `<label>Initials <input type="text" id="sticky-initials" size="4" maxlength="6" value="${esc(settings.initials)}"></label> <span class="muted">default colour:</span> <span id="sticky-colors" class="swatches"></span>`,
};

function wireParams() {
  const bindNum = (id, fn) => { const e = el(id); if (e) e.addEventListener("change", async () => { fn(parseFloat(e.value)); await saveSettings(); }); };
  bindNum("p-nudge", (v) => { settings.nudge = v; });
  bindNum("p-mrows", (v) => { settings.matrix.rows = v; }); bindNum("p-mcols", (v) => { settings.matrix.cols = v; }); bindNum("p-mgapx", (v) => { settings.matrix.gapX = v; }); bindNum("p-mgapy", (v) => { settings.matrix.gapY = v; });
  bindNum("p-factor", (v) => { settings.resizeFactor = v; });
  bindNum("p-srows", (v) => { settings.slice.rows = v; }); bindNum("p-scols", (v) => { settings.slice.cols = v; }); bindNum("p-sgapx", (v) => { settings.slice.gapX = v; }); bindNum("p-sgapy", (v) => { settings.slice.gapY = v; });
  bindNum("p-ml", (v) => { settings.margins.left = v; }); bindNum("p-mr", (v) => { settings.margins.right = v; }); bindNum("p-mt", (v) => { settings.margins.top = v; }); bindNum("p-mb", (v) => { settings.margins.bottom = v; });
  bindNum("p-fsize", (v) => { settings.fontSize = v; });
  bindNum("p-goto", (v) => { settings.gotoSlide = v; });
  bindNum("p-mleft", (v) => { settings.masterPos.left = v; }); bindNum("p-mtop", (v) => { settings.masterPos.top = v; }); bindNum("p-mw", (v) => { settings.masterPos.width = v; }); bindNum("p-mh", (v) => { settings.masterPos.height = v; });
  for (const [id, key] of [["p-rfont", "resizeFont"], ["p-rline", "resizeLine"]]) { const e = el(id); if (e) e.addEventListener("change", async () => { settings[key] = e.checked; await saveSettings(); }); }
  const pal = el("p-palette");
  if (pal) pal.addEventListener("change", async () => {
    settings.palette = pal.value.split(",").map((s) => s.trim()).filter((s) => /^#[0-9a-f]{6}$/i.test(s)).map((s) => s.toUpperCase());
    await saveSettings(); renderToolbar();
  });
  const ml = el("p-mlabel"); if (ml) ml.addEventListener("change", async () => { settings.masterLabel = ml.value; await saveSettings(); });
  const ag = el("p-agenda"); if (ag) ag.addEventListener("change", async () => { settings.agenda = ag.value; await saveSettings(); });
  const ini = el("sticky-initials");
  if (ini) ini.addEventListener("change", async () => { settings.initials = ini.value.trim().toUpperCase() || DEFAULT_SETTINGS.initials; ini.value = settings.initials; await saveSettings(); });
  const sc = el("sticky-colors");
  if (sc) { renderStickySwatches(); sc.addEventListener("click", async (e) => { const sw = e.target.closest("[data-color]"); if (!sw) return; settings.stickyColor = sw.dataset.color; await saveSettings(); renderStickySwatches(); }); }
  const da = el("btn-dims-apply"); if (da) da.addEventListener("click", () => applyDimensions().catch((err) => log("dimensions FAILED: " + err.message)));
  renderFormats();
}

function renderStickySwatches() {
  const root = el("sticky-colors"); if (!root) return;
  root.innerHTML = STICKY.COLORS.map((c) => `<button class="swatch${c.name === settings.stickyColor ? " selected" : ""}" data-color="${c.name}" style="background:${c.hex}" title="${c.name}"></button>`).join("");
}

function renderFormats() {
  const root = el("formats"); if (!root) return;
  root.innerHTML = `<div class="muted">Clipboard: ${formatClipboard ? "picked up ✓" : "empty"} · Painter: ${painterOn ? "<b>ON</b>" : "off"}</div>
    <div class="row"><input type="text" id="myformat-name" placeholder="name for Save as My Format"></div>
    <ol class="myformats">${(settings.myFormats || []).map((f, i) => `<li><b>${i + 1}</b> ${esc(f.name)} <button data-apply="${i}">apply</button> <button data-del="${i}" class="x">×</button></li>`).join("") || "<li class='muted'>no saved formats yet</li>"}</ol>`;
  root.onclick = async (e) => {
    const a = e.target.closest("[data-apply]"), d = e.target.closest("[data-del]");
    if (a) { formatApply(settings.myFormats[+a.dataset.apply].props).catch((err) => log("apply FAILED: " + err.message)); }
    if (d) { settings.myFormats.splice(+d.dataset.del, 1); await saveSettings(); renderFormats(); }
  };
}

// Physical layout for the key map: number row on top, then the three QWERTY rows.
const KEY_ROWS = [["1","2","3","4","5","6","7","8","9","0"], ["Q","W","E","R","T","Y","U","I","O","P"], ["A","S","D","F","G","H","J","K","L"], ["Z","X","C","V","B","N","M"]];
function describeCombo(combo) {
  const id = keymap[combo];
  const native = NATIVE_BY_COMBO[combo] || [];
  if (id) return { kind: "assigned", label: COMMAND_BY_ID[id].label, detail: native.map((n) => `also ${n.source}: ${n.what}`) };
  if (native.length) return { kind: "conflict", label: native[0].source, detail: native.map((n) => `${n.source}: ${n.what}`) };
  return { kind: "free", label: "", detail: ["available"] };
}
function renderKeyGrid() {
  const root = el("keymap-grid"); if (!root) return;
  const inRows = new Set(KEY_ROWS.flat());
  const extra = KEY_BANK_KEYS.filter((k) => !inRows.has(k));
  const rows = extra.length ? [...KEY_ROWS, extra] : KEY_ROWS;
  root.innerHTML = KEY_BANK_MODIFIER_SETS.map((mods) => {
    const html = rows.map((row, r) => `<div class="krow r${r}">` + row.filter((k) => KEY_BANK_KEYS.includes(k)).map((k) => {
      const combo = `${mods}+${k}`, d = describeCombo(combo);
      const active = recorder.active && keymap[combo] && recorder.commandId === keymap[combo];
      return `<button class="cell ${d.kind}${active ? " active" : ""}" data-combo="${combo}" aria-label="${displayCombo(combo)}: ${d.detail.join("; ")}"><b>${k}</b><i>${esc(d.label)}</i></button>`;
    }).join("") + "</div>").join("");
    return `<div class="mods">${displayCombo(mods + "+")}</div><div class="keyboard">${html}</div>`;
  }).join("");
}
function showTip(cell) {
  const tip = el("keytip"); if (!tip || !cell) return;
  const combo = cell.dataset.combo, d = describeCombo(combo);
  tip.innerHTML = `<b>${displayCombo(combo)}</b> ${d.kind === "assigned" ? "→ " + esc(d.label) : ""}<br>` + (d.kind === "free" ? "<span class='muted'>available</span>" : d.detail.map((t) => `<span>${esc(t)}</span>`).join("<br>"));
  tip.hidden = false;
  const r = cell.getBoundingClientRect(), w = tip.offsetWidth;
  tip.style.left = Math.max(8, Math.min(window.innerWidth - w - 8, r.left + r.width / 2 - w / 2)) + window.scrollX + "px";
  tip.style.top = r.bottom + 6 + window.scrollY + "px";
}
function hideTip() { const tip = el("keytip"); if (tip) tip.hidden = true; }

function renderExport() { const ta = el("keymap-json"); if (ta && document.activeElement !== ta) ta.value = JSON.stringify(keymap, null, 2); }
function renderPane() { renderToolbar(); renderKeyGrid(); renderExport(); }

function renderDiagnostics() {
  const d = el("diagnostics"); if (!d) return;
  const info = Office.context.diagnostics || {};
  const sets = ["PowerPointApi 1.4", "PowerPointApi 1.5", "PowerPointApi 1.8", "PowerPointApi 1.10", "SharedRuntime 1.1", "KeyboardShortcuts 1.1"]
    .map((s) => { const [name, ver] = s.split(" "); return `${s}: ${supports(name, ver) ? "yes" : "no"}`; });
  d.innerHTML = `<div><b>Host</b> ${info.host || "?"} · <b>Platform</b> ${info.platform || "?"} · <b>Version</b> ${info.version || "?"} · <b>Build</b> ${BUILD} · <b>Bank</b> ${KEY_BANK.length} keys · <b>Commands</b> ${COMMANDS.length}</div><div class="muted">${sets.join(" · ")}</div>`;
}

function setMode(assign) {
  assignMode = assign;
  document.body.classList.toggle("assign", assign);
  const b1 = el("mode-run"), b2 = el("mode-assign");
  if (b1) b1.classList.toggle("on", !assign);
  if (b2) b2.classList.toggle("on", assign);
  if (!assign && recorder.active) recorder.stop();
  const km = el("keymap-section"); if (km) km.open = assign;
  setStatus(assign ? "Assign mode: click any icon, then press the key you want for it." : "", assign ? "info" : "muted");
}

function wireTaskPane() {
  const on = (id, evt, fn) => { const e = el(id); if (e) e.addEventListener(evt, fn); };

  on("toolbar", "click", async (e) => {
    const t = e.target.closest(".tool");
    if (!t) return;
    if (assignMode) recorder.start(t.dataset.id);
    else await runCommand(t.dataset.id);
  });
  on("toolbar", "contextmenu", (e) => { const t = e.target.closest(".tool"); if (!t) return; e.preventDefault(); setMode(true); recorder.start(t.dataset.id); });
  on("mode-run", "click", () => setMode(false));
  on("mode-assign", "click", () => setMode(true));
  on("btn-compact", "click", () => setCompact(!compact));
  setCompact(compact);

  on("keymap-grid", "click", (e) => {
    const cell = e.target.closest("[data-combo]"); if (!cell) return;
    const combo = cell.dataset.combo;
    if (recorder.active) { recorder.capture(combo, "grid"); return; }
    if (keymap[combo]) { setMode(true); recorder.start(keymap[combo]); return; }
    const d = describeCombo(combo);
    setStatus(`${displayCombo(combo)} — ${d.detail.join("; ")}. Switch to Assign, click an icon, then click a key.`, "muted");
  });
  on("keymap-grid", "mouseover", (e) => showTip(e.target.closest("[data-combo]")));
  on("keymap-grid", "mouseout", (e) => { if (!e.relatedTarget || !e.relatedTarget.closest || !e.relatedTarget.closest("[data-combo]")) hideTip(); });
  on("keymap-grid", "focusin", (e) => showTip(e.target.closest("[data-combo]")));
  on("keymap-grid", "focusout", hideTip);
  document.addEventListener("keydown", onKeyDown, true);

  on("btn-apply-json", "click", async () => {
    try {
      const { clean, dropped } = sanitizeKeymap(JSON.parse(el("keymap-json").value));
      keymap = clean; await saveKeymap(); renderPane();
      setStatus(dropped.length ? `Applied; ignored: ${dropped.join("; ")}` : "Applied.", dropped.length ? "warn" : "ok");
    } catch (err) { setStatus("Not valid JSON: " + err.message, "warn"); }
  });
  on("btn-copy-json", "click", async () => {
    try { await navigator.clipboard.writeText(JSON.stringify(keymap, null, 2)); setStatus("Copied.", "ok"); }
    catch (_) { el("keymap-json").select(); setStatus("Select-all + copy from the box.", "muted"); }
  });
  on("btn-reset-json", "click", async () => { keymap = { ...DEFAULT_KEYMAP }; await saveKeymap(); renderPane(); setStatus("Reset to defaults.", "ok"); });

  on("btn-dump", "click", () => dumpSelection().catch((e) => log("dumpSelection FAILED: " + e.message)));
  on("btn-probe-shapes", "click", () => createProbeShapes().catch((e) => log("createProbeShapes FAILED: " + e.message)));
  on("btn-acceptance-shapes", "click", () => createAcceptanceShapes().catch((e) => log("createAcceptanceShapes FAILED: " + e.message)));
  on("btn-shape-api", "click", () => dumpShapeApi().catch((e) => log("dumpShapeApi FAILED: " + e.message)));
  on("btn-addin-refs", "click", () => dumpAddinRefs().catch((e) => log("dumpAddinRefs FAILED: " + e.message)));
  on("btn-clear-log", "click", () => { logBuffer.length = 0; log("log cleared"); });

  for (const key of ["RECENTER_ON_REF", "KEEP_CENTER", "SWAP_ZORDER", "SWAP_SIZE"]) {
    const box = el("cfg-" + key);
    if (box) { box.checked = CONFIG[key]; box.addEventListener("change", () => { CONFIG[key] = box.checked; log(`${key} = ${box.checked}`); }); }
  }

  renderDiagnostics();
  renderPane();
  setMode(false);
  const l = el("log"); if (l) l.textContent = logBuffer.join("\n");
}

// ---- selection-changed: dimensions panel + format painter ----
async function onSelectionChanged() {
  if (painterOn && formatClipboard) {
    try { await formatApply(); } catch (err) { log("painter: " + err.message); }
  }
  refreshDimensions();
}

// ---- probes (Phase 0 / Phase 1 / API) ----
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
async function createProbeShapes() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const specs = [{ name: "A", left: 60, top: 100, width: 120, height: 80 }, { name: "B", left: 240, top: 100, width: 120, height: 80 }, { name: "C", left: 420, top: 100, width: 120, height: 80 }];
    const made = specs.map((spec) => { const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, spec); shape.name = spec.name; shape.textFrame.textRange.text = spec.name; return shape; });
    await context.sync();
    if (supports("PowerPointApi", "1.8")) { made[0].setZOrder(PowerPoint.ShapeZOrder.bringToFront); await context.sync(); }
    log("Created probe shapes A, B, C (A brought to front).");
  });
}
async function createAcceptanceShapes() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const target = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: 60, top: 80, width: 100, height: 200 });
    target.name = "Target 100x200"; target.textFrame.textRange.text = "target";
    const ref = slide.shapes.addGeometricShape("RoundRectangle" /* enum key is roundRectangle; the literal avoids an undefined type → line */, { left: 300, top: 80, width: 300, height: 150 });
    ref.name = "Reference 300x150"; ref.textFrame.textRange.text = "reference";
    await context.sync();
    log("Created acceptance shapes. Expect fitInside → 75×150, fillOutside → 300×600.");
  });
}
/** Read this deck's package and log every part that references add-ins (webextensions, custom XML). */
async function dumpAddinRefs() {
  const bytes = await getDeckBytes();
  if (typeof JSZip === "undefined") throw new Error("JSZip not loaded");
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => /webextension|customXml|_rels\/presentation\.xml\.rels|presProps|\[Content_Types\]/i.test(n)).sort();
  log(`add-in refs: ${Math.round(bytes.length / 1024)} KB package, ${Object.keys(zip.files).length} parts; candidates: ${names.length}`);
  const rows = [];
  for (const n of names) {
    const text = await zip.file(n).async("string");
    const hit = /4e27ba64|webextension|we:reference/i.test(text);
    rows.push({ part: n, size: text.length, mentionsAddin: hit ? "yes" : "" });
    if (hit) log(`  ${n}:\n${text.replace(/></g, ">\n<").slice(0, 1500)}`);
  }
  renderTable("probe-output", rows);
}

async function dumpShapeApi() {
  const CANDIDATES = ["shadow", "shadowFormat", "effects", "effectFormat", "glow", "glowFormat", "reflection", "reflectionFormat", "softEdge", "softEdges", "softEdgeFormat", "threeDFormat", "style"];
  await PowerPoint.run(async (context) => {
    const sel = context.presentation.getSelectedShapes();
    sel.load("items/id,items/name");
    await context.sync();
    if (!sel.items.length) throw new Error("select a shape first");
    const shape = sel.items[0];
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

// ---- log / formatting ----
const logBuffer = [];
function log(message) {
  const line = `${new Date().toLocaleTimeString()}  ${message}`;
  logBuffer.push(line); if (logBuffer.length > 200) logBuffer.shift();
  console.log("[ExpropriatedElements] " + message);
  const l = el("log"); if (l) { l.textContent = logBuffer.join("\n"); l.scrollTop = l.scrollHeight; }
}
function fmt(g) { const n = (v) => (Math.round(v * 100) / 100).toString(); return `${n(g.width)}×${n(g.height)} @ (${n(g.left)}, ${n(g.top)})`; }
function renderTable(elementId, rows) {
  const root = el(elementId); if (!root) return;
  if (!rows.length) { root.innerHTML = "<p class='muted'>Nothing selected.</p>"; return; }
  const cols = Object.keys(rows[0]);
  root.innerHTML = "<table><thead><tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr></thead><tbody>" +
    rows.map((r) => "<tr>" + cols.map((c) => `<td>${esc(r[c])}</td>`).join("") + "</tr>").join("") + "</tbody></table>";
}

// ---------------------------------------------------------------------------
// 7. Wiring
// ---------------------------------------------------------------------------
Office.onReady(async (info) => {
  if (Office.actions && typeof Office.actions.associate === "function") {
    for (const combo of KEY_BANK) Office.actions.associate(slotId(combo), () => onSlot(combo));
    // Ribbon buttons (manifest FunctionName = "ribbon_<command id>"). In Assign mode a ribbon
    // click records a key for that command instead of running it.
    for (const c of COMMANDS) {
      Office.actions.associate("ribbon_" + c.id, async (event) => {
        try {
          if (assignMode && paneLooksVisible()) recorder.start(c.id);
          else await runCommand(c.id);
        } finally { if (event && event.completed) event.completed(); }
      });
    }
  } else {
    log("Office.actions.associate unavailable — not running inside an Office shared runtime.");
  }
  if (Office.addin && Office.addin.onVisibilityModeChanged) {
    try { await Office.addin.onVisibilityModeChanged((args) => { paneVisible = args.visibilityMode === "Taskpane"; log(`pane visibility → ${args.visibilityMode}`); }); }
    catch (err) { log("onVisibilityModeChanged unavailable: " + (err.message || err)); }
  }
  // Ask Office to start this runtime silently the next time the document opens, so the
  // shortcuts work without first opening the pane. Stored per document by Office.
  // The document stores the add-in's exact id for this. Reset to "none" first so a stale
  // reference (from an earlier manifest version) is replaced rather than left alongside.
  try {
    if (Office.addin && Office.addin.setStartupBehavior) {
      let before = "?";
      try { before = Office.addin.getStartupBehavior ? await Office.addin.getStartupBehavior() : "n/a"; } catch (_) { /* ignore */ }
      const SB = Office.StartupBehavior || { none: "none", load: "load" };
      await Office.addin.setStartupBehavior(SB.none);
      await Office.addin.setStartupBehavior(SB.load);
      log(`startup behavior: was ${before} → reset → load (save the deck so the cloud copy gets the clean reference)`);
    }
  } catch (err) { log("setStartupBehavior unavailable: " + (err.message || err)); }
  try {
    if (Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged, (r) => { if (r.status !== Office.AsyncResultStatus.Succeeded) log("selection-changed handler: " + (r.error && r.error.message)); });
    }
  } catch (err) { log("selection-changed handler unavailable: " + (err.message || err)); }

  await loadKeymap();
  await loadSettings();
  log(`ready: build ${BUILD} · host=${info.host} platform=${info.platform} version=${(Office.context.diagnostics || {}).version} · ${COMMANDS.length} commands · ${KEY_BANK.length} slots · ${Object.keys(keymap).length} bound`);

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wireTaskPane);
  else wireTaskPane();
});
