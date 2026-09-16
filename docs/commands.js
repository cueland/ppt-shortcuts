/* global Office, PowerPoint */
/*
 * ppt-shortcuts — keyboard-bound shape commands for PowerPoint on macOS.
 *
 * Runs in the add-in's shared runtime (lifetime="long"), so the same JS context serves
 * both the keyboard-shortcut actions and the task pane. That's why command handlers can
 * write straight into the task pane's #log element when the pane happens to be open.
 *
 * Layout of this file:
 *   1. CONFIG            — behaviour switches from the build brief (§5)
 *   2. geometry          — pure functions, no Office API (unit-testable in any JS runtime)
 *   3. resolveReference  — THE swappable piece; Phase 0 decides which branch is real
 *   4. commands          — PowerPoint.run wrappers, one per action ID
 *   5. probe / task pane — Phase 0 selection-order dump and debugging helpers
 *   6. Office.onReady    — Office.actions.associate wiring
 */
"use strict";

// ---------------------------------------------------------------------------
// 1. CONFIG
// ---------------------------------------------------------------------------
const CONFIG = {
  // How the reference ("master") shape is chosen. Phase 0 decides this:
  //   "lastSelected" — reference = last shape returned by getSelectedShapes().
  //                    Only valid if the API returns shapes in SELECTION order.
  //   "pickup"       — reference = shape stored by the pickupReference action.
  //                    Required if the API returns shapes in Z-order.
  REFERENCE_MODE: "lastSelected",

  // fitInside / fillOutside: move the result onto the reference's centre (true) or
  // scale in place around the target's own centre (false).
  RECENTER_ON_REF: true,

  // matchWidth / matchHeight / matchBoth: grow/shrink around the target's own centre
  // (true) instead of PowerPoint's native top-left anchoring (false).
  KEEP_CENTER: true,
};

// Runtime override of REFERENCE_MODE so the task pane can flip it without a redeploy.
let referenceMode = CONFIG.REFERENCE_MODE;

// Reference picked up by the pickupReference action. Lives as long as the shared runtime.
// { slideId, shapeId, name, left, top, width, height }
let pickedReference = null;

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
// 3. Reference resolution — the only place that knows about REFERENCE_MODE.
// ---------------------------------------------------------------------------

/** Plain-object snapshot of a loaded shape proxy. */
function snapshot(shape) {
  return {
    id: shape.id,
    name: shape.name,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
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
 * Returns { reference: snapshot, targets: [{...snapshot, proxy}] }.
 * Throws with a human-readable message if no reference can be determined.
 */
async function resolveReference(context, selected) {
  if (referenceMode === "lastSelected") {
    if (selected.length < 2) {
      throw new Error("Select at least two shapes: targets first, reference last.");
    }
    const reference = selected[selected.length - 1];
    return { reference, targets: selected.slice(0, -1) };
  }

  if (referenceMode === "pickup") {
    if (!pickedReference) {
      throw new Error("No reference picked up yet. Select a shape and press the pickupReference key.");
    }
    // Prefer the live shape (it may have been moved/resized since pickup).
    let reference = selected.find((s) => s.id === pickedReference.shapeId);
    if (!reference) {
      reference = await lookupShape(context, pickedReference.slideId, pickedReference.shapeId);
    }
    if (!reference) {
      log(`Reference "${pickedReference.name}" no longer found on its slide; using stored geometry.`);
      reference = { ...pickedReference, id: pickedReference.shapeId };
    }
    const targets = selected.filter((s) => s.id !== reference.id);
    if (targets.length === 0) {
      throw new Error("Selection contains only the reference shape; select the targets.");
    }
    return { reference, targets };
  }

  throw new Error(`Unknown REFERENCE_MODE "${referenceMode}".`);
}

/** Find a shape by slide id + shape id. Returns a snapshot or null. */
async function lookupShape(context, slideId, shapeId) {
  const slide = context.presentation.slides.getItemOrNullObject(slideId);
  const shape = slide.shapes.getItemOrNullObject(shapeId);
  shape.load("id,name,left,top,width,height,isNullObject");
  await context.sync();
  return shape.isNullObject ? null : snapshot(shape);
}

// ---------------------------------------------------------------------------
// 4. Commands
// ---------------------------------------------------------------------------

/**
 * Shared driver: load selection, resolve reference, compute new geometry for each target
 * with `compute(target, reference)`, write it back, then re-read and report.
 */
async function applyToTargets(actionId, compute) {
  return PowerPoint.run(async (context) => {
    const selected = await loadSelection(context);
    const { reference, targets } = await resolveReference(context, selected);

    log(`${actionId}: reference "${reference.name}" ${fmt(reference)} → ${targets.length} target(s)`);

    for (const target of targets) {
      const next = compute(target, reference);
      // Set all four explicitly and in this order. If PowerPoint honours "lock aspect
      // ratio" on width/height writes, the post-sync verification below will show it.
      target.proxy.left = next.left;
      target.proxy.top = next.top;
      target.proxy.width = next.width;
      target.proxy.height = next.height;
      target.expected = next;
    }
    await context.sync();

    // Verify: re-read and compare, so Phase 1 acceptance numbers are visible in the log.
    for (const target of targets) {
      target.proxy.load("left,top,width,height");
    }
    await context.sync();
    for (const target of targets) {
      const got = snapshot(target.proxy);
      const ok = approxEqual(got, target.expected);
      log(`  "${target.name}" ${fmt(target)} → ${fmt(got)}${ok ? "" : "  ⚠ expected " + fmt(target.expected)}`);
    }
  });
}

const commands = {
  matchWidth: () => applyToTargets("matchWidth", (t, r) => geometry.match(t, r, { width: true })),
  matchHeight: () => applyToTargets("matchHeight", (t, r) => geometry.match(t, r, { height: true })),
  matchBoth: () => applyToTargets("matchBoth", (t, r) => geometry.match(t, r, { width: true, height: true })),
  fitInside: () => applyToTargets("fitInside", (t, r) => geometry.scale(t, r, "contain")),
  fillOutside: () => applyToTargets("fillOutside", (t, r) => geometry.scale(t, r, "cover")),

  /** Store the (first) selected shape as the reference for "pickup" mode. */
  pickupReference: () =>
    PowerPoint.run(async (context) => {
      const selected = await loadSelection(context);
      if (selected.length === 0) throw new Error("Select a shape to pick up as reference.");
      const slides = context.presentation.getSelectedSlides();
      const slide = slides.getItemAt(0);
      slide.load("id");
      await context.sync();
      const s = selected[0];
      pickedReference = { slideId: slide.id, shapeId: s.id, name: s.name, left: s.left, top: s.top, width: s.width, height: s.height };
      log(`pickupReference: "${s.name}" ${fmt(s)} on slide ${slide.id}`);
      renderReferenceStatus();
    }),
};

/** Wrap a command so errors are logged instead of vanishing inside the shortcut runtime. */
function guarded(actionId) {
  return async () => {
    const t0 = performance.now();
    try {
      await commands[actionId]();
      log(`${actionId} done in ${Math.round(performance.now() - t0)} ms`);
    } catch (err) {
      log(`${actionId} FAILED: ${err && err.message ? err.message : err}`);
      if (err && err.debugInfo) log("  debugInfo: " + JSON.stringify(err.debugInfo));
    }
  };
}

// ---------------------------------------------------------------------------
// 5. Phase 0 probe and task-pane helpers
// ---------------------------------------------------------------------------

const supports = (set, ver) =>
  !!(Office.context && Office.context.requirements && Office.context.requirements.isSetSupported(set, ver));

/**
 * Phase 0: dump the selection as returned by getSelectedShapes(), with zOrderPosition
 * when PowerPointApi 1.8 is available. Compare the array order against your click order.
 */
async function dumpSelection() {
  const hasZ = supports("PowerPointApi", "1.8");
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/name" + (hasZ ? ",items/zOrderPosition" : ""));
    await context.sync();
    const rows = shapes.items.map((s, i) => ({
      index: i,
      name: s.name,
      id: s.id,
      zOrderPosition: hasZ ? s.zOrderPosition : "(needs PowerPointApi 1.8)",
    }));
    renderTable("probe-output", rows);
    const order = rows.map((r) => r.name).join(" → ");
    const zOrder = hasZ ? [...rows].sort((a, b) => a.zOrderPosition - b.zOrderPosition).map((r) => r.name).join(" → ") : "n/a";
    log(`dumpSelection: array order [${order}]; z-order (back→front) [${zOrder}]`);
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
      made[0].setZOrder(PowerPoint.ShapeZOrder.bringToFront); // A on top, so z ≠ creation order
      await context.sync();
    }
    log("Created probe shapes A, B, C (A brought to front). Now select C → A → B and press Dump.");
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

async function showRegisteredShortcuts() {
  if (!supports("KeyboardShortcuts", "1.1")) {
    log("KeyboardShortcuts 1.1 not reported as supported; getShortcuts() unavailable.");
    return;
  }
  try {
    const map = await Office.actions.getShortcuts();
    const rows = Object.keys(map).map((action) => ({ action, key: map[action] === null ? "(overridden / conflict)" : map[action] }));
    renderTable("shortcuts-output", rows);
    const inUse = await Office.actions.areShortcutsInUse(Object.values(map).filter(Boolean));
    const conflicts = inUse.filter((s) => s.inUse).map((s) => s.shortcut);
    log(conflicts.length ? `Shortcuts already in use elsewhere: ${conflicts.join(", ")}` : "No shortcut conflicts reported.");
  } catch (err) {
    log("getShortcuts failed: " + (err.message || err));
  }
}

// ---- Task pane rendering (no-ops when the pane isn't open) ----

const logBuffer = [];
function log(message) {
  const line = `${new Date().toLocaleTimeString()}  ${message}`;
  logBuffer.push(line);
  if (logBuffer.length > 200) logBuffer.shift();
  console.log("[ppt-shortcuts] " + message);
  const el = document.getElementById("log");
  if (el) {
    el.textContent = logBuffer.join("\n");
    el.scrollTop = el.scrollHeight;
  }
}

function fmt(g) {
  const n = (v) => (Math.round(v * 100) / 100).toString();
  return `${n(g.width)}×${n(g.height)} @ (${n(g.left)}, ${n(g.top)})`;
}

function approxEqual(a, b, eps = 0.05) {
  return ["left", "top", "width", "height"].every((k) => Math.abs(a[k] - b[k]) < eps);
}

function renderTable(elementId, rows) {
  const el = document.getElementById(elementId);
  if (!el) return;
  if (rows.length === 0) {
    el.innerHTML = "<p class='muted'>Nothing selected.</p>";
    return;
  }
  const cols = Object.keys(rows[0]);
  const esc = (v) => String(v).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
  el.innerHTML =
    "<table><thead><tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr></thead><tbody>" +
    rows.map((r) => "<tr>" + cols.map((c) => `<td>${esc(r[c])}</td>`).join("") + "</tr>").join("") +
    "</tbody></table>";
}

function renderReferenceStatus() {
  const el = document.getElementById("reference-status");
  if (!el) return;
  el.textContent = pickedReference
    ? `Picked up: "${pickedReference.name}" ${fmt(pickedReference)}`
    : "No reference picked up.";
}

function renderDiagnostics() {
  const el = document.getElementById("diagnostics");
  if (!el) return;
  const d = Office.context.diagnostics || {};
  const sets = ["PowerPointApi 1.4", "PowerPointApi 1.5", "PowerPointApi 1.8", "PowerPointApi 1.10", "SharedRuntime 1.1", "KeyboardShortcuts 1.1"]
    .map((s) => { const [name, ver] = s.split(" "); return `${s}: ${supports(name, ver) ? "yes" : "no"}`; });
  el.innerHTML =
    `<div><b>Host</b> ${d.host || "?"} · <b>Platform</b> ${d.platform || "?"} · <b>Version</b> ${d.version || "?"}</div>` +
    `<div class="muted">${sets.join(" · ")}</div>`;
}

function wireTaskPane() {
  const bind = (id, fn) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("click", () => Promise.resolve(fn()).catch((e) => log(`${id} FAILED: ${e.message || e}`)));
  };
  bind("btn-dump", dumpSelection);
  bind("btn-probe-shapes", createProbeShapes);
  bind("btn-acceptance-shapes", createAcceptanceShapes);
  bind("btn-shortcuts", showRegisteredShortcuts);
  bind("btn-clear-log", () => { logBuffer.length = 0; log("log cleared"); });
  for (const actionId of Object.keys(commands)) bind("btn-" + actionId, guarded(actionId));

  const modeSelect = document.getElementById("reference-mode");
  if (modeSelect) {
    modeSelect.value = referenceMode;
    modeSelect.addEventListener("change", () => {
      referenceMode = modeSelect.value;
      log(`REFERENCE_MODE (runtime) = ${referenceMode}`);
    });
  }
  for (const key of ["RECENTER_ON_REF", "KEEP_CENTER"]) {
    const box = document.getElementById("cfg-" + key);
    if (box) {
      box.checked = CONFIG[key];
      box.addEventListener("change", () => { CONFIG[key] = box.checked; log(`${key} = ${box.checked}`); });
    }
  }

  renderDiagnostics();
  renderReferenceStatus();
  const el = document.getElementById("log");
  if (el) el.textContent = logBuffer.join("\n");
}

// ---------------------------------------------------------------------------
// 6. Wiring
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  // The first argument must match "actions[].id" in shortcuts.json exactly.
  if (Office.actions && typeof Office.actions.associate === "function") {
    for (const actionId of Object.keys(commands)) {
      Office.actions.associate(actionId, guarded(actionId));
    }
  } else {
    log("Office.actions.associate unavailable — not running inside an Office shared runtime.");
  }
  log(`ready: host=${info.host} platform=${info.platform} version=${(Office.context.diagnostics || {}).version} mode=${referenceMode}`);

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wireTaskPane);
  } else {
    wireTaskPane();
  }
});
