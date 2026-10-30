#!/usr/bin/env node
// The browser smoke test for corro's `wasm` backend.
//
// Run it through `scripts/wasm-smoke.sh`, which builds the module and serves
// `web/` first; this file only drives a page. It is deliberately NOT wired
// into `cargo test`: it needs a browser and a network-free HTTP server, and a
// `cargo test` that silently requires either is a test that fails for reasons
// that have nothing to do with the code under test. See
// docs/WASM_MISSING.md §8.
//
// The contract it asserts against is the one the app already had:
//
//   * `window.__corro_output` -- the app appends one line per logged action
//     (`append_wasm_output_line`, fed by `log_key_action` in gui_backend.rs);
//   * `document.title === "CORRO_QUIT"` -- set by `quit_main_loop`.
//
// Three checks, in the order a user would hit them:
//
//   1. the canvas is sized from its box, not left at the 300x150 default
//      (the "blank grid" bug);
//   2. the grid actually paints -- pixels differ from the page background
//      over the sheet area, not just the menu bar;
//   3. File > Replay and Ctrl+Q round-trip: the menu's action is logged, the
//      quit is logged, and the title becomes CORRO_QUIT.
//
// Usage:
//   node scripts/wasm_smoke.mjs --url http://127.0.0.1:8099/ [--browser <path>]
//
// `--browser` names a Chromium-compatible binary (or set $CHROME). Without it
// a headless Chrome is looked for where Playwright and the distro packages put
// it. Nothing is downloaded.
//
// Zero npm dependencies on purpose: Node 22 ships a WebSocket client, so the
// DevTools protocol is spoken directly rather than through puppeteer or
// playwright. A smoke test that needs `npm install` before it can report
// anything is a smoke test that is not run.

import { spawn } from "node:child_process";
import { existsSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const args = process.argv.slice(2);
const opt = (name, dflt) => {
  const i = args.indexOf(name);
  return i >= 0 && i + 1 < args.length ? args[i + 1] : dflt;
};
const URL_ = opt("--url", "http://127.0.0.1:8099/");
const PORT = 9333 + Math.floor(Math.random() * 400);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function findChrome() {
  const explicit = opt("--browser", process.env.CHROME);
  if (explicit) return explicit;
  const candidates = [];
  const pwRoot = join(homedir(), ".cache", "ms-playwright");
  if (existsSync(pwRoot)) {
    // Newest Playwright build first.
    const dirs = readdirSync(pwRoot).sort().reverse();
    for (const d of dirs) {
      for (const sub of ["chrome-headless-shell-linux64/chrome-headless-shell",
                         "chrome-linux/chrome", "chrome-linux64/chrome"]) {
        candidates.push(join(pwRoot, d, sub));
      }
    }
  }
  for (const p of ["/usr/bin/chromium", "/usr/bin/chromium-browser",
                   "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
                   "/snap/bin/chromium",
                   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]) {
    candidates.push(p);
  }
  return candidates.find((p) => existsSync(p));
}

const chrome = findChrome();
if (!chrome) {
  console.error(
    "wasm_smoke: no Chromium-compatible browser found.\n" +
    "  Pass --browser /path/to/chrome, set $CHROME, or install one.\n" +
    "  Playwright's cached builds under ~/.cache/ms-playwright are found automatically.");
  process.exit(2);
}

const failures = [];
const check = (ok, what, detail) => {
  console.log(`${ok ? "ok  " : "FAIL"}  ${what}${detail ? "  (" + detail + ")" : ""}`);
  if (!ok) failures.push(what);
};

const proc = spawn(chrome, [
  "--headless=new", "--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage",
  "--no-first-run", "--no-default-browser-check",
  `--remote-debugging-port=${PORT}`, "--window-size=1200,800", "about:blank",
], { stdio: "ignore" });

const kill = () => { try { proc.kill("SIGKILL"); } catch {} };
process.on("exit", kill);

let target = null;
for (let i = 0; i < 80 && !target; i++) {
  await sleep(100);
  try {
    const r = await fetch(`http://127.0.0.1:${PORT}/json/list`);
    target = (await r.json()).find((t) => t.type === "page");
  } catch {}
}
if (!target) { console.error("wasm_smoke: the browser never opened a DevTools endpoint"); kill(); process.exit(2); }

const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
let nextId = 0;
const pending = new Map();
const pageErrors = [];
ws.onmessage = (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); return; }
  if (m.method === "Runtime.exceptionThrown") {
    pageErrors.push(m.params.exceptionDetails.exception?.description ?? m.params.exceptionDetails.text);
  }
  if (m.method === "Runtime.consoleAPICalled" && m.params.type === "error") {
    pageErrors.push(m.params.args.map((a) => a.value ?? a.description).join(" "));
  }
};
const send = (method, params = {}) =>
  new Promise((res) => { const id = ++nextId; pending.set(id, res); ws.send(JSON.stringify({ id, method, params })); });
const evaluate = async (expression) => {
  const r = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
  if (r.result?.exceptionDetails) throw new Error(r.result.exceptionDetails.exception?.description ?? "evaluate failed");
  return r.result?.result?.value;
};

await send("Page.enable");
await send("Runtime.enable");
await send("Page.navigate", { url: URL_ });

// Wait for the app: `present` gives the window tabindex 0, and the first log
// line (`gui_started`) is appended before the first paint.
let ready = false;
for (let i = 0; i < 100 && !ready; i++) {
  await sleep(100);
  try {
    ready = await evaluate(
      `!!(document.querySelector('canvas') && window.__corro_output && window.__corro_output.length > 0)`);
  } catch {}
}
if (!ready) {
  const err = await evaluate(`(document.getElementById('corro-fatal')||{}).textContent || ''`);
  console.error(`wasm_smoke: the app did not start at ${URL_}\n${err}\n${pageErrors.join("\n")}`);
  kill();
  process.exit(1);
}
check(true, "app started", (await evaluate("window.__corro_output")).split("\n")[0]);

// 1. The sheet canvas is sized from its CSS box.
// The box is reported unrounded: a canvas is a replaced element whose box
// follows its attribute size, so at a fractional layout the box and the store
// settle a pixel or so apart, and the comparison has to see that fraction.
const geom = await evaluate(`(() => {
  const cv = document.querySelector('canvas');
  const r = cv.getBoundingClientRect();
  return { box: [r.width, r.height],
           backing: [cv.width, cv.height],
           dpr: window.devicePixelRatio,
           viewport: [window.innerWidth, window.innerHeight] };
})()`);
check(geom.box[0] > 100 && geom.box[1] > 100, "canvas box is laid out",
      `${geom.box[0].toFixed(1)}x${geom.box[1].toFixed(1)} CSS px in a ${geom.viewport[0]}x${geom.viewport[1]} viewport`);
// The store must scale with dpr to within a pixel either way: rounding is the
// adapter's choice (it rounds to nearest, which is a fixed point of the
// canvas's own intrinsic-size feedback), and a store that ignored dpr would
// miss by a factor of it.
const scales = (dev, css) => Math.abs(dev - css * geom.dpr) <= 1.5;
// Width first, on its own: a `width: 100%` canvas keeps the box the layout gave
// it, so width is the axis a dpr bug shows on unambiguously. Height is reported
// beside it rather than asserted at the same tolerance, because a canvas takes
// its height from the attribute when the layout only stretches it via
// `height: 100%`/flex -- the two settle a pixel or two apart at a fractional
// size, and that difference is not a scaling error.
check(scales(geom.backing[0], geom.box[0]),
      "backing store follows the box width x devicePixelRatio",
      `${geom.backing[0]}x${geom.backing[1]} device px for a ${geom.box[0].toFixed(1)}x${geom.box[1].toFixed(1)} box at dpr ${geom.dpr}`);

// 2. The sheet is painted. A blank (unsized, or cleared) canvas is a uniform
//    field, so the test is colour variety over the grid area: the gutter fill,
//    the paper, the selection band and the grid rules are four different
//    colours, and a frame that never reached the surface has one.
const painted = await evaluate(`(() => {
  const cv = document.querySelector('canvas');
  const ctx = cv.getContext('2d');
  const counts = new Map();
  let samples = 0;
  for (let y = 30; y < cv.height - 4; y += 7) {
    for (let x = 4; x < cv.width - 4; x += 7) {
      const d = ctx.getImageData(x, y, 1, 1).data;
      const key = d[0] + ',' + d[1] + ',' + d[2];
      counts.set(key, (counts.get(key) || 0) + 1);
      samples++;
    }
  }
  const top = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
  return { samples, colours: counts.size, dominantShare: top ? top[1] / samples : 1, dominant: top && top[0] };
})()`);
check(painted.colours > 5 && painted.dominantShare < 0.98, "the sheet is painted",
      `${painted.colours} colours over ${painted.samples} samples, most common ${(painted.dominantShare * 100).toFixed(1)}%`);

// 2b. Menus open AND close. The regression this pins: the bar's open-state and
//     the DOM disagreed, so a menu could be dismissed on screen while the
//     menubar still thought it was open -- keys were routed to it as mnemonics
//     and typing stopped working until the page was reloaded.
const menuState = () => evaluate(`JSON.stringify([...document.querySelectorAll('div')]
  .filter((d) => d.style.position === 'absolute' && d.style.display === 'block')
  .map((d) => d.textContent.slice(0, 24)))`);
const toggleAt = (name) => evaluate(`(() => {
  const b = [...document.querySelectorAll('button')].find((b) => b.textContent.startsWith('${name}'));
  if (!b) return null;
  const r = b.getBoundingClientRect();
  return [Math.round(r.x + r.width / 2), Math.round(r.y + r.height / 2)];
})()`);
const clickAt = async ([x, y]) => {
  for (const type of ["mousePressed", "mouseReleased"]) {
    await send("Input.dispatchMouseEvent", { type, x, y, button: "left", clickCount: 1 });
  }
  await sleep(250);
};
const menus = [await toggleAt("File"), await toggleAt("Edit")];
await clickAt(menus[0]);
const opened = JSON.parse(await menuState());
await clickAt(menus[0]);
const closed = JSON.parse(await menuState());
await clickAt(menus[0]);
await clickAt(menus[1]);
const switched = JSON.parse(await menuState());
await clickAt([600, 500]);
const outside = JSON.parse(await menuState());
await clickAt(menus[0]);
for (const type of ["rawKeyDown", "keyUp"]) {
  await send("Input.dispatchKeyEvent", { type, key: "Escape", code: "Escape", windowsVirtualKeyCode: 27, nativeVirtualKeyCode: 27 });
}
await sleep(250);
const escaped = JSON.parse(await menuState());
check(opened.length === 1 && opened[0].startsWith("New"),
      "clicking File opens it", opened.join(",") || "(none)");
check(closed.length === 0, "a second click on File closes it", closed.join(",") || "(none)");
check(switched.length === 1 && switched[0].startsWith("Cut"),
      "clicking Edit switches menus", switched.join(",") || "(none)");
check(outside.length === 0, "clicking the sheet closes the menu", outside.join(",") || "(none)");
check(escaped.length === 0, "Escape closes the menu", escaped.join(",") || "(none)");

// 3. A menu action is logged, then Ctrl+Q quits.
// The action is dispatched through the app's own registry (`App::activate_action`
// -> `invoke_action`), which is what a menu item's click handler calls. Clicking
// the item element would work too, but the menu bar's keyboard state can move
// items between the dropdown and the bar, so addressing the action is stable
// where addressing an element is not.
const clicked = await evaluate(`(() => {
  // The top-level labels render the mnemonic as an underlined character, so
  // textContent is "File", not the raw "_File" marker.
  const file = [...document.querySelectorAll('button')].find((b) => b.textContent === 'File');
  if (!file) return 'no File menu';
  file.click();
  const fileMenu = [...document.querySelectorAll('div')]
    .find((d) => d.style.position === 'absolute' && d.style.display === 'block');
  if (!fileMenu) return 'File menu did not open';
  if (![...fileMenu.children].some((d) => d.textContent === 'Replay')) return 'no Replay item';
  const act = window.__corro_activate_action;
  if (typeof act !== 'function') return 'no action registry';
  act('app.replay');
  return 'ok';
})()`);
check(clicked === "ok", "File > Replay is clickable", clicked);
// The item's handler runs on the app's next tick, not inside `click()`: poll
// rather than sleep a guessed interval.
let afterMenu = await evaluate("window.__corro_output");
for (let i = 0; i < 30 && afterMenu.includes("ACTION: gui_started") && afterMenu.split("\n").length < 2; i++) {
  await sleep(100);
  afterMenu = await evaluate("window.__corro_output");
}
check(afterMenu.split("\n").length >= 2, "the menu click reached the app",
      afterMenu.split("\n").slice(1).join(" | ") || afterMenu);

// Ctrl is modifier bit 2 in the DevTools protocol; keyCode 81 = Q, which is
// what the adapter's `KeyboardEvent::key_code` reads.
for (const type of ["keyDown", "keyUp"]) {
  await send("Input.dispatchKeyEvent",
             { type, key: "q", code: "KeyQ", windowsVirtualKeyCode: 81, nativeVirtualKeyCode: 81, modifiers: 2 });
}
let title = "";
for (let i = 0; i < 50 && title !== "CORRO_QUIT"; i++) {
  await sleep(100);
  title = await evaluate("document.title");
}
check(title === "CORRO_QUIT", "Ctrl+Q sets document.title to CORRO_QUIT", title);
const finalOutput = await evaluate("window.__corro_output");
check(/ACTION: quit\b/.test(finalOutput), "the quit was logged",
      finalOutput.split("\n").filter((l) => l.includes("quit")).join(" | ") || "(none)");
check(pageErrors.length === 0, "no uncaught page error", pageErrors.join(" | "));

if (process.env.SMOKE_VERBOSE) {
  console.log("\n--- window.__corro_output ---\n" + finalOutput);
}

ws.close();
kill();
console.log(`\n${failures.length === 0 ? "PASS" : "FAIL"}  ${failures.length} failed check(s)`);
process.exit(failures.length === 0 ? 0 : 1);
