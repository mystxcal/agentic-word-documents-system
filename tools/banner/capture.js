// Deterministic frame capture for the project banner.
//
// The page is stepped, never timed: `window.__step(t)` draws frame `t` from
// scratch, so the same source always produces the same frames and the loop
// closes exactly. Nothing here depends on wall-clock speed.
//
// Usage: node capture.js <page.html> <outDir> <theme> <fps> <phase> <scale>

"use strict";
const path = require("path");

function loadPuppeteer() {
  for (const id of ["puppeteer", "puppeteer-core", "rebrowser-puppeteer-core"]) {
    try { return require(id); } catch { /* try the next one */ }
  }
  const extra = process.env.BANNER_PUPPETEER;
  if (extra) return require(extra);
  console.error(
    "No puppeteer package found. Install one of:\n" +
    "  npm install puppeteer            (downloads its own Chrome)\n" +
    "  npm install puppeteer-core       (then set CHROME_PATH)\n" +
    "or set BANNER_PUPPETEER to a module path.");
  process.exit(2);
}

const puppeteer = loadPuppeteer();
const [page, outDir, theme, fpsArg, phaseArg, scaleArg] = process.argv.slice(2);
const W = 1280, H = 400, LOOP = 12.0;
const FPS = Number(fpsArg || 25);
const PHASE = Number(phaseArg || 10.8);   // frame 0 shows a build already proven
const SCALE = Number(scaleArg || 2);
const N = Math.round(LOOP * FPS);

(async () => {
  const launch = {
    headless: "new",
    args: ["--no-sandbox", "--disable-dev-shm-usage", "--hide-scrollbars",
           "--force-device-scale-factor=1",
           `--window-size=${W * SCALE + 40},${H * SCALE + 40}`]
  };
  if (process.env.CHROME_PATH) launch.executablePath = process.env.CHROME_PATH;

  const browser = await puppeteer.launch(launch);
  const p = await browser.newPage();
  await p.setViewport({ width: W * SCALE + 40, height: H * SCALE + 40, deviceScaleFactor: 1 });

  const errs = [];
  p.on("pageerror", e => errs.push(e.message));
  p.on("console", m => { if (m.type() === "error") errs.push("console: " + m.text()); });

  await p.evaluateOnNewDocument(() => { window.__manual = true; });
  await p.goto("file://" + path.resolve(page), { waitUntil: "load" });
  await p.waitForFunction(() => window.__ready === true, { timeout: 30000 });
  await p.evaluate(t => window.__setTheme(t), theme);
  await p.evaluate(k => window.__setScale(k), SCALE);

  const el = await p.$("#c");
  for (let i = 0; i < N; i++) {
    await p.evaluate(t => window.__step(t), (PHASE + i / FPS) % LOOP);
    await el.screenshot({ path: path.join(outDir, `f_${String(i).padStart(4, "0")}.png`) });
    if (i % 50 === 0) process.stdout.write(`${i}/${N} `);
  }
  await browser.close();

  if (errs.length) {
    console.error("\nPAGE ERRORS:\n" + errs.join("\n"));
    process.exit(1);
  }
  console.log(`\ncaptured ${N} frames  ${W * SCALE}x${H * SCALE}  ${theme}`);
})();
