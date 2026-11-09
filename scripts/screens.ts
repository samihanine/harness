/**
 * What each screen shows, kept to compare before / after a change that must not change it (a
 * refactor): every route opened in a browser (the dev server running: bun dev), its accessibility tree
 * saved in .local/screens/<screen>.txt with a screenshot. `--check`: compared with the saved ones, the
 * differences printed (exit 1 when any).
 *   bun run screens            save
 *   bun run screens --check    compare
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";

const { chromium } = require("playwright") as typeof import("playwright");
const BASE = process.env.APP_URL ?? "http://localhost:3300";
const DIR = ".local/screens";
const check = process.argv.includes("--check");
const routes = ["", ...readdirSync("src/routes").filter((f) => f.endsWith(".tsx") && !f.startsWith("__") && f !== "index.tsx").map((f) => f.replace(/\.tsx$/, ""))];

mkdirSync(DIR, { recursive: true });
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
const errors: string[] = [];
page.on("pageerror", (e) => errors.push(`${page.url()}: ${e.message}`));
let changed = 0;
for (const route of routes) {
  await page.goto(`${BASE}/${route}`, { waitUntil: "networkidle" });
  await page.waitForTimeout(800);
  const name = route || "home";
  const tree = await page.locator("main").ariaSnapshot();
  const file = `${DIR}/${name}.txt`;
  if (check && existsSync(file)) {
    const before = readFileSync(file, "utf8");
    if (before !== tree) {
      changed++;
      const a = before.split("\n");
      const b = tree.split("\n");
      console.log(`✗ ${name}\n${[...a.filter((l) => !b.includes(l)).map((l) => `  - ${l}`), ...b.filter((l) => !a.includes(l)).map((l) => `  + ${l}`)].slice(0, 30).join("\n")}`);
    } else console.log(`✓ ${name}`);
  } else {
    writeFileSync(file, tree);
    await page.screenshot({ path: `${DIR}/${name}.png` });
    console.log(`saved ${name}`);
  }
}
await browser.close();
if (errors.length) console.log(`Errors in the page:\n${errors.join("\n")}`);
process.exit(changed || errors.length ? 1 : 0);
