/** UI check: opens screens in a headless browser and saves screenshots to .local/ui-*.png. bun scripts/checks/ui.ts [route…] */
import { chromium } from "playwright-core";
const routes = process.argv.slice(2).length ? process.argv.slice(2) : ["/", "/trainer", "/settings"];
const b = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" }).catch(() => chromium.launch());
const p = await b.newPage({ viewport: { width: 1440, height: 900 } });
p.on("pageerror", (e) => console.log("pageerror", e.message));
p.on("console", (m) => m.type() === "error" && console.log("console", m.text().slice(0, 300)));
for (const r of routes) {
  await p.goto(`http://localhost:3300${r}`);
  await p.waitForTimeout(2500);
  await p.screenshot({ path: `.local/ui${r.replace(/\W+/g, "-")}.png` });
  console.log("shot", r);
}
await b.close();
