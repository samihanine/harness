/**
 * bun run smoke (the dev server running: bun dev): opens every screen of the app in a browser, each of
 * its tabs and every tab of Settings, and fails on a screen that cannot be drawn ("Something went
 * wrong", "could not be drawn"), an error of the page, or a local server that stopped. Microsoft is
 * simulated (nothing is read or written anywhere): it checks that the app stands, not its data.
 */
const { chromium } = require("playwright") as typeof import("playwright");
const BASE = process.env.APP_URL ?? "http://localhost:3300";
const SERVER = process.env.SERVER_URL ?? "http://127.0.0.1:3301";

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
await ctx.route("**/get-token**", (r) => r.fulfill({ json: { token: "fake", expiresAt: Date.now() + 3600e3, account: "smoke@test" } }));
await ctx.route("**/ms/**", (r) => r.fulfill({ status: 404, json: { error: { message: "simulated" } } }));
const page = await ctx.newPage();
const errors: string[] = [];
page.on("pageerror", (e) => errors.push(`page error: ${e.message.slice(0, 300)}`));
page.on("console", (m) => m.type() === "error" && !/Failed to load resource|ERR_/.test(m.text()) && errors.push(`console: ${m.text().slice(0, 300)}`));
const broken = async () => /Something went wrong|could not be drawn/.test(await page.locator("body").innerText());

let failed = 0;
const check = async (label: string) => {
  if (await broken()) (failed++, console.log(`✗ ${label}: the screen cannot be drawn`));
};
await page.goto(`${BASE}/`, { waitUntil: "load" });
await page.waitForTimeout(1500);
for (const name of ["Trainer", "Builder", "Data Scientist", "Presenter", "Reports", "Datasets", "Excels", "Confluence", "Settings"]) {
  const link = page.locator("aside a").filter({ hasText: new RegExp(name.replace(" ", "\\s+")) }).first();
  if (!(await link.count())) {
    console.log(`- ${name}: hidden (its flag is off)`);
    continue;
  }
  await link.click();
  await page.waitForTimeout(1000);
  await check(name);
  const tabs = (await page.getByRole("tab").allInnerTexts()).map((t) => t.replace(/\s+\d+$/, "").trim());
  for (const t of tabs) {
    await page.getByRole("tab", { name: new RegExp(`^${t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}`) }).first().click().catch(() => undefined);
    await page.waitForTimeout(400);
    await check(`${name} › ${t}`);
  }
  console.log(`${failed ? "✗" : "✓"} ${name}: ${tabs.length} tab(s)`);
}
const alive = await fetch(`${SERVER}/get-token`, { headers: { "X-Hub-Client": "1" } }).then((r) => r.status < 500, () => false);
if (!alive) (failed++, console.log("✗ the local server does not answer"));
if (errors.length) (failed++, console.log(`✗ errors in the page:\n${[...new Set(errors)].slice(0, 8).join("\n")}`));
await browser.close();
console.log(failed ? `\n${failed} problem(s)` : "\nSmoke test: ok");
process.exit(failed ? 1 : 0);
