/**
 * Round 4 UI: the Trainer from its address (?item=), a conversation, the history (all items, the item switched
 * with it, the view restored), feedback, the assistant resized, the Structure tab, the Data Scientist's tabs,
 * Settings › Confluence projects. Screenshots in .local/r4-*.png.
 */
import { answerAsk, ask, openApp, shot } from "./browser";
const REPORT = process.argv[2] ?? "2066ecc8-ead6-425b-a465-2475384528f1";
const { ctx, page } = await openApp();
await page.goto("http://localhost:3300/builder");
await page.waitForTimeout(1500);
// The report added from its browser address.
await page.goto("http://localhost:3300/trainer");
await page.waitForTimeout(1500);
// The picker's menu (the current report's name, or "Choose a report"), then "Add a report…".
await page.locator("h1").first().locator("xpath=..").locator("button").first().click();
await page.getByText("Add a report…").click();
await page.waitForTimeout(800);
const input = page.getByPlaceholder(/paste the report's address/);
if (await input.count()) {
  await input.fill(`https://app.powerbi.com/groups/me/reports/${REPORT}/ReportSection?experience=power-bi`);
  await page.keyboard.press("Enter");
  await page.waitForTimeout(5000);
}
console.log("URL after add:", page.url());
await page.goto(`http://localhost:3300/trainer?item=${REPORT}`);
await page.waitForTimeout(12000);
await shot(page, "r4-trainer");
await ask(page, "Quel est le revenu total ? Réponds simplement.", 300_000);
await answerAsk(page);
await page.waitForTimeout(1000);
console.log("URL with conversation:", page.url());
await shot(page, "r4-answer");
// Feedback menu
await page.getByRole("button", { name: "This conversation" }).click();
await page.getByRole("button", { name: /Good/ }).click();
await page.keyboard.press("Escape");
// History
await page.getByRole("button", { name: "Conversations" }).click();
await page.waitForTimeout(500);
await shot(page, "r4-history");
await page.keyboard.press("Escape");
// Resize the assistant
const handle = page.locator('[title^="Drag to resize"]');
const box = await handle.boundingBox();
if (box) {
  await page.mouse.move(box.x + 3, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + 180, box.y + box.height / 2, { steps: 5 });
  await page.mouse.up();
}
await page.waitForTimeout(1500);
await shot(page, "r4-resized");
// Structure tab
await page.getByRole("button", { name: /Structure/ }).first().click();
await page.waitForTimeout(8000);
await shot(page, "r4-structure");
// Data Scientist tabs
await page.goto("http://localhost:3300/scientist");
await page.waitForTimeout(2500);
await page.getByRole("button", { name: /Semantic models/ }).first().click().catch(() => undefined);
await page.waitForTimeout(800);
await shot(page, "r4-scientist-models");
await page.getByRole("button", { name: /^Excel/ }).first().click().catch(() => undefined);
await page.waitForTimeout(800);
await shot(page, "r4-scientist-excel");
// Settings › Confluence
await page.goto("http://localhost:3300/settings");
await page.waitForTimeout(1500);
await page.getByRole("button", { name: /Confluence/ }).first().click();
await page.waitForTimeout(6000);
await shot(page, "r4-settings-confluence");
await ctx.close();
