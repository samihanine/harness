import { answerAsk, lastAnswer, openApp, shot } from "./browser";
const { ctx, page } = await openApp();
await page.goto("http://localhost:3300/builder");
await page.waitForTimeout(2000);
const picker = page.locator("header button").filter({ hasText: /report|Hub v2|Revenue/i }).first();
await picker.click();
const existing = page.getByRole("menuitem", { name: /Hub v2 check — Builder/ });
if (await existing.count()) await existing.click();
else {
  await page.getByText("Add a report…").click();
  await page.getByPlaceholder("Search, or paste a Power BI link").fill("Hub v2 check");
  await page.getByRole("button", { name: /Hub v2 check — Builder/ }).first().click();
}
await page.waitForTimeout(12000);
await page.locator('input[type=file][accept="image/*"]').setInputFiles(".local/target-screen.png");
await page.waitForTimeout(1500);
await page.getByPlaceholder(/./).last().fill("Refais ce screen sur une nouvelle page « Overview » de mon rapport, au plus proche (visuels, champs, disposition, couleurs).");
await page.keyboard.press("Enter");
await page.waitForTimeout(1500);
await page.getByRole("button", { name: "Stop" }).waitFor({ state: "detached", timeout: 400_000 });
console.log("ANSWER1:", (await lastAnswer(page)).slice(0, 700));
await shot(page, "builder-ui-1");
await answerAsk(page, 600_000);
console.log("ANSWER2:", (await lastAnswer(page)).slice(-900));
await page.waitForTimeout(8000);
await shot(page, "builder-ui-2");
await ctx.close();
