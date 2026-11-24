import { answerAsk, ask, lastAnswer, openApp, shot } from "./browser";
const { ctx, page } = await openApp();
await page.goto("http://localhost:3300/presenter");
await page.waitForTimeout(1500);
const create = page.getByRole("button", { name: "Create from template" });
if (await create.isVisible()) await create.click();
const empty = page.getByRole("button", { name: "New deck" });
if (await empty.count()) await empty.last().click();
else {
  await page.locator("header button").filter({ hasText: /Presentation|deck/i }).first().click();
  await page.getByRole("menuitem", { name: "New deck (template)" }).click();
}
await page.waitForTimeout(15000);
await shot(page, "presenter-1");
await ask(page, "Fais 4 slides sur la répartition du pipeline par région (modèle Revenue Opportunities) avec un tableau et un visuel Power BI.");
await answerAsk(page, 600_000);
console.log("ANSWER:", (await lastAnswer(page)).slice(-600));
await page.waitForTimeout(20000);
await shot(page, "presenter-2");
await page.getByRole("button", { name: /Edit/ }).first().click();
await page.waitForTimeout(1000);
await shot(page, "presenter-3");
await ctx.close();
