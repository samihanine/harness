import { answerAsk, ask, lastAnswer, openApp, shot } from "./browser";
const { ctx, page } = await openApp();
await page.goto("http://localhost:3300/trainer");
await page.waitForTimeout(1500);
if (await page.getByText("Choose a report", { exact: true }).first().isVisible()) {
  await page.getByRole("button", { name: /Choose a report/ }).first().click();
  await page.getByText("Add a report…").click();
  await page.getByPlaceholder("Search, or paste a Power BI link").fill("Revenue Opportunities");
  await page.getByRole("button", { name: /^Revenue Opportunities\s*My workspace$/ }).first().click();
}
await page.waitForTimeout(12000);
await page.getByRole("button", { name: "New conversation" }).first().click();
await shot(page, "trainer-1");
await ask(page, process.argv[2] ?? "Montre-moi uniquement les opportunités de la région ouest");
console.log("ANSWER1:", await lastAnswer(page));
await shot(page, "trainer-2");
await answerAsk(page);
console.log("ANSWER2:", await lastAnswer(page));
await page.waitForTimeout(3000);
await shot(page, "trainer-3");
await ctx.close();
