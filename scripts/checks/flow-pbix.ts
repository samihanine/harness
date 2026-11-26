/**
 * The Builder on a dropped .pbix: imported, shown embedded, changed by the assistant, downloaded (screenshots
 * .local/pbix-*.png, the download .local/pbix-download.pbix).  bun scripts/checks/flow-pbix.ts <file.pbix>
 */
import { answerAsk, ask, lastAnswer, openApp, shot } from "./browser";
const file = process.argv[2] ?? ".local/thin.pbix";
const { ctx, page } = await openApp();
await page.goto("http://localhost:3300/builder?item=");
await page.waitForTimeout(2000);
await page.locator('input[type=file][accept=".pbix"]').setInputFiles(file);
await page.waitForTimeout(20000);
await shot(page, "pbix-1");
await ask(page, "Ajoute en bas de la page un texte doré 'Préparé par le Builder' en gras, taille 14.", 400_000);
await answerAsk(page, 400_000);
console.log("ANSWER:", (await lastAnswer(page)).slice(-500));
await page.waitForTimeout(8000);
await shot(page, "pbix-2");
const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 120_000 }), page.getByRole("button", { name: /Download .pbix/ }).click()]);
await dl.saveAs(".local/pbix-download.pbix");
console.log("downloaded", dl.suggestedFilename());
await ctx.close();
