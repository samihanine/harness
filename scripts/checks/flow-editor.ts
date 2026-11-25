/** Builder in editor mode: a report forced to "editor" (as a PBIR-Legacy one would be), the AI changes it in the embedded editor. */
import { answerAsk, ask, lastAnswer, openApp, shot } from "./browser";
const ID = "75cbbf42-6159-4023-9053-1485050722b6";
const { ctx, page } = await openApp();
await page.goto("http://localhost:3300/builder");
await page.waitForTimeout(2000);
await page.getByRole("button", { name: /Add a report/ }).first().click().catch(() => undefined);
await page.getByPlaceholder(/Search, or paste/).fill("https://app.powerbi.com/groups/me/reports/" + ID);
await page.getByRole("button", { name: "Add from the link" }).click();
await page.waitForTimeout(4000);
// Forced into editor mode (the store of the page).
await page.evaluate(async (id) => {
  const db: IDBDatabase = await new Promise((res, rej) => { const r = indexedDB.open("hub-v2"); r.onsuccess = () => res(r.result); r.onerror = rej; });
  const tx = db.transaction("items", "readwrite");
  const store = tx.objectStore("items");
  const item: any = await new Promise((res) => { const g = store.get(`report:${id}`); g.onsuccess = () => res(g.result); });
  store.put({ ...item, access: "editor" }, `report:${id}`);
  await new Promise((res) => (tx.oncomplete = res));
}, ID);
await page.reload();
await page.waitForTimeout(15000);
await shot(page, "editor-1");
await ask(page, "Sur la page Synthèse, ajoute une carte avec le revenu pondéré (Factored Revenue) et mets le titre du graphique mensuel en doré.");
await answerAsk(page, 600_000);
console.log("ANSWER:", (await lastAnswer(page)).slice(-600));
await page.waitForTimeout(4000);
await shot(page, "editor-2");
await ctx.close();
