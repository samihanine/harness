/**
 * The Data Scientist's workbook: a new one in OneDrive, then agent turns on it (CUBE formulas on a model,
 * formulas, formats), then its sheet as the screen shows it (.local/workbook.png).
 *   bun scripts/checks/workbook.ts '<message>' ['<answer>'…]
 */
import { createWorkbook, ensureConnections } from "@server/workbook/excel";
import { sheetHtml, sheetPreview } from "@server/workbook/preview";
import { closeBrowser, renderHtml } from "@server/capture/browser";
import { runChat } from "@server/agent/run";
const model = { id: "eda77871-2401-4255-ba6c-2fc3c66902a1", name: "Revenue Opportunities" };
const wb = process.env.ITEM ? { itemId: process.env.ITEM } : await createWorkbook(`check-${Date.now()}`, "Check");
console.log("workbook", wb);
await ensureConnections(wb.itemId, [model]);
const scope = { datasets: [model], doc: wb, files: [] };
const conv = `wb-${Date.now()}`;
for (const message of process.argv.slice(2)) {
  let text = "";
  console.log(`\n=== USER: ${message}`);
  await runChat({ conversationId: conv, screen: "scientist", message, scope }, (e) => {
    if (e.type === "text") text += e.text;
    else if (e.type === "tool" && e.status !== "running") console.log(`[${e.name} ${e.status} ${e.ms}ms] ${JSON.stringify(e.args).slice(0, 200)}\n   → ${(e.result ?? "").slice(0, 300).replace(/\n/g, "\n     ")}`);
    else if (e.type === "effect" && e.name === "ask") console.log(`[ASK] ${JSON.stringify(e.data).slice(0, 400)}`);
    else if (e.type === "error") console.log(`[ERROR] ${e.error}`);
  }, new AbortController().signal);
  console.log(`--- answer ---\n${text}`);
}
const p = await sheetPreview(wb.itemId, process.env.SHEET);
await Bun.write(".local/workbook.png", await renderHtml(sheetHtml(p), 1400));
await closeBrowser();
