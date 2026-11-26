/**
 * The Presenter: agent turns on a new presentation (the company's templates), then its pictures, PowerPoint and
 * PDF (.local/slides-*.png, .local/slides.pptx, .local/slides.pdf).  bun scripts/checks/slides.ts '<message>' ['<answer>'…]
 */
import { DEFAULT_TEMPLATES } from "@shared/slide-templates";
import { runChat } from "@server/agent/run";
import { deckPdf, deckPptx, slidePictures } from "@server/slides/render";
import { closeBrowser } from "@server/capture/browser";
const scope: any = { doc: { title: "Check", markdown: process.env.MD ?? "" }, templates: DEFAULT_TEMPLATES, references: [{ id: "2066ecc8-ead6-425b-a465-2475384528f1", name: "Revenue Opportunities" }], datasets: [{ id: "eda77871-2401-4255-ba6c-2fc3c66902a1", name: "Revenue Opportunities" }] };
const conv = `slides-${Date.now()}`;
for (const message of process.argv.slice(2)) {
  let text = "";
  console.log(`\n=== USER: ${message}`);
  await runChat({ conversationId: conv, screen: "presenter", message, scope }, (e) => {
    if (e.type === "text") text += e.text;
    else if (e.type === "tool" && e.status !== "running") console.log(`[${e.name} ${e.status} ${e.ms}ms${e.images ? ` 🖼${e.images}` : ""}] ${JSON.stringify(e.args).slice(0, 160)}\n   → ${(e.result ?? "").slice(0, 240).replace(/\n/g, " ")}`);
    else if (e.type === "effect" && e.name === "deck") scope.doc = { ...scope.doc, markdown: e.data };
    else if (e.type === "effect" && e.name === "templates") scope.templates = e.data;
    else if (e.type === "effect" && e.name === "ask") console.log(`[ASK] ${JSON.stringify(e.data).slice(0, 300)}`);
    else if (e.type === "error") console.log(`[ERROR] ${e.error}`);
  }, new AbortController().signal);
  console.log(`--- answer ---\n${text}`);
}
const md = scope.doc.markdown;
await Bun.write(".local/slides.md", md);
const t = Date.now();
const pics = await slidePictures(md, scope.templates);
console.log(`${pics.pictures.length} pictures in ${Date.now() - t} ms`);
for (const [i, p] of pics.pictures.entries()) await Bun.write(`.local/slides-${i + 1}.png`, p);
await Bun.write(".local/slides.pptx", await deckPptx(md, scope.templates, "Check"));
await Bun.write(".local/slides.pdf", await deckPdf(md, scope.templates));
await closeBrowser();
