/**
 * How long slides take to draw: the whole deck, the same again (kept), then after one text change (only that
 * slide drawn, from a light copy).  bun scripts/checks/pptx-speed.ts <file.pptx | stored file id>
 */
import { readStoredFile } from "@server/files";
import { closeBrowser } from "@server/capture/browser";
import { Pptx } from "@server/pptx/ooxml";
import { slidePictures } from "@server/pptx/render";
const arg = process.argv[2];
const data = arg.endsWith(".pptx") ? new Uint8Array(await Bun.file(arg).arrayBuffer()) : (await readStoredFile(arg)).data;
const time = async (label: string, f: () => Promise<unknown[]>) => {
  const t = Date.now();
  const r = await f();
  console.log(`${label}: ${r.length} pictures in ${((Date.now() - t) / 1000).toFixed(1)} s`);
};
await time("whole deck", () => slidePictures(data));
await time("the same again", () => slidePictures(data, [1, 2, 3]));
const deck = await Pptx.load(data);
const hits = await deck.replaceText((await deck.structure())[1]?.shapes.find((s) => s.text)?.text?.split(/\s+/)[0] ?? "a", "CHANGED", { slide: 2 });
console.log("replaced:", JSON.stringify(hits));
const changed = await deck.save();
await time("after a change on slide 2 (slides 1-3 asked)", () => slidePictures(changed, [1, 2, 3], 1280));
const small = await (await Pptx.load(changed)).only([2]);
console.log(`light copy with slide 2: ${(small.byteLength / 1024).toFixed(0)} KB (deck ${(data.byteLength / 1024).toFixed(0)} KB)`);
await closeBrowser();
if (process.argv[3]) {
  const pics = await slidePictures(changed, [2], 1280);
  await Bun.write(process.argv[3], pics[0]);
  await closeBrowser();
}
