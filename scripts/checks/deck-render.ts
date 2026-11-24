import { deckToPptx } from "@server/pptx/deck";
import { slidePictures } from "@server/pptx/render";
import { closeBrowser } from "@server/capture/browser";
const md = JSON.parse(await Bun.file(process.argv[2] ?? ".local/deck-doc.json").text());
const pics = await slidePictures(new Uint8Array(await deckToPptx(md)));
for (const [i, p] of pics.entries()) await Bun.write(`.local/d-${i + 1}.png`, p);
console.log(pics.length);
await closeBrowser();
