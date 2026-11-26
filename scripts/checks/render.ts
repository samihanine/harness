import { slidePictures } from "@server/pptx/render";
import { closeBrowser } from "@server/capture/browser";
const pics = await slidePictures(new Uint8Array(await Bun.file(process.argv[2]).arrayBuffer()), [1, 3, 4]);
for (const [i, p] of pics.entries()) await Bun.write(`.local/r-${i + 1}.png`, p);
await closeBrowser();
