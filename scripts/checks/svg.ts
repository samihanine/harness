import { loadPicture } from "@server/pptx/images";
import { closeBrowser } from "@server/capture/browser";
for (const src of ["https://upload.wikimedia.org/wikipedia/commons/a/a7/React-icon.svg", "https://www.gstatic.com/webp/gallery/1.webp"]) {
  const p = await loadPicture(src).catch((e) => ({ error: String(e) }));
  console.log(src.slice(-30), "error" in p ? p.error : `${p.mime} ${p.width}x${p.height}`);
}
await closeBrowser();
