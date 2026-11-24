/** The uploaded PowerPoint being edited: each change makes a new version (the screen can go back). */
import { Pptx } from "@server/pptx/ooxml";
import { readStoredFile, saveFile } from "@server/files";
import { loadPicture } from "@server/pptx/images";
import { docOf, setDoc } from "../helpers";
import type { ToolContext } from "../tool";

export type PptxDoc = { fileId: string; name: string };

export const pptxDoc = (ctx: ToolContext) => {
  const d = docOf<PptxDoc>(ctx);
  if (!d?.fileId) throw new Error("No PowerPoint is open: ask the user to upload one.");
  return d;
};

export async function openPptx(ctx: ToolContext) {
  const d = pptxDoc(ctx);
  return { doc: d, pptx: await Pptx.load((await readStoredFile(d.fileId)).data) };
}

export async function savePptx(ctx: ToolContext, doc: PptxDoc, pptx: Pptx) {
  const file = await saveFile(await pptx.save(), doc.name, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  setDoc(ctx, { ...doc, fileId: file.id });
}

/** A picture's bytes from a source (a capture file:<id>, a file id, a URL), as PNG / JPEG / GIF (SVG, WebP converted). */
export async function pictureBytes(src: string): Promise<Uint8Array> {
  const pic = await loadPicture(/^(file:|https?:|data:)/.test(src) ? src : `file:${src}`);
  return new Uint8Array(Buffer.from(pic.data.split("base64,")[1], "base64"));
}
