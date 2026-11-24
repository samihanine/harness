/** The uploaded PowerPoint being edited: each change makes a new version (the screen can go back). */
import { Pptx } from "@server/pptx/ooxml";
import { readStoredFile, saveFile } from "@server/files";
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

/** A picture's bytes from a source: a capture (file:<id>), a file id, a URL. */
export async function pictureBytes(src: string): Promise<Uint8Array> {
  if (src.startsWith("file:")) return (await readStoredFile(src.slice(5))).data;
  if (/^https?:\/\//.test(src)) {
    const r = await fetch(src, { signal: AbortSignal.timeout(15_000) });
    if (!r.ok) throw new Error(`HTTP ${r.status} for ${src}`);
    return new Uint8Array(await r.arrayBuffer());
  }
  return (await readStoredFile(src)).data;
}
