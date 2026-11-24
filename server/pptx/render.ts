/**
 * Pictures of a PowerPoint's slides, as PowerPoint draws them: the file is put in the user's OneDrive
 * (a hidden "Hub previews" folder), converted to PDF by Microsoft (Graph `?format=pdf`), removed, and the
 * PDF's pages drawn by pdf.js in the headless browser. Only Microsoft is called. Same file: same pictures (cache).
 */
import { createHash } from "node:crypto";
import { graph, request } from "../ms/api";
import { renderPdf } from "../capture/browser";

const cache = new Map<string, Promise<Buffer>>();
const FOLDER = "Hub previews";

async function upload(name: string, data: Uint8Array) {
  const path = `/me/drive/root:/${encodeURIComponent(FOLDER)}/${encodeURIComponent(name)}`;
  if (data.byteLength < 4 * 1024 * 1024) {
    const r = await request("graph", `${path}:/content?@microsoft.graph.conflictBehavior=replace`, { method: "PUT", headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" }, body: data as unknown as BodyInit });
    const item = await r.json();
    if (!r.ok) throw new Error(`OneDrive refused the file (${r.status}): ${item.error?.message ?? ""}`);
    return item.id as string;
  }
  // Large file: an upload session, in chunks.
  const session = await graph(`${path}:/createUploadSession`, "POST", { item: { "@microsoft.graph.conflictBehavior": "replace" } });
  const CHUNK = 5 * 1024 * 1024;
  let item: any;
  for (let at = 0; at < data.byteLength; at += CHUNK) {
    const part = data.slice(at, at + CHUNK);
    const r = await fetch(session.uploadUrl, { method: "PUT", headers: { "Content-Length": String(part.byteLength), "Content-Range": `bytes ${at}-${at + part.byteLength - 1}/${data.byteLength}` }, body: part });
    item = await r.json();
    if (!r.ok) throw new Error(`OneDrive upload failed (${r.status})`);
  }
  return item.id as string;
}

/** The PDF of a PowerPoint file, converted by Microsoft. */
export async function pptxToPdf(data: Uint8Array): Promise<Buffer> {
  const hash = createHash("sha1").update(data).digest("hex");
  if (!cache.has(hash)) {
    const run = (async () => {
      const id = await upload(`${hash}.pptx`, data);
      try {
        // The conversion can lag the upload a little.
        for (let attempt = 0; ; attempt++) {
          const r = await request("graph", `/me/drive/items/${id}/content?format=pdf`);
          if (r.ok) return Buffer.from(await r.arrayBuffer());
          if (attempt >= 3 || ![404, 406, 409, 423, 500, 503].includes(r.status)) throw new Error(`Microsoft could not convert the PowerPoint to PDF (${r.status}): ${(await r.text()).slice(0, 200)}`);
          await new Promise((res) => setTimeout(res, 1500 * (attempt + 1)));
        }
      } finally {
        void request("graph", `/me/drive/items/${id}`, { method: "DELETE" }).catch(() => undefined);
      }
    })();
    cache.set(hash, run);
    run.catch(() => cache.delete(hash));
    if (cache.size > 30) cache.delete(cache.keys().next().value!);
  }
  return cache.get(hash)!;
}

/** Pictures of slides (1-based numbers; all when none given). */
export async function slidePictures(data: Uint8Array, slides?: number[], width = 1280) {
  const pdf = await pptxToPdf(data);
  return renderPdf(new Uint8Array(pdf), slides, width);
}
