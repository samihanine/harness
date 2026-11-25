/**
 * Pictures of a PowerPoint's slides, as PowerPoint draws them: the file is put in the user's OneDrive
 * (a hidden "Hub previews" folder), converted to PDF by Microsoft (Graph `?format=pdf`), removed, and the
 * PDF's pages drawn by pdf.js in the headless browser. Only Microsoft is called. Each slide's picture is kept by
 * a fingerprint of what it shows: after a change only the slides that changed are drawn again, from a light copy
 * of the deck holding just them (a deck full of pictures is long to send and convert whole).
 */
import { createHash } from "node:crypto";
import { graph, request } from "../ms/api";
import { renderPdf } from "../capture/browser";
import { Pptx } from "./ooxml";

const cache = new Map<string, Promise<Buffer>>();
const FOLDER = "Hub previews";

/** A Graph call again after the wait Microsoft asks for when it is busy (429, 503), a few times. */
async function patient(call: () => Promise<Response>) {
  for (let attempt = 0; ; attempt++) {
    const r = await call();
    if (![429, 503].includes(r.status) || attempt >= 4) return r;
    const wait = Math.min(30, Number(r.headers.get("Retry-After")) || 2 ** attempt * 2);
    await new Promise((res) => setTimeout(res, wait * 1000));
  }
}

async function upload(name: string, data: Uint8Array) {
  const path = `/me/drive/root:/${encodeURIComponent(FOLDER)}/${encodeURIComponent(name)}`;
  if (data.byteLength < 4 * 1024 * 1024) {
    const r = await patient(() => request("graph", `${path}:/content?@microsoft.graph.conflictBehavior=replace`, { method: "PUT", headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" }, body: data as unknown as BodyInit }));
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
          const r = await patient(() => request("graph", `/me/drive/items/${id}/content?format=pdf`));
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

/** Pictures of slides already drawn, by fingerprint and width (the latest few hundred). */
const pictures = new Map<string, Buffer>();
const remember = (key: string, png: Buffer) => {
  pictures.set(key, png);
  if (pictures.size > 400) pictures.delete(pictures.keys().next().value!);
};

/** Pictures of slides (1-based numbers; all when none given): drawn only for the slides not seen as they are. */
export async function slidePictures(data: Uint8Array, slides?: number[], width = 1280) {
  const deck = await Pptx.load(data);
  const count = (await deck.slideFiles()).length;
  const wanted = (slides?.length ? slides : Array.from({ length: count }, (_, i) => i + 1)).filter((n) => n >= 1 && n <= count);
  const keys = new Map<number, string>();
  for (const n of new Set(wanted)) keys.set(n, `${await deck.slideHash(n)}@${width}`);
  const missing = [...keys.keys()].filter((n) => !pictures.has(keys.get(n)!)).sort((a, b) => a - b);
  if (missing.length) {
    const whole = missing.length === count;
    const file = whole ? data : await (await Pptx.load(data)).only(missing);
    const pngs = await renderPdf(new Uint8Array(await pptxToPdf(file)), whole ? missing : missing.map((_, i) => i + 1), width);
    missing.forEach((n, i) => pngs[i] && remember(keys.get(n)!, pngs[i]));
  }
  return wanted.map((n) => pictures.get(keys.get(n)!)).filter((x): x is Buffer => !!x);
}
