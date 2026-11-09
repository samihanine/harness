/**
 * The Presenter's decks: an entry in the library (name, where the file is, the sources the AI may use)
 * and the file itself, kept in this browser (IndexedDB) with its earlier versions (undo). The file
 * comes from this computer (opened and saved back in place where the browser allows it) or from
 * SharePoint / OneDrive (read and written through Graph). The template deck (Settings › Slides) is
 * kept the same way.
 */
import { del, get, set } from "idb-keyval";
import { graph, readItem, writeItem } from "@/services/ms";
import { shareId } from "@/services/excel";
import { PPTX_TYPE } from "./pptx/deck";

export type PresentationSource = { kind: "local"; fileName: string } | { kind: "sharepoint"; link: string; driveId: string; itemId: string; webUrl?: string; eTag?: string };
export type PresentationEntry = {
  id: string;
  name: string;
  source: PresentationSource;
  /** The copy here differs from the file (not saved to it yet). */
  unsaved?: boolean;
  updatedAt: number;
  sources?: { datasetIds: string[]; reportIds: string[] };
  contextIds: string[];
};
declare module "@/lib/store" {
  interface Collections {
    presentations: PresentationEntry;
  }
}

const fileKey = (id: string) => `pptx:${id}`;
const undoKey = (id: string) => `pptx-undo:${id}`;
const handleKey = (id: string) => `pptx-handle:${id}`;
const UNDO_STEPS = 15;

export const readDeckFile = async (id: string) => (await get<Uint8Array>(fileKey(id))) ?? undefined;

/** Keeps a new version of the deck here (the one before stays for undo). */
export async function keepDeckFile(id: string, bytes: Uint8Array) {
  const before = await get<Uint8Array>(fileKey(id));
  if (before) await set(undoKey(id), [before, ...((await get<Uint8Array[]>(undoKey(id))) ?? [])].slice(0, UNDO_STEPS));
  await set(fileKey(id), bytes);
}

/** Puts the version before the last change back; false when there is none. */
export async function undoDeckFile(id: string) {
  const [last, ...rest] = (await get<Uint8Array[]>(undoKey(id))) ?? [];
  if (!last) return false;
  await set(fileKey(id), last);
  await set(undoKey(id), rest);
  return true;
}

export const forgetDeck = (id: string) => Promise.all([del(fileKey(id)), del(undoKey(id)), del(handleKey(id))]);

/* ----------------------------------- Local ----------------------------------- */

type Handle = { name: string; getFile(): Promise<File>; createWritable(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>; requestPermission?(o: { mode: string }): Promise<string> };
const picker = () => (window as unknown as { showOpenFilePicker?: (o: unknown) => Promise<Handle[]> }).showOpenFilePicker;
const PPTX_PICKER = { types: [{ description: "PowerPoint", accept: { [PPTX_TYPE]: [".pptx"] } }], multiple: false };

/** A .pptx of this computer: through the file picker that can save back in place (Chrome, Edge), else a plain file input. */
export async function pickLocalFile(): Promise<{ file: File; handle?: Handle } | undefined> {
  const open = picker();
  if (open) {
    const [handle] = await open(PPTX_PICKER).catch(() => [] as Handle[]);
    return handle ? { file: await handle.getFile(), handle } : undefined;
  }
  return new Promise((resolve) => {
    const input = Object.assign(document.createElement("input"), { type: "file", accept: ".pptx" });
    input.onchange = () => resolve(input.files?.[0] ? { file: input.files[0] } : undefined);
    input.click();
  });
}

export const keepHandle = (id: string, handle: unknown) => set(handleKey(id), handle);

/** Writes the deck into its local file (in place when the browser kept its handle; else downloaded). */
export async function saveLocal(id: string, name: string, data: Blob) {
  const handle = await get<Handle>(handleKey(id));
  if (handle) {
    const ok = (await handle.requestPermission?.({ mode: "readwrite" })) ?? "granted";
    if (ok === "granted") {
      const w = await handle.createWritable();
      await w.write(data);
      await w.close();
      return `Saved in ${handle.name}.`;
    }
  }
  download(data, name.endsWith(".pptx") ? name : `${name}.pptx`);
  return "Downloaded (this browser cannot write the file in place).";
}

export function download(data: Blob, name: string) {
  const url = URL.createObjectURL(data);
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/* --------------------------------- SharePoint --------------------------------- */

/** A .pptx on SharePoint / OneDrive from its sharing link: its reference and its content. */
export async function readSharePoint(link: string) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,parentReference,webUrl`);
  if (!/\.pptx$/i.test(item.name)) throw new Error(`"${item.name}" is not a PowerPoint file (.pptx).`);
  const ref = { driveId: item.parentReference.driveId as string, itemId: item.id as string };
  const { data, eTag } = await readItem(ref);
  return { source: { kind: "sharepoint" as const, link, ...ref, webUrl: item.webUrl as string, eTag }, name: (item.name as string).replace(/\.pptx$/i, ""), data };
}

/** Writes the deck to its SharePoint file, only if nobody saved it meanwhile; returns its new version. */
export async function saveSharePoint(source: Extract<PresentationSource, { kind: "sharepoint" }>, data: Blob) {
  await writeItem({ driveId: source.driveId, itemId: source.itemId, data, eTag: source.eTag });
  const m = await graph(`/drives/${source.driveId}/items/${source.itemId}?$select=eTag`);
  return m.eTag as string;
}

/* ---------------------------------- Template ---------------------------------- */

const TEMPLATE = "presenter-template";
export type TemplateFile = { name: string; bytes: Uint8Array; at: number };
export const readTemplate = () => get<TemplateFile>(TEMPLATE);
export const keepTemplate = (t: TemplateFile | undefined) => (t ? set(TEMPLATE, t) : del(TEMPLATE));
