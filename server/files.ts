/**
 * Files the server works on: what the user sends (images, Excel, CSV, PowerPoint) and what it makes
 * (screenshots, generated decks). Kept on disk in .local/files (the browser keeps its own copy and sends a
 * file again when the server no longer has it), cleaned after a few days. No other storage.
 */
import { mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { FileRef } from "@shared/types";
import { DATA_DIR } from "./settings";

const DIR = join(DATA_DIR, "files");
const KEEP_DAYS = 7;
const ID = /^[a-z0-9-]{8,64}$/i;

export async function saveFile(data: ArrayBuffer | Uint8Array, name: string, type: string, id: string = crypto.randomUUID()): Promise<FileRef> {
  if (!ID.test(id)) throw new Error("Bad file id");
  await mkdir(DIR, { recursive: true });
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const ref: FileRef = { id, name: name.replace(/[\\/]/g, "_").slice(0, 200), type: type || "application/octet-stream", size: bytes.byteLength };
  await writeFile(join(DIR, id), bytes);
  await writeFile(join(DIR, `${id}.json`), JSON.stringify(ref));
  return ref;
}

export async function fileInfo(id: string): Promise<FileRef | undefined> {
  if (!ID.test(id)) return undefined;
  return readFile(join(DIR, `${id}.json`), "utf8").then(JSON.parse).catch(() => undefined);
}

export async function readStoredFile(id: string): Promise<{ ref: FileRef; data: Uint8Array }> {
  const ref = await fileInfo(id);
  if (!ref) throw new Error(`File ${id} is no longer on the server: send it again.`);
  return { ref, data: new Uint8Array(await readFile(join(DIR, id))) };
}

/** Removes files older than a few days (at start, then daily). */
export async function cleanFiles() {
  const names = await readdir(DIR).catch(() => [] as string[]);
  const limit = Date.now() - KEEP_DAYS * 86_400_000;
  for (const n of names) {
    const s = await stat(join(DIR, n)).catch(() => undefined);
    if (s && s.mtimeMs < limit) await rm(join(DIR, n)).catch(() => undefined);
  }
}
