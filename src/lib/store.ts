/**
 * What the browser keeps (IndexedDB): the items of each screen (reports, models, analyses, decks,
 * PowerPoints), the conversations, and the files sent to the server (sent again when the server no
 * longer has them). Everything can be exported and imported from Settings. Context links (Confluence
 * pages linked to an item) and small choices live in localStorage.
 */
import { createStore, del, entries, get, set } from "idb-keyval";
import { useEffect, useState, useSyncExternalStore } from "react";
import type { Analysis } from "@shared/analysis";
import type { ChatItem, ContextLink, DatasetRef, FileRef, ReportRef, ScreenId } from "@shared/types";
import { upload } from "./api";

const db = createStore("hub-v2", "items");

export type Stored = {
  report: ReportRef & { addedAt: number };
  dataset: DatasetRef & { addedAt: number };
  analysis: Analysis & { id: string; datasetIds: string[]; files: FileRef[] };
  deck: { id: string; title: string; markdown: string; updatedAt: string };
  pptx: { id: string; name: string; fileId: string; history: string[]; updatedAt: string };
  conversation: { id: string; screen: ScreenId; item: string; title: string; items: ChatItem[]; updatedAt: string };
};
export type Kind = keyof Stored;

const listeners = new Map<Kind, Set<() => void>>();
const cache = new Map<Kind, Stored[Kind][]>();
const notify = (k: Kind) => listeners.get(k)?.forEach((l) => l());

async function load<K extends Kind>(kind: K): Promise<Stored[K][]> {
  if (!cache.has(kind)) {
    const all = await entries(db);
    cache.set(kind, all.filter(([k]) => String(k).startsWith(`${kind}:`)).map(([, v]) => v as Stored[K]));
  }
  return cache.get(kind) as Stored[K][];
}

export async function list<K extends Kind>(kind: K) {
  return load(kind);
}

export async function save<K extends Kind>(kind: K, value: Stored[K]) {
  const items = (await load(kind)).filter((x) => (x as { id: string }).id !== (value as { id: string }).id);
  cache.set(kind, [...items, value]);
  await set(`${kind}:${(value as { id: string }).id}`, value, db);
  notify(kind);
}

export async function remove(kind: Kind, id: string) {
  cache.set(kind, (await load(kind)).filter((x) => (x as { id: string }).id !== id));
  await del(`${kind}:${id}`, db);
  notify(kind);
}

export async function item<K extends Kind>(kind: K, id: string) {
  return (await load(kind)).find((x) => (x as { id: string }).id === id) as Stored[K] | undefined;
}

/** The items of a kind, kept up to date. */
export function useItems<K extends Kind>(kind: K): Stored[K][] | undefined {
  const [items, setItems] = useState<Stored[K][]>();
  useEffect(() => {
    const update = () => void load(kind).then((x) => setItems([...x]));
    if (!listeners.has(kind)) listeners.set(kind, new Set());
    listeners.get(kind)!.add(update);
    update();
    return () => void listeners.get(kind)!.delete(update);
  }, [kind]);
  return items;
}

/* ---------------------------------- Files ---------------------------------- */

const blobs = createStore("hub-v2-files", "files");

/** Sends a file to the server and keeps it here (to send it again later). */
export async function keepFile(file: File | Blob, name: string): Promise<FileRef> {
  const ref = await upload(file, name);
  await set(ref.id, { ref, blob: file }, blobs);
  return ref;
}

/** Makes sure the server has these files (sent again from this browser when it lost them). */
export async function ensureFiles(files: FileRef[]) {
  await Promise.all(
    files.map(async (f) => {
      const r = await fetch(`/api/files/${f.id}`, { method: "HEAD" });
      if (r.ok) return;
      const kept = (await get(f.id, blobs)) as { blob: Blob } | undefined;
      if (kept) await upload(kept.blob, f.name, f.id);
    }),
  );
}

/** Keeps a server file in this browser too (e.g. a PowerPoint the AI changed). */
export async function keepServerFile(ref: FileRef) {
  if (await get(ref.id, blobs)) return;
  const blob = await (await fetch(`/api/files/${ref.id}`)).blob();
  await set(ref.id, { ref, blob }, blobs);
}

/* ------------------------------ Local choices ------------------------------ */

const local = new Set<() => void>();
export function readLocal<T>(key: string, fallback: T): T {
  try {
    const v = localStorage.getItem(`hub:${key}`);
    return v === null ? fallback : (JSON.parse(v) as T);
  } catch {
    return fallback;
  }
}
export function writeLocal(key: string, value: unknown) {
  try {
    localStorage.setItem(`hub:${key}`, JSON.stringify(value));
  } catch {}
  local.forEach((l) => l());
}
export function useLocal<T>(key: string, fallback: T): [T, (v: T) => void] {
  const value = useSyncExternalStore(
    (l) => (local.add(l), () => void local.delete(l)),
    () => localStorage.getItem(`hub:${key}`),
  );
  let parsed = fallback;
  try {
    if (value !== null) parsed = JSON.parse(value);
  } catch {}
  return [parsed, (v: T) => writeLocal(key, v)];
}

/** Confluence pages linked to an item as its context (as in v1: kept in this browser). */
export const useContexts = (itemKey: string) => useLocal<ContextLink[]>(`contexts:${itemKey}`, []);
export const contextsOf = (itemKey: string) => readLocal<ContextLink[]>(`contexts:${itemKey}`, []);

/** Everything of this browser as one JSON file (Settings › General), and back. */
export async function exportAll() {
  const all = await entries(db);
  const ls = Object.fromEntries(Object.keys(localStorage).filter((k) => k.startsWith("hub:")).map((k) => [k, localStorage.getItem(k)]));
  return new Blob([JSON.stringify({ version: 2, items: all, local: ls }, null, 1)], { type: "application/json" });
}
export async function importAll(file: File) {
  const data = JSON.parse(await file.text());
  for (const [k, v] of data.items ?? []) await set(k, v, db);
  for (const [k, v] of Object.entries(data.local ?? {})) if (k.startsWith("hub:") && typeof v === "string") localStorage.setItem(k, v);
  cache.clear();
  for (const k of listeners.keys()) notify(k);
}
