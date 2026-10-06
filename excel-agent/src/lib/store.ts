/**
 * Local persistence: schemas and file links in localStorage (small JSON), file and folder
 * handles in IndexedDB (localStorage cannot hold them).
 */
import { useSyncExternalStore } from "react";
import { createStore, del, get, set } from "idb-keyval";
import type { Schema } from "./schema";

/** A local Excel file edited with a schema. */
export type FileLink = {
  id: string;
  name: string;
  schemaId: string;
  fileName: string;
  imagesFolder?: string;
  openedAt: number;
};

function localList<T extends { id: string }>(key: string) {
  const listeners = new Set<() => void>();
  let cache: T[] | undefined;
  const read = () => (cache ??= JSON.parse(localStorage.getItem(key) ?? "[]") as T[]);
  const write = (items: T[]) => {
    cache = items;
    localStorage.setItem(key, JSON.stringify(items));
    listeners.forEach((l) => l());
  };
  window.addEventListener("storage", (e) => {
    if (e.key !== key) return;
    cache = undefined;
    listeners.forEach((l) => l());
  });
  return {
    all: read,
    get: (id: string) => read().find((item) => item.id === id),
    put: (item: T) => write(read().some((i) => i.id === item.id) ? read().map((i) => (i.id === item.id ? item : i)) : [...read(), item]),
    remove: (id: string) => write(read().filter((i) => i.id !== id)),
    use: () =>
      useSyncExternalStore(
        (listener) => {
          listeners.add(listener);
          return () => listeners.delete(listener);
        },
        read,
      ),
  };
}

export const schemas = localList<Schema>("excel-agent:schemas");
export const links = localList<FileLink>("excel-agent:files");

const handlesDb = createStore("excel-agent", "handles");
export const handles = {
  file: (id: string) => get<FileSystemFileHandle>(`file:${id}`, handlesDb),
  images: (id: string) => get<FileSystemDirectoryHandle>(`images:${id}`, handlesDb),
  setFile: (id: string, handle: FileSystemFileHandle) => set(`file:${id}`, handle, handlesDb),
  setImages: (id: string, handle: FileSystemDirectoryHandle | undefined) =>
    handle ? set(`images:${id}`, handle, handlesDb) : del(`images:${id}`, handlesDb),
  remove: async (id: string) => {
    await del(`file:${id}`, handlesDb);
    await del(`images:${id}`, handlesDb);
  },
};

/** Read/write permission on a stored handle; asking requires a user gesture. */
export async function permission(handle: FileSystemHandle, ask: boolean) {
  const mode = { mode: "readwrite" as const };
  if ((await handle.queryPermission(mode)) === "granted") return true;
  return ask && (await handle.requestPermission(mode)) === "granted";
}
