/**
 * The team's catalog of reports and models (Confluence, read by the server), as the screens list it: the
 * items this user hid are left out (a choice of this browser). An item is opened on first use: its details
 * (model, rights) and structure are read then and kept in this browser — the user's access is checked only
 * when it opens.
 */
import { useEffect, useState } from "react";
import type { DatasetRef, LibraryEntry, LibraryKind, ReportRef } from "@shared/types";
import { del, get, patch, post } from "./api";
import { prepareDataset, prepareReport } from "./items";
import { item, save, useLocal } from "./store";

const loaded = new Map<LibraryKind, Promise<LibraryEntry[]>>();
const listeners = new Map<LibraryKind, Set<(e: LibraryEntry[]) => void>>();
const publish = (kind: LibraryKind, p: Promise<LibraryEntry[]>) => {
  loaded.set(kind, p);
  p.then((e) => listeners.get(kind)?.forEach((l) => l(e)), () => loaded.delete(kind));
  return p;
};

export function useLibrary(kind: LibraryKind) {
  const [entries, setEntries] = useState<LibraryEntry[]>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    if (!listeners.has(kind)) listeners.set(kind, new Set());
    listeners.get(kind)!.add(setEntries);
    (loaded.get(kind) ?? publish(kind, get<LibraryEntry[]>(`/library/${kind}`))).then(setEntries, (e) => setError(e instanceof Error ? e.message : String(e)));
    return () => void listeners.get(kind)!.delete(setEntries);
  }, [kind]);
  return { entries, error };
}

/** Reads the catalog again from Confluence. */
export const refreshLibrary = (kind: LibraryKind) => publish(kind, get<LibraryEntry[]>(`/library/${kind}?fresh=1`));

/** Changes of the catalog (Settings, when the catalog is editable): the list is read again after each. */
export const libraryEdit = {
  add: async (kind: LibraryKind, link: string, description?: string) => (await post(`/library/${kind}`, { link, description }), void publish(kind, get(`/library/${kind}`))),
  change: async (kind: LibraryKind, pageId: string, b: { name?: string; description?: string }) => (await patch(`/library/${kind}/${pageId}`, b), void publish(kind, get(`/library/${kind}`))),
  remove: async (kind: LibraryKind, pageId: string) => (await del(`/library/${kind}/${pageId}`), void publish(kind, get(`/library/${kind}`))),
};

/** The items this user hid (this browser only). */
export const useHidden = (kind: LibraryKind) => useLocal<string[]>(`hidden:${kind}`, []);

/** A report of the catalog opened: its details read (and kept) the first time; refused when the user cannot open it. */
export async function openReport(e: LibraryEntry, fresh = false): Promise<ReportRef> {
  const kept = await item("report", e.id!);
  if (kept && !fresh) return { ...kept, name: e.name };
  const q = new URLSearchParams({ id: e.id!, ...(e.groupId ? { groupId: e.groupId } : {}), ...(e.appId ? { appId: e.appId } : {}) });
  const r = await get<ReportRef>(`/powerbi/report?${q}`);
  const saved = { ...r, name: e.name, ...(kept && !fresh ? { structure: kept.structure, structureAt: kept.structureAt } : {}), addedAt: kept?.addedAt ?? Date.now() };
  await save("report", saved);
  prepareReport(saved);
  return saved;
}

/** A model of the catalog as the screens use it (its ids are in its link); its structure read in the background. */
export function datasetOf(e: LibraryEntry): DatasetRef {
  const d = { id: e.id!, groupId: e.groupId, name: e.name };
  prepareDataset(d);
  return d;
}
