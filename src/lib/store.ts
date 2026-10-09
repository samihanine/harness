/** Local storage (IndexedDB through idb-keyval): one array per collection, with a React hook. */
import { useEffect, useState } from "react";
import { get, set } from "idb-keyval";
import type { ModelInfo } from "./model";

export type ReportEntry = {
  id: string;
  groupId?: string;
  name: string;
  url: string;
  datasetId: string;
  datasetGroupId?: string;
  /** Free text given to the AI with this report. */
  context: string;
  /** What could be read of the report (pages, visuals, fields, filters), as text for the AI. */
  snapshot?: string;
  /** Can the signed-in user edit / clone it? */
  editable?: boolean;
  /** Excel files (ids of `excels`) holding the info links and the guides shown in the viewer. */
  infoExcelId?: string;
  guidesExcelId?: string;
};

export type DatasetEntry = {
  id: string;
  groupId?: string;
  name: string;
  context: string;
  /** Model structure: summary always given to the AI, details read by its model tools. */
  model?: string;
  info?: ModelInfo;
  /** Generated from this Excel file (dataset builder). */
  excelId?: string;
};

export type ExcelEntry = {
  id: string;
  name: string;
  link: string;
  driveId: string;
  itemId: string;
  /** Server-relative URL used by Power BI to read the file. */
  fileUrl: string;
  table: string;
  context: string;
  /** Folder receiving uploaded images (default: "<file> images" next to the file). */
  imagesFolder?: { link: string; driveId: string; itemId: string; name: string };
  /** Named views of the dataset builder (layout, search, filters, sort, kanban column). */
  savedViews?: { name: string; view: Record<string, unknown> }[];
};

export type Conversation = { id: string; scope: string; title: string; messages: unknown[]; updatedAt: number };
export type Settings = { aiKey?: string; model?: string };

type Collections = { reports: ReportEntry; datasets: DatasetEntry; excels: ExcelEntry; conversations: Conversation };
type Name = keyof Collections;

const cache = new Map<Name, unknown[]>();
const listeners = new Map<Name, Set<() => void>>();

export async function all<N extends Name>(name: N): Promise<Collections[N][]> {
  if (!cache.has(name)) cache.set(name, (await get(name)) ?? []);
  return cache.get(name) as Collections[N][];
}

async function write<N extends Name>(name: N, items: Collections[N][]) {
  cache.set(name, items);
  await set(name, items);
  listeners.get(name)?.forEach((l) => l());
}

export async function put<N extends Name>(name: N, item: Collections[N]) {
  const items = await all(name);
  await write(name, [...items.filter((i) => i.id !== item.id), item]);
  return item;
}

export async function patch<N extends Name>(name: N, id: string, values: Partial<Collections[N]>) {
  const items = await all(name);
  await write(name, items.map((i) => (i.id === id ? { ...i, ...values } : i)));
}

export async function remove(name: Name, id: string) {
  await write(name, (await all(name)).filter((i) => i.id !== id) as never);
}

export const find = async <N extends Name>(name: N, id?: string) => (id ? (await all(name)).find((i) => i.id === id) : undefined);

export function useCollection<N extends Name>(name: N) {
  const [items, setItems] = useState<Collections[N][]>((cache.get(name) as Collections[N][]) ?? []);
  useEffect(() => {
    const update = () => void all(name).then((list) => setItems([...list]));
    if (!listeners.has(name)) listeners.set(name, new Set());
    listeners.get(name)!.add(update);
    update();
    return () => void listeners.get(name)!.delete(update);
  }, [name]);
  return items;
}

export const settings = {
  get: async (): Promise<Settings> => (await get("settings")) ?? {},
  set: async (values: Settings) => set("settings", { ...(await settings.get()), ...values }),
};
