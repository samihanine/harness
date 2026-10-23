/**
 * Local storage (IndexedDB through idb-keyval): one array per collection, with a React hook. This is
 * the library (only in this browser: each entry can be downloaded as JSON and imported again).
 */
import { useEffect, useState } from "react";
import { get, set } from "idb-keyval";
import type { ModelInfo } from "./model";
import type { Field, FormSection } from "./schema";

/** What the signed-in user may do: edit (definition readable and writable), build (DAX queries), view. */
export type Access = "edit" | "build" | "view";

/** A markdown file given to the AI with the entries it is linked to. */
export type ContextEntry = { id: string; name: string; text: string; updatedAt: number };

export type ReportEntry = {
  id: string;
  groupId?: string;
  /** Report of a Power BI app (view only). */
  appId?: string;
  name: string;
  url: string;
  datasetId: string;
  datasetGroupId?: string;
  access?: Access;
  /** Why the definition is not readable (e.g. PBIR-Legacy format, no edit right). */
  accessNote?: string;
  contextIds: string[];
  /** The name was changed in the app: reading the report again keeps it. */
  renamed?: boolean;
  /** Its descriptive schema (pages, visuals, fields), kept: made again with "Regenerate the schema". */
  schema?: import("./report-schema").ReportSchema;
  /** Builder: reports given to the AI as examples (their schemas). */
  exampleIds?: string[];
};

export type DatasetEntry = {
  id: string;
  groupId?: string;
  name: string;
  /** Model structure (tables, columns, measures, relationships, sample values). */
  info?: ModelInfo;
  access?: Access;
  /** Why its structure could not be read. */
  error?: string;
  contextIds: string[];
  /** The name was changed in the app: reading the model again keeps it. */
  renamed?: boolean;
};

/** An Excel table on SharePoint / OneDrive, as found from its link. */
export type ExcelRef = { link: string; driveId: string; itemId: string; name: string; table: string; fileUrl: string; webUrl?: string };
/** A Microsoft Form (its id holds the tenant and the owner). */
export type FormRef = { url: string; id: string; title: string };

/**
 * A schema: an Excel table (required), the Microsoft Form filling it (optional) and the definition of
 * its columns. `builtIn` schemas (the report viewer's guides) have fixed fields: only their
 * Excel file can be set.
 */
export type SchemaEntry = {
  id: string;
  title: string;
  builtIn?: "guides";
  excel?: ExcelRef;
  form?: FormRef;
  fields: Field[];
  /** Sections of the form (and where each goes when done); none: one page. */
  sections?: FormSection[];
  /** Folder receiving uploaded files (chosen by the user; required to upload). */
  filesFolder?: { link: string; driveId: string; itemId: string; name: string };
  /** Named views of the Excel editor. */
  views?: { name: string; view: Record<string, unknown> }[];
  contextIds: string[];
};

/** A research (kept in this browser): see research.ts. */
export type ResearchEntry = import("./research").ResearchEntry;

/** `viewState`: the report as the user left it in this conversation (filters, slicers, page: a bookmark state). */
export type Conversation = { id: string; scope: string; title: string; messages: unknown[]; updatedAt: number; viewState?: string };
/** `configLink`: the SharePoint JSON file holding the app's options (read at start). */
export type Settings = { aiKey?: string; model?: string; configLink?: string };

/** A slide deck (its JSON: see src/slides). */
export type DeckEntry = { id: string; name: string; deck: import("../slides").Deck; updatedAt: number; sources?: { datasetIds: string[]; reportIds: string[]; schemaIds: string[] } };

type Collections = { contexts: ContextEntry; reports: ReportEntry; datasets: DatasetEntry; schemas: SchemaEntry; researches: ResearchEntry; decks: DeckEntry; conversations: Conversation };
export type Name = keyof Collections;

const cache = new Map<Name, unknown[]>();
const loading = new Map<Name, Promise<void>>();
const listeners = new Map<Name, Set<() => void>>();

export async function all<N extends Name>(name: N): Promise<Collections[N][]> {
  // One read of IndexedDB per collection: concurrent first reads used to overwrite a newer write
  // with the stale stored list (an added entry then vanished from the screen and from later writes).
  if (!loading.has(name)) loading.set(name, get(name).then((items) => void (cache.has(name) || cache.set(name, items ?? []))));
  await loading.get(name);
  return cache.get(name) as Collections[N][];
}

/** Writes one at a time per collection (each starts from the result of the previous one). */
const queues = new Map<Name, Promise<unknown>>();
function queued<T>(name: Name, task: () => Promise<T>) {
  const run = (queues.get(name) ?? Promise.resolve()).then(task);
  queues.set(name, run.catch(() => undefined));
  return run;
}

async function write<N extends Name>(name: N, items: Collections[N][]) {
  cache.set(name, items);
  listeners.get(name)?.forEach((l) => l());
  await set(name, items);
}

export const put = <N extends Name>(name: N, item: Collections[N]) =>
  queued(name, async () => {
    const items = await all(name);
    const index = items.findIndex((i) => i.id === item.id);
    // Kept in place (lists keep their order when an entry is updated).
    await write(name, index >= 0 ? items.map((i) => (i.id === item.id ? item : i)) : [...items, item]);
    return item;
  });

export const patch = <N extends Name>(name: N, id: string, values: Partial<Collections[N]>) =>
  queued(name, async () => write(name, (await all(name)).map((i) => (i.id === id ? { ...i, ...values } : i))));

export const remove = (name: Name, id: string) => queued(name, async () => write(name, (await all(name)).filter((i) => i.id !== id) as never));

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

/** Asks the browser not to evict this site's data when the disk is full (the library vanished twice). */
export const persist = () => void navigator.storage?.persist?.().catch(() => undefined);

export const settings = {
  get: async (): Promise<Settings> => (await get("settings")) ?? {},
  set: async (values: Settings) => set("settings", { ...(await settings.get()), ...values }),
};
