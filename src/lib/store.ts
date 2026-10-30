/**
 * Local storage (IndexedDB through idb-keyval): one array per collection, with a React hook. This is
 * the library (only in this browser: each entry can be downloaded as JSON and imported again).
 */
import { useEffect, useState } from "react";
import { get, set } from "idb-keyval";
import type { ModelInfo } from "@/services/model";
import type { Field, FormSection } from "@/features/excel/schema";
import { upgradeSlide } from "@/features/presenter/slides/upgrade";

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
  schema?: import("@/services/report-schema").ReportSchema;
  /** Builder: reports given to the AI as examples (their schemas). */
  exampleIds?: string[];
  /** Viewer: other semantic models the AI may query (DAX) besides the report's own. */
  datasetIds?: string[];
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
export type ResearchEntry = import("@/features/data-scientist/research").ResearchEntry;

/** The user's opinion of an answer or of a whole conversation (optional), with an optional note. */
export type Rating = { score: "good" | "bad"; note?: string; at: number };

/**
 * `viewState`: the report as the user left it in this conversation (filters, slicers, page: a bookmark
 * state). `rating`: of the whole conversation; `ratings`: of answers, by message id.
 */
export type Conversation = { id: string; scope: string; title: string; messages: unknown[]; updatedAt: number; viewState?: string; rating?: Rating; ratings?: Record<string, Rating> };
/** `model`: the planner's (and of the one-call features); `implementerModel`: the implementer's (default: `model`). `maxSteps`: tool turns of the implementer (the planner gets half). `configLink`: the SharePoint JSON file holding the app's options (read at start). */
export type Settings = { aiKey?: string; model?: string; implementerModel?: string; maxSteps?: number; configLink?: string; confluence?: { url: string; key: string; folderId: string } };

/** A slide deck (its JSON: see src/features/presenter/slides) and the models and reports it is written from. */
export type DeckEntry = { id: string; name: string; deck: import("@/features/presenter/slides").Deck; updatedAt: number; sources?: { datasetIds: string[]; reportIds: string[]; schemaIds?: string[] } };

/** The guide of a visual of a report (Trainer › guide of the clicked visual), in markdown. */
export type GuideEntry = { id: string; reportId: string; page: string; visual: string; title: string; content: string; updatedAt: number };

type Collections = { contexts: ContextEntry; reports: ReportEntry; datasets: DatasetEntry; schemas: SchemaEntry; researches: ResearchEntry; decks: DeckEntry; conversations: Conversation; guides: GuideEntry };
export type Name = keyof Collections;

const cache = new Map<Name, unknown[]>();
const loading = new Map<Name, Promise<void>>();
const listeners = new Map<Name, Set<() => void>>();

const list = (v: unknown) => (Array.isArray(v) ? v : []);

/**
 * An entry as this version reads it: the lists an older version may have left out are there (empty),
 * so no screen breaks on an old library (e.g. a report without contextIds, a model without relationships).
 */
function normalize(name: Name, item: any) {
  if (!item || typeof item !== "object") return item;
  const withContexts = { ...item, contextIds: list(item.contextIds) };
  switch (name) {
    case "reports":
      return { ...withContexts, ...(item.schema && { schema: { ...item.schema, pages: list(item.schema.pages), visuals: list(item.schema.visuals).map((v: any) => ({ ...v, fields: v?.fields ?? {} })), filters: list(item.schema.filters), measures: list(item.schema.measures), settings: item.schema.settings ?? {} } }) };
    case "datasets":
      return {
        ...withContexts,
        ...(item.info && {
          info: { ...item.info, tables: list(item.info.tables).map((t: any) => ({ ...t, columns: list(t?.columns), measures: list(t?.measures), hierarchies: list(t?.hierarchies) })), relationships: list(item.info.relationships), roles: list(item.info.roles), warnings: list(item.info.warnings) },
        }),
      };
    case "schemas":
      return { ...withContexts, fields: list(item.fields) };
    case "conversations":
      return { ...item, messages: list(item.messages) };
    case "decks":
      return item.deck ? { ...item, deck: { ...item.deck, slides: list(item.deck.slides).map(upgradeSlide) } } : item;
    default:
      return item;
  }
}

export async function all<N extends Name>(name: N): Promise<Collections[N][]> {
  // One read of IndexedDB per collection: concurrent first reads used to overwrite a newer write
  // with the stale stored list (an added entry then vanished from the screen and from later writes).
  if (!loading.has(name)) loading.set(name, get(name).then((items) => void (cache.has(name) || cache.set(name, list(items).map((i) => normalize(name, i))))));
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

export const put = <N extends Name>(name: N, given: Collections[N]) =>
  queued(name, async () => {
    // An imported entry may come from an older version.
    const item = normalize(name, given) as Collections[N];
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

/** Whether a collection has been read from IndexedDB (before that, useCollection gives an empty list). */
export function useLoaded(name: Name) {
  const [loaded, setLoaded] = useState(() => cache.has(name));
  useEffect(() => {
    let live = true;
    if (!loaded) void all(name).then(() => live && setLoaded(true));
    return () => void (live = false);
  }, [name, loaded]);
  return loaded;
}

/** Asks the browser not to evict this site's data when the disk is full (the library vanished twice). */
export const persist = () => void navigator.storage?.persist?.().catch(() => undefined);

export const settings = {
  get: async (): Promise<Settings> => (await get("settings")) ?? {},
  set: async (values: Settings) => set("settings", { ...(await settings.get()), ...values }),
};
