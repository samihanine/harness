/**
 * The catalog of the reports and models everyone works with, kept in Confluence: one folder for the reports
 * (CONFLUENCE_REPORT_FOLDER_ID), one for the models (CONFLUENCE_DATASET_FOLDER_ID); each page is one item —
 * its title the name, its content the Power BI link and a description (what it is for: the Trainer uses it to
 * pick the right report). Read once and kept (.local/library-*.json), "Refresh" reads it again. Changed from
 * Settings only when CONFLUENCE_CATALOG_EDIT is on. Structures are not kept here: each browser reads and keeps
 * them the first time it opens an item.
 */
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import { datasetUrl, parseLink, reportUrl } from "@shared/links";
import type { LibraryEntry, LibraryKind } from "@shared/types";
import { createPage, deletePage, folderChildren, readPageMarkdown, writePage, readPage } from "./confluence";
import { catalogEditable, DATA_DIR, settings } from "./settings";
import { resolveDataset, resolveReport } from "./pbi/resolve";

const fileOf = (kind: LibraryKind) => join(DATA_DIR, `library-${kind}.json`);
const memory = new Map<LibraryKind, Promise<LibraryEntry[]>>();

async function folderOf(kind: LibraryKind) {
  const c = (await settings()).confluence;
  const id = kind === "report" ? c.reportsFolderId : c.datasetsFolderId;
  if (!id) throw new Error(`No Confluence folder for the ${kind === "report" ? "reports" : "models"} (${kind === "report" ? "CONFLUENCE_REPORT_FOLDER_ID" : "CONFLUENCE_DATASET_FOLDER_ID"}, Settings › Confluence).`);
  return id;
}

/** A catalog page's content: the link on its own line, then the description. */
const pageText = (link: string, description = "") => `Power BI: ${link}\n\n${description.trim()}`.trim();

/** An item from its page: the first Power BI link is its address, the rest its description. */
function entryOf(kind: LibraryKind, page: { id: string; title: string; markdown: string; webUrl?: string }): LibraryEntry {
  const link = page.markdown.match(/https:\/\/app\.powerbi\.com[^\s)>\]"]*/)?.[0] ?? "";
  const ids = parseLink(link);
  const description = page.markdown
    .replace(/^.*https:\/\/app\.powerbi\.com\S*.*$/m, "")
    .replace(/^\s*(\*\*)?Power BI:?(\*\*)?\s*$/gim, "")
    .trim();
  const id = kind === "report" ? (ids?.report ?? ids?.bare) : (ids?.dataset ?? ids?.bare);
  return { pageId: page.id, name: page.title, url: link, description, pageUrl: page.webUrl, ...(id ? { id, groupId: ids?.groupId && ids.groupId !== "me" ? ids.groupId : undefined, appId: ids?.appId } : { error: "No Power BI link on the page." }) };
}

async function read(kind: LibraryKind): Promise<LibraryEntry[]> {
  const pages = (await folderChildren(await folderOf(kind))).filter((p) => p.type === "page");
  const out: LibraryEntry[] = [];
  // A few at a time: Confluence limits the calls per second.
  for (let i = 0; i < pages.length; i += 8)
    out.push(...(await Promise.all(pages.slice(i, i + 8).map((p) => readPageMarkdown(p.id).then((x) => entryOf(kind, x), (e) => ({ pageId: p.id, name: p.title, url: "", description: "", error: e instanceof Error ? e.message : String(e) }))))));
  await mkdir(DATA_DIR, { recursive: true });
  await Bun.write(fileOf(kind), JSON.stringify(out));
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

/** The catalog: kept (memory, then disk), else read from Confluence; `fresh` reads it again. */
export function library(kind: LibraryKind, fresh = false): Promise<LibraryEntry[]> {
  if (!fresh && memory.has(kind)) return memory.get(kind)!;
  const p = (async () => {
    if (!fresh) {
      const f = Bun.file(fileOf(kind));
      if (await f.exists()) return (await f.json()) as LibraryEntry[];
    }
    return read(kind);
  })();
  memory.set(kind, p);
  p.catch(() => memory.delete(kind));
  return p;
}

const editable = () => {
  if (!catalogEditable()) throw new Error("The catalog is read-only here (CONFLUENCE_CATALOG_EDIT is off).");
};
/** After a change: read again from Confluence (the copy kept on disk is stale). */
const changed = (kind: LibraryKind) => (memory.delete(kind), void Bun.file(fileOf(kind)).delete().catch(() => undefined));

/** Adds an item from any Power BI link (a report's link in the models' catalog adds its model). */
export async function addToLibrary(kind: LibraryKind, b: { link: string; description?: string; name?: string }) {
  editable();
  const l = parseLink(b.link);
  if (!l) throw new Error("This is not a Power BI link.");
  let name: string;
  let link: string;
  if (kind === "report") {
    const id = l.report ?? l.bare;
    if (!id) throw new Error("This is the link of a semantic model, not of a report.");
    const r = await resolveReport({ id, groupId: l.groupId, appId: l.appId });
    name = r.name;
    link = reportUrl(r);
  } else if (l.dataset || (l.bare && !l.report)) {
    const d = await resolveDataset({ id: (l.dataset ?? l.bare)!, groupId: l.groupId });
    name = d.name;
    link = datasetUrl(d);
  } else {
    // A report's link: its model, in the model's own workspace (an app's report sits elsewhere).
    const r = await resolveReport({ id: l.report!, groupId: l.groupId, appId: l.appId });
    if (!r.datasetId) throw new Error(`The model of "${r.name}" cannot be read with your rights.`);
    const d = await resolveDataset({ id: r.datasetId, groupId: r.datasetGroupId });
    name = d.name;
    link = datasetUrl(d);
  }
  const current = await library(kind, true);
  if (current.some((e) => e.url === link)) throw new Error(`"${name}" is already in the catalog.`);
  // Page titles are unique in a Confluence space: a taken name gets the kind after it.
  const title = b.name?.trim() || name;
  const folder = await folderOf(kind);
  const page = await createPage(folder, title, pageText(link, b.description)).catch((e) => {
    if (/title|exists/i.test(String(e))) return createPage(folder, `${title} (${kind === "report" ? "report" : "model"})`, pageText(link, b.description));
    throw e;
  });
  changed(kind);
  return entryOf(kind, { id: page.id, title: page.title, markdown: pageText(link, b.description), webUrl: page.webUrl });
}

export async function changeLibrary(kind: LibraryKind, pageId: string, b: { name?: string; description?: string }) {
  editable();
  const page = await readPage(pageId);
  const entry = entryOf(kind, { ...page, markdown: (await readPageMarkdown(pageId)).markdown });
  const saved = await writePage({ id: page.id, title: b.name?.trim() || page.title, version: page.version }, pageText(entry.url, b.description ?? entry.description));
  changed(kind);
  return entryOf(kind, { id: saved.id, title: saved.title, markdown: pageText(entry.url, b.description ?? entry.description), webUrl: saved.webUrl });
}

export async function removeFromLibrary(kind: LibraryKind, pageId: string) {
  editable();
  await deletePage(pageId);
  changed(kind);
  return { ok: true };
}

/** The reports of the catalog as told to the AI (names and what they are for): to pick the right one. */
export async function libraryText(kind: LibraryKind) {
  const list = await library(kind).catch(() => []);
  return list.filter((e) => e.id).map((e) => `- "${e.name}" (id ${e.id})${e.description ? `: ${e.description.replace(/\s+/g, " ").slice(0, 300)}` : ""}`).join("\n");
}
