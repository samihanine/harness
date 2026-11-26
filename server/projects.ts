/**
 * Projects: the sub-folders of the Confluence projects folder (CONFLUENCE_PROJECTS_FOLDER_ID), each named as
 * its project. The pages of the projects chosen in the chat are given whole to the conversation at its start
 * (up to MAX_CHARS; past it, the rest by title, read with read_project_page). Pages are read once and kept on
 * disk (.local/projects): "Refresh" reads a project again.
 */
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import type { Project } from "@shared/types";
import { folderChildren, folderUrl, listPages, readPageMarkdown } from "./confluence";
import { DATA_DIR, settings } from "./settings";

export type ProjectPage = { project: string; id: string; title: string; markdown: string };

const MAX_CHARS = 150_000;
const MAX_PAGES = 200;

type Kept = { at: string; folderId: string; url?: string; pages: Omit<ProjectPage, "project">[] };
const DIR = join(DATA_DIR, "projects");
const memory = new Map<string, Promise<Kept>>();
const fileOf = (folderId: string) => join(DIR, `${folderId.replace(/[^\w-]/g, "_")}.json`);

/** Reads a folder's pages (all its descendants) from Confluence and keeps them. */
async function fetchFolder(folderId: string): Promise<Kept> {
  const list = (await listPages(folderId)).slice(0, MAX_PAGES);
  const read: (Kept["pages"][number] | undefined)[] = [];
  // A few at a time: Confluence limits the calls per second.
  for (let i = 0; i < list.length; i += 8) read.push(...(await Promise.all(list.slice(i, i + 8).map((p) => readPageMarkdown(p.id).then((x) => ({ id: p.id, title: x.title, markdown: x.markdown }), () => undefined)))));
  const kept: Kept = { at: new Date().toISOString(), folderId, url: await folderUrl(folderId).catch(() => undefined), pages: read.filter((x): x is Kept["pages"][number] => !!x) };
  await mkdir(DIR, { recursive: true });
  await Bun.write(fileOf(folderId), JSON.stringify(kept));
  return kept;
}

/** A folder's pages: kept (memory, then disk), else read once from Confluence; `fresh` reads them again. */
function folder(folderId: string, fresh = false): Promise<Kept> {
  if (!fresh && memory.has(folderId)) return memory.get(folderId)!;
  const p = (async () => {
    if (!fresh) {
      const f = Bun.file(fileOf(folderId));
      if (await f.exists()) return (await f.json()) as Kept;
    }
    return fetchFolder(folderId);
  })();
  memory.set(folderId, p);
  p.catch(() => memory.delete(folderId));
  return p;
}

/** The projects: the sub-folders of the projects folder (read once a session, then on "Refresh"). */
let listed: Promise<{ id: string; name: string }[]> | undefined;
async function projectFolders(fresh = false) {
  const root = (await settings()).confluence.projectsFolderId;
  if (!root) return [];
  if (fresh || !listed) {
    listed = folderChildren(root).then((c) => c.map((x) => ({ id: x.id, name: x.title })));
    listed.catch(() => (listed = undefined));
  }
  return listed;
}

/** The projects with what is kept of each (pages, when read, the folder's address); `refresh`: one read again ("*": the list too). */
export async function projectsList(refresh?: string): Promise<Project[]> {
  const all = await projectFolders(refresh === "*");
  return Promise.all(
    all.map(async (p) => {
      // Not read yet: said as such (reading every project's pages to list them would be slow).
      const kept = p.id === refresh || memory.has(p.id) || (await Bun.file(fileOf(p.id)).exists());
      if (!kept) return { id: p.id, name: p.name };
      try {
        const k = await folder(p.id, p.id === refresh);
        return { id: p.id, name: p.name, pages: k.pages.length, at: k.at, url: k.url };
      } catch (e) {
        return { id: p.id, name: p.name, error: e instanceof Error ? e.message : String(e) };
      }
    }),
  );
}

/** The pages of the chosen projects (errors said, not thrown: a conversation starts anyway). */
export async function projectPages(ids: string[]) {
  const chosen = (await projectFolders()).filter((p) => ids.includes(p.id));
  const pages: ProjectPage[] = [];
  const errors: string[] = [];
  for (const p of chosen) await folder(p.id).then((k) => pages.push(...k.pages.map((x) => ({ ...x, project: p.name }))), (e) => errors.push(`${p.name}: ${e instanceof Error ? e.message : e}`));
  return { pages, errors, names: chosen.map((p) => p.name) };
}

/** The context text of a conversation: every page whole while they fit, then the titles of the rest. */
export async function projectContext(ids: string[]) {
  if (!ids.length) return undefined;
  const { pages, errors, names } = await projectPages(ids);
  if (!pages.length && !errors.length) return undefined;
  let used = 0;
  const whole: string[] = [];
  const listedPages: string[] = [];
  for (const p of pages) {
    const text = `### ${p.title} (project ${p.project}, page ${p.id})\n${p.markdown}`;
    if (used + text.length <= MAX_CHARS) (whole.push(text), (used += text.length));
    else listedPages.push(`- "${p.title}" (page ${p.id}, project ${p.project})`);
  }
  return [
    `Context of the projects ${names.map((n) => `"${n}"`).join(", ")} (Confluence pages: the company's rules, definitions and vocabulary — use them):`,
    ...whole,
    listedPages.length ? `Other pages of these projects (read them with read_project_page when useful):\n${listedPages.join("\n")}` : "",
    errors.length ? `Pages not readable: ${errors.join("; ")}` : "",
  ]
    .filter(Boolean)
    .join("\n\n");
}
