/**
 * Projects: a name and a Confluence folder. The pages of the folders of the projects chosen in the chat
 * are the context of a conversation: given whole at its start up to a size, the others by their title
 * (the AI reads them with read_project_page). Pages are read once and kept on disk (.local/projects):
 * "Refresh" (Settings › Confluence, the chat's projects menu) reads them again.
 */
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import { folderUrl, listPages, readPageMarkdown } from "./confluence";
import { DATA_DIR, projects } from "./settings";

export type ProjectPage = { project: string; id: string; title: string; markdown: string };

const MAX_CHARS = 80_000;
const MAX_PAGES = 60;
type Kept = { at: string; folderId: string; url?: string; pages: Omit<ProjectPage, "project">[] };
const DIR = join(DATA_DIR, "projects");
const memory = new Map<string, Promise<Kept>>();
const fileOf = (folderId: string) => join(DIR, `${folderId.replace(/[^\w-]/g, "_")}.json`);

/** Reads a folder's pages from Confluence and keeps them. */
async function fetchFolder(folderId: string): Promise<Kept> {
  const list = (await listPages(folderId)).slice(0, MAX_PAGES);
  const read = await Promise.all(list.map((p) => readPageMarkdown(p.id).then((x) => ({ id: p.id, title: x.title, markdown: x.markdown }), () => undefined)));
  const kept: Kept = { at: new Date().toISOString(), folderId, url: await folderUrl(folderId).catch(() => undefined), pages: read.filter((x) => !!x) };
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

const folderPages = async (projectName: string, folderId: string): Promise<ProjectPage[]> => (await folder(folderId)).pages.map((p) => ({ ...p, project: projectName }));

/** What is kept of each project: pages, when they were read, the folder's address in Confluence. */
export async function projectsStatus(refresh?: string) {
  const all = await projects();
  return Promise.all(
    all.map(async (p) => {
      if (!p.folderId) return { id: p.id, pages: 0 };
      try {
        const k = await folder(p.folderId, p.id === refresh);
        return { id: p.id, pages: k.pages.length, at: k.at, url: k.url };
      } catch (e) {
        return { id: p.id, pages: 0, error: e instanceof Error ? e.message : String(e) };
      }
    }),
  );
}

/** The pages of the chosen projects (errors said, not thrown: a conversation starts anyway). */
export async function projectPages(ids: string[]) {
  const all = await projects();
  const chosen = all.filter((p) => ids.includes(p.id));
  const pages: ProjectPage[] = [];
  const errors: string[] = [];
  for (const p of chosen) await folderPages(p.name, p.folderId).then((x) => pages.push(...x), (e) => errors.push(`${p.name}: ${e instanceof Error ? e.message : e}`));
  return { pages, errors, names: chosen.map((p) => p.name) };
}

/** The context text of a conversation: the pages whole while they fit, then their titles. */
export async function projectContext(ids: string[]) {
  if (!ids.length) return undefined;
  const { pages, errors, names } = await projectPages(ids);
  if (!pages.length && !errors.length) return undefined;
  let used = 0;
  const whole: string[] = [];
  const listed: string[] = [];
  for (const p of pages) {
    const text = `### ${p.title} (project ${p.project}, page ${p.id})\n${p.markdown}`;
    if (used + text.length <= MAX_CHARS) (whole.push(text), (used += text.length));
    else listed.push(`- "${p.title}" (page ${p.id}, project ${p.project})`);
  }
  return [
    `Context of the projects ${names.map((n) => `"${n}"`).join(", ")} (Confluence pages: the company's rules, definitions and vocabulary — use them):`,
    ...whole,
    listed.length ? `Other pages of these projects (read them with read_project_page when useful):\n${listed.join("\n")}` : "",
    errors.length ? `Pages not readable: ${errors.join("; ")}` : "",
  ]
    .filter(Boolean)
    .join("\n\n");
}
