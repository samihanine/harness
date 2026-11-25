/**
 * Projects: a name and a Confluence folder. The pages of the folders of the projects chosen in the chat
 * are the context of a conversation: given whole at its start up to a size, the others by their title
 * (the AI reads them with read_project_page). Pages are read again after a few minutes.
 */
import { listPages, readPageMarkdown } from "./confluence";
import { projects } from "./settings";

export type ProjectPage = { project: string; id: string; title: string; markdown: string };

const MAX_CHARS = 80_000;
const MAX_PAGES = 60;
const cache = new Map<string, { at: number; pages: Promise<ProjectPage[]> }>();

async function folderPages(projectName: string, folderId: string): Promise<ProjectPage[]> {
  const key = folderId;
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < 10 * 60_000) return hit.pages;
  const pages = (async () => {
    const list = (await listPages(folderId)).slice(0, MAX_PAGES);
    const read = await Promise.all(list.map((p) => readPageMarkdown(p.id).then((x) => ({ project: projectName, id: p.id, title: x.title, markdown: x.markdown }), () => undefined)));
    return read.filter((x): x is ProjectPage => !!x);
  })();
  cache.set(key, { at: Date.now(), pages });
  pages.catch(() => cache.delete(key));
  return pages;
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
