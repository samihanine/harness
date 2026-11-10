/**
 * Confluence pages, read and written through the Confluence REST API (v1 content API: Cloud and Data
 * Center). The address, email, key and folder are set in Settings › Confluence and kept by the local
 * server (scripts/token-server.ts), in the system's credential store: the key never stays in the
 * browser, and every call goes through the server, which adds it (Confluence Cloud allows no call
 * from a browser page anyway).
 */
import { settings } from "@/lib/store";
import { fromStorage } from "@/lib/markdown/storage";
import { server } from "./ms";

/** The settings as the app sees them: never the key, only whether one is saved. */
export type ConfluenceConfig = { url: string; email?: string; folderId: string; hasKey: boolean };
export type ConfluencePage = { id: string; title: string; version: number; html: string; webUrl?: string };

let config: Promise<ConfluenceConfig | null> | undefined;

/** A key saved in this browser by an earlier version: handed to the local server once, then removed from here. */
async function migrate() {
  const old = (await settings.get()).confluence;
  if (!old) return;
  if (old.key) await (await server("/confluence-settings", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(old) })).text();
  await settings.set({ confluence: undefined });
}

/** The settings (asked once to the local server), or null when Confluence is not set up. */
export const readConfluenceConfig = () =>
  (config ??= (async () => {
    await migrate();
    const r = await server("/confluence-settings");
    return r.ok ? ((await r.json()) as ConfluenceConfig | null) : null;
  })().catch((e) => {
    config = undefined;
    throw e;
  }));

/** Saves the settings in the local server (an empty key keeps the one saved). */
export async function saveConfluenceConfig(values: { url: string; email?: string; key?: string; folderId: string }) {
  const r = await server("/confluence-settings", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(values) });
  if (!r.ok) throw new Error((await r.json().catch(() => undefined))?.error ?? `The settings were not saved (${r.status}).`);
  config = undefined;
}

export async function confluenceConfig(): Promise<ConfluenceConfig> {
  const c = await readConfluenceConfig();
  if (!c?.url || !c.hasKey || !c.folderId) throw new Error("Confluence is not set up: give its address, your email, an API token and a folder id in Settings › Confluence.");
  return c;
}

/** One call (through the local server); its errors said plainly. */
async function call<T = any>(path: string, init: RequestInit = {}): Promise<T> {
  await confluenceConfig();
  // A form (file upload) sets its own Content-Type (with its boundary).
  const headers: Record<string, string> = { ...(init.headers as Record<string, string> | undefined) };
  if (init.body && !(init.body instanceof FormData)) headers["Content-Type"] = "application/json";
  const r = await server(`/confluence${path}`, { ...init, headers });
  const text = await r.text();
  const body = text ? (() => { try { return JSON.parse(text); } catch { return text; } })() : null;
  if (r.status === 401 || r.status === 403) throw new Error(body?.error && r.status === 403 && /Origin|header|localhost/.test(body.error) ? body.error : `Confluence refused the key (HTTP ${r.status}): check the email (your Atlassian account's) and the token in Settings › Confluence.`);
  if (!r.ok) throw new Error(`Confluence error ${r.status}: ${typeof body === "string" ? body.slice(0, 300) : (body?.error ?? body?.message ?? JSON.stringify(body).slice(0, 300))}`);
  return body as T;
}

const link = (base: string, page: { _links?: { webui?: string; base?: string } }) => (page._links?.webui ? `${page._links.base ?? base.replace(/\/+$/, "")}${page._links.webui}` : undefined);

/** The pages of the folder set in the settings (by title). */
export async function listPages(): Promise<{ id: string; title: string }[]> {
  const { folderId } = await confluenceConfig();
  const cql = encodeURIComponent(`ancestor=${folderId} and type=page order by title`);
  const found = await call<{ results: { id: string; title: string }[] }>(`/rest/api/content/search?cql=${cql}&limit=250`).catch(() => undefined);
  const results = found?.results ?? (await call<{ results: { id: string; title: string }[] }>(`/rest/api/content/${folderId}/child/page?limit=250`)).results;
  return results.map((p) => ({ id: String(p.id), title: p.title })).sort((a, b) => a.title.localeCompare(b.title));
}

/** A page with its body (Confluence storage format: XHTML) and its version (needed to write it). */
export async function readPage(id: string): Promise<ConfluencePage> {
  const { url } = await confluenceConfig();
  const p = await call(`/rest/api/content/${id}?expand=body.storage,version`);
  return { id: String(p.id), title: p.title, version: p.version?.number ?? 1, html: p.body?.storage?.value ?? "", webUrl: link(url, p) };
}

/** Writes a page (title and body in storage format) as the next version of the one read. */
export async function writePage(page: Pick<ConfluencePage, "id" | "title" | "version">, html: string): Promise<ConfluencePage> {
  const { url } = await confluenceConfig();
  const p = await call(`/rest/api/content/${page.id}`, {
    method: "PUT",
    body: JSON.stringify({ id: page.id, type: "page", title: page.title, version: { number: page.version + 1 }, body: { storage: { value: html, representation: "storage" } } }),
  });
  return { id: String(p.id), title: p.title, version: p.version?.number ?? page.version + 1, html, webUrl: link(url, p) };
}

/**
 * A new page in the folder (in the folder's space). Under a Confluence Cloud folder (not a page), through
 * the v2 API, the one that takes a folder as parent; under a page, the v1 API (also Data Center).
 */
export async function createPage(title: string, html = ""): Promise<ConfluencePage> {
  const { url, folderId } = await confluenceConfig();
  const folder = await call(`/rest/api/content/${folderId}?expand=space`);
  const p =
    folder.type === "folder"
      ? await call(`/api/v2/pages`, { method: "POST", body: JSON.stringify({ spaceId: String(folder.space?.id), status: "current", title, parentId: folderId, body: { representation: "storage", value: html } }) })
      : await call(`/rest/api/content`, {
          method: "POST",
          body: JSON.stringify({ type: "page", title, space: { key: folder.space?.key }, ancestors: [{ id: folderId }], body: { storage: { value: html, representation: "storage" } } }),
        });
  return { id: String(p.id), title: p.title, version: typeof p.version === "object" ? (p.version?.number ?? 1) : 1, html, webUrl: link(url, p) };
}

/* -------------------------------- Attachments -------------------------------- */

/** The names of the files attached to a page. */
export async function attachmentNames(pageId: string): Promise<Set<string>> {
  const r = await call<{ results: { title: string }[] }>(`/rest/api/content/${pageId}/child/attachment?limit=250`);
  return new Set(r.results.map((a) => a.title));
}

/** Attaches a file to a page (a file of the same name gets a new version). */
export async function attachFile(pageId: string, name: string, data: Blob) {
  const form = new FormData();
  form.append("file", new File([data], name, { type: data.type || "application/octet-stream" }));
  form.append("minorEdit", "true");
  // PUT creates the file or adds a version (Cloud); Data Center only has POST (a new file).
  await call(`/rest/api/content/${pageId}/child/attachment`, { method: "PUT", body: form }).catch(() => call(`/rest/api/content/${pageId}/child/attachment`, { method: "POST", body: form }));
}

/* ------------------------------ Markdown sources ------------------------------ */

/**
 * The markdown each page was last written from, by version (in this browser). A page not changed in
 * Confluence since gives back exactly that markdown (no conversion, so no drift); a page changed there
 * is converted from its storage format.
 */
const SOURCES = "hub-confluence-markdown";
const sources = (): Record<string, { version: number; markdown: string }> => {
  try {
    return JSON.parse(localStorage.getItem(SOURCES) ?? "{}") ?? {};
  } catch {
    return {};
  }
};

export function markdownOf(page: Pick<ConfluencePage, "id" | "version" | "html">) {
  const kept = sources()[page.id];
  return kept?.version === page.version ? kept.markdown : fromStorage(page.html);
}

export function keepMarkdown(page: Pick<ConfluencePage, "id" | "version">, markdown: string) {
  try {
    localStorage.setItem(SOURCES, JSON.stringify({ ...sources(), [page.id]: { version: page.version, markdown } }));
  } catch {}
}
