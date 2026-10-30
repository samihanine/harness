/**
 * Confluence pages, read and written from the browser through the Confluence REST API (v1 content
 * API: Cloud and Data Center). The address, key and folder are set in Settings › Confluence and kept
 * in this browser only. A key "email:token" is sent as Basic auth (Cloud), any other as a Bearer
 * token (a personal access token of Data Center).
 */
import { settings } from "@/lib/store";

export type ConfluenceConfig = { url: string; key: string; folderId: string };
export type ConfluencePage = { id: string; title: string; version: number; html: string; webUrl?: string };

export async function confluenceConfig(): Promise<ConfluenceConfig> {
  const c = (await settings.get()).confluence;
  if (!c?.url || !c.key || !c.folderId) throw new Error("Confluence is not set up: give its address, an API key and a folder id in Settings › Confluence.");
  return c;
}

/** One call; its errors said plainly (a browser call refused by Confluence's CORS cannot be read: said too). */
async function call<T = any>(path: string, init: RequestInit = {}): Promise<T> {
  const { url, key } = await confluenceConfig();
  const base = url.trim().replace(/\/+$/, "");
  const auth = key.includes(":") ? `Basic ${btoa(key.trim())}` : `Bearer ${key.trim()}`;
  let r: Response;
  try {
    r = await fetch(`${base}${path}`, { ...init, headers: { Accept: "application/json", "Content-Type": "application/json", Authorization: auth, ...(init.headers ?? {}) } });
  } catch {
    throw new Error(`Confluence cannot be reached from the browser at ${base} (network, or the route does not allow calls from this app: CORS).`);
  }
  const text = await r.text();
  const body = text ? (() => { try { return JSON.parse(text); } catch { return text; } })() : null;
  if (r.status === 401 || r.status === 403) throw new Error(`Confluence refused the key (HTTP ${r.status}): check it in Settings › Confluence.`);
  if (!r.ok) throw new Error(`Confluence error ${r.status}: ${typeof body === "string" ? body.slice(0, 300) : (body?.message ?? JSON.stringify(body).slice(0, 300))}`);
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

/** A new page in the folder (in the folder's space). */
export async function createPage(title: string, html = ""): Promise<ConfluencePage> {
  const { url, folderId } = await confluenceConfig();
  const folder = await call(`/rest/api/content/${folderId}?expand=space`);
  const p = await call(`/rest/api/content`, {
    method: "POST",
    body: JSON.stringify({ type: "page", title, space: { key: folder.space?.key }, ancestors: [{ id: folderId }], body: { storage: { value: html, representation: "storage" } } }),
  });
  return { id: String(p.id), title: p.title, version: p.version?.number ?? 1, html, webUrl: link(url, p) };
}
