/**
 * Confluence pages, read and written from the browser through the Confluence REST API (v1 content
 * API: Cloud and Data Center). The address, key and folder are set in Settings › Confluence and kept
 * in this browser only. With an email, email + token are sent as Basic auth (Cloud); without, the
 * token as a Bearer token (a personal access token of Data Center). Confluence Cloud allows no call from a browser page:
 * its sites (*.atlassian.net) are reached through the local token server (scripts/token-server.ts).
 */
import { TOKEN_SERVER } from "@/config";
import { settings } from "@/lib/store";
import { fromStorage } from "@/lib/markdown/storage";

export type ConfluenceConfig = { url: string; email?: string; key: string; folderId: string };
export type ConfluencePage = { id: string; title: string; version: number; html: string; webUrl?: string };

export async function confluenceConfig(): Promise<ConfluenceConfig> {
  const c = (await settings.get()).confluence;
  if (!c?.url || !c.key || !c.folderId) throw new Error("Confluence is not set up: give its address, your email, an API token and a folder id in Settings › Confluence.");
  return c;
}

/** A Confluence Cloud address (https://<site>.atlassian.net/…) as the token server's route to it; any other as given. */
const viaTokenServer = (url: string) => url.replace(/^https?:\/\/([a-z0-9-]+\.atlassian\.net)(?=\/|$)/i, (_, host) => `${TOKEN_SERVER.replace(/\/+$/, "")}/confluence/${host}`);

/** One call; its errors said plainly (a browser call refused by Confluence's CORS cannot be read: said too). */
async function call<T = any>(path: string, init: RequestInit = {}): Promise<T> {
  const { url, email, key } = await confluenceConfig();
  const base = viaTokenServer(url.trim().replace(/\/+$/, ""));
  const cloud = base !== url.trim().replace(/\/+$/, "");
  // "email:token" in the key alone: the form of earlier versions.
  const basic = email?.trim() ? `${email.trim()}:${key.trim()}` : key.includes(":") ? key.trim() : undefined;
  if (cloud && !basic) throw new Error("Confluence Cloud needs the email of your Atlassian account with the token (Settings › Confluence).");
  const auth = basic ? `Basic ${btoa(basic)}` : `Bearer ${key.trim()}`;
  let r: Response;
  try {
    // A form (file upload) sets its own Content-Type (with its boundary).
    const headers: Record<string, string> = { Accept: "application/json", Authorization: auth, "X-Atlassian-Token": "no-check", ...(init.headers as Record<string, string> | undefined) };
    if (!(init.body instanceof FormData)) headers["Content-Type"] = "application/json";
    r = await fetch(`${base}${path}`, { ...init, headers });
  } catch {
    throw new Error(cloud ? `Confluence cannot be reached through the token server (${TOKEN_SERVER}): is it running (bun dev)?` : `Confluence cannot be reached from the browser at ${base} (network, or the route does not allow calls from this app: CORS).`);
  }
  const text = await r.text();
  const body = text ? (() => { try { return JSON.parse(text); } catch { return text; } })() : null;
  if (r.status === 401 || r.status === 403) throw new Error(`Confluence refused the key (HTTP ${r.status}): check the email (your Atlassian account's) and the token in Settings › Confluence.`);
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
