/**
 * Confluence (Cloud or Data Center), called by the server: the pages of one folder, read and written as
 * markdown (converted from / to the storage format). With an email, email + token as Basic auth (Cloud);
 * without, the token as a Bearer token (Data Center personal access token).
 */
import { DOMParser } from "@xmldom/xmldom";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { settings } from "./settings";

export type ConfluencePage = { id: string; title: string; version: number; html: string; webUrl?: string };

async function config() {
  const c = (await settings()).confluence;
  if (!c.url || !c.key) throw new Error("Confluence is not set up: give its address, your email and an API token in Settings › Confluence.");
  return c;
}

async function call<T = any>(path: string, init: RequestInit = {}): Promise<T> {
  const { url, email, key } = await config();
  const base = url.trim().replace(/\/+$/, "");
  const auth = email.trim() ? `Basic ${Buffer.from(`${email.trim()}:${key.trim()}`).toString("base64")}` : `Bearer ${key.trim()}`;
  const r = await fetch(`${base}${path}`, { ...init, signal: AbortSignal.timeout(20_000), headers: { Accept: "application/json", "Content-Type": "application/json", "X-Atlassian-Token": "no-check", Authorization: auth, ...(init.headers ?? {}) } });
  const text = await r.text();
  const body = text ? (() => { try { return JSON.parse(text); } catch { return text; } })() : null;
  if (r.status === 401 || r.status === 403) throw new Error(`Confluence refused the key (HTTP ${r.status}${typeof body === "object" && body?.message ? `: ${String(body.message).replace(/^.*Exception:\s*/, "")}` : ""}). Check the email and the token in Settings › Confluence.`);
  if (!r.ok) throw new Error(`Confluence error ${r.status}: ${typeof body === "string" ? body.slice(0, 300) : (body?.message ?? JSON.stringify(body).slice(0, 300))}`);
  return body as T;
}

const link = (base: string, page: { _links?: { webui?: string; base?: string } }) => (page._links?.webui ? `${page._links.base ?? base.replace(/\/+$/, "")}${page._links.webui}` : undefined);

/** The pages of a folder (its descendants; the settings' folder when none is given), by title. */
export async function listPages(folder?: string): Promise<{ id: string; title: string }[]> {
  const folderId = folder ?? (await config()).folderId;
  if (!folderId) throw new Error("No Confluence folder set (Settings › Confluence).");
  if (!/^\d+$/.test(folderId)) throw new Error(`"${folderId}" is not a Confluence folder id.`);
  const cql = encodeURIComponent(`ancestor=${folderId} and type=page order by title`);
  const found = await call<{ results: { id: string; title: string }[] }>(`/rest/api/content/search?cql=${cql}&limit=250`).catch(() => undefined);
  const results = found?.results ?? (await call<{ results: { id: string; title: string }[] }>(`/rest/api/content/${folderId}/child/page?limit=250`)).results;
  return results.map((p) => ({ id: String(p.id), title: p.title })).sort((a, b) => a.title.localeCompare(b.title));
}

export async function readPage(id: string): Promise<ConfluencePage> {
  if (!/^\d+$/.test(id)) throw new Error("A Confluence page id is a number.");
  const { url } = await config();
  const p = await call(`/rest/api/content/${id}?expand=body.storage,version`);
  return { id: String(p.id), title: p.title, version: p.version?.number ?? 1, html: p.body?.storage?.value ?? "", webUrl: link(url, p) };
}

export async function readPageMarkdown(id: string) {
  const p = await readPage(id);
  return { ...p, markdown: toMarkdown(p.html) };
}

export async function writePage(page: { id: string; title: string; version: number }, markdown: string): Promise<ConfluencePage> {
  const { url } = await config();
  const html = toStorage(markdown);
  const p = await call(`/rest/api/content/${page.id}`, {
    method: "PUT",
    body: JSON.stringify({ id: page.id, type: "page", title: page.title, version: { number: page.version + 1 }, body: { storage: { value: html, representation: "storage" } } }),
  });
  return { id: String(p.id), title: p.title, version: p.version?.number ?? page.version + 1, html, webUrl: link(url, p) };
}

/** A new page in the folder (v2 API under a Cloud folder, v1 under a page). */
export async function createPage(title: string, markdown = ""): Promise<ConfluencePage> {
  const { url, folderId } = await config();
  const html = toStorage(markdown);
  const folder = await call(`/rest/api/content/${folderId}?expand=space`);
  const p =
    folder.type === "folder"
      ? await call(`/api/v2/pages`, { method: "POST", body: JSON.stringify({ spaceId: String(folder.space?.id), status: "current", title, parentId: folderId, body: { representation: "storage", value: html } }) })
      : await call(`/rest/api/content`, { method: "POST", body: JSON.stringify({ type: "page", title, space: { key: folder.space?.key }, ancestors: [{ id: folderId }], body: { storage: { value: html, representation: "storage" } } }) });
  return { id: String(p.id), title: p.title, version: typeof p.version === "object" ? (p.version?.number ?? 1) : 1, html, webUrl: link(url, p) };
}

/* ---------------------------- Markdown ⇄ storage ---------------------------- */

export const toStorage = (markdown: string) => renderToStaticMarkup(createElement(ReactMarkdown, { remarkPlugins: [remarkGfm] }, markdown));

type N = { nodeType: number; nodeName: string; textContent: string | null; childNodes: ArrayLike<N>; getAttribute?: (n: string) => string | null };
const children = (n: N) => Array.from(n.childNodes ?? []);
const tag = (n: N) => n.nodeName.toLowerCase();

export function toMarkdown(html: string) {
  const escaped = html.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, (_, t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"));
  // Confluence's own tags (ac:, ri:) come without their namespace declarations: added, or the parser refuses them.
  const doc = new DOMParser({ onError: () => undefined } as never).parseFromString(`<body xmlns:ac="http://atlassian.com/content" xmlns:ri="http://atlassian.com/resource/identifier" xmlns:at="http://atlassian.com/template">${escaped}</body>`, "text/html") as unknown as { getElementsByTagName: (t: string) => ArrayLike<N> };
  const body = doc.getElementsByTagName("body")[0];
  return body ? block(body).replace(/\n{3,}/g, "\n\n").trim() : "";
}

function inline(node: N): string {
  if (node.nodeType === 3) return (node.textContent ?? "").replace(/\s+/g, " ");
  if (node.nodeType !== 1) return "";
  const inner = () => children(node).map(inline).join("");
  switch (tag(node)) {
    case "strong":
    case "b":
      return `**${inner().trim()}**`;
    case "em":
    case "i":
      return `*${inner().trim()}*`;
    case "code":
      return `\`${node.textContent ?? ""}\``;
    case "a":
      return `[${inner().trim()}](${node.getAttribute?.("href") ?? ""})`;
    case "br":
      return "  \n";
    case "img":
      return `![${node.getAttribute?.("alt") ?? ""}](${node.getAttribute?.("src") ?? ""})`;
    default:
      return inner();
  }
}

function block(node: N, depth = 0): string {
  return children(node)
    .map((c) => {
      if (c.nodeType !== 1) return inline(c).trim() ? `${inline(c).trim()}\n\n` : "";
      const t = tag(c);
      if (/^h[1-6]$/.test(t)) return `${"#".repeat(Number(t[1]))} ${inline(c).trim()}\n\n`;
      if (t === "p") return `${inline(c).trim()}\n\n`;
      if (t === "hr") return "---\n\n";
      if (t === "blockquote") return `${block(c).trim().split("\n").map((l) => `> ${l}`).join("\n")}\n\n`;
      if (t === "pre") return `\`\`\`\n${(c.textContent ?? "").replace(/\n+$/, "")}\n\`\`\`\n\n`;
      if (t === "ul" || t === "ol") return `${list(c, depth)}\n`;
      if (t === "table") return `${table(c)}\n\n`;
      if (t === "ac:structured-macro") {
        const body = findAll(c, "ac:plain-text-body")[0]?.textContent;
        return body != null ? `\`\`\`\n${body.replace(/\n+$/, "")}\n\`\`\`\n\n` : block(c, depth);
      }
      return block(c, depth);
    })
    .join("");
}

function findAll(n: N, name: string): N[] {
  const out: N[] = [];
  const walk = (x: N) => children(x).forEach((c) => (c.nodeType === 1 && tag(c) === name && out.push(c), walk(c)));
  walk(n);
  return out;
}

function list(node: N, depth: number): string {
  const ordered = tag(node) === "ol";
  return children(node)
    .filter((li) => li.nodeType === 1 && tag(li) === "li")
    .map((li, i) => {
      const nested = children(li).filter((c) => c.nodeType === 1 && ["ul", "ol"].includes(tag(c)));
      const text = children(li).filter((c) => !(c.nodeType === 1 && ["ul", "ol"].includes(tag(c)))).map(inline).join("").trim();
      return `${"  ".repeat(depth)}${ordered ? `${i + 1}.` : "-"} ${text}\n${nested.map((n) => list(n, depth + 1)).join("")}`;
    })
    .join("");
}

function table(node: N): string {
  const rows = findAll(node, "tr").map((tr) => children(tr).filter((c) => c.nodeType === 1).map((c) => inline(c).trim().replace(/\|/g, "\\|")));
  if (!rows.length) return "";
  const width = Math.max(...rows.map((r) => r.length));
  const line = (r: string[]) => `| ${Array.from({ length: width }, (_, i) => r[i] ?? "").join(" | ")} |`;
  return [line(rows[0]), `| ${Array.from({ length: width }, () => "---").join(" | ")} |`, ...rows.slice(1).map(line)].join("\n");
}
