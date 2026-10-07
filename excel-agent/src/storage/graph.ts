/** Microsoft Graph client: tokens from the local API (server/), retries, JSON batches, share links. */
import type { DriveRef } from "./types";

export const GRAPH = "https://graph.microsoft.com/v1.0";

export class SignInRequiredError extends Error {
  constructor() {
    super("Sign in with your Microsoft account to use SharePoint / OneDrive files.");
  }
}

export type AuthStatus = {
  signedIn: boolean;
  account?: string;
  pending: { userCode: string; verificationUri: string; expiresAt: number; error?: string } | null;
};

let cached: { token: string; expiresAt: number } | null = null;

export async function getToken() {
  if (cached && cached.expiresAt > Date.now()) return cached.token;
  const response = await fetch("/api/token").catch(() => {
    throw new Error("The local API is not running: start the app with `bun dev` or `bun start`.");
  });
  if (response.status === 401) throw new SignInRequiredError();
  if (!response.ok) throw new Error(`Token error (${response.status})`);
  cached = (await response.json()) as { token: string; expiresAt: number };
  return cached.token;
}

const api = async <T>(path: string, method = "GET") => {
  const response = await fetch(`/api${path}`, { method });
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error ?? `API error ${response.status}`);
  return (await response.json()) as T;
};
export const auth = {
  status: () => api<AuthStatus>("/status"),
  login: () => api<AuthStatus["pending"]>("/login", "POST"),
  logout: async () => {
    cached = null;
    await api("/logout", "POST");
  },
};

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code = "",
  ) {
    super(message);
  }
}

const errorOf = (status: number, body: { error?: { code?: string; message?: string; innerError?: { code?: string } } }) =>
  new GraphError(body.error?.message ?? `Microsoft Graph error ${status}`, status, body.error?.innerError?.code ?? body.error?.code ?? "");

/** Graph request (path relative to /v1.0 or absolute URL); retries throttling. */
export async function graph<T = unknown>(path: string, init: RequestInit & { json?: unknown } = {}): Promise<T> {
  const { json, ...rest } = init;
  for (let attempt = 0; ; attempt++) {
    const response = await fetch(path.startsWith("http") ? path : `${GRAPH}${path}`, {
      ...rest,
      headers: {
        Authorization: `Bearer ${await getToken()}`,
        ...(json !== undefined ? { "Content-Type": "application/json" } : {}),
        ...rest.headers,
      },
      body: json !== undefined ? JSON.stringify(json) : rest.body,
    });
    if ((response.status === 429 || response.status === 503) && attempt < 4) {
      await new Promise((r) => setTimeout(r, (Number(response.headers.get("retry-after")) || 2 ** attempt) * 1000));
      continue;
    }
    if (response.status === 204) return undefined as T;
    const body = response.headers.get("content-type")?.includes("json") ? await response.json() : await response.blob();
    if (!response.ok) throw errorOf(response.status, body as never);
    return body as T;
  }
}

export type BatchRequest = { method: string; url: string; body?: unknown; headers?: Record<string, string> };

/**
 * Several requests in one round trip (JSON batching, 20 per call), run in order
 * (each depends on the previous one, as workbook writes must be).
 */
export async function batch(requests: BatchRequest[]) {
  const results: { status: number; body: unknown }[] = [];
  for (let start = 0; start < requests.length; start += 20) {
    const chunk = requests.slice(start, start + 20);
    const { responses } = await graph<{ responses: { id: string; status: number; body: unknown }[] }>("/$batch", {
      method: "POST",
      json: {
        requests: chunk.map((r, i) => ({
          id: String(i + 1),
          method: r.method,
          url: r.url,
          ...(i ? { dependsOn: [String(i)] } : {}),
          headers: { ...(r.body !== undefined ? { "Content-Type": "application/json" } : {}), ...r.headers },
          ...(r.body !== undefined ? { body: r.body } : {}),
        })),
      },
    });
    const sorted = [...responses].sort((a, b) => Number(a.id) - Number(b.id));
    for (const r of sorted) {
      if (r.status >= 400) throw errorOf(r.status, r.body as never);
      results.push({ status: r.status, body: r.body });
    }
  }
  return results;
}

type Item = { id: string; name: string; webUrl: string; folder?: unknown; parentReference: { driveId: string } };
const SELECT = "$select=id,name,webUrl,folder,parentReference";
const refOf = (item: Item) => ({ driveId: item.parentReference.driveId, itemId: item.id, name: item.name, webUrl: item.webUrl, isFolder: !!item.folder });

/**
 * Server-relative path of a SharePoint / OneDrive address bar URL ("…/my?id=/personal/…/Documents/Main",
 * "…/AllItems.aspx?id=…" or "?RootFolder=…"), which /shares does not resolve.
 */
const browserPath = (url: URL) => {
  const path = url.searchParams.get("id") ?? url.searchParams.get("RootFolder");
  return path?.startsWith("/") ? path : undefined;
};

/** "/personal/x/Documents/Main" → its site, then the library whose address starts the path, then the item. */
async function resolvePath(host: string, path: string): Promise<Item> {
  const [, kind, name] = path.split("/");
  if (!["personal", "sites", "teams"].includes(kind) || !name) throw new Error("Unrecognized SharePoint / OneDrive link: use a sharing link (Share › Copy link).");
  const site = await graph<{ id: string }>(`/sites/${host}:/${kind}/${name}?$select=id`);
  const { value: drives } = await graph<{ value: { id: string; webUrl: string }[] }>(`/sites/${site.id}/drives?$select=id,webUrl`);
  const full = decodeURIComponent(path).toLowerCase();
  const drive = drives
    .map((d) => ({ ...d, path: decodeURIComponent(new URL(d.webUrl).pathname).toLowerCase() }))
    .filter((d) => full === d.path || full.startsWith(`${d.path}/`))
    .sort((x, y) => y.path.length - x.path.length)[0];
  if (!drive) throw new Error("Library not found for this link: use a sharing link (Share › Copy link).");
  const rest = decodeURIComponent(path).slice(drive.path.length).replace(/^\/|\/$/g, "");
  const encoded = rest.split("/").map(encodeURIComponent).join("/");
  return graph<Item>(rest ? `/drives/${drive.id}/root:/${encoded}?${SELECT}` : `/drives/${drive.id}/root?${SELECT}`);
}

/** Share link or address bar URL of a SharePoint / OneDrive item → its drive reference. */
export async function resolveLink(link: string): Promise<DriveRef & { isFolder: boolean }> {
  const url = new URL(link.trim());
  const path = browserPath(url);
  if (path) return refOf(await resolvePath(url.host, path));
  // UTF-8 first: btoa() only takes Latin-1 (links with accents, e.g. a "Médias" folder).
  const encoded = `u!${btoa(String.fromCharCode(...new TextEncoder().encode(link.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-")}`;
  return refOf(await graph<Item>(`/shares/${encoded}/driveItem?${SELECT}`));
}

export const isDriveLink = (value: string) => /^https:\/\/[^/]*(sharepoint\.com|onedrive\.live\.com|1drv\.ms)\//i.test(value);
