/**
 * The Microsoft APIs, called from the browser (the app is front-end only). Access tokens come from the
 * local token server (scripts/token-server.ts: one device-code sign-in, no app registration), kept in
 * localStorage until they expire. Power BI, Fabric and Graph accept browser calls (CORS).
 */
import { TOKEN_SERVER } from "@/config";

export type Service = "powerbi" | "fabric" | "graph";
const BASES: Record<Service, string> = {
  powerbi: "https://api.powerbi.com/v1.0/myorg",
  fabric: "https://api.fabric.microsoft.com/v1",
  graph: "https://graph.microsoft.com/v1.0",
};

/** Nobody is signed in: the URL and the code to sign in with (the token server finishes the sign-in by itself). */
export class SignInRequired extends Error {
  constructor(readonly url: string, readonly code: string) {
    super(`Sign in to Microsoft: open ${url} and enter the code ${code} (Settings › Power BI).`);
  }
}

type Token = { token: string; expiresAt: number; account?: string };
const KEY = (s: Service) => `ms-token:${s}`;
const read = (s: Service): Token | undefined => {
  try {
    const t = JSON.parse(localStorage.getItem(KEY(s)) ?? "null");
    return t && t.expiresAt > Date.now() + 60_000 ? t : undefined;
  } catch {
    return undefined;
  }
};
const pendingTokens = new Map<Service, Promise<Token>>();

/** An access token for a service: the stored one while valid, else asked to the token server (once at a time). */
export async function token(service: Service): Promise<Token> {
  const stored = read(service);
  if (stored) return stored;
  if (!pendingTokens.has(service))
    pendingTokens.set(
      service,
      (async () => {
        let r: Response;
        try {
          r = await fetch(`${TOKEN_SERVER}/get-token?scope=${service}`);
        } catch {
          throw new Error(`The token server is not running (${TOKEN_SERVER}): start it with "bun dev" (or "bun run token").`);
        }
        const body = await r.json();
        if (body.signIn) throw new SignInRequired(body.signIn.url, body.signIn.code);
        if (!r.ok || !body.token) throw new Error(body.error ?? `Token server error ${r.status}`);
        try {
          localStorage.setItem(KEY(service), JSON.stringify(body));
          if (body.account) localStorage.setItem("ms-account", body.account);
        } catch {}
        return body as Token;
      })().finally(() => pendingTokens.delete(service)),
    );
  return pendingTokens.get(service)!;
}

/** Who is signed in, or how to sign in, or why it cannot be known (the token server is down). */
export async function tokenStatus(): Promise<{ signedIn: boolean; account?: string; signIn?: { url: string; code: string }; error?: string }> {
  try {
    const t = await token("powerbi");
    return { signedIn: true, account: t.account ?? localStorage.getItem("ms-account") ?? undefined };
  } catch (e) {
    if (e instanceof SignInRequired) return { signedIn: false, signIn: { url: e.url, code: e.code } };
    return { signedIn: false, error: errorText(e) };
  }
}

/** Signs out: the token server forgets the sign-in, this browser its tokens. */
export async function signOut() {
  for (const s of Object.keys(BASES) as Service[]) localStorage.removeItem(KEY(s));
  localStorage.removeItem("ms-account");
  await fetch(`${TOKEN_SERVER}/sign-out`, { method: "POST" }).catch(() => undefined);
}

export const powerBiToken = async () => (await token("powerbi")).token;

/** The useful message of a Microsoft error (Power BI puts it under "pbi.error", names wrapped in <oii> tags). */
function errorOf(service: Service, status: number, text: string, body: any) {
  const e = body?.error;
  const pbiError = e?.["pbi.error"] ?? e?.pbi?.error;
  const detail = pbiError?.details?.find((d: { code?: string }) => d.code === "DetailsMessage")?.detail?.value ?? pbiError?.details?.[0]?.detail?.value;
  const message = detail ?? e?.message ?? body?.message ?? `${service} error ${status}: ${String(text).slice(0, 600)}`;
  return new Error(String(message).replace(/<\/?(oii|ccon)>/g, ""));
}

/** A raw request with the service's token (absolute URL or path relative to its base). */
export async function request(service: Service, path: string, init: RequestInit = {}) {
  const url = path.startsWith("http") ? path : `${BASES[service]}${path}`;
  const { token: t } = await token(service);
  return fetch(url, { ...init, headers: { Authorization: `Bearer ${t}`, ...(init.headers ?? {}) } });
}

/**
 * One JSON call (path relative to the service base, or absolute URL). Throttling is retried; Fabric
 * long-running operations (202 + Location) are awaited and their result returned.
 */
export async function call<T = any>(service: Service, path: string, method = "GET", body?: unknown, signal?: AbortSignal): Promise<T> {
  const send = (u: string, m = "GET", b?: unknown) => request(service, u, { method: m, signal, ...(b !== undefined && { headers: { "Content-Type": "application/json" }, body: JSON.stringify(b) }) });
  let response = await send(path, method, body);
  // Throttling, and Excel Online's passing failures: retried when repeating is harmless (reads, cell writes).
  const safe = method === "GET" || method === "PATCH";
  for (let i = 0; i < 4 && (response.status === 429 || (safe && [409, 500, 502, 503, 504].includes(response.status))); i++) {
    await new Promise((r) => setTimeout(r, (Number(response.headers.get("retry-after")) || 2 ** i) * 1000));
    response = await send(path, method, body);
  }
  const location = response.headers.get("location");
  if (response.status === 202 && location && service === "fabric") {
    for (let i = 0; i < 90; i++) {
      await new Promise((r) => setTimeout(r, 2000));
      const op = await (await send(location)).json();
      if (op.status === "Failed") throw new Error(op.error?.message ?? "Operation failed");
      if (op.status === "Succeeded") {
        const result = await send(`${location}/result`);
        return result.ok ? JSON.parse((await result.text()) || "null") : (null as T);
      }
    }
    throw new Error("Operation timed out");
  }
  const text = await response.text();
  let json: any = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = text;
  }
  if (!response.ok) throw errorOf(service, response.status, text, json);
  return json as T;
}

export const pbi = <T = any>(path: string, method?: string, body?: unknown, signal?: AbortSignal) => call<T>("powerbi", path, method, body, signal);
export const fabric = <T = any>(path: string, method?: string, body?: unknown) => call<T>("fabric", path, method, body);
export const graph = <T = any>(path: string, method?: string, body?: unknown) => call<T>("graph", path, method, body);

/** Refreshes a semantic model (started by the user, never by the AI) and waits for its result. */
export async function refreshModel({ datasetId, groupId }: { datasetId: string; groupId?: string }) {
  const base = `${groupId ? `/groups/${groupId}` : ""}/datasets/${datasetId}`;
  await pbi(`${base}/refreshes`, "POST", { notifyOption: "NoNotification" });
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 3000));
    const last = (await pbi(`${base}/refreshes?$top=1`)).value?.[0];
    if (last && last.status !== "Unknown") return { status: last.status as string, error: last.serviceExceptionJson ? (JSON.parse(last.serviceExceptionJson).errorDescription ?? last.serviceExceptionJson) : undefined };
  }
  return { status: "Running" };
}

/** The report as a .pbix ("LiveConnect": connected to the online model, the download that always works). */
export async function exportPbix({ id, groupId }: { id: string; groupId?: string }) {
  const get = (g?: string) => {
    const base = `${g ? `/groups/${g}` : ""}/reports/${id}/Export`;
    return request("powerbi", `${base}?downloadType=LiveConnect`).then((r) => (r.ok ? r : request("powerbi", base)));
  };
  let r = await get(groupId);
  // My workspace has an id (Fabric) but no /groups path in the Power BI API: asked again without it.
  if (!r.ok && groupId && /GroupNotAccessible|personal workspace/i.test(await r.clone().text())) r = await get();
  if (!r.ok) throw new Error(`The report cannot be downloaded (${r.status}). Downloads may be disabled by the admin.`);
  return r.blob();
}

/** A drive item's content and its version (eTag), to change the file and send it back. */
export async function readItem({ driveId, itemId }: { driveId: string; itemId: string }) {
  const meta = await graph(`/drives/${driveId}/items/${itemId}?$select=eTag`);
  const r = await request("graph", `/drives/${driveId}/items/${itemId}/content`);
  if (!r.ok) throw new Error(`Cannot read the file (${r.status})`);
  return { data: await r.arrayBuffer(), eTag: meta.eTag as string };
}

/** Replaces a drive item's content; with `eTag`, only if it did not change since it was read. */
export async function writeItem({ driveId, itemId, data, eTag }: { driveId: string; itemId: string; data: ArrayBuffer | Blob; eTag?: string }) {
  const put = () => request("graph", `/drives/${driveId}/items/${itemId}/content`, { method: "PUT", headers: { "Content-Type": "application/octet-stream", ...(eTag && { "If-Match": eTag }) }, body: data });
  let r = await put();
  // Excel for the web keeps the file locked a little after workbook calls: waited for.
  for (let i = 0; i < 5 && r.status === 423; i++) {
    await new Promise((res) => setTimeout(res, 2000 * (i + 1)));
    r = await put();
  }
  if (r.status === 412) throw new Error("The file changed meanwhile: read it again.");
  if (!r.ok) throw new Error(`Cannot write the file (${r.status})`);
}

/** Uploads a file to a drive folder (missing folders created; an existing file renamed, or replaced). */
export async function uploadFile({ driveId, parentId, path, file, replace }: { driveId: string; parentId: string; path: string; file: Blob; replace?: boolean }) {
  const p = path.split("/").map(encodeURIComponent).join("/");
  const r = await request("graph", `/drives/${driveId}/items/${parentId}:/${p}:/content?@microsoft.graph.conflictBehavior=${replace ? "replace" : "rename"}`, { method: "PUT", headers: { "Content-Type": file.type || "application/octet-stream" }, body: file });
  const item = await r.json();
  if (!r.ok) throw new Error(item.error?.message ?? `Upload failed (${r.status})`);
  return { id: item.id as string, webUrl: item.webUrl as string };
}

export const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));
