/**
 * Microsoft sign-in and API calls, server side (no app registration, no CORS issues).
 * One device-code sign-in with the Microsoft Office public client gives tokens for Power BI, Fabric,
 * Graph and SharePoint. The refresh token stays in .local/tokens.json on this machine.
 */
import { createServerFn } from "@tanstack/react-start";
import { mkdir, readFile, writeFile } from "node:fs/promises";

// Production server (`bun start`, which loads .env itself): same TLS option as in vite.config.ts.
if (process.env.IGNORE_TLS_ERRORS === "true") process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office (public client)
const LOGIN = "https://login.microsoftonline.com/organizations/oauth2/v2.0";
const STORE = ".local/tokens.json";

const BASES = {
  powerbi: "https://api.powerbi.com/v1.0/myorg",
  fabric: "https://api.fabric.microsoft.com/v1",
  graph: "https://graph.microsoft.com/v1.0",
} as const;
export type Service = keyof typeof BASES | "sharepoint";

const scopeOf = (service: Service, host?: string) =>
  ({
    powerbi: "https://analysis.windows.net/powerbi/api/.default",
    fabric: "https://api.fabric.microsoft.com/.default",
    graph: "https://graph.microsoft.com/.default",
    sharepoint: `https://${host}/.default`,
  })[service] + " offline_access openid profile";

type Pending = { userCode: string; verificationUri: string; expiresAt: number; error?: string };
let state: { refreshToken?: string; account?: string } | undefined;
const load = async () => (state ??= await readFile(STORE, "utf8").then(JSON.parse).catch(() => ({})))!;
let pending: Pending | null = null;
const cache = new Map<string, { token: string; expiresAt: number }>();

const post = async (endpoint: string, params: Record<string, string>) =>
  (await (await fetch(`${LOGIN}/${endpoint}`, { method: "POST", body: new URLSearchParams(params) })).json()) as Record<string, any>;

const save = async () => {
  await mkdir(".local", { recursive: true });
  await writeFile(STORE, JSON.stringify(state));
};

async function accessToken(service: Service, host?: string) {
  const key = `${service}:${host ?? ""}`;
  const hit = cache.get(key);
  if (hit && hit.expiresAt > Date.now()) return hit;
  const s = await load();
  if (!s.refreshToken) throw new Error("SIGN_IN_REQUIRED");
  const t = await post("token", { client_id: CLIENT_ID, grant_type: "refresh_token", refresh_token: s.refreshToken, scope: scopeOf(service, host) });
  if (!t.access_token) throw new Error(/invalid_grant|interaction_required/.test(t.error) ? "SIGN_IN_REQUIRED" : t.error_description);
  if (t.refresh_token && t.refresh_token !== s.refreshToken) {
    s.refreshToken = t.refresh_token;
    void save();
  }
  const entry = { token: t.access_token as string, expiresAt: Date.now() + (t.expires_in - 300) * 1000 };
  cache.set(key, entry);
  return entry;
}

export const authStatus = createServerFn().handler(async () => {
  if (pending && pending.expiresAt < Date.now()) pending = null;
  const signedIn = await accessToken("powerbi").then(() => true, () => false);
  return { signedIn, account: signedIn ? (await load()).account : undefined, pending };
});

export const signIn = createServerFn({ method: "POST" }).handler(async () => {
  const d = await post("devicecode", { client_id: CLIENT_ID, scope: scopeOf("powerbi") });
  if (!d.device_code) throw new Error(d.error_description ?? "Could not start the sign-in");
  const p: Pending = (pending = { userCode: d.user_code, verificationUri: d.verification_uri, expiresAt: Date.now() + d.expires_in * 1000 });
  void (async () => {
    let interval = Number(d.interval ?? 5);
    while (pending === p && p.expiresAt > Date.now()) {
      await new Promise((r) => setTimeout(r, interval * 1000));
      const t = await post("token", { client_id: CLIENT_ID, grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: d.device_code });
      if (t.error === "authorization_pending") continue;
      if (t.error === "slow_down") interval += 5;
      else if (t.access_token) {
        const claims = JSON.parse(Buffer.from(t.access_token.split(".")[1], "base64url").toString());
        state = { refreshToken: t.refresh_token, account: claims.upn ?? claims.unique_name };
        cache.clear();
        await save();
        pending = null;
        return;
      } else return void (p.error = t.error_description ?? "Sign-in failed");
    }
  })();
  return p;
});

export const signOut = createServerFn({ method: "POST" }).handler(async () => {
  state = {};
  cache.clear();
  pending = null;
  await save();
});

/** Access token for the browser (Power BI embedding). */
export const getToken = createServerFn()
  .inputValidator((d: { service: Service; host?: string }) => d)
  .handler(({ data }) => accessToken(data.service, data.host));

export type MsRequest = { service: keyof typeof BASES; path: string; method?: string; body?: unknown };

/**
 * One Microsoft API call (path relative to the service base, or absolute URL). Fabric long-running
 * operations (202 + Location) are awaited and their result returned. Errors carry the service message.
 */
export const ms = createServerFn({ method: "POST" })
  .inputValidator((d: MsRequest) => d)
  .handler(async ({ data }) => {
    const { token } = await accessToken(data.service);
    const url = data.path.startsWith("http") ? data.path : `${BASES[data.service]}${data.path}`;
    const send = (u: string, method = "GET", body?: unknown) =>
      fetch(u, {
        method,
        headers: { Authorization: `Bearer ${token}`, ...(body !== undefined ? { "Content-Type": "application/json" } : {}) },
        body: body === undefined ? undefined : JSON.stringify(body),
      });
    const method = data.method ?? "GET";
    let response = await send(url, method, data.body);
    // Throttling, and Excel Online's passing failures ("Sorry… we ran into a problem"): retried when
    // repeating is harmless (reads and cell writes; never row inserts / deletes).
    const safe = method === "GET" || method === "PATCH";
    for (let i = 0; i < 4 && (response.status === 429 || (safe && [409, 500, 502, 503, 504].includes(response.status))); i++) {
      await new Promise((r) => setTimeout(r, (Number(response.headers.get("retry-after")) || 2 ** i) * 1000));
      response = await send(url, method, data.body);
    }
    const location = response.headers.get("location");
    if (response.status === 202 && location && data.service === "fabric") {
      for (let i = 0; i < 90; i++) {
        await new Promise((r) => setTimeout(r, 2000));
        const op = await (await send(location)).json();
        if (op.status === "Failed") throw new Error(op.error?.message ?? "Operation failed");
        if (op.status === "Succeeded") {
          const result = await send(`${location}/result`);
          return result.ok ? JSON.parse((await result.text()) || "null") : null;
        }
      }
      throw new Error("Operation timed out");
    }
    const text = await response.text();
    let body: any = null;
    try {
      body = text ? JSON.parse(text) : null;
    } catch {
      body = text;
    }
    if (!response.ok) {
      const e = body?.error;
      throw new Error(e?.pbi?.error?.details?.[0]?.detail?.value ?? e?.message ?? body?.message ?? `${data.service} error ${response.status}: ${String(text).slice(0, 200)}`);
    }
    return body;
  });

/**
 * Gives a semantic model fresh OAuth credentials for its SharePoint / OneDrive sources, then refreshes it
 * (credentials set by API expire with the token, so they are renewed before each refresh).
 */
export const refreshModel = createServerFn({ method: "POST" })
  .inputValidator((d: { datasetId: string; groupId?: string }) => d)
  .handler(async ({ data }) => {
    const base = `${BASES.powerbi}${data.groupId ? `/groups/${data.groupId}` : ""}/datasets/${data.datasetId}`;
    const pbi = (await accessToken("powerbi")).token;
    const call = async (url: string, method = "GET", body?: unknown) => {
      const r = await fetch(url, { method, headers: { Authorization: `Bearer ${pbi}`, "Content-Type": "application/json" }, body: body === undefined ? undefined : JSON.stringify(body) });
      const text = await r.text();
      if (!r.ok) throw new Error(`${r.status} ${text.slice(0, 300)}`);
      return text ? JSON.parse(text) : null;
    };
    const { value: sources } = await call(`${base}/datasources`);
    for (const s of sources) {
      const url = s.connectionDetails?.url ?? s.connectionDetails?.path;
      if (!url || !/sharepoint\.com/i.test(url)) continue;
      const { token } = await accessToken("sharepoint", new URL(url).host);
      await call(`${BASES.powerbi}/gateways/${s.gatewayId}/datasources/${s.datasourceId}`, "PATCH", {
        credentialDetails: {
          credentialType: "OAuth2",
          credentials: JSON.stringify({ credentialData: [{ name: "accessToken", value: token }] }),
          encryptedConnection: "Encrypted",
          encryptionAlgorithm: "None",
          privacyLevel: "Organizational",
        },
      });
    }
    await call(`${base}/refreshes`, "POST", { notifyOption: "NoNotification" });
    for (let i = 0; i < 60; i++) {
      await new Promise((r) => setTimeout(r, 3000));
      const last = (await call(`${base}/refreshes?$top=1`)).value?.[0];
      if (last && last.status !== "Unknown")
        return { status: last.status as string, error: last.serviceExceptionJson ? JSON.parse(last.serviceExceptionJson).errorDescription ?? last.serviceExceptionJson : undefined };
    }
    return { status: "Running" };
  });

/** Uploads a file (base64) to a drive folder, created when missing; returns the drive item. */
export const uploadFile = createServerFn({ method: "POST" })
  .inputValidator((d: { driveId: string; parentId: string; path: string; base64: string; type?: string }) => d)
  .handler(async ({ data }) => {
    const { token } = await accessToken("graph");
    const path = data.path.split("/").map(encodeURIComponent).join("/");
    const r = await fetch(`${BASES.graph}/drives/${data.driveId}/items/${data.parentId}:/${path}:/content?@microsoft.graph.conflictBehavior=rename`, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": data.type || "application/octet-stream" },
      body: Buffer.from(data.base64, "base64"),
    });
    const item = await r.json();
    if (!r.ok) throw new Error(item.error?.message ?? `Upload failed (${r.status})`);
    return { id: item.id as string };
  });

/**
 * The report as a .pbix (base64). "LiveConnect": the file is connected to the online semantic model —
 * the only download that works for reports made in the service on a model of another workspace.
 */
export const exportPbix = createServerFn({ method: "POST" })
  .inputValidator((d: { id: string; groupId?: string }) => d)
  .handler(async ({ data }) => {
    const { token } = await accessToken("powerbi");
    const base = `${BASES.powerbi}${data.groupId ? `/groups/${data.groupId}` : ""}/reports/${data.id}/Export`;
    let r = await fetch(`${base}?downloadType=LiveConnect`, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) r = await fetch(base, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) {
      const body = await r.text();
      throw new Error(`The report cannot be downloaded (${r.status}${body ? `: ${body.slice(0, 200)}` : ""}). Downloads may be disabled by the admin.`);
    }
    return Buffer.from(await r.arrayBuffer()).toString("base64");
  });

/** Content of a SharePoint / OneDrive file from a link, as a data URL (images shown in the app). */
export const downloadFile = createServerFn({ method: "POST" })
  .inputValidator((d: { url: string }) => d)
  .handler(async ({ data }) => {
    const { token } = await accessToken("graph");
    const id = "u!" + Buffer.from(data.url.trim()).toString("base64url");
    const r = await fetch(`${BASES.graph}/shares/${id}/driveItem/content`, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) throw new Error(`Cannot read the file (${r.status})`);
    const type = r.headers.get("content-type") ?? "application/octet-stream";
    return `data:${type};base64,${Buffer.from(await r.arrayBuffer()).toString("base64")}`;
  });
