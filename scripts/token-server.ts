/**
 * The local server beside the app (docs/security.md). It holds what must not live in a web page:
 * - the Microsoft sign-in: one device-code sign-in with the Microsoft Office public client (no app
 *   registration); the refresh token stays in the system's credential store (scripts/server/secrets.ts);
 * - the calls to Microsoft: /ms/<service>/… forwards to Power BI REST, Fabric and Graph with the
 *   user's token, so those tokens never reach the browser. Only the Power BI token is given to the
 *   page (GET /get-token), because the embedding SDK needs it; it is kept in memory there;
 * - Confluence: its key is kept here, /confluence/… forwards to the configured site;
 * - the preview of PowerPoint decks: /office/render (scripts/server/office.ts).
 * It answers the app only (scripts/server/guard.ts): loopback interface, Host, Origin and a custom
 * header checked. The AI API is called by the page itself, never from here.
 *
 *   bun scripts/token-server.ts          the server (port TOKEN_PORT, default 3301)
 *   bun scripts/token-server.ts login    signs in from the terminal
 */
import { HOST, PORT, cors, json, refusal } from "./server/guard";
import { renderDeck, warmEngine } from "./server/office";
import { getSecret, migrateOldTokens, secretStore, setSecret } from "./server/secrets";

const CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office (public client)
const LOGIN = "https://login.microsoftonline.com/organizations/oauth2/v2.0";
const SCOPES = {
  powerbi: "https://analysis.windows.net/powerbi/api/.default",
  fabric: "https://api.fabric.microsoft.com/.default",
  graph: "https://graph.microsoft.com/.default",
} as const;
type Scope = keyof typeof SCOPES;
/** Where each service's calls go, and the hosts an absolute URL of that service may name (operation and download links). */
const BASES: Record<Scope, string> = { powerbi: "https://api.powerbi.com/v1.0/myorg", fabric: "https://api.fabric.microsoft.com/v1", graph: "https://graph.microsoft.com/v1.0" };
const HOSTS: Record<Scope, RegExp> = {
  powerbi: /(^|\.)powerbi\.com$|\.analysis\.windows\.net$/i,
  fabric: /(^|\.)fabric\.microsoft\.com$|\.analysis\.windows\.net$|(^|\.)powerbi\.com$/i,
  graph: /^graph\.microsoft\.com$/i,
};

type Session = { refreshToken?: string; account?: string };
type Pending = { userCode: string; verificationUri: string; expiresAt: number; error?: string };
let state: Session | undefined;
let pending: Pending | null = null;
const cache = new Map<Scope, { token: string; expiresAt: number }>();

const load = async () => (state ??= JSON.parse((await getSecret("session")) ?? "{}") as Session);
const save = async () => setSecret("session", state?.refreshToken ? JSON.stringify(state) : undefined);
const post = async (endpoint: string, params: Record<string, string>) => (await (await fetch(`${LOGIN}/${endpoint}`, { method: "POST", body: new URLSearchParams(params) })).json()) as Record<string, any>;
const scopeOf = (s: Scope) => `${SCOPES[s]} offline_access openid profile`;

/** An access token for a service, from the refresh token (undefined: sign-in needed). */
async function accessToken(scope: Scope) {
  const hit = cache.get(scope);
  if (hit && hit.expiresAt > Date.now()) return hit;
  const s = await load();
  if (!s.refreshToken) return undefined;
  const t = await post("token", { client_id: CLIENT_ID, grant_type: "refresh_token", refresh_token: s.refreshToken, scope: scopeOf(scope) });
  if (!t.access_token) {
    if (/invalid_grant|interaction_required/.test(t.error)) return undefined;
    throw new Error(t.error_description ?? t.error ?? "Token refused");
  }
  if (t.refresh_token && t.refresh_token !== s.refreshToken) (s.refreshToken = t.refresh_token), void save();
  const entry = { token: t.access_token as string, expiresAt: Date.now() + (t.expires_in - 300) * 1000 };
  cache.set(scope, entry);
  return entry;
}

/** Starts a device-code sign-in (once at a time); it finishes in the background. */
async function signIn() {
  if (pending && pending.expiresAt > Date.now() && !pending.error) return pending;
  const d = await post("devicecode", { client_id: CLIENT_ID, scope: scopeOf("powerbi") });
  if (!d.device_code) throw new Error(d.error_description ?? "Could not start the sign-in");
  const p: Pending = (pending = { userCode: d.user_code, verificationUri: d.verification_uri, expiresAt: Date.now() + d.expires_in * 1000 });
  const done = (async () => {
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
        return true;
      } else return void (p.error = t.error_description ?? "Sign-in failed"), false;
    }
    return false;
  })();
  return Object.assign(p, { done });
}

// A failure in the background (a stopped engine, a closed connection) is logged, never a reason to stop the server.
process.on("unhandledRejection", (e) => console.error("Unhandled:", e instanceof Error ? e.message : e));
process.on("uncaughtException", (e) => console.error("Uncaught:", e instanceof Error ? e.message : e));

await migrateOldTokens();

if (process.argv[2] === "login") {
  if (!(await accessToken("powerbi"))) {
    const p = (await signIn()) as Pending & { done: Promise<boolean> };
    console.log(`Sign in: open ${p.verificationUri} and enter the code ${p.userCode}`);
    if (!(await p.done)) throw new Error(pending?.error ?? "Sign-in expired");
  }
  console.log(`Signed in as ${(await load()).account}. The session is kept in ${await secretStore()}.`);
  process.exit(0);
}

/** "Sign in" as an answer: the app shows the code (the sign-in then finishes by itself). */
async function signInAnswer(req: Request, status = 200) {
  const p = await signIn();
  return json(req, { signIn: { url: p.verificationUri, code: p.userCode, expiresAt: p.expiresAt, error: p.error } }, status);
}

/* -------------------------------- Microsoft ---------------------------------- */

/** Request headers passed on to Microsoft, response headers passed back. */
const PASS_UP = ["content-type", "if-match", "prefer", "accept"];
const PASS_DOWN = ["content-type", "location", "retry-after", "etag", "content-disposition", "x-ms-request-id"];

/**
 * Forwards a call of the app to a Microsoft service with the user's token: /ms/<service>/<path> (under
 * the service's base), or /ms/<service>?url=<absolute URL> for the links the services give back
 * (long-running operations, downloads), on that service's own hosts only — never an open proxy.
 */
async function microsoft(req: Request, url: URL) {
  const [service, ...rest] = url.pathname.replace(/^\/ms\//, "").split("/");
  if (!(service in SCOPES)) return json(req, { error: `Unknown service "${service}".` }, 400);
  const scope = service as Scope;
  const absolute = url.searchParams.get("url");
  let target: URL;
  if (absolute && !rest.length) {
    target = new URL(absolute);
    if (target.protocol !== "https:" || !HOSTS[scope].test(target.hostname)) return json(req, { error: `${target.hostname} is not a host of ${service}.` }, 400);
  } else target = new URL(`${BASES[scope]}/${rest.join("/")}${url.search}`);
  const token = await accessToken(scope);
  if (!token) return signInAnswer(req, 401);
  const headers: Record<string, string> = { Authorization: `Bearer ${token.token}` };
  for (const h of PASS_UP) if (req.headers.get(h)) headers[h] = req.headers.get(h)!;
  const body = req.method === "GET" || req.method === "HEAD" ? undefined : await req.arrayBuffer();
  const r = await fetch(target, { method: req.method, headers, body, redirect: "follow" });
  const out: Record<string, string> = { "Cache-Control": "no-store", ...cors(req) };
  for (const h of PASS_DOWN) if (r.headers.get(h)) out[h] = r.headers.get(h)!;
  return new Response(r.body, { status: r.status, headers: out });
}

/* -------------------------------- Confluence --------------------------------- */

type ConfluenceConfig = { url: string; email?: string; key: string; folderId: string };
const confluenceConfig = async () => JSON.parse((await getSecret("confluence")) ?? "null") as ConfluenceConfig | null;

/** Forwards a Confluence call to the configured site with the key kept here (Basic with an email, else a Bearer token). */
async function confluence(req: Request, url: URL) {
  const c = await confluenceConfig();
  if (!c?.url || !c.key) return json(req, { error: "Confluence is not set up: give its address, your email, an API token and a folder id in Settings › Confluence." }, 400);
  const basic = c.email?.trim() ? `${c.email.trim()}:${c.key}` : c.key.includes(":") ? c.key : undefined;
  const headers: Record<string, string> = { Accept: "application/json", "X-Atlassian-Token": "no-check", Authorization: basic ? `Basic ${Buffer.from(basic).toString("base64")}` : `Bearer ${c.key}` };
  if (req.headers.get("content-type")) headers["Content-Type"] = req.headers.get("content-type")!;
  const target = `${c.url.replace(/\/+$/, "")}${url.pathname.replace(/^\/confluence/, "")}${url.search}`;
  const r = await fetch(target, { method: req.method, headers, body: req.method === "GET" || req.method === "HEAD" ? undefined : await req.arrayBuffer() });
  return new Response(r.body, { status: r.status, headers: { "Content-Type": r.headers.get("content-type") ?? "application/json", "Cache-Control": "no-store", ...cors(req) } });
}

/** The Confluence settings: read without the key (only whether there is one), written whole (an empty key keeps the one saved). */
async function confluenceSettings(req: Request) {
  if (req.method === "GET") {
    const c = await confluenceConfig();
    return json(req, c ? { url: c.url, email: c.email ?? "", folderId: c.folderId, hasKey: !!c.key } : null);
  }
  const given = (await req.json()) as Partial<ConfluenceConfig>;
  const old = await confluenceConfig();
  const url = String(given.url ?? "").trim();
  if (!/^https:\/\/[^/]+/i.test(url) && !/^http:\/\/(localhost|127\.0\.0\.1)/i.test(url)) return json(req, { error: "The Confluence address must start with https://." }, 400);
  await setSecret("confluence", JSON.stringify({ url, email: String(given.email ?? "").trim(), key: String(given.key ?? "").trim() || old?.key || "", folderId: String(given.folderId ?? "").trim() }));
  return json(req, { ok: true });
}

/* ---------------------------------- Office ----------------------------------- */

/** POST /office/render?width=1280&pages=1,2&fonts=Century Gothic|Arial — the deck in the body, its slides back as PNG (base64). */
async function office(req: Request, url: URL) {
  const deck = new Uint8Array(await req.arrayBuffer());
  const pages = (url.searchParams.get("pages") ?? "").split(",").map(Number).filter((n) => Number.isInteger(n) && n > 0);
  const width = Math.max(160, Math.min(2400, Number(url.searchParams.get("width")) || 1280));
  const fonts = (url.searchParams.get("fonts") ?? "").split("|").map((f) => f.trim()).filter(Boolean).slice(0, 40);
  try {
    return json(req, await renderDeck(deck, { width, pages, fonts }));
  } catch (e) {
    const missing = e instanceof Error && e.message === "ENGINE_MISSING";
    return json(req, { error: missing ? "The preview engine is not installed (run: bun install)." : e instanceof Error ? e.message : String(e), missing }, missing ? 501 : 500);
  }
}

/* ---------------------------------- Server ----------------------------------- */

/** The exported app (dist/hub.html, or hub.html beside the server), served here so it has a real origin. */
async function exportedApp() {
  for (const path of ["dist/hub.html", "hub.html"]) if (await Bun.file(path).exists()) return Bun.file(path);
  return undefined;
}
const PAGE_HEADERS = { "Content-Type": "text/html; charset=utf-8", "X-Content-Type-Options": "nosniff", "X-Frame-Options": "DENY", "Referrer-Policy": "no-referrer", "Cache-Control": "no-cache" };

Bun.serve({
  port: PORT,
  hostname: HOST,
  // Decks and Excel files go through (a deck with pictures: tens of MB).
  maxRequestBodySize: 300 * 1024 * 1024,
  async fetch(req) {
    const url = new URL(req.url);
    // The exported app itself: a page, not a call of the app.
    if (req.method === "GET" && (url.pathname === "/" || url.pathname === "/hub.html")) {
      const no = refusal(req, false);
      const page = no ? undefined : await exportedApp();
      return page ? new Response(page, { headers: PAGE_HEADERS }) : new Response(no ?? "Hub local server. The app runs on http://localhost:3300 (bun dev); bun run export:html writes dist/hub.html, served here.", { status: no ? 403 : 200 });
    }
    const no = refusal(req);
    if (no) return new Response(JSON.stringify({ error: no }), { status: 403, headers: { "Content-Type": "application/json" } });
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(req) });
    try {
      if (url.pathname === "/get-token") {
        // Only the Power BI token leaves the server (the embedding SDK draws reports with it); Fabric and Graph are called through /ms.
        if ((url.searchParams.get("scope") ?? "powerbi") !== "powerbi") return json(req, { error: "Only the Power BI token is given to the app: call the other services through /ms/<service>/…" }, 403);
        const token = await accessToken("powerbi");
        return token ? json(req, { token: token.token, expiresAt: token.expiresAt, account: (await load()).account }) : signInAnswer(req);
      }
      if (url.pathname.startsWith("/ms/")) return await microsoft(req, url);
      if (url.pathname === "/confluence-settings") return await confluenceSettings(req);
      if (url.pathname.startsWith("/confluence/")) return await confluence(req, url);
      if (url.pathname === "/office/render" && req.method === "POST") return await office(req, url);
      if (url.pathname === "/office/warm" && req.method === "POST") return warmEngine(), json(req, { ok: true });
      if (url.pathname === "/sign-out" && req.method === "POST") {
        state = {};
        cache.clear();
        pending = null;
        await save();
        return json(req, { ok: true });
      }
      return json(req, { error: "Not found." }, 404);
    } catch (e) {
      return json(req, { error: e instanceof Error ? e.message : String(e) }, 500);
    }
  },
});
console.log(`Local server on http://localhost:${PORT} (answers the app only; session kept in ${await secretStore()})`);
