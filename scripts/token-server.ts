/**
 * The local token server: the app is front-end only, the Microsoft sign-in lives here. One device-code
 * sign-in with the Microsoft Office public client (no app registration); the refresh token stays in
 * .local/tokens.json. GET /get-token?scope=powerbi|fabric|graph returns an access token for that
 * service, or, when nobody is signed in, the URL and the code to sign in with (the sign-in then
 * finishes by itself).
 *
 *   bun scripts/token-server.ts          the server (port TOKEN_PORT, default 3301)
 *   bun scripts/token-server.ts login    signs in from the terminal and writes .local/access-tokens.json
 */
import { mkdir, readFile, writeFile } from "node:fs/promises";

const CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office (public client)
const LOGIN = "https://login.microsoftonline.com/organizations/oauth2/v2.0";
const STORE = ".local/tokens.json";
const PORT = Number(process.env.TOKEN_PORT ?? 3301);
const SCOPES = {
  powerbi: "https://analysis.windows.net/powerbi/api/.default",
  fabric: "https://api.fabric.microsoft.com/.default",
  graph: "https://graph.microsoft.com/.default",
} as const;
type Scope = keyof typeof SCOPES;

type Pending = { userCode: string; verificationUri: string; expiresAt: number; error?: string };
let state: { refreshToken?: string; account?: string } | undefined;
let pending: Pending | null = null;
const cache = new Map<Scope, { token: string; expiresAt: number }>();

const load = async () => (state ??= await readFile(STORE, "utf8").then(JSON.parse).catch(() => ({})))!;
const save = async () => (await mkdir(".local", { recursive: true }), writeFile(STORE, JSON.stringify(state)));
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
  const entry = { token: t.access_token as string, expiresAt: Date.now() + (t.expires_in - 300) * 1000, account: s.account };
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

/** Writes every service's access token to .local/access-tokens.json (for scripts and tests). */
async function writeAccessTokens() {
  const out: Record<string, unknown> = { account: (await load()).account };
  for (const s of Object.keys(SCOPES) as Scope[]) out[s] = await accessToken(s);
  await writeFile(".local/access-tokens.json", JSON.stringify(out, null, 2));
  return out;
}

if (process.argv[2] === "login") {
  if (!(await accessToken("powerbi"))) {
    const p = (await signIn()) as Pending & { done: Promise<boolean> };
    console.log(`Sign in: open ${p.verificationUri} and enter the code ${p.userCode}`);
    if (!(await p.done)) throw new Error(pending?.error ?? "Sign-in expired");
  }
  const out = await writeAccessTokens();
  console.log(`Signed in as ${out.account}. Tokens written to .local/access-tokens.json`);
  process.exit(0);
}

const cors = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "GET, POST, OPTIONS", "Access-Control-Allow-Headers": "*" };
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json", ...cors } });

Bun.serve({
  port: PORT,
  async fetch(req) {
    const url = new URL(req.url);
    if (req.method === "OPTIONS") return new Response(null, { headers: cors });
    try {
      if (url.pathname === "/get-token") {
        const scope = (url.searchParams.get("scope") ?? "powerbi") as Scope;
        if (!(scope in SCOPES)) return json({ error: `Unknown scope "${scope}": ${Object.keys(SCOPES).join(", ")}` }, 400);
        const token = await accessToken(scope);
        if (token) return json({ token: token.token, expiresAt: token.expiresAt, account: (await load()).account });
        const p = await signIn();
        return json({ signIn: { url: p.verificationUri, code: p.userCode, expiresAt: p.expiresAt, error: p.error } });
      }
      if (url.pathname === "/sign-out" && req.method === "POST") {
        state = {};
        cache.clear();
        pending = null;
        await save();
        return json({ ok: true });
      }
      return json({ error: "Not found. Routes: GET /get-token?scope=powerbi|fabric|graph, POST /sign-out" }, 404);
    } catch (e) {
      return json({ error: e instanceof Error ? e.message : String(e) }, 500);
    }
  },
});
console.log(`Token server on http://localhost:${PORT} (GET /get-token?scope=powerbi|fabric|graph)`);
