/**
 * Microsoft sign-in, on the server: one device-code sign-in with the Microsoft Office public client (no
 * app registration, no admin consent, any tenant). The refresh token stays in .local/tokens.json; access
 * tokens for Power BI, Fabric and Graph are asked from it when needed. The browser only receives a
 * Power BI token to embed reports (the embedding SDK needs it).
 */
import { chmod, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { DATA_DIR } from "../settings";

const CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office (public client)
const LOGIN = "https://login.microsoftonline.com/organizations/oauth2/v2.0";
const STORE = join(DATA_DIR, "tokens.json");
export const SCOPES = {
  powerbi: "https://analysis.windows.net/powerbi/api/.default",
  fabric: "https://api.fabric.microsoft.com/.default",
  graph: "https://graph.microsoft.com/.default",
} as const;
export type Service = keyof typeof SCOPES;

type Pending = { userCode: string; verificationUri: string; expiresAt: number; error?: string };
let state: { refreshToken?: string; account?: string } | undefined;
let pending: Pending | null = null;
const cache = new Map<Service, { token: string; expiresAt: number }>();

const load = async () => (state ??= await readFile(STORE, "utf8").then(JSON.parse).catch(() => ({})))!;
const save = async () => {
  await mkdir(DATA_DIR, { recursive: true });
  await writeFile(STORE, JSON.stringify(state));
  await chmod(STORE, 0o600).catch(() => undefined);
};
const post = async (endpoint: string, params: Record<string, string>) => (await (await fetch(`${LOGIN}/${endpoint}`, { method: "POST", body: new URLSearchParams(params) })).json()) as Record<string, any>;
const scopeOf = (s: Service) => `${SCOPES[s]} offline_access openid profile`;

export class SignInRequired extends Error {
  constructor() {
    super("Not signed in to Microsoft: sign in from Settings › Power BI.");
  }
}

/** An access token for a service (throws SignInRequired when nobody is signed in). */
export async function accessToken(service: Service): Promise<string> {
  const hit = cache.get(service);
  if (hit && hit.expiresAt > Date.now()) return hit.token;
  const s = await load();
  if (!s.refreshToken) throw new SignInRequired();
  const t = await post("token", { client_id: CLIENT_ID, grant_type: "refresh_token", refresh_token: s.refreshToken, scope: scopeOf(service) });
  if (!t.access_token) {
    if (/invalid_grant|interaction_required/.test(t.error)) throw new SignInRequired();
    throw new Error(t.error_description ?? t.error ?? "Token refused");
  }
  if (t.refresh_token && t.refresh_token !== s.refreshToken) (s.refreshToken = t.refresh_token), void save();
  cache.set(service, { token: t.access_token, expiresAt: Date.now() + (t.expires_in - 300) * 1000 });
  return t.access_token;
}

/** The Power BI token and its expiry, for the browser's embeds. */
export async function embedToken() {
  const token = await accessToken("powerbi");
  return { token, expiresAt: cache.get("powerbi")!.expiresAt };
}

export async function authStatus(): Promise<{ signedIn: boolean; account?: string; signIn?: Pending; error?: string }> {
  try {
    await accessToken("powerbi");
    return { signedIn: true, account: (await load()).account };
  } catch (e) {
    if (e instanceof SignInRequired) return { signedIn: false, signIn: pending && pending.expiresAt > Date.now() ? pending : undefined };
    return { signedIn: false, error: e instanceof Error ? e.message : String(e) };
  }
}

/** Starts a device-code sign-in (one at a time); it finishes in the background. */
export async function signIn() {
  if (pending && pending.expiresAt > Date.now() && !pending.error) return pending;
  const d = await post("devicecode", { client_id: CLIENT_ID, scope: scopeOf("powerbi") });
  if (!d.device_code) throw new Error(d.error_description ?? "Could not start the sign-in");
  const p: Pending = (pending = { userCode: d.user_code, verificationUri: d.verification_uri, expiresAt: Date.now() + d.expires_in * 1000 });
  void (async () => {
    let interval = Number(d.interval ?? 5);
    while (pending === p && p.expiresAt > Date.now()) {
      await new Promise((r) => setTimeout(r, interval * 1000));
      const t = await post("token", { client_id: CLIENT_ID, grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: d.device_code }).catch(() => ({ error: "authorization_pending" }) as Record<string, any>);
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
}

export async function signOut() {
  state = {};
  cache.clear();
  pending = null;
  await save();
}
