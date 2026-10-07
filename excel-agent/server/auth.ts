/**
 * Microsoft 365 sign-in (SharePoint / OneDrive files through Microsoft Graph) for a generic
 * tenant with the Microsoft Office public client (no app registration): device code once,
 * then refresh tokens renewed here. They live in .local/tokens.json (gitignored), on this machine only.
 */
import { mkdir, readFile, writeFile } from "node:fs/promises";

const TENANT = "organizations";
const CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"; // Microsoft Office
const LOGIN = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0`;
const STORE = ".local/tokens.json";

/** Token audiences: Microsoft Graph (files, Excel workbooks). */
export const SCOPES = {
  graph: "https://graph.microsoft.com/.default offline_access openid profile",
} as const;
export type Audience = keyof typeof SCOPES;

type Tokens = { access_token?: string; refresh_token?: string; expires_in?: number; error?: string; error_description?: string };
export type Pending = { userCode: string; verificationUri: string; expiresAt: number; error?: string };

let state: { refreshToken?: string; account?: string } = await readFile(STORE, "utf8")
  .then((text) => JSON.parse(text))
  .catch(() => ({}));
const access = new Map<Audience, { token: string; expiresAt: number }>();
let pending: Pending | null = null;

const save = async () => {
  await mkdir(".local", { recursive: true });
  await writeFile(STORE, JSON.stringify(state));
};

async function post(endpoint: string, params: Record<string, string>) {
  const response = await fetch(`${LOGIN}/${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
  });
  return (await response.json()) as Tokens & Record<string, unknown>;
}

/** Account name from the id/access token (display only). */
const accountOf = (token: string) => {
  try {
    const claims = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
    return String(claims.upn ?? claims.unique_name ?? claims.name ?? "");
  } catch {
    return "";
  }
};

function remember(audience: Audience, tokens: Tokens) {
  if (tokens.refresh_token) state.refreshToken = tokens.refresh_token;
  if (audience === "graph") state.account = accountOf(tokens.access_token!);
  access.set(audience, { token: tokens.access_token!, expiresAt: Date.now() + ((tokens.expires_in ?? 3600) - 300) * 1000 });
  void save();
}

/** Valid access token (cached, renewed with the refresh token); null when a sign-in is needed. */
export async function accessToken(audience: Audience = "graph") {
  const cached = access.get(audience);
  if (cached && cached.expiresAt > Date.now()) return { token: cached.token, expiresAt: cached.expiresAt };
  if (!state.refreshToken) return null;
  const tokens = await post("token", {
    client_id: CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: state.refreshToken,
    scope: SCOPES[audience],
  });
  if (!tokens.access_token) {
    if (audience === "graph" && /invalid_grant|interaction_required/.test(String(tokens.error))) state = {};
    return null;
  }
  remember(audience, tokens);
  return { token: tokens.access_token, expiresAt: access.get(audience)!.expiresAt };
}

export async function startLogin() {
  const device = await post("devicecode", { client_id: CLIENT_ID, scope: SCOPES.graph });
  if (!device.device_code) throw new Error(String(device.error_description ?? "Could not start the sign-in"));
  pending = {
    userCode: String(device.user_code),
    verificationUri: String(device.verification_uri ?? "https://microsoft.com/devicelogin"),
    expiresAt: Date.now() + Number(device.expires_in ?? 900) * 1000,
  };
  void poll(String(device.device_code), Number(device.interval ?? 5), pending.userCode);
  return pending;
}

async function poll(deviceCode: string, interval: number, code: string) {
  while (pending?.userCode === code && pending.expiresAt > Date.now()) {
    await new Promise((r) => setTimeout(r, interval * 1000));
    const tokens = await post("token", {
      client_id: CLIENT_ID,
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: deviceCode,
    });
    if (tokens.error === "authorization_pending") continue;
    if (tokens.error === "slow_down") {
      interval += 5;
      continue;
    }
    if (tokens.access_token) {
      remember("graph", tokens);
      pending = null;
    } else if (pending) pending.error = tokens.error_description ?? "Sign-in failed";
    return;
  }
}

export async function status() {
  if (pending && pending.expiresAt < Date.now()) pending = null;
  const token = await accessToken("graph").catch(() => null);
  return { signedIn: !!token, account: token ? state.account : undefined, pending };
}

export async function logout() {
  state = {};
  access.clear();
  pending = null;
  await save();
}
