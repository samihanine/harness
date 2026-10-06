/** Tokens come from the local API (server/), which signs in and renews them. */
export class LoginRequiredError extends Error {
  constructor() {
    super("Sign in to Power BI first (Settings).");
  }
}

export type AuthStatus = {
  signedIn: boolean;
  account?: string;
  pending: { userCode: string; verificationUri: string; expiresAt: number; error?: string } | null;
};

const cache = new Map<string, { token: string; expiresAt: number }>();

export async function getToken(audience: "powerbi" | "fabric" = "powerbi") {
  const cached = cache.get(audience);
  if (cached && cached.expiresAt > Date.now()) return cached.token;
  const response = await fetch(`/api/token?audience=${audience}`);
  if (response.status === 401) throw new LoginRequiredError();
  if (!response.ok) throw new Error(`Token error (${response.status})`);
  const data = (await response.json()) as { token: string; expiresAt: number };
  cache.set(audience, data);
  return data.token;
}

const api = async <T>(path: string, method = "GET") => {
  const response = await fetch(`/api${path}`, { method });
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error ?? `API error ${response.status}`);
  return (await response.json()) as T;
};

export const getStatus = () => api<AuthStatus>("/status");
export const startLogin = () => api<AuthStatus["pending"]>("/login", "POST");
export const logout = async () => {
  cache.clear();
  await api("/logout", "POST");
};
