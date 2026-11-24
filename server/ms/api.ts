/** The Microsoft REST APIs (Power BI, Fabric, Graph), called by the server with the signed-in user's tokens. */
import { accessToken, type Service } from "./auth";

const BASES: Record<Service, string> = {
  powerbi: "https://api.powerbi.com/v1.0/myorg",
  fabric: "https://api.fabric.microsoft.com/v1",
  graph: "https://graph.microsoft.com/v1.0",
};

/** The useful message of a Microsoft error (Power BI puts it under "pbi.error", names wrapped in <oii> tags). */
function errorOf(service: Service, status: number, text: string, body: any) {
  const e = body?.error;
  const pbiError = e?.["pbi.error"] ?? e?.pbi?.error;
  const detail = pbiError?.details?.find((d: { code?: string }) => d.code === "DetailsMessage")?.detail?.value ?? pbiError?.details?.[0]?.detail?.value;
  const message = detail ?? e?.message ?? body?.message ?? `${service} error ${status}: ${String(text).slice(0, 600)}`;
  return Object.assign(new Error(String(message).replace(/<\/?(oii|ccon)>/g, "")), { status });
}

/** A raw request with the service's token (absolute URL or path relative to its base). */
export async function request(service: Service, path: string, init: RequestInit = {}) {
  const url = path.startsWith("http") ? path : `${BASES[service]}${path}`;
  const token = await accessToken(service);
  return fetch(url, { ...init, headers: { Authorization: `Bearer ${token}`, ...(init.headers ?? {}) } });
}

/** One JSON call. Throttling is retried; Fabric long-running operations (202 + Location) are awaited. */
export async function call<T = any>(service: Service, path: string, method = "GET", body?: unknown, signal?: AbortSignal): Promise<T> {
  const send = (u: string, m = "GET", b?: unknown) => request(service, u, { method: m, signal, ...(b !== undefined && { headers: { "Content-Type": "application/json" }, body: JSON.stringify(b) }) });
  let response = await send(path, method, body);
  const safe = method === "GET" || method === "PATCH";
  for (let i = 0; i < 4 && (response.status === 429 || (safe && [409, 500, 502, 503, 504].includes(response.status))); i++) {
    await new Promise((r) => setTimeout(r, (Number(response.headers.get("retry-after")) || 2 ** i) * 1000));
    response = await send(path, method, body);
  }
  const location = response.headers.get("location");
  if (response.status === 202 && location && service === "fabric") {
    for (let i = 0; i < 90; i++) {
      await new Promise((r) => setTimeout(r, i < 5 ? 1000 : 2000));
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

export const scope = (groupId?: string) => (groupId ? `/groups/${groupId}` : "");

/** Fabric needs a workspace id: "My workspace" has one too. */
let personal: string | undefined;
export async function workspaceId(groupId?: string) {
  if (groupId) return groupId;
  if (!personal) personal = (await fabric("/workspaces")).value.find((w: { type: string }) => w.type === "Personal").id as string;
  return personal!;
}

export const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));
