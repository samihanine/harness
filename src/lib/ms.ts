/** Browser side of the Microsoft APIs: everything goes through the server functions (tokens never stored here). */
import { getToken as serverToken, ms } from "@/server/ms";
import type { MsRequest } from "@/server/ms";

const call = <T = any>(service: MsRequest["service"], path: string, method?: string, body?: unknown) =>
  ms({ data: { service, path, method, body } }) as Promise<T>;

export const pbi = <T = any>(path: string, method?: string, body?: unknown) => call<T>("powerbi", path, method, body);
export const fabric = <T = any>(path: string, method?: string, body?: unknown) => call<T>("fabric", path, method, body);
export const graph = <T = any>(path: string, method?: string, body?: unknown) => call<T>("graph", path, method, body);

let embedToken: { token: string; expiresAt: number } | undefined;
export async function powerBiToken() {
  if (!embedToken || embedToken.expiresAt < Date.now()) embedToken = await serverToken({ data: { service: "powerbi" } });
  return embedToken.token;
}

export const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));
