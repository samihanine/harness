/** Calls to the Hub server (/api): JSON in and out, errors as readable messages. */
import type { FileRef } from "@shared/types";

export class ApiError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
  }
}

export async function api<T = any>(path: string, init: RequestInit & { json?: unknown } = {}): Promise<T> {
  const { json, ...rest } = init;
  const r = await fetch(`/api${path}`, { ...rest, ...(json !== undefined ? { method: rest.method ?? "POST", headers: { "Content-Type": "application/json", ...(rest.headers ?? {}) }, body: JSON.stringify(json) } : {}) });
  const type = r.headers.get("content-type") ?? "";
  if (!r.ok) {
    const body = type.includes("json") ? await r.json().catch(() => ({})) : { error: await r.text() };
    throw new ApiError(body.error ?? `Error ${r.status}`, r.status);
  }
  return (type.includes("json") ? r.json() : r.blob()) as Promise<T>;
}

export const get = <T = any>(path: string) => api<T>(path);
export const post = <T = any>(path: string, json: unknown) => api<T>(path, { json });
export const put = <T = any>(path: string, json: unknown) => api<T>(path, { method: "PUT", json });
export const patch = <T = any>(path: string, json: unknown) => api<T>(path, { method: "PATCH", json });
export const del = <T = any>(path: string) => api<T>(path, { method: "DELETE" });

/** A file sent to the server (kept there a few days); `id` to send the same file again under its id. */
export async function upload(file: Blob, name: string, id?: string): Promise<FileRef> {
  const form = new FormData();
  form.append("file", file, name);
  if (id) form.append("id", id);
  return api<FileRef>("/files", { method: "POST", body: form });
}

export const fileUrl = (id: string) => `/api/files/${id}`;

/** Downloads a blob as a file. */
export function download(blob: Blob, name: string) {
  const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10_000);
}

export const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));
