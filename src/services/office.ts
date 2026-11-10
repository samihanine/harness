/**
 * PowerPoint slides drawn as pictures by the local server's LibreOffice engine
 * (scripts/server/office.ts): the deck is sent to the server on this computer only, its slides come
 * back as PNG. `missing`: the engine is not installed there (the app then draws its own simple preview).
 */
import { server } from "./ms";

export class OfficeEngineError extends Error {
  constructor(message: string, readonly missing = false) {
    super(message);
  }
}

export type RenderedPage = { page: number; width: number; height: number; url: string };

export async function renderDeck(deck: Uint8Array, { width, pages, fonts = [] }: { width: number; pages?: number[]; fonts?: string[] }): Promise<{ count: number; pages: RenderedPage[] }> {
  const query = new URLSearchParams({ width: String(width), ...(pages?.length && { pages: pages.join(",") }), ...(fonts.length && { fonts: fonts.join("|") }) });
  const r = await server(`/office/render?${query}`, { method: "POST", headers: { "Content-Type": "application/octet-stream" }, body: deck as BodyInit });
  const body = await r.json().catch(() => undefined);
  if (!r.ok || !body?.pages) throw new OfficeEngineError(body?.error ?? `The preview engine answered ${r.status}.`, r.status === 501 || r.status === 404);
  return { count: body.count, pages: (body.pages as { page: number; width: number; height: number; png: string }[]).map((p) => ({ page: p.page, width: p.width, height: p.height, url: `data:image/png;base64,${p.png}` })) };
}

/** Starts the engine ahead of the first deck (it takes a moment to start: a second with Node, more than a minute with Bun alone). */
export const warmOffice = () => void server("/office/warm", { method: "POST" }).then((r) => r.text(), () => undefined);
