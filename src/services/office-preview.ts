/**
 * A PowerPoint deck drawn by Microsoft Office itself: the deck is copied into a work folder of the
 * user's own OneDrive ("Hub previews", one file per deck, overwritten at each change) and converted to
 * a PDF by Microsoft 365 (Graph: the file's content with ?format=pdf) — the same rendering PowerPoint
 * gives, fonts, themes, SmartArt and charts included, whether the deck came from this computer or
 * from SharePoint. The calls go through the local server like every Microsoft call (services/ms.ts).
 * The copy stays in the user's Microsoft 365 account; it is deleted when the deck leaves the app.
 */
import { request } from "./ms";

const FOLDER = "Hub previews";
const PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const itemPath = (id: string) => `/me/drive/root:/${encodeURIComponent(FOLDER)}/${encodeURIComponent(id)}.pptx:`;

/** `unavailable`: this account cannot do it at all (no OneDrive, no right, a blocked tenant): not tried again in this session. */
export class OfficePreviewError extends Error {
  constructor(message: string, readonly unavailable = false) {
    super(message);
  }
}

async function failure(r: Response, what: string) {
  const body = await r.json().catch(() => undefined);
  const detail = body?.error?.message ?? body?.error ?? `HTTP ${r.status}`;
  // Not allowed, no OneDrive, nothing to convert with: the account cannot use this; a passing failure (429, 5xx) can.
  return new OfficePreviewError(`Microsoft could not ${what}: ${String(detail).slice(0, 200)}`, [400, 401, 403, 404].includes(r.status) && !/throttl|timeout|temporar/i.test(String(detail)));
}

/** The deck as the PDF Microsoft Office draws from it. */
export async function officePdf(id: string, deck: Uint8Array): Promise<Uint8Array> {
  const up = await request("graph", `${itemPath(id)}/content?@microsoft.graph.conflictBehavior=replace`, { method: "PUT", headers: { "Content-Type": PPTX }, body: deck as BodyInit });
  if (!up.ok) throw await failure(up, "keep a work copy in your OneDrive");
  // The conversion is asked right after the upload: Microsoft 365 may need a moment to index a new file.
  let last: Response | undefined;
  for (let attempt = 0; attempt < 4; attempt++) {
    last = await request("graph", `${itemPath(id)}/content?format=pdf`);
    if (last.ok) return new Uint8Array(await last.arrayBuffer());
    if (![404, 409, 423, 429, 500, 502, 503, 504].includes(last.status)) break;
    await new Promise((r) => setTimeout(r, 1500 * (attempt + 1)));
  }
  throw await failure(last!, "draw the deck");
}

/** Deletes the work copy of a deck (best effort: nothing to say when it is already gone). */
export const deleteOfficeCopy = (id: string) => void request("graph", itemPath(id), { method: "DELETE" }).then((r) => r.text(), () => undefined);
