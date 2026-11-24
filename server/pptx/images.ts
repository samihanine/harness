/** Pictures for PowerPoint: read from a URL, a server file (`file:<id>`), a data URL or the template's assets, with their size. */
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { readStoredFile } from "../files";

export type Picture = { data: string; mime: string; width: number; height: number };

/** Natural size of a PNG, JPEG or GIF (undefined for other formats). */
export function imageSize(b: Uint8Array): { width: number; height: number } | undefined {
  if (b[0] === 0x89 && b[1] === 0x50) return { width: (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19], height: (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23] };
  if (b[0] === 0x47 && b[1] === 0x49) return { width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i < b.length) {
      if (b[i] !== 0xff) return undefined;
      const marker = b[i + 1];
      const len = (b[i + 2] << 8) | b[i + 3];
      if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) return { height: (b[i + 5] << 8) | b[i + 6], width: (b[i + 7] << 8) | b[i + 8] };
      i += 2 + len;
    }
  }
  return undefined;
}

const mimeOf = (b: Uint8Array, fallback = "image/png") => (b[0] === 0x89 ? "image/png" : b[0] === 0xff ? "image/jpeg" : b[0] === 0x47 ? "image/gif" : fallback);

export const TEMPLATE_DIR = join(process.cwd(), "assets", "template");
const cache = new Map<string, Promise<Picture>>();

/** A picture by its source. Errors say what failed (the export puts a placeholder instead). */
export function loadPicture(src: string): Promise<Picture> {
  if (!cache.has(src)) {
    const p = (async () => {
      let bytes: Uint8Array;
      let mime: string | undefined;
      if (src.startsWith("data:")) {
        const m = src.match(/^data:([^;,]+)?(;base64)?,(.*)$/s);
        if (!m) throw new Error("Bad data URL");
        bytes = m[2] ? Buffer.from(m[3], "base64") : Buffer.from(decodeURIComponent(m[3]));
        mime = m[1];
      } else if (src.startsWith("file:")) {
        const f = await readStoredFile(src.slice(5));
        bytes = f.data;
        mime = f.ref.type;
      } else if (src.startsWith("template:")) {
        bytes = new Uint8Array(await readFile(join(TEMPLATE_DIR, src.slice(9).replace(/[^\w.-]/g, ""))));
      } else if (/^https?:\/\//.test(src)) {
        const r = await fetch(src, { signal: AbortSignal.timeout(15_000), headers: { "User-Agent": "Mozilla/5.0 (Hub presenter)" } });
        if (!r.ok) throw new Error(`HTTP ${r.status} for ${src}`);
        bytes = new Uint8Array(await r.arrayBuffer());
        mime = r.headers.get("content-type") ?? undefined;
      } else throw new Error(`Unknown picture source "${src}" (use https://…, file:<id> or data:)`);
      const size = imageSize(bytes);
      if (!size) throw new Error(`Unsupported picture format for "${src.slice(0, 80)}" (PNG, JPEG or GIF)`);
      const type = mimeOf(bytes, mime);
      return { data: `${type};base64,${Buffer.from(bytes).toString("base64")}`, mime: type, ...size };
    })();
    cache.set(src, p);
    p.catch(() => cache.delete(src));
    // Keep the cache small.
    if (cache.size > 200) cache.delete(cache.keys().next().value!);
  }
  return cache.get(src)!;
}
