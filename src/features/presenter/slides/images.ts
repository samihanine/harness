/**
 * The images of a deck: given by a public https link or inline ("data:image/…;base64,…") in its JSON.
 * Read once in the browser before a PowerPoint export.
 */
import type { Deck } from ".";

/* -------------------------------- PowerPoint -------------------------------- */

/**
 * The images of a PowerPoint export, read before it: each drawn once in the browser into a PNG (an SVG
 * logo too) with its real size. PowerPoint then gets the data — not a link it would load itself — and
 * the size keeps its proportions. An image the browser may not read (no CORS) is left out, the export
 * goes on.
 */
const ready = new Map<string, { data: string; width: number; height: number } | null>();

/** An image as a PNG data URL with its size, or null when it cannot be read. */
function rasterize(src: string) {
  return new Promise<{ data: string; width: number; height: number } | null>((resolve) => {
    const img = new Image();
    if (!src.startsWith("data:")) img.crossOrigin = "anonymous";
    img.onload = () => {
      // An SVG without a size reports 0 (or 150): drawn at a size good for a logo.
      const width = img.naturalWidth || 512;
      const height = img.naturalHeight || 512;
      const k = Math.min(4, 2000 / Math.max(width, height));
      const canvas = Object.assign(document.createElement("canvas"), { width: Math.round(width * k), height: Math.round(height * k) });
      try {
        canvas.getContext("2d")!.drawImage(img, 0, 0, canvas.width, canvas.height);
        // Logos keep their transparency (PNG); photos are much lighter in JPEG.
        const transparent = /\.(svg|png|gif)(\?|$)|^data:image\/(svg|png|gif)|simpleicons/i.test(src);
        resolve({ data: canvas.toDataURL(transparent ? "image/png" : "image/jpeg", 0.88).slice(5), width, height });
      } catch {
        resolve(null);
      }
    };
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

/** A picture drawn for the export under a key of its own (e.g. "mascot:trainer", "icon:shield-check"). */
export const setPptxPicture = (key: string, picture: { data: string; width: number; height: number } | undefined) => void ready.set(key, picture ?? null);

/** Reads the images of a deck for its export; returns those that could not be read. */
export async function preparePptxImages(deck: Deck) {
  const srcs = new Set<string>();
  const walk = (v: unknown, key = "") => {
    if (typeof v === "string" && (key === "image" || key === "logo") && v.trim()) srcs.add(v.trim());
    else if (Array.isArray(v)) v.forEach((x) => walk(x));
    else if (v && typeof v === "object") Object.entries(v).forEach(([k, x]) => walk(x, k));
  };
  deck.slides.forEach((s) => walk(s));
  await Promise.all([...srcs].filter((s) => !ready.has(s)).map(async (s) => ready.set(s, await rasterize(s))));
  return [...srcs].filter((s) => !ready.get(s));
}

/**
 * An image placed in a box of a PowerPoint slide (inches), in its proportions: "contain" fits it whole
 * (centered), "cover" fills the box and crops the rest. Undefined when the image could not be read.
 */
export function pptxImage(src: string, box: { x: number; y: number; w: number; h: number }, fit: "cover" | "contain" = "contain") {
  const image = ready.get(src.trim());
  if (!image) return undefined;
  const ratio = image.width / image.height;
  if (fit === "contain") {
    const w = Math.min(box.w, box.h * ratio);
    const h = w / ratio;
    return { data: image.data, x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h };
  }
  // Cover: the image's own size (only its proportions count) and the box it is cropped to.
  const k = Math.max(box.w / image.width, box.h / image.height);
  return { data: image.data, x: box.x, y: box.y, w: image.width * k, h: image.height * k, sizing: { type: "cover" as const, w: box.w, h: box.h } };
}
