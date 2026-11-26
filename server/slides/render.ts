/**
 * A presentation (templates × markdown content, shared/slides.ts) drawn on the server: pictures of its slides
 * and its PDF from the same HTML as the app (the headless browser: a second, no Microsoft call), and its
 * PowerPoint with native, editable shapes and texts (pptxgenjs). Pictures come from a public link, the template's
 * files (asset:…) or the app's files (file:…); a Power BI view is a picture of the report taken now.
 */
import PptxGenJS from "pptxgenjs";
import { parseLink } from "@shared/links";
import { deckHtml, parseSlides, resolveSlide, slideHtml, SLIDE, templateOf, type Template, type TextStyle } from "@shared/slides";
import { captureReport, rasterize, renderHtml, renderPdfOf } from "../capture/browser";
import { readStoredFile } from "../files";
import { staticFile } from "../static";
import { resolveReport } from "../pbi/resolve";

type Bytes = { data: Uint8Array; mime: string };
const pictures = new Map<string, Promise<Bytes | undefined>>();

const mimeOf = (b: Uint8Array, name = "") => (b[0] === 0x89 ? "image/png" : b[0] === 0xff ? "image/jpeg" : b[0] === 0x47 ? "image/gif" : /\.svg$/i.test(name) || new TextDecoder().decode(b.slice(0, 200)).includes("<svg") ? "image/svg+xml" : /\.webp$/i.test(name) ? "image/webp" : "image/png");

/** A picture's bytes from its source (kept for the session). */
export function pictureOf(src: string): Promise<Bytes | undefined> {
  if (!pictures.has(src)) {
    const p = (async () => {
      if (src.startsWith("asset:")) {
        const f = staticFile(`assets/template/${src.slice(6).replace(/[^\w.-]/g, "")}`);
        if (!f) return undefined;
        const data = new Uint8Array(await f.arrayBuffer());
        return { data, mime: mimeOf(data, src) };
      }
      if (src.startsWith("file:")) {
        const f = await readStoredFile(src.slice(5));
        return { data: f.data, mime: f.ref.type || mimeOf(f.data) };
      }
      if (/^https?:\/\//.test(src)) {
        const r = await fetch(src, { signal: AbortSignal.timeout(15_000) });
        if (!r.ok) return undefined;
        const data = new Uint8Array(await r.arrayBuffer());
        return { data, mime: (r.headers.get("content-type") ?? "").split(";")[0] || mimeOf(data, src) };
      }
      return undefined;
    })().catch(() => undefined);
    pictures.set(src, p);
  }
  return pictures.get(src)!;
}

const dataUrl = (b: Bytes) => `data:${b.mime};base64,${Buffer.from(b.data).toString("base64")}`;

/** A Power BI view of a slide (a report page or a visual, by its link): a picture taken now, kept 10 minutes. */
const views = new Map<string, { at: number; png: Promise<Buffer> }>();
export function powerbiPicture(url: string): Promise<Buffer> {
  const hit = views.get(url);
  if (hit && Date.now() - hit.at < 10 * 60_000) return hit.png;
  const png = (async () => {
    const l = parseLink(url);
    const id = l?.report ?? l?.bare;
    if (!id) throw new Error(`Not a Power BI report link: ${url}`);
    const r = await resolveReport({ id, groupId: l?.groupId, appId: l?.appId });
    return (await captureReport({ id: r.id, groupId: r.groupId, appId: r.appId }, { page: l?.page, visual: l?.visual, width: 1600, height: 900 })).png;
  })();
  views.set(url, { at: Date.now(), png });
  png.catch(() => views.delete(url));
  return png;
}

/** Every picture of the slides as data (for the HTML drawn on the server). */
async function inlined(markdown: string, templates: Template[]) {
  const slides = parseSlides(markdown);
  const src = new Map<string, string>();
  const pbi = new Map<string, string>();
  await Promise.all(
    slides.flatMap((c) =>
      resolveSlide(templateOf(templates, c.template), c).map(async (d) => {
        if (d.type === "image" && !src.has(d.src)) {
          const b = await pictureOf(d.src);
          if (b) src.set(d.src, dataUrl(b));
        }
        if (d.type === "powerbi" && !pbi.has(d.url)) {
          const png = await powerbiPicture(d.url).catch(() => undefined);
          if (png) pbi.set(d.url, `data:image/png;base64,${png.toString("base64")}`);
        }
      }),
    ),
  );
  return { slides, html: (n: number) => slideHtml(templateOf(templates, slides[n].template), slides[n], { src: (s) => src.get(s) ?? s, powerbi: (u) => pbi.get(u) }) };
}

/** Pictures of slides (1-based; all when none given). */
export async function slidePictures(markdown: string, templates: Template[], numbers?: number[], width = 1280) {
  const { slides, html } = await inlined(markdown, templates);
  const wanted = (numbers?.length ? numbers : slides.map((_, i) => i + 1)).filter((n) => n >= 1 && n <= slides.length);
  const out: Buffer[] = [];
  void width;
  for (const n of wanted) out.push(await renderHtml(deckHtml([html(n - 1)]), 1280, 720));
  return { pictures: out, numbers: wanted, count: slides.length };
}

/** The presentation as a PDF: one page per slide, as the app draws it. */
export async function deckPdf(markdown: string, templates: Template[]) {
  const { slides, html } = await inlined(markdown, templates);
  return renderPdfOf(deckHtml(slides.map((_, i) => html(i))), { width: "1280px", height: "720px" });
}

/* --------------------------------- PowerPoint --------------------------------- */

const hex = (c?: string) => (c ?? "#000000").replace("#", "").slice(0, 6).toUpperCase();

/** Inline markdown as pptxgenjs runs (bold, italic, links). */
function runs(text: string, base: Record<string, unknown>) {
  const out: { text: string; options: Record<string, unknown> }[] = [];
  const re = /\*\*(.+?)\*\*|(^|[^*])\*([^*]+)\*|\[([^\]]+)\]\(([^)]+)\)/g;
  let at = 0;
  for (const m of text.matchAll(re)) {
    const start = m.index! + (m[2]?.length ?? 0);
    if (start > at) out.push({ text: text.slice(at, start), options: { ...base } });
    if (m[1]) out.push({ text: m[1], options: { ...base, bold: true } });
    else if (m[3]) out.push({ text: m[3], options: { ...base, italic: true } });
    else out.push({ text: m[4], options: { ...base, hyperlink: { url: m[5] } } });
    at = m.index! + m[0].length;
  }
  if (at < text.length) out.push({ text: text.slice(at), options: { ...base } });
  return out.length ? out : [{ text: "", options: base }];
}

function textOptions(s: TextStyle) {
  return { fontFace: s.font ?? "Century Gothic", fontSize: s.size ?? 16, color: hex(s.color), bold: s.bold, italic: s.italic, align: s.align ?? "left", valign: s.valign ?? "top", charSpacing: s.spacing, margin: 0, lineSpacingMultiple: s.lineHeight ?? 1.1 };
}

export async function deckPptx(markdown: string, templates: Template[], title = "Presentation"): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE", width: SLIDE.w, height: SLIDE.h });
  pptx.layout = "WIDE";
  pptx.title = title;
  const slides = parseSlides(markdown);
  for (const c of slides) {
    const t = templateOf(templates, c.template);
    const slide = pptx.addSlide();
    if (t.background) slide.background = { color: hex(t.background) };
    for (const d of resolveSlide(t, c)) {
      const box = { x: d.x, y: d.y, w: Math.max(0.01, d.w), h: Math.max(0.01, d.h) };
      if (d.type === "shape") {
        if (d.shape === "line") slide.addShape("line", { ...box, h: 0, line: { color: hex(d.line?.color ?? d.fill), width: d.line?.width ?? 1 } });
        else slide.addShape(d.shape === "ellipse" ? "ellipse" : d.shape === "roundRect" ? "roundRect" : "rect", { ...box, fill: d.fill ? { color: hex(d.fill) } : { type: "none" } as never, line: d.line ? { color: hex(d.line.color), width: d.line.width ?? 1 } : { type: "none" } as never, ...(d.shape === "roundRect" && { rectRadius: Math.min(0.5, (d.radius ?? 0.15) / Math.min(box.w, box.h)) }) });
      } else if (d.type === "text") {
        const o = textOptions(d.style);
        if (d.fill) slide.addShape("rect", { ...box, fill: { color: hex(d.fill) }, line: { type: "none" } as never });
        const value = d.style.upper ? d.text.toUpperCase() : d.text;
        if (!d.rich) slide.addText(runs(value, {}) as never, { ...box, ...o, fit: "shrink" } as never);
        else {
          const paragraphs = value.split("\n").filter((l) => l.trim()).flatMap((line, i, all) => {
            const m = line.match(/^(\s*)(?:([-*+])|(\d+)[.)])\s+(.*)$/);
            const level = m ? Math.min(2, Math.floor(m[1].length / 2)) : 0;
            const parts = runs(m ? m[4] : line.trim(), {});
            parts[0].options = { ...parts[0].options, ...(m ? { bullet: m[3] ? { type: "number" } : { code: level ? "2013" : "25AA", color: hex("#B39251") }, indentLevel: level } : {}), paraSpaceAfter: 6 };
            if (i < all.length - 1) parts[parts.length - 1].options.breakLine = true;
            return parts;
          });
          slide.addText(paragraphs as never, { ...box, ...o, fit: "shrink" } as never);
        }
      } else if (d.type === "image") {
        const b = await pictureOf(d.src);
        if (!b) continue;
        const data = /svg|webp|avif/.test(b.mime) ? await rasterize(b.data, b.mime, 1600) : Buffer.from(b.data);
        const mime = /svg|webp|avif/.test(b.mime) ? "image/png" : b.mime;
        slide.addImage({ data: `data:${mime};base64,${data.toString("base64")}`, ...box, sizing: { type: d.fit === "contain" ? "contain" : "cover", w: box.w, h: box.h } });
      } else {
        const png = await powerbiPicture(d.url).catch(() => undefined);
        if (png) slide.addImage({ data: `data:image/png;base64,${png.toString("base64")}`, ...box, sizing: { type: "contain", w: box.w, h: box.h }, hyperlink: { url: d.url } });
        else slide.addText(`Power BI: ${d.url}`, { ...box, fontSize: 10, color: "71717A", align: "center", valign: "middle" });
      }
    }
    if (c.notes) slide.addNotes(c.notes);
  }
  return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
}

