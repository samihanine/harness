/**
 * Presentations: the look and the content apart.
 * - A template (the look) is a slide's elements placed in inches on a 13.333 × 7.5 in slide: texts, shapes,
 *   pictures, Power BI views, and repeated items (agenda lines, key-figure cards). Elements are bound to slots
 *   the content fills (title, subtitle, body, image, powerbi, items…), or fixed (a logo, a band, a label).
 * - The content is markdown, one slide per `---` block: a `template:` line (and other `key: value` lines), then
 *   `# Title`, `## Subtitle`, text and bullets (the body), `![](https://…)` pictures, Power BI links (a report
 *   page, or a visual: …/reports/<id>/<page>/<visual>), list items `**Value** label` for repeated items.
 * The same resolved elements are drawn as HTML (the app, the AI's pictures, the PDF) and as PowerPoint shapes.
 */

export const SLIDE = { w: 13.333, h: 7.5, dpi: 96 };

export type TextStyle = { font?: string; size?: number; color?: string; bold?: boolean; italic?: boolean; align?: "left" | "center" | "right"; valign?: "top" | "middle" | "bottom"; upper?: boolean; spacing?: number; lineHeight?: number; bullets?: boolean };
export type Box = { x: number; y: number; w: number; h: number };
export type Element =
  | (Box & { type: "text"; slot?: string; text?: string; style?: TextStyle; fill?: string })
  | (Box & { type: "shape"; shape?: "rect" | "roundRect" | "ellipse" | "line"; fill?: string; line?: { color?: string; width?: number }; radius?: number })
  | (Box & { type: "image"; slot?: string; src?: string; fit?: "cover" | "contain" })
  | (Box & { type: "powerbi"; slot?: string })
  | (Box & { type: "items"; slot?: string; direction?: "row" | "column" | "grid"; columns?: number; gap?: number; max?: number; item: Element[] });
export type Template = { id: string; name: string; description?: string; favorite?: boolean; builtin?: boolean; background?: string; elements: Element[] };

export type Item = { value?: string; label?: string; text: string; index: string };
export type SlideContent = { template: string; fields: Record<string, string>; title?: string; subtitle?: string; body?: string; images: string[]; powerbi: string[]; items: Item[]; notes?: string; source: string };

/* ---------------------------------- Content ---------------------------------- */

const FIELD = /^([A-Za-z][\w-]*)\s*:\s*(.*)$/;
const PBI = /https:\/\/app\.powerbi\.com\/\S+/;

/** The slides' markdown blocks (split on `---` lines; a block of `key: value` lines belongs to the next one). */
export function slideBlocks(markdown: string): string[] {
  const parts = markdown.replace(/\r/g, "").split(/^---[ \t]*$/m);
  const out: string[] = [];
  let pending = "";
  for (const p of parts) {
    const lines = p.split("\n").filter((l) => l.trim());
    if (!lines.length) continue;
    if (lines.every((l) => FIELD.test(l.trim()) && !/^https?:/.test(l.trim()))) pending += `${p.trim()}\n`;
    else (out.push(`${pending}${p.trim()}`.trim()), (pending = ""));
  }
  if (pending) out.push(pending.trim());
  return out;
}

/** The deck's markdown from its blocks (each block's fields as its front matter). */
export function joinBlocks(blocks: string[]) {
  return blocks
    .map((b) => {
      const lines = b.split("\n");
      const head: string[] = [];
      while (lines.length && FIELD.test(lines[0].trim()) && !/^https?:/.test(lines[0].trim()) && !lines[0].startsWith("#")) head.push(lines.shift()!);
      return `---\n${head.join("\n")}${head.length ? "\n---\n" : ""}${lines.join("\n").trim()}`;
    })
    .join("\n\n");
}

/** One slide's content: its fields, title, subtitle, body, pictures, Power BI links, items, notes. */
export function slideContent(block: string): SlideContent {
  const lines = block.split("\n");
  const fields: Record<string, string> = {};
  while (lines.length && (FIELD.test(lines[0].trim()) || !lines[0].trim()) && !/^https?:/.test(lines[0].trim()) && !lines[0].startsWith("#")) {
    const m = lines.shift()!.trim().match(FIELD);
    if (m) fields[m[1].toLowerCase()] = m[2].trim();
  }
  let text = lines.join("\n");
  let notes: string | undefined;
  text = text.replace(/<!--\s*notes?:?([\s\S]*?)-->/i, (_, n) => ((notes = n.trim()), ""));
  const images: string[] = [];
  const powerbi: string[] = [];
  text = text.replace(/!\[[^\]]*\]\(([^)\s]+)[^)]*\)/g, (_, src) => (images.push(src), ""));
  text = text.replace(/\[[^\]]*\]\((https:\/\/app\.powerbi\.com[^)\s]+)\)/g, (_, u) => (powerbi.push(u), ""));
  text = text.replace(new RegExp(`^\\s*(${PBI.source})\\s*$`, "gm"), (_, u) => (powerbi.push(u), ""));
  let title: string | undefined;
  let subtitle: string | undefined;
  text = text.replace(/^#\s+(.+)$/m, (_, t) => ((title = t.trim()), ""));
  text = text.replace(/^##\s+(.+)$/m, (_, t) => ((subtitle = t.trim()), ""));
  // List items: for repeated elements ("**1,97 Md$** Revenue", "Agenda line").
  const items: Item[] = [];
  for (const m of text.matchAll(/^(?:[-*+]|\d+[.)])\s+(.+)$/gm)) {
    const line = m[1].trim();
    const b = line.match(/^\*\*(.+?)\*\*\s*[:—–-]?\s*(.*)$/);
    items.push({ value: b?.[1], label: b ? b[2] : undefined, text: line, index: String(items.length + 1).padStart(2, "0") });
  }
  const body = text.replace(/\n{3,}/g, "\n\n").trim();
  return { template: fields.template ?? fields.layout ?? "content", fields, title: fields.title ?? title, subtitle: fields.subtitle ?? subtitle, body: body || undefined, images: [...(fields.image ? [fields.image] : []), ...images], powerbi: [...(fields.powerbi ? [fields.powerbi] : []), ...powerbi], items, notes: fields.notes ?? notes, source: block };
}

export const parseSlides = (markdown: string) => slideBlocks(markdown).map(slideContent);

/* ------------------------------ Template × content ------------------------------ */

/** A concrete element of a slide: what is drawn (its text as inline markdown, its picture, its Power BI link). */
export type Drawn =
  | (Box & { type: "text"; text: string; style: TextStyle; fill?: string; rich: boolean })
  | (Box & { type: "shape"; shape: string; fill?: string; line?: { color?: string; width?: number }; radius?: number })
  | (Box & { type: "image"; src: string; fit: "cover" | "contain" })
  | (Box & { type: "powerbi"; url: string });

/** The value of a slot from the content (`image2`: the second picture, `items`: the list). */
function slotValue(c: SlideContent, slot: string): string | undefined {
  const n = Number(slot.match(/(\d+)$/)?.[1] ?? 1) - 1;
  const base = slot.replace(/\d+$/, "");
  if (base === "image") return c.images[n];
  if (base === "powerbi") return c.powerbi[n];
  if (slot === "title") return c.title;
  if (slot === "subtitle") return c.subtitle;
  if (slot === "body") return c.body;
  if (slot === "notes") return c.notes;
  return c.fields[slot.toLowerCase()];
}

/** The elements of a slide: the template's, bound to the content (empty slots left out, items repeated). */
export function resolveSlide(t: Template, c: SlideContent): Drawn[] {
  const out: Drawn[] = [];
  // With repeated items on the template, the list lines are the items, not the body.
  const listed = t.elements.some((e) => e.type === "items");
  const body = listed ? c.body?.replace(/^(?:[-*+]|\d+[.)])\s+.+$/gm, "").replace(/\n{3,}/g, "\n\n").trim() : c.body;
  const add = (e: Element, item?: Item, dx = 0, dy = 0) => {
    const box = { x: e.x + dx, y: e.y + dy, w: e.w, h: e.h };
    if (e.type === "shape") return out.push({ ...box, type: "shape", shape: e.shape ?? "rect", fill: e.fill, line: e.line, radius: e.radius });
    if (e.type === "text") {
      const value = e.slot ? (item ? (item as Record<string, string | undefined>)[e.slot] : e.slot === "body" ? body : slotValue(c, e.slot)) : e.text;
      if (!value?.trim()) return;
      return out.push({ ...box, type: "text", text: value, style: e.style ?? {}, fill: e.fill, rich: e.slot === "body" || e.slot === "text" });
    }
    if (e.type === "image") {
      const src = e.slot ? slotValue(c, e.slot) : e.src;
      if (src) out.push({ ...box, type: "image", src, fit: e.fit ?? "cover" });
      return;
    }
    if (e.type === "powerbi") {
      const url = slotValue(c, e.slot ?? "powerbi");
      if (url) out.push({ ...box, type: "powerbi", url });
      return;
    }
    if (e.type === "items") {
      const list = c.items.slice(0, e.max ?? 12);
      if (!list.length) return;
      const gap = e.gap ?? 0.2;
      const dir = e.direction ?? "column";
      const cols = dir === "row" ? list.length : dir === "grid" ? Math.max(1, e.columns ?? 3) : 1;
      const rows = Math.ceil(list.length / cols);
      const cw = (e.w - gap * (cols - 1)) / cols;
      const ch = dir === "row" ? e.h : (e.h - gap * (rows - 1)) / rows;
      list.forEach((it, i) => {
        const cx = e.x + (i % cols) * (cw + gap);
        const cy = e.y + Math.floor(i / cols) * (ch + gap);
        // Inner elements are placed in a cell of 1 × 1: their boxes are fractions of the cell.
        for (const inner of e.item) add({ ...inner, x: inner.x * cw, y: inner.y * ch, w: inner.w * cw, h: inner.h * ch } as Element, it, cx, cy);
      });
    }
  };
  for (const e of t.elements) add(e);
  return out;
}

/* ------------------------------------ HTML ------------------------------------ */

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
/** Inline markdown (bold, italic, code, links) as HTML. */
export const inlineHtml = (s: string) =>
  esc(s)
    .replace(/\*\*(.+?)\*\*/g, "<b>$1</b>")
    .replace(/(^|[^*])\*([^*]+)\*/g, "$1<i>$2</i>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<u>$1</u>');

/** A body's markdown (paragraphs, bullets) as HTML lines. */
function bodyHtml(md: string, style: TextStyle, accent: string) {
  return md
    .split(/\n/)
    .filter((l) => l.trim())
    .map((l) => {
      const m = l.match(/^(\s*)(?:[-*+]|(\d+)[.)])\s+(.*)$/);
      if (!m) return `<p style="margin:0 0 0.35em">${inlineHtml(l.trim())}</p>`;
      const level = Math.min(2, Math.floor(m[1].length / 2));
      const mark = m[2] ? `${m[2]}.` : level ? "–" : "■";
      return `<p style="margin:0 0 0.35em;padding-left:${1.1 + level * 1.1}em;text-indent:-1.1em"><span style="color:${level || m[2] ? "inherit" : accent};display:inline-block;width:1.1em;text-indent:0;font-size:${m[2] ? "1em" : "0.6em"};vertical-align:${m[2] ? "baseline" : "0.15em"}">${mark}</span>${inlineHtml(m[3])}</p>`;
    })
    .join("");
}

export type HtmlOptions = { src: (src: string) => string; powerbi: (url: string) => string | undefined; accent?: string; scale?: number };

/** One slide as HTML: a 1280 × 720 box (scaled by `scale`), its elements absolutely placed. */
export function slideHtml(t: Template, c: SlideContent, o: HtmlOptions) {
  const px = (v: number) => `${Math.round(v * SLIDE.dpi * 100) / 100}px`;
  const accent = o.accent ?? "#B39251";
  const parts = resolveSlide(t, c).map((d) => {
    const box = `position:absolute;left:${px(d.x)};top:${px(d.y)};width:${px(d.w)};height:${px(d.h)};`;
    if (d.type === "shape") {
      if (d.shape === "line") return `<div style="${box}border-top:${d.line?.width ?? 1}px solid ${d.line?.color ?? d.fill ?? "#000"};height:0"></div>`;
      return `<div style="${box}background:${d.fill ?? "transparent"};${d.line ? `border:${d.line.width ?? 1}px solid ${d.line.color ?? "#000"};` : ""}border-radius:${d.shape === "ellipse" ? "50%" : d.shape === "roundRect" ? px(d.radius ?? 0.15) : "0"}"></div>`;
    }
    if (d.type === "image") return `<img src="${esc(o.src(d.src))}" style="${box}object-fit:${d.fit}" alt="">`;
    if (d.type === "powerbi") {
      const pic = o.powerbi(d.url);
      return pic ? `<img src="${esc(pic)}" data-powerbi="${esc(d.url)}" style="${box}object-fit:contain;background:#fff" alt="">` : `<div data-powerbi="${esc(d.url)}" style="${box}background:#f4f4f5;border:1px dashed #d4d4d8;display:flex;align-items:center;justify-content:center;font:12px sans-serif;color:#71717a">Power BI view</div>`;
    }
    const s = d.style;
    const size = (s.size ?? 16) * (SLIDE.dpi / 72);
    const v = s.valign ?? "top";
    const content = d.rich ? bodyHtml(d.text, s, accent) : inlineHtml(d.text).replace(/\n/g, "<br>");
    return `<div style="${box}display:flex;flex-direction:column;justify-content:${v === "middle" ? "center" : v === "bottom" ? "flex-end" : "flex-start"};${d.fill ? `background:${d.fill};` : ""}overflow:hidden"><div style="font-family:${esc(s.font ?? "Century Gothic")},'Questrial','Avenir',sans-serif;font-size:${size}px;color:${s.color ?? "#000"};font-weight:${s.bold ? 700 : 400};font-style:${s.italic ? "italic" : "normal"};text-align:${s.align ?? "left"};${s.upper ? "text-transform:uppercase;" : ""}${s.spacing ? `letter-spacing:${s.spacing}px;` : ""}line-height:${s.lineHeight ?? 1.2}">${content}</div></div>`;
  });
  const scale = o.scale ?? 1;
  return `<div class="slide" style="position:relative;width:1280px;height:720px;overflow:hidden;background:${t.background ?? "#fff"};transform:scale(${scale});transform-origin:0 0">${parts.join("")}</div>`;
}

/** A whole document of slides (for pictures and the PDF): one 1280 × 720 page per slide. */
export function deckHtml(slides: string[]) {
  return `<!doctype html><html><head><meta charset="utf-8"><style>@page{size:1280px 720px;margin:0}body{margin:0}.page{width:1280px;height:720px;page-break-after:always;overflow:hidden}</style></head><body>${slides.map((s) => `<div class="page">${s}</div>`).join("")}</body></html>`;
}

/** The template a slide asks for (by id or name), else the first one. */
export const templateOf = (templates: Template[], name: string): Template => templates.find((t) => t.id === name) ?? templates.find((t) => t.name.toLowerCase() === name.toLowerCase()) ?? templates.find((t) => t.id === "content") ?? templates[0] ?? BLANK;
/** Drawn when no template exists at all: the title and the body. */
const BLANK: Template = { id: "blank", name: "Blank", elements: [{ type: "text", slot: "title", x: 0.6, y: 0.4, w: 12.13, h: 0.8, style: { size: 26 } }, { type: "text", slot: "body", x: 0.6, y: 1.5, w: 12.13, h: 5.4, style: { size: 16 } }] };

/** The slots a template takes (for the AI and the Visual tab). */
export function templateSlots(t: Template) {
  const slots = new Set<string>();
  const walk = (els: Element[]) => els.forEach((e) => (e.type === "items" ? (slots.add(`${e.slot ?? "items"} (list items: ${[...new Set(e.item.map((i) => ("slot" in i ? i.slot : undefined)).filter(Boolean))].join(", ")})`)) : "slot" in e && e.slot && slots.add(e.slot)));
  walk(t.elements);
  return [...slots];
}
