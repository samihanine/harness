/**
 * A real PowerPoint deck (any .pptx, no template required) read and changed in the browser through
 * its XML (lib/ooxml): the AI never sees the slides, it reads them as text — each shape with its id,
 * kind, place (in % of the slide), text and how much text fits — and changes them with small
 * operations written by code (texts, slides added from a template deck, pictures, order). Its raw XML
 * stays reachable for anything else. Everything not touched is kept as it was.
 */
import { OoxmlPackage, REL, all, element, kid, kids, path, relativePart, resolvePart, type XmlDoc, type XmlEl } from "@/lib/ooxml/package";
import { below, fit, freeRect, inset, inside, rightOf, type Box } from "./place";
import { colorOf, linesOf, paragraphsOf, styleRun, writeParagraphs, type Paragraph, type TextStyle } from "./text";

export type { Box } from "./place";
export type { Paragraph, TextStyle } from "./text";
export const PPTX_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";
const EMU_PER_PT = 12_700;

export type ShapeKind = "text" | "picture" | "table" | "chart" | "group" | "line" | "other";
export type Shape = {
  id: string;
  name: string;
  kind: ShapeKind;
  /** Placeholder type when it is one (title, body, subTitle, pic…). */
  placeholder?: string;
  box?: Box;
  paragraphs: Paragraph[];
  /** A table's cells, row by row. */
  rows?: string[][];
  /** Font size in points (the first run's, else the placeholder's usual). */
  size: number;
  /** Characters that fit in its box at that size (an estimate). */
  budget?: number;
  fill?: string;
  /** A picture's file in the package. */
  image?: string;
  /**
   * What an empty shape is for: a "frame" drawn around other shapes (a place the AI can name:
   * "inside #12"), or a "decoration" (a line, a colored square). Shapes with content have none.
   */
  role?: "frame" | "decoration";
  /** The frame it lies in (its id). */
  frame?: string;
};
export type SlideInfo = { number: number; part: string; layout: string; title: string; shapes: Shape[]; background?: string; notes: string };
/** Where a new element goes: in a frame or over a shape, under or beside a shape, or a box in % of the slide; nothing: the largest free area. */
export type Where = { inside?: string; below?: string; right_of?: string; box?: Box };

const num = (v: string | null | undefined) => (v ? Number(v) : 0);
const nonVisual = (el: XmlEl) => kids(el).find((c) => c.nodeName.startsWith("p:nv"));
const cNvPr = (el: XmlEl) => (nonVisual(el) ? kid(nonVisual(el)!, "p:cNvPr") : undefined);
const phOf = (el: XmlEl) => {
  const nv = nonVisual(el);
  const ph = nv && path(nv, "p:nvPr/p:ph");
  // (A missing attribute reads as "": || rather than ??.)
  return ph ? { type: ph.getAttribute("type") || "body", idx: ph.getAttribute("idx") || undefined } : undefined;
};
const xfrmOf = (el: XmlEl) => path(el, "p:spPr/a:xfrm") ?? path(el, "p:xfrm") ?? path(el, "p:grpSpPr/a:xfrm");
const boxOf = (xfrm: XmlEl | undefined): Box | undefined => {
  const off = xfrm && kid(xfrm, "a:off");
  const ext = xfrm && kid(xfrm, "a:ext");
  return off && ext ? { x: num(off.getAttribute("x")), y: num(off.getAttribute("y")), w: num(ext.getAttribute("cx")), h: num(ext.getAttribute("cy")) } : undefined;
};
/** The usual size of a placeholder's text (points) when the slide does not say. */
const DEFAULT_SIZE: Record<string, number> = { title: 36, ctrTitle: 44, subTitle: 20, body: 18 };

export class Deck {
  private constructor(
    readonly pkg: OoxmlPackage,
    /** Slide size in EMU. */
    readonly size: { w: number; h: number },
  ) {}

  static async open(data: ArrayBuffer | Uint8Array | Blob) {
    const pkg = await OoxmlPackage.open(data);
    if (!pkg.has("ppt/presentation.xml")) throw new Error("This file is not a PowerPoint presentation (.pptx).");
    const pres = (await pkg.xml("ppt/presentation.xml")).documentElement as XmlEl;
    const sz = kid(pres, "p:sldSz");
    return new Deck(pkg, { w: num(sz?.getAttribute("cx")) || 12_192_000, h: num(sz?.getAttribute("cy")) || 6_858_000 });
  }

  /* ---------------------------------- Reading ---------------------------------- */

  private async presentation() {
    return (await this.pkg.xml("ppt/presentation.xml")).documentElement as XmlEl;
  }
  /** The slides' parts, in their order. */
  async slideParts() {
    const pres = await this.presentation();
    const list = kid(pres, "p:sldIdLst");
    const rels = await this.pkg.rels("ppt/presentation.xml");
    return (list ? kids(list, "p:sldId") : []).map((s) => {
      const rel = rels.find((r) => r.id === s.getAttribute("r:id"));
      return resolvePart("ppt/presentation.xml", rel!.target);
    });
  }
  async count() {
    return (await this.slideParts()).length;
  }
  async partOf(number: number) {
    const parts = await this.slideParts();
    const part = parts[number - 1];
    if (!part) throw new Error(`No slide ${number}: the deck has ${parts.length} slides.`);
    return part;
  }
  private async layoutPart(slidePart: string) {
    const rel = (await this.pkg.rels(slidePart)).find((r) => r.type === REL.slideLayout);
    return rel ? resolvePart(slidePart, rel.target) : undefined;
  }
  private async masterPart(layoutPart: string) {
    const rel = (await this.pkg.rels(layoutPart)).find((r) => r.type === REL.slideMaster);
    return rel ? resolvePart(layoutPart, rel.target) : undefined;
  }
  async layoutName(slidePart: string) {
    const layout = await this.layoutPart(slidePart);
    return layout ? (kid((await this.pkg.xml(layout)).documentElement as XmlEl, "p:cSld")?.getAttribute("name") || "") : "";
  }

  /** A placeholder's box and text size as its layout (or master) gives them, when the slide does not. */
  private async inherited(slidePart: string, ph: { type: string; idx?: string }) {
    const layout = await this.layoutPart(slidePart);
    const find = async (part: string | undefined, byIdx: boolean) => {
      if (!part) return undefined;
      const tree = path((await this.pkg.xml(part)).documentElement as XmlEl, "p:cSld/p:spTree");
      return tree && kids(tree).find((s) => {
        const p = phOf(s);
        return p && (byIdx && ph.idx ? p.idx === ph.idx : p.type === ph.type || (ph.type === "ctrTitle" && p.type === "title"));
      });
    };
    const fromLayout = (await find(layout, true)) ?? (await find(layout, false));
    const fromMaster = layout ? await find(await this.masterPart(layout), false) : undefined;
    const box = boxOf(fromLayout && xfrmOf(fromLayout)) ?? boxOf(fromMaster && xfrmOf(fromMaster));
    const sizeOf = (el?: XmlEl) => {
      const sz = el && (all(el, "a:defRPr")[0]?.getAttribute("sz") || all(el, "a:rPr")[0]?.getAttribute("sz"));
      return sz ? num(sz) / 100 : undefined;
    };
    return { box, size: sizeOf(fromLayout) ?? sizeOf(fromMaster) };
  }

  private async shape(slidePart: string, el: XmlEl): Promise<Shape> {
    const pr = cNvPr(el);
    const ph = phOf(el);
    const kind: ShapeKind =
      el.nodeName === "p:pic" ? "picture" : el.nodeName === "p:grpSp" ? "group" : el.nodeName === "p:cxnSp" ? "line" : el.nodeName === "p:graphicFrame" ? (all(el, "a:tbl").length ? "table" : all(el, "c:chart").length ? "chart" : "other") : kid(el, "p:txBody") || ph ? "text" : "other";
    let box = boxOf(xfrmOf(el));
    let size: number | undefined;
    if (ph) {
      const inh = await this.inherited(slidePart, ph);
      box ??= inh.box;
      size = inh.size;
    }
    const rows = kind === "table" ? all(el, "a:tr").map((tr) => kids(tr, "a:tc").map((tc) => paragraphsOf(kid(tc, "a:txBody")).map((p) => p.marked).join(" / "))) : undefined;
    const paragraphs: Paragraph[] = rows ? rows.flat().map((text) => ({ text, marked: text, level: 0, bullet: false })) : paragraphsOf(kid(el, "p:txBody"));
    const pt = paragraphs.find((p) => p.size)?.size ?? size ?? DEFAULT_SIZE[ph?.type ?? ""] ?? 18;
    const budget = box && kind === "text" ? Math.max(10, Math.round((box.w / EMU_PER_PT / (pt * 0.52)) * (box.h / EMU_PER_PT / (pt * 1.25)))) : undefined;
    const blip = kind === "picture" ? all(el, "a:blip")[0]?.getAttribute("r:embed") : undefined;
    return {
      id: pr?.getAttribute("id") ?? "",
      name: pr?.getAttribute("name") ?? "",
      kind,
      placeholder: ph?.type,
      box,
      paragraphs,
      rows,
      size: pt,
      budget,
      fill: colorOf(path(el, "p:spPr")),
      image: blip ? await this.pkg.target(slidePart, blip) : undefined,
    };
  }

  /** The shapes of a slide-like part (slide, layout, master), groups flattened into their shapes. */
  async shapesOf(part: string, placeholders = true) {
    const tree = path((await this.pkg.xml(part)).documentElement as XmlEl, "p:cSld/p:spTree");
    const out: Shape[] = [];
    for (const el of tree ? kids(tree).filter((c) => /^p:(sp|pic|graphicFrame|grpSp|cxnSp)$/.test(c.nodeName)) : []) {
      if (!placeholders && phOf(el)) continue;
      out.push(await this.shape(part, el));
    }
    // Empty shapes: frames around others, or decoration; each shape knows the (smallest) frame it lies in.
    const empty = (s: Shape) => !s.placeholder && (s.kind === "text" || s.kind === "other" || s.kind === "line") && !s.paragraphs.some((p) => p.text.trim());
    for (const s of out) if (empty(s)) s.role = s.box && out.some((o) => o !== s && o.box && inside(o.box, s.box!) && o.box.w * o.box.h < s.box!.w * s.box!.h * 0.9) ? "frame" : "decoration";
    const frames = out.filter((s) => s.role === "frame").sort((a, b) => a.box!.w * a.box!.h - b.box!.w * b.box!.h);
    for (const s of out) s.frame = s.box ? frames.find((f) => f !== s && inside(s.box!, f.box!) && s.box!.w * s.box!.h < f.box!.w * f.box!.h * 0.9)?.id : undefined;
    return out;
  }

  private async notesOf(slidePart: string) {
    const rel = (await this.pkg.rels(slidePart)).find((r) => r.type === REL.notesSlide);
    if (!rel) return "";
    const notes = await this.pkg.xml(resolvePart(slidePart, rel.target));
    const body = all(notes, "p:sp").find((s) => phOf(s)?.type === "body");
    return paragraphsOf(body && kid(body, "p:txBody")).map((p) => p.text).join("\n").trim();
  }

  async slide(number: number): Promise<SlideInfo> {
    const part = await this.partOf(number);
    const shapes = await this.shapesOf(part);
    const title = shapes.find((s) => s.placeholder === "title" || s.placeholder === "ctrTitle")?.paragraphs.map((p) => p.text).join(" ") ?? shapes.find((s) => s.kind === "text" && s.paragraphs.some((p) => p.text.trim()))?.paragraphs[0]?.text ?? "";
    const bg = path((await this.pkg.xml(part)).documentElement as XmlEl, "p:cSld/p:bg/p:bgPr");
    return { number, part, layout: await this.layoutName(part), title: title.trim(), shapes, background: colorOf(bg), notes: await this.notesOf(part) };
  }

  /** What the slide's layout and master draw under it (their shapes that are not placeholders) and its background. */
  async chromeOf(number: number) {
    const part = await this.partOf(number);
    const layout = await this.layoutPart(part);
    const master = layout && (await this.masterPart(layout));
    const shapes = [...(master ? await this.shapesOf(master, false) : []), ...(layout ? await this.shapesOf(layout, false) : [])];
    const bgOf = async (p?: string) => (p ? colorOf(path((await this.pkg.xml(p)).documentElement as XmlEl, "p:cSld/p:bg/p:bgPr")) : undefined);
    return { shapes, background: (await bgOf(layout)) ?? (await bgOf(master)) };
  }

  async slides() {
    const out: SlideInfo[] = [];
    for (let i = 1; i <= (await this.count()); i++) out.push(await this.slide(i));
    return out;
  }

  /** A box in % of the slide, for the AI. */
  pct(b?: Box) {
    if (!b) return "inherited";
    const p = (v: number, of: number) => Math.round((v / of) * 100);
    return `x ${p(b.x, this.size.w)}% y ${p(b.y, this.size.h)}% w ${p(b.w, this.size.w)}% h ${p(b.h, this.size.h)}%`;
  }

  /** One shape as the AI reads it: id, kind, place, size, room, and its text (lines) or rows. */
  private shapeText(sh: Shape, full: boolean) {
    const head = `#${sh.id} ${sh.kind}${sh.placeholder ? `/${sh.placeholder}` : ""} "${sh.name}" (${this.pct(sh.box)}${sh.kind === "text" ? `, ${sh.size}pt${sh.budget ? `, room ≈ ${sh.budget} chars` : ""}` : ""}${sh.frame ? `, in frame #${sh.frame}` : ""})`;
    if (sh.rows) return `${head}: ${sh.rows.length} rows × ${sh.rows[0]?.length ?? 0} columns\n${sh.rows.map((r) => `      | ${r.join(" | ")} |`).join("\n")}`;
    const lines = linesOf(sh.paragraphs).filter((l, i, list) => l.trim() || (i > 0 && i < list.length - 1));
    if (!lines.some((l) => l.trim())) return head;
    return full ? `${head}:\n${lines.map((l) => `      ${l}`).join("\n")}` : `${head}: ${lines.join(" ⏎ ").slice(0, 160)}`;
  }

  /**
   * A slide as text for the AI: its shapes with content in reading order (top to bottom, left to
   * right), then its frames (empty boxes drawn around content: places to name), the decoration only
   * counted (`all`: listed too) — so the AI is not lost among empty rectangles.
   */
  describe(s: SlideInfo, { full = true, all: everything = false } = {}) {
    const order = (a: Shape, b: Shape) => (a.box?.y ?? 0) - (b.box?.y ?? 0) || (a.box?.x ?? 0) - (b.box?.x ?? 0);
    const content = s.shapes.filter((x) => !x.role).sort(order);
    const frames = s.shapes.filter((x) => x.role === "frame").sort(order);
    const decoration = s.shapes.filter((x) => x.role === "decoration");
    return [
      `Slide ${s.number} — layout "${s.layout}"${s.title ? ` — "${s.title}"` : ""}${content.length ? "" : " — empty (add_text, add_table, insert_visual put content on it)"}`,
      ...content.map((x) => `  · ${this.shapeText(x, full)}`),
      ...frames.map((f) => `  · frame #${f.id} (${this.pct(f.box)}) around ${s.shapes.filter((x) => x.frame === f.id && !x.role).map((x) => `#${x.id}`).join(" ") || "nothing"}`),
      everything ? decoration.map((x) => `  · decoration #${x.id} "${x.name}" (${this.pct(x.box)})`).join("\n") : decoration.length ? `  · ${decoration.length} decorative shape(s) without text not listed` : "",
      full && s.notes ? `  notes: ${s.notes}` : "",
    ].filter(Boolean).join("\n");
  }

  /** The whole deck as text for the AI (texts cut short; `full`: whole). */
  async outline(full = false) {
    return (await this.slides()).map((s) => this.describe(s, { full })).join("\n");
  }

  /* --------------------------------- Changing --------------------------------- */

  /** A shape by its id ("#12", "12"), its name, or — a placeholder — its role: "title", "subtitle", "body" / "content". */
  private async shapeEl(number: number, id: string) {
    const part = await this.partOf(number);
    const doc = await this.pkg.xml(part);
    const want = String(id).trim().replace(/^#/, "");
    const shapes = [...all(doc, "p:sp"), ...all(doc, "p:pic"), ...all(doc, "p:graphicFrame"), ...all(doc, "p:grpSp"), ...all(doc, "p:cxnSp")];
    const roles: Record<string, string[]> = { title: ["title", "ctrTitle"], subtitle: ["subTitle"], body: ["body", "obj"], content: ["body", "obj"], text: ["body", "obj"] };
    const el =
      shapes.find((e) => cNvPr(e)?.getAttribute("id") === want) ??
      shapes.find((e) => cNvPr(e)?.getAttribute("name")?.toLowerCase() === want.toLowerCase()) ??
      shapes.find((e) => roles[want.toLowerCase()]?.includes(phOf(e)?.type ?? ""));
    if (!el) throw new Error(`No shape #${want} on slide ${number}. Its shapes: ${shapes.map((e) => `#${cNvPr(e)?.getAttribute("id")}`).join(" ") || "none (an empty slide: add_text, add_table or insert_visual put content on it)"}.`);
    return { part, doc, el };
  }

  /**
   * Replaces a shape's text by lines (pptx/text.ts: one paragraph per line, two spaces deeper, "- " a
   * bullet, **bold**); each line takes the look of the paragraph it replaces. A table cell: `cell` [row, column] from 1.
   */
  async setText(number: number, id: string, text: string) {
    const { doc, el } = await this.shapeEl(number, id);
    if (el.nodeName === "p:graphicFrame" && all(el, "a:tbl").length) throw new Error(`#${id.replace(/^#/, "")} is a table: use set_table (its rows).`);
    let body = kid(el, "p:txBody");
    if (!body) {
      if (el.nodeName !== "p:sp") throw new Error(`Shape #${id.replace(/^#/, "")} holds no text (a ${el.nodeName.slice(2)}): add_text puts a text over or beside it.`);
      body = element(doc, "p:txBody", {}, [element(doc, "a:bodyPr"), element(doc, "a:lstStyle")]);
      el.appendChild(body);
    }
    writeParagraphs(doc, body, text);
  }

  /** Replaces a text everywhere (or on the given slides), also across the runs of a paragraph; returns how many. */
  async replaceText(find: string, replace: string, numbers?: number[]) {
    let count = 0;
    const parts = await this.slideParts();
    for (const [i, part] of parts.entries()) {
      if (numbers && !numbers.includes(i + 1)) continue;
      const doc = await this.pkg.xml(part);
      for (const p of all(doc, "a:p")) {
        const runs = kids(p, "a:r");
        const texts = runs.map((r) => kid(r, "a:t")?.textContent ?? "");
        const joined = texts.join("");
        if (!joined.includes(find)) continue;
        const inOne = texts.findIndex((t) => t.includes(find));
        if (inOne >= 0) {
          const t = kid(runs[inOne], "a:t")!;
          count += t.textContent!.split(find).length - 1;
          t.textContent = t.textContent!.split(find).join(replace);
        } else {
          // Across runs: the paragraph's text goes into its first run (its look kept).
          count += joined.split(find).length - 1;
          kid(runs[0], "a:t")!.textContent = joined.split(find).join(replace);
          for (const r of runs.slice(1)) p.removeChild(r);
        }
      }
    }
    return count;
  }

  /** Moves / resizes a shape (box in % of the slide). */
  async setBox(number: number, id: string, box: Box) {
    const { doc, el } = await this.shapeEl(number, id);
    const emu = { x: Math.round((box.x / 100) * this.size.w), y: Math.round((box.y / 100) * this.size.h), w: Math.round((box.w / 100) * this.size.w), h: Math.round((box.h / 100) * this.size.h) };
    let spPr = kid(el, "p:spPr") ?? kid(el, "p:grpSpPr");
    if (!spPr) el.insertBefore((spPr = element(doc, "p:spPr")), kids(el)[1] ?? null);
    let xfrm = kid(spPr, "a:xfrm") ?? kid(el, "p:xfrm");
    if (!xfrm) spPr.insertBefore((xfrm = element(doc, "a:xfrm")), spPr.firstChild);
    for (const c of kids(xfrm).filter((c) => c.nodeName === "a:off" || c.nodeName === "a:ext")) xfrm.removeChild(c);
    xfrm.appendChild(element(doc, "a:off", { x: emu.x, y: emu.y }));
    xfrm.appendChild(element(doc, "a:ext", { cx: emu.w, cy: emu.h }));
  }

  async deleteShape(number: number, id: string) {
    const { el } = await this.shapeEl(number, id);
    el.parentNode!.removeChild(el);
  }

  /** A free shape id on a slide. */
  private nextId(doc: XmlDoc) {
    return Math.max(1, ...all(doc, "p:cNvPr").map((e) => num(e.getAttribute("id")))) + 1;
  }
  private tree(doc: XmlDoc) {
    return path(doc.documentElement as XmlEl, "p:cSld/p:spTree")!;
  }
  private emu(box: Box): Box {
    return { x: Math.round((box.x / 100) * this.size.w), y: Math.round((box.y / 100) * this.size.h), w: Math.round((box.w / 100) * this.size.w), h: Math.round((box.h / 100) * this.size.h) };
  }

  /**
   * The place of a new element (EMU), computed here so the AI gives no coordinates: `inside` a frame
   * (its largest free area) or over a shape, `below` / `right_of` a shape (the room up to the next
   * one), a `box` in % — or nothing: the largest free area of the slide. `aspect` (width / height):
   * the element is fitted in that place in its proportions.
   */
  async placeFor(number: number, where: Where = {}, aspect?: number): Promise<Box> {
    const info = await this.slide(number);
    const margin = Math.round(this.size.w * 0.03);
    const gap = Math.round(this.size.w * 0.012);
    const slide = inset({ x: 0, y: 0, w: this.size.w, h: this.size.h }, margin);
    const find = (id: string) => {
      const want = String(id).trim().replace(/^#/, "").toLowerCase();
      const roles: Record<string, string[]> = { title: ["title", "ctrTitle"], subtitle: ["subTitle"], body: ["body", "obj"], content: ["body", "obj"] };
      const sh = info.shapes.find((x) => x.id === want) ?? info.shapes.find((x) => x.name.toLowerCase() === want) ?? info.shapes.find((x) => roles[want]?.includes(x.placeholder ?? ""));
      if (!sh?.box) throw new Error(`No shape #${want} with a place on slide ${number}. Its shapes: ${info.shapes.filter((x) => x.role !== "decoration").map((x) => `#${x.id}`).join(" ") || "none"}.`);
      return sh;
    };
    // What a new element must not cover: every shape with content.
    const obstacles = (but?: Shape) => info.shapes.filter((x) => x !== but && !x.role && x.box).map((x) => x.box!);
    let room: Box | undefined;
    if (where.box) room = this.emu(where.box);
    else if (where.inside) {
      const target = find(where.inside);
      const bounds = inset(target.box!, target.role === "frame" ? gap : 0);
      room = target.role === "frame" ? (freeRect(bounds, obstacles(target), { aspect }) ?? bounds) : bounds;
    } else if (where.below) {
      const ref = find(where.below);
      const bounds = ref.frame ? inset(find(ref.frame).box!, gap) : slide;
      room = below(ref.box!, bounds, obstacles(ref), gap);
    } else if (where.right_of) {
      const ref = find(where.right_of);
      const bounds = ref.frame ? inset(find(ref.frame).box!, gap) : slide;
      room = rightOf(ref.box!, bounds, obstacles(ref), gap);
    } else room = freeRect(slide, obstacles(), { aspect });
    if (!room || (!where.box && (room.w < this.size.w * 0.05 || room.h < this.size.h * 0.04)))
      throw new Error(`No room ${where.below ? `below #${where.below}` : where.right_of ? `right of #${where.right_of}` : where.inside ? `inside #${where.inside}` : `left on slide ${number}`}: give another place (inside a frame, below / right_of another shape, or a box in %), move or shorten shapes (set_box), or use a new slide.`);
    return aspect ? fit(room, aspect) : { x: Math.round(room.x), y: Math.round(room.y), w: Math.round(room.w), h: Math.round(room.h) };
  }

  /** The look of the texts around: of a shape (`like`), else of the slide's most common body text. */
  private async styleLike(number: number, like?: string): Promise<TextStyle> {
    const info = await this.slide(number);
    const texts = info.shapes.filter((x) => x.kind === "text" && x.paragraphs.some((p) => p.text.trim()));
    const ref = like ? texts.find((x) => x.id === like.replace(/^#/, "") || x.name === like) : undefined;
    // The body text: the smallest size written the most.
    const written = (x: Shape) => x.paragraphs.reduce((n, q) => n + q.text.length, 0);
    const body = ref ?? [...texts].sort((a, b) => written(b) - written(a))[0];
    const p = body?.paragraphs.find((x) => x.text.trim());
    return { size: p?.size ?? body?.size ?? 14, color: p?.color, font: p?.font ?? texts.flatMap((x) => x.paragraphs).find((x) => x.font)?.font };
  }

  /**
   * Adds a text box (lines as in setText). Its place is computed (placeFor), its look taken from a
   * shape (`like`) or from the slide's body text, unless `style` says otherwise. Returns its id.
   */
  async addText(number: number, text: string, where: Where = {}, { like, style }: { like?: string; style?: TextStyle } = {}) {
    const doc = await this.pkg.xml(await this.partOf(number));
    const look = { ...(await this.styleLike(number, like)), ...style };
    const room = await this.placeFor(number, where);
    // As high as its lines need at that size (not the whole free area), within the room.
    const lines = text.split("\n");
    const perLine = Math.max(8, Math.floor(room.w / EMU_PER_PT / ((look.size ?? 14) * 0.52)));
    const rows = lines.reduce((n, l) => n + Math.max(1, Math.ceil(l.length / perLine)), 0);
    const h = where.box ? room.h : Math.min(room.h, Math.round((rows * (look.size ?? 14) * 1.3 + 8) * EMU_PER_PT));
    const id = this.nextId(doc);
    const body = element(doc, "p:txBody", {}, [element(doc, "a:bodyPr", { wrap: "square", rtlCol: 0 }, [element(doc, "a:normAutofit")]), element(doc, "a:lstStyle")]);
    writeParagraphs(doc, body, text, look);
    this.tree(doc).appendChild(
      element(doc, "p:sp", {}, [
        element(doc, "p:nvSpPr", {}, [element(doc, "p:cNvPr", { id, name: `TextBox ${id}` }), element(doc, "p:cNvSpPr", { txBox: 1 }), element(doc, "p:nvPr")]),
        element(doc, "p:spPr", {}, [element(doc, "a:xfrm", {}, [element(doc, "a:off", { x: room.x, y: room.y }), element(doc, "a:ext", { cx: room.w, cy: h })]), element(doc, "a:prstGeom", { prst: "rect" }, [element(doc, "a:avLst")]), element(doc, "a:noFill")]),
        body,
      ]),
    );
    return String(id);
  }

  /** Adds a plain rectangle (a frame, a colored band): fill and outline as #rrggbb, none: no fill, a thin gray outline. Returns its id. */
  async addBox(number: number, where: Where, { fill, line }: { fill?: string; line?: string } = {}) {
    const doc = await this.pkg.xml(await this.partOf(number));
    const room = await this.placeFor(number, where);
    const id = this.nextId(doc);
    const color = (c: string) => element(doc, "a:solidFill", {}, [element(doc, "a:srgbClr", { val: c.replace("#", "").toUpperCase() })]);
    const shape = element(doc, "p:sp", {}, [
      element(doc, "p:nvSpPr", {}, [element(doc, "p:cNvPr", { id, name: `Rectangle ${id}` }), element(doc, "p:cNvSpPr"), element(doc, "p:nvPr")]),
      element(doc, "p:spPr", {}, [element(doc, "a:xfrm", {}, [element(doc, "a:off", { x: room.x, y: room.y }), element(doc, "a:ext", { cx: room.w, cy: room.h })]), element(doc, "a:prstGeom", { prst: "rect" }, [element(doc, "a:avLst")]), fill ? color(fill) : element(doc, "a:noFill"), element(doc, "a:ln", { w: 9525 }, [color(line ?? (fill ? fill : "#A1A1AA"))])]),
    ]);
    // Under the content it frames: first in the drawing order after the group's own properties.
    const tree = this.tree(doc);
    tree.insertBefore(shape, kids(tree).find((c) => /^p:(sp|pic|graphicFrame|grpSp|cxnSp)$/.test(c.nodeName)) ?? null);
    return String(id);
  }

  /** Changes the look of a shape: its text (size in points, bold, italic, color, font, align) and its fill / outline (#rrggbb, "none"). */
  async setStyle(number: number, id: string, style: TextStyle & { fill?: string; line?: string }) {
    const { doc, el } = await this.shapeEl(number, id);
    const text: TextStyle = { size: style.size, bold: style.bold, italic: style.italic, color: style.color, font: style.font };
    if (Object.values(text).some((v) => v !== undefined)) for (const rPr of [...all(el, "a:rPr"), ...all(el, "a:endParaRPr")]) styleRun(doc, rPr, text);
    if (style.align)
      for (const p of all(el, "a:p")) {
        let pPr = kid(p, "a:pPr");
        if (!pPr) p.insertBefore((pPr = element(doc, "a:pPr")), p.firstChild);
        pPr.setAttribute("algn", style.align);
      }
    const spPr = kid(el, "p:spPr");
    const solid = (c: string) => element(doc, "a:solidFill", {}, [element(doc, "a:srgbClr", { val: c.replace("#", "").toUpperCase() })]);
    if (spPr && style.fill) {
      for (const f of ["a:noFill", "a:solidFill", "a:gradFill", "a:pattFill", "a:blipFill"]) for (const c of kids(spPr, f)) spPr.removeChild(c);
      spPr.insertBefore(style.fill === "none" ? element(doc, "a:noFill") : solid(style.fill), kids(spPr).find((c) => ["a:ln", "a:effectLst", "a:scene3d", "a:sp3d", "a:extLst"].includes(c.nodeName)) ?? null);
    }
    if (spPr && style.line) {
      for (const c of kids(spPr, "a:ln")) spPr.removeChild(c);
      spPr.insertBefore(element(doc, "a:ln", style.line === "none" ? {} : { w: 9525 }, [style.line === "none" ? element(doc, "a:noFill") : solid(style.line)]), kids(spPr).find((c) => ["a:effectLst", "a:scene3d", "a:sp3d", "a:extLst"].includes(c.nodeName)) ?? null);
    }
  }

  /* ----------------------------------- Tables ---------------------------------- */

  /** One table cell (its text as lines), in the look of a model cell when given. */
  private cell(doc: XmlDoc, text: string, model?: XmlEl, style?: TextStyle) {
    const tc = model ? (model.cloneNode(true) as XmlEl) : element(doc, "a:tc", {}, [element(doc, "a:txBody", {}, [element(doc, "a:bodyPr"), element(doc, "a:lstStyle"), element(doc, "a:p")]), element(doc, "a:tcPr")]);
    for (const a of ["gridSpan", "rowSpan", "hMerge", "vMerge"]) tc.removeAttribute(a);
    writeParagraphs(doc, kid(tc, "a:txBody")!, text, model ? undefined : style);
    return tc;
  }

  /**
   * Replaces a table's content by rows of cells (the first row: its header). Rows are added or removed
   * as needed, each in the look of the row it replaces (the last body row for added ones); the number
   * of columns is the table's own.
   */
  async setTable(number: number, id: string, rows: string[][]) {
    const { doc, el } = await this.shapeEl(number, id);
    const tbl = all(el, "a:tbl")[0];
    if (!tbl) throw new Error(`#${id.replace(/^#/, "")} is not a table.`);
    const old = kids(tbl, "a:tr");
    const columns = kids(kid(tbl, "a:tblGrid")!, "a:gridCol").length;
    const wide = rows.find((r) => r.length > columns);
    if (wide) throw new Error(`The table has ${columns} columns: a row has ${wide.length}. Give ${columns} cells per row (or add_table for another table).`);
    const heightBefore = old.reduce((n, r) => n + num(r.getAttribute("h")), 0);
    for (const r of old) tbl.removeChild(r);
    rows.forEach((cells, i) => {
      const model = old[Math.min(i, old.length - 1)] ?? old[0];
      const tr = element(doc, "a:tr", { h: model?.getAttribute("h") || 370_840 });
      const cellsOf = model ? kids(model, "a:tc") : [];
      for (let c = 0; c < columns; c++) tr.appendChild(this.cell(doc, cells[c] ?? "", cellsOf[Math.min(c, cellsOf.length - 1)]));
      tbl.appendChild(tr);
    });
    // The frame follows its rows.
    const ext = path(el, "p:xfrm/a:ext");
    const heightAfter = kids(tbl, "a:tr").reduce((n, r) => n + num(r.getAttribute("h")), 0);
    if (ext && heightBefore) ext.setAttribute("cy", String(Math.round((num(ext.getAttribute("cy")) * heightAfter) / heightBefore)));
  }

  /**
   * Adds a table (rows of cells, the first row its header). `like`: a table of the slide whose look is
   * copied; otherwise a plain one in the slide's text style, header in bold on a light gray. Returns its id.
   */
  async addTable(number: number, rows: string[][], where: Where = {}, like?: string) {
    const part = await this.partOf(number);
    const doc = await this.pkg.xml(part);
    const columns = Math.max(1, ...rows.map((r) => r.length));
    const look = await this.styleLike(number);
    const size = Math.min(look.size ?? 12, 14);
    const rowH = Math.round((size * 1.3 + 12) * EMU_PER_PT);
    const room = await this.placeFor(number, where);
    const frame = { ...room, h: where.box ? room.h : Math.min(room.h, rowH * rows.length) };
    const id = this.nextId(doc);
    const model = like ? (await this.shapeEl(number, like)).el : all(doc, "p:graphicFrame").find((g) => all(g, "a:tbl").length);
    let graphic: XmlEl;
    if (model && all(model, "a:tbl").length) {
      // The same table, at the new place, then its content replaced.
      graphic = model.cloneNode(true) as XmlEl;
      const pr = cNvPr(graphic)!;
      pr.setAttribute("id", String(id));
      pr.setAttribute("name", `Table ${id}`);
      const xfrm = kid(graphic, "p:xfrm")!;
      kid(xfrm, "a:off")!.setAttribute("x", String(frame.x));
      kid(xfrm, "a:off")!.setAttribute("y", String(frame.y));
      const grid = kid(all(graphic, "a:tbl")[0], "a:tblGrid")!;
      const first = kids(grid, "a:gridCol")[0];
      for (const c of kids(grid, "a:gridCol")) grid.removeChild(c);
      for (let c = 0; c < columns; c++) {
        const col = first.cloneNode(true) as XmlEl;
        col.setAttribute("w", String(Math.round(frame.w / columns)));
        grid.appendChild(col);
      }
      kid(xfrm, "a:ext")!.setAttribute("cx", String(frame.w));
      this.tree(doc).appendChild(graphic);
    } else {
      const border = (name: string) => element(doc, name, { w: 6350 }, [element(doc, "a:solidFill", {}, [element(doc, "a:srgbClr", { val: "A1A1AA" })])]);
      const tbl = element(doc, "a:tbl", {}, [element(doc, "a:tblPr", { firstRow: 1 }), element(doc, "a:tblGrid", {}, Array.from({ length: columns }, () => element(doc, "a:gridCol", { w: Math.round(frame.w / columns) })))]);
      rows.forEach((_, i) => {
        const tr = element(doc, "a:tr", { h: rowH });
        for (let c = 0; c < columns; c++) {
          const tc = this.cell(doc, "", undefined, { ...look, size, bold: i === 0 });
          const tcPr = kid(tc, "a:tcPr")!;
          for (const b of ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]) tcPr.appendChild(border(b));
          tcPr.appendChild(i === 0 ? element(doc, "a:solidFill", {}, [element(doc, "a:srgbClr", { val: "E5E7EB" })]) : element(doc, "a:noFill"));
          tr.appendChild(tc);
        }
        tbl.appendChild(tr);
      });
      graphic = element(doc, "p:graphicFrame", {}, [
        element(doc, "p:nvGraphicFramePr", {}, [element(doc, "p:cNvPr", { id, name: `Table ${id}` }), element(doc, "p:cNvGraphicFramePr", {}, [element(doc, "a:graphicFrameLocks", { noGrp: 1 })]), element(doc, "p:nvPr")]),
        element(doc, "p:xfrm", {}, [element(doc, "a:off", { x: frame.x, y: frame.y }), element(doc, "a:ext", { cx: frame.w, cy: frame.h })]),
        element(doc, "a:graphic", {}, [element(doc, "a:graphicData", { uri: TABLE_URI }, [tbl])]),
      ]);
      this.tree(doc).appendChild(graphic);
    }
    await this.setTable(number, String(id), rows.map((r) => Array.from({ length: columns }, (_, c) => r[c] ?? "")));
    return String(id);
  }

  /* ---------------------------------- Pictures --------------------------------- */

  /**
   * Puts a picture (PNG) on a slide, never stretched: `replace` a shape (a picture keeps its frame,
   * the new one fitted in it; any other shape gives its place), or at a computed place (placeFor).
   * Returns its id.
   */
  async putPicture(number: number, png: Uint8Array, size: { width: number; height: number }, target: Where & { replace?: string } = {}, name = "Power BI visual") {
    const part = await this.partOf(number);
    const aspect = size.width / size.height;
    let frame: Box;
    if (target.replace) {
      const { el } = await this.shapeEl(number, target.replace);
      const sh = await this.shape(part, el);
      if (!sh.box) throw new Error(`Shape #${target.replace.replace(/^#/, "")} has no place on the slide.`);
      frame = fit(sh.box, aspect);
      el.parentNode!.removeChild(el);
    } else frame = await this.placeFor(number, target, aspect);
    const media = this.pkg.freeName("ppt/media/hub-image", "png");
    this.pkg.setBinary(media, png);
    await this.pkg.ensureDefault("png", "image/png");
    const rid = await this.pkg.addRel(part, REL.image, media);
    const doc = await this.pkg.xml(part);
    const id = this.nextId(doc);
    this.tree(doc).appendChild(
      element(doc, "p:pic", {}, [
        element(doc, "p:nvPicPr", {}, [element(doc, "p:cNvPr", { id, name: `${name} ${id}`.slice(0, 80) }), element(doc, "p:cNvPicPr", {}, [element(doc, "a:picLocks", { noChangeAspect: 1 })]), element(doc, "p:nvPr")]),
        element(doc, "p:blipFill", {}, [element(doc, "a:blip", { "r:embed": rid }), element(doc, "a:stretch", {}, [element(doc, "a:fillRect")])]),
        element(doc, "p:spPr", {}, [element(doc, "a:xfrm", {}, [element(doc, "a:off", { x: frame.x, y: frame.y }), element(doc, "a:ext", { cx: frame.w, cy: frame.h })]), element(doc, "a:prstGeom", { prst: "rect" }, [element(doc, "a:avLst")])]),
      ]),
    );
    return String(id);
  }

  /** The fonts the deck names (its texts' and its theme's), for the preview engine to load them from the computer. */
  async fonts() {
    const out = new Set<string>();
    for (const part of this.pkg.parts("ppt/").filter((p) => /^ppt\/(slides|slideLayouts|slideMasters|theme)\/[^/]+\.xml$/.test(p))) for (const m of (await this.pkg.text(part)).matchAll(/<a:(?:latin|ea|cs)\s+typeface="([^"+][^"]*)"/g)) out.add(m[1]);
    return [...out];
  }

  /* ---------------------------------- Slides ---------------------------------- */

  private async slideIds() {
    const pres = await this.presentation();
    let list = kid(pres, "p:sldIdLst");
    if (!list) {
      list = element(await this.pkg.xml("ppt/presentation.xml"), "p:sldIdLst");
      pres.insertBefore(list, kid(pres, "p:sldSz") ?? kid(pres, "p:notesSz") ?? null);
    }
    return list;
  }

  /** Registers a slide part in the deck at a position (1 = first; none: the end). */
  private async register(part: string, at?: number) {
    const list = await this.slideIds();
    const rid = await this.pkg.addRel("ppt/presentation.xml", REL.slide, part);
    const ids = kids(list, "p:sldId");
    const id = Math.max(255, ...ids.map((s) => num(s.getAttribute("id")))) + 1;
    const el = element(await this.pkg.xml("ppt/presentation.xml"), "p:sldId", { id, "r:id": rid });
    list.insertBefore(el, at && at <= ids.length ? ids[at - 1] : null);
    await this.pkg.setContentType(part, SLIDE_TYPE);
  }

  async deleteSlide(number: number) {
    const part = await this.partOf(number);
    const list = await this.slideIds();
    const sldId = kids(list, "p:sldId")[number - 1];
    const rid = sldId.getAttribute("r:id")!;
    list.removeChild(sldId);
    await this.pkg.removeRel("ppt/presentation.xml", rid);
    // Its notes go with it (nothing else points to them).
    const notes = (await this.pkg.rels(part)).find((r) => r.type === REL.notesSlide);
    if (notes) {
      const np = resolvePart(part, notes.target);
      this.pkg.remove(np);
      await this.pkg.removeContentType(np);
    }
    this.pkg.remove(part);
    await this.pkg.removeContentType(part);
  }

  async moveSlide(number: number, to: number) {
    const list = await this.slideIds();
    const ids = kids(list, "p:sldId");
    const el = ids[number - 1];
    if (!el) throw new Error(`No slide ${number}.`);
    list.removeChild(el);
    const rest = kids(list, "p:sldId");
    list.insertBefore(el, rest[Math.max(0, Math.min(to - 1, rest.length))] ?? null);
  }

  /**
   * Copies a slide — of this deck or of another one (a template) — to a position. From another deck,
   * its pictures come with it, and it takes this deck's layout of the same name (else of the same
   * type, else the first one); its placeholders get their template's places written on them, so it
   * looks the same whatever the layout. Charts and embedded objects are not copied (said in the result).
   */
  async copySlide(from: Deck, number: number, at?: number) {
    const src = await from.partOf(number);
    const part = this.pkg.freeName("ppt/slides/slide", "xml");
    const doc = (await from.pkg.xml(src)).cloneNode(true) as XmlDoc;
    const dropped: string[] = [];
    const srcRels = await from.pkg.rels(src);
    // Placeholders: their places and text sizes written on them (the layout here may differ).
    if (from !== this)
      for (const el of [...all(doc, "p:sp"), ...all(doc, "p:pic")]) {
        const ph = phOf(el);
        if (!ph || boxOf(xfrmOf(el))) continue;
        const { box } = await from.inherited(src, ph);
        let spPr = kid(el, "p:spPr");
        if (!spPr) el.insertBefore((spPr = element(doc, "p:spPr")), kids(el)[1] ?? null);
        if (box) spPr.insertBefore(element(doc, "a:xfrm", {}, [element(doc, "a:off", { x: box.x, y: box.y }), element(doc, "a:ext", { cx: box.w, cy: box.h })]), spPr.firstChild);
      }
    const remap = new Map<string, string>();
    for (const rel of srcRels) {
      if (rel.type === REL.notesSlide) continue;
      if (rel.external) {
        remap.set(rel.id, await this.pkg.addRel(part, rel.type, rel.target, true));
        continue;
      }
      const target = resolvePart(src, rel.target);
      if (rel.type === REL.slideLayout) {
        remap.set(rel.id, await this.pkg.addRel(part, rel.type, from === this ? target : await this.layoutFor(from, target)));
      } else if (rel.type === REL.image || /\/(media|image|audio|video)$/.test(rel.type)) {
        if (from === this) remap.set(rel.id, await this.pkg.addRel(part, rel.type, target));
        else {
          const ext = target.split(".").pop() ?? "png";
          const copy = this.pkg.freeName("ppt/media/hub-media", ext);
          this.pkg.setBinary(copy, await from.pkg.binary(target));
          const type = await from.pkg.contentType(target);
          if (type) await this.pkg.ensureDefault(ext.toLowerCase(), type);
          remap.set(rel.id, await this.pkg.addRel(part, rel.type, copy));
        }
      } else if (from === this) remap.set(rel.id, await this.pkg.addRel(part, rel.type, target));
      else dropped.push(rel.type.split("/").pop()!);
    }
    // References to relationships renamed; those of parts not copied: their shapes removed.
    for (const el of Array.from(doc.getElementsByTagName("*") as ArrayLike<XmlEl>)) {
      for (const a of Array.from(el.attributes)) {
        if (!a.name.startsWith("r:")) continue;
        const next = remap.get(a.value);
        if (next) el.setAttribute(a.name, next);
        else if (/^a:hlink/.test(el.nodeName)) el.parentNode?.removeChild(el);
        else {
          const shape = closestShape(el);
          if (shape?.parentNode) shape.parentNode.removeChild(shape);
        }
      }
    }
    this.pkg.setXml(part, doc);
    await this.register(part, at);
    return dropped.length ? `Not copied (not supported): ${[...new Set(dropped)].join(", ")}.` : "";
  }

  /** The deck's own layouts: their names and the placeholders a slide made from them gets. */
  async layouts() {
    const out: { name: string; part: string; placeholders: string[] }[] = [];
    for (const part of this.pkg.parts("ppt/slideLayouts/").filter((p) => /slideLayout\d+\.xml$/.test(p))) {
      const root = (await this.pkg.xml(part)).documentElement as XmlEl;
      const tree = path(root, "p:cSld/p:spTree");
      const placeholders = (tree ? kids(tree) : []).map(phOf).filter((ph) => ph && !["dt", "ftr", "sldNum", "hdr"].includes(ph.type)).map((ph) => ph!.type);
      out.push({ name: kid(root, "p:cSld")?.getAttribute("name") || part, part, placeholders });
    }
    return out;
  }

  /**
   * Adds an empty slide made from one of the deck's layouts (by name; none: the layout with a title and
   * a body, else the first): its placeholders (title, body…) are there to fill, in the layout's look.
   */
  async addSlideFromLayout(name?: string, at?: number) {
    const layouts = await this.layouts();
    const key = (v: string) => v.toLowerCase().replace(/\s+/g, " ").trim();
    const layout = name ? layouts.find((l) => key(l.name) === key(name)) : (layouts.find((l) => l.placeholders.includes("title") && l.placeholders.some((p) => p === "body" || p === "obj")) ?? layouts[0]);
    if (!layout) throw new Error(`No layout "${name}" in this deck: ${layouts.map((l) => `"${l.name}"`).join(", ")}.`);
    const source = path((await this.pkg.xml(layout.part)).documentElement as XmlEl, "p:cSld/p:spTree");
    const part = this.pkg.freeName("ppt/slides/slide", "xml");
    this.pkg.setXml(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
    const doc = await this.pkg.xml(part);
    let id = 1;
    for (const el of source ? kids(source, "p:sp") : []) {
      const ph = phOf(el);
      if (!ph || ["dt", "ftr", "sldNum", "hdr"].includes(ph.type)) continue;
      // The placeholder alone: its place and look come from the layout.
      const nv = doc.importNode(nonVisual(el)!, true) as XmlEl;
      kid(nv, "p:cNvPr")!.setAttribute("id", String(++id));
      this.tree(doc).appendChild(element(doc, "p:sp", {}, [nv, element(doc, "p:spPr"), element(doc, "p:txBody", {}, [element(doc, "a:bodyPr"), element(doc, "a:lstStyle"), element(doc, "a:p")])]));
    }
    await this.pkg.addRel(part, REL.slideLayout, layout.part);
    await this.register(part, at);
    return layout.name;
  }

  /** This deck's layout for a layout of another deck: same name, else same type, else the first. */
  private async layoutFor(from: Deck, layoutPart: string) {
    const src = (await from.pkg.xml(layoutPart)).documentElement as XmlEl;
    const name = kid(src, "p:cSld")?.getAttribute("name");
    const type = src.getAttribute("type");
    const layouts = this.pkg.parts("ppt/slideLayouts/").filter((p) => /slideLayout\d+\.xml$/.test(p));
    const info = await Promise.all(layouts.map(async (p) => ({ p, root: (await this.pkg.xml(p)).documentElement as XmlEl })));
    return (info.find((l) => kid(l.root, "p:cSld")?.getAttribute("name") === name) ?? info.find((l) => type && l.root.getAttribute("type") === type) ?? info.find((l) => l.root.getAttribute("type") === "blank") ?? info[0]).p;
  }

  /* ------------------------------------ XML ------------------------------------ */

  /** A part's XML (a slide by number, or any part by name: layouts, masters, theme…). */
  async xml(target: number | string) {
    return this.pkg.text(typeof target === "number" ? await this.partOf(target) : target);
  }
  async setXml(target: number | string, xml: string) {
    const part = typeof target === "number" ? await this.partOf(target) : target;
    if (!this.pkg.has(part)) throw new Error(`No part ${part}.`);
    this.pkg.setXml(part, xml);
  }
  /** The parts of the deck (for the AI's raw edits). */
  parts() {
    return this.pkg.parts("ppt/").filter((p) => p.endsWith(".xml") && !p.includes("_rels"));
  }

  async bytes() {
    return this.pkg.bytes();
  }
  async blob() {
    return this.pkg.save(PPTX_TYPE);
  }
  /** A picture of the deck as a data URL (the preview). */
  async imageUrl(part: string) {
    const bytes = await this.pkg.binary(part);
    const type = (await this.pkg.contentType(part)) ?? "image/png";
    return URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  }
  relative(from: string, to: string) {
    return relativePart(from, to);
  }
}

/** The slide-level shape holding an element (a picture inside a group: the picture). */
function closestShape(el: XmlEl): XmlEl | undefined {
  for (let n: Node | null = el; n; n = n.parentNode) if (n.nodeType === 1 && /^p:(sp|pic|graphicFrame|grpSp|cxnSp)$/.test(n.nodeName)) return n as XmlEl;
  return undefined;
}
