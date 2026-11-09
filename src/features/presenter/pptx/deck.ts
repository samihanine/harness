/**
 * A real PowerPoint deck (any .pptx, no template required) read and changed in the browser through
 * its XML (lib/ooxml): the AI never sees the slides, it reads them as text — each shape with its id,
 * kind, place (in % of the slide), text and how much text fits — and changes them with small
 * operations written by code (texts, slides added from a template deck, pictures, order). Its raw XML
 * stays reachable for anything else. Everything not touched is kept as it was.
 */
import { OoxmlPackage, REL, all, element, kid, kids, path, relativePart, resolvePart, type XmlDoc, type XmlEl } from "@/lib/ooxml/package";

export const PPTX_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const EMU_PER_PT = 12_700;

export type Box = { x: number; y: number; w: number; h: number };
export type ShapeKind = "text" | "picture" | "table" | "chart" | "group" | "line" | "other";
export type Paragraph = { text: string; level: number; size?: number; bold?: boolean; color?: string; align?: string };
export type Shape = {
  id: string;
  name: string;
  kind: ShapeKind;
  /** Placeholder type when it is one (title, body, subTitle, pic…). */
  placeholder?: string;
  box?: Box;
  paragraphs: Paragraph[];
  /** Font size in points (the first run's, else the placeholder's usual). */
  size: number;
  /** Characters that fit in its box at that size (an estimate). */
  budget?: number;
  fill?: string;
  /** A picture's file in the package. */
  image?: string;
};
export type SlideInfo = { number: number; part: string; layout: string; title: string; shapes: Shape[]; background?: string; notes: string };

const num = (v: string | null | undefined) => (v ? Number(v) : 0);
const nonVisual = (el: XmlEl) => kids(el).find((c) => c.nodeName.startsWith("p:nv"));
const cNvPr = (el: XmlEl) => (nonVisual(el) ? kid(nonVisual(el)!, "p:cNvPr") : undefined);
const phOf = (el: XmlEl) => {
  const nv = nonVisual(el);
  const ph = nv && path(nv, "p:nvPr/p:ph");
  return ph ? { type: ph.getAttribute("type") ?? "body", idx: ph.getAttribute("idx") ?? undefined } : undefined;
};
const xfrmOf = (el: XmlEl) => path(el, "p:spPr/a:xfrm") ?? path(el, "p:xfrm") ?? path(el, "p:grpSpPr/a:xfrm");
const boxOf = (xfrm: XmlEl | undefined): Box | undefined => {
  const off = xfrm && kid(xfrm, "a:off");
  const ext = xfrm && kid(xfrm, "a:ext");
  return off && ext ? { x: num(off.getAttribute("x")), y: num(off.getAttribute("y")), w: num(ext.getAttribute("cx")), h: num(ext.getAttribute("cy")) } : undefined;
};
const colorOf = (el: XmlEl | undefined) => {
  const c = el && path(el, "a:solidFill/a:srgbClr");
  return c ? `#${c.getAttribute("val")}` : undefined;
};
/** The usual size of a placeholder's text (points) when the slide does not say. */
const DEFAULT_SIZE: Record<string, number> = { title: 36, ctrTitle: 44, subTitle: 20, body: 18 };

function paragraphsOf(txBody: XmlEl | undefined): Paragraph[] {
  if (!txBody) return [];
  return kids(txBody, "a:p").map((p) => {
    const runs = kids(p).filter((c) => c.nodeName === "a:r" || c.nodeName === "a:fld" || c.nodeName === "a:br");
    const text = runs.map((r) => (r.nodeName === "a:br" ? "\n" : (kid(r, "a:t")?.textContent ?? ""))).join("");
    const rPr = runs.map((r) => kid(r, "a:rPr")).find(Boolean) ?? kid(p, "a:endParaRPr");
    const pPr = kid(p, "a:pPr");
    return { text, level: num(pPr?.getAttribute("lvl")), size: rPr?.getAttribute("sz") ? num(rPr.getAttribute("sz")) / 100 : undefined, bold: rPr?.getAttribute("b") === "1" || undefined, color: colorOf(rPr), align: pPr?.getAttribute("algn") ?? undefined };
  });
}

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
    return layout ? (kid((await this.pkg.xml(layout)).documentElement as XmlEl, "p:cSld")?.getAttribute("name") ?? "") : "";
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
      const sz = el && (all(el, "a:defRPr")[0]?.getAttribute("sz") ?? all(el, "a:rPr")[0]?.getAttribute("sz"));
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
    const paragraphs: Paragraph[] = kind === "table" ? all(el, "a:tc").map((tc) => ({ text: paragraphsOf(kid(tc, "a:txBody")).map((p) => p.text).join(" "), level: 0 })) : paragraphsOf(kid(el, "p:txBody"));
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

  /** The whole deck as text for the AI: each slide, its layout and its shapes (ids, kinds, places, texts, room). */
  async outline(full = false) {
    const lines: string[] = [];
    for (const s of await this.slides()) {
      lines.push(`Slide ${s.number} — layout "${s.layout}"${s.title ? ` — "${s.title}"` : ""}`);
      for (const sh of s.shapes) {
        const text = sh.paragraphs.map((p) => `${"  ".repeat(p.level)}${p.text}`).join(" ⏎ ");
        lines.push(`  · #${sh.id} ${sh.kind}${sh.placeholder ? `/${sh.placeholder}` : ""} "${sh.name}" (${this.pct(sh.box)}, ${sh.size}pt${sh.budget ? `, room ≈ ${sh.budget} chars` : ""})${text ? `: ${full ? text : text.slice(0, 160)}` : ""}`);
      }
      if (full && s.notes) lines.push(`  notes: ${s.notes}`);
    }
    return lines.join("\n");
  }

  /* --------------------------------- Changing --------------------------------- */

  private async shapeEl(number: number, id: string) {
    const part = await this.partOf(number);
    const doc = await this.pkg.xml(part);
    const want = id.replace(/^#/, "");
    const el = [...all(doc, "p:sp"), ...all(doc, "p:pic"), ...all(doc, "p:graphicFrame"), ...all(doc, "p:grpSp"), ...all(doc, "p:cxnSp")].find((e) => cNvPr(e)?.getAttribute("id") === want || cNvPr(e)?.getAttribute("name") === id);
    if (!el) throw new Error(`No shape #${want} on slide ${number}: read the slide first.`);
    return { part, doc, el };
  }

  /**
   * Replaces a shape's text by paragraphs (lines; "- " or tabs before a line: one level deeper). Each
   * new paragraph takes the look of the shape's paragraph at the same level (bullets, size, color, font).
   */
  async setText(number: number, id: string, text: string) {
    const { doc, el } = await this.shapeEl(number, id);
    let body = kid(el, "p:txBody");
    if (!body) {
      if (el.nodeName !== "p:sp") throw new Error(`Shape #${id} holds no text (a ${el.nodeName.slice(2)}).`);
      body = element(doc, "p:txBody", {}, [element(doc, "a:bodyPr"), element(doc, "a:lstStyle")]);
      el.appendChild(body);
    }
    const old = kids(body, "a:p");
    const model = (level: number) => old.find((p) => num(kid(p, "a:pPr")?.getAttribute("lvl")) === level && kids(p, "a:r").length) ?? old.find((p) => kids(p, "a:r").length) ?? old[0];
    const lines = text.replace(/\r/g, "").split("\n");
    for (const p of old) body.removeChild(p);
    for (const line of lines) {
      const m = line.match(/^(\t+|(?: {2})+)?(?:[-•]\s+)?(.*)$/)!;
      const indent = m[1] ?? "";
      const level = Math.min(8, indent.includes("\t") ? indent.length : indent.length / 2);
      const from = model(level);
      const p = element(doc, "a:p");
      const pPr = from && kid(from, "a:pPr");
      if (pPr) p.appendChild(pPr.cloneNode(true));
      else if (level) p.appendChild(element(doc, "a:pPr", { lvl: level }));
      if (level && p.firstChild) (p.firstChild as XmlEl).setAttribute("lvl", String(level));
      const rPr = from && (kid(kids(from, "a:r")[0] ?? from, "a:rPr") ?? kid(from, "a:endParaRPr"));
      if (m[2]) {
        const r = element(doc, "a:r");
        if (rPr) {
          const copy = rPr.cloneNode(true) as XmlEl;
          if (copy.nodeName === "a:endParaRPr") {
            const renamed = element(doc, "a:rPr");
            for (const a of Array.from(copy.attributes)) renamed.setAttribute(a.name, a.value);
            for (const c of Array.from(copy.childNodes)) renamed.appendChild(c);
            r.appendChild(renamed);
          } else r.appendChild(copy);
        }
        r.appendChild(element(doc, "a:t", {}, [m[2]]));
        p.appendChild(r);
      }
      const end = from && kid(from, "a:endParaRPr");
      if (end) p.appendChild(end.cloneNode(true));
      body.appendChild(p);
    }
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

  /**
   * Puts a picture (PNG) on a slide: in place of a shape (a picture keeps its frame, a placeholder or
   * text box is replaced by the picture at its place, fitted in its proportions) or at a box (% of the slide).
   */
  async putPicture(number: number, png: Uint8Array, size: { width: number; height: number }, target: { shape?: string; box?: Box }, name = "Power BI visual") {
    const part = await this.partOf(number);
    const media = this.pkg.freeName("ppt/media/hub-image", "png");
    this.pkg.setBinary(media, png);
    await this.pkg.ensureDefault("png", "image/png");
    const rid = await this.pkg.addRel(part, REL.image, media);
    const doc = await this.pkg.xml(part);
    let frame: Box;
    if (target.shape) {
      const { el } = await this.shapeEl(number, target.shape);
      const blip = el.nodeName === "p:pic" ? all(el, "a:blip")[0] : undefined;
      if (blip) {
        blip.setAttribute("r:embed", rid);
        return;
      }
      const sh = await this.shape(part, el);
      if (!sh.box) throw new Error(`Shape #${target.shape} has no place on the slide.`);
      frame = sh.box;
      el.parentNode!.removeChild(el);
    } else if (target.box) frame = { x: (target.box.x / 100) * this.size.w, y: (target.box.y / 100) * this.size.h, w: (target.box.w / 100) * this.size.w, h: (target.box.h / 100) * this.size.h };
    else frame = { x: this.size.w * 0.08, y: this.size.h * 0.22, w: this.size.w * 0.84, h: this.size.h * 0.7 };
    // Fitted in its frame, centered, in its own proportions.
    const k = Math.min(frame.w / size.width, frame.h / size.height);
    const w = Math.round(size.width * k);
    const h = Math.round(size.height * k);
    const x = Math.round(frame.x + (frame.w - w) / 2);
    const y = Math.round(frame.y + (frame.h - h) / 2);
    const id = this.nextId(doc);
    const pic = element(doc, "p:pic", {}, [
      element(doc, "p:nvPicPr", {}, [element(doc, "p:cNvPr", { id, name: `${name} ${id}` }), element(doc, "p:cNvPicPr", {}, [element(doc, "a:picLocks", { noChangeAspect: 1 })]), element(doc, "p:nvPr")]),
      element(doc, "p:blipFill", {}, [element(doc, "a:blip", { "r:embed": rid }), element(doc, "a:stretch", {}, [element(doc, "a:fillRect")])]),
      element(doc, "p:spPr", {}, [element(doc, "a:xfrm", {}, [element(doc, "a:off", { x, y }), element(doc, "a:ext", { cx: w, cy: h })]), element(doc, "a:prstGeom", { prst: "rect" }, [element(doc, "a:avLst")])]),
    ]);
    path(doc.documentElement as XmlEl, "p:cSld/p:spTree")!.appendChild(pic);
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
