/**
 * An existing PowerPoint, read and changed in place (its XML): everything not touched stays as it was —
 * masters, layouts, pictures, animations. The AI reads the structure (slides, shapes with their id, kind,
 * place and text, pictures, tables) and changes shapes by id: text, pictures (a screenshot replaced by a
 * fresh one, cropped to the same frame), table cells, slides added / moved / removed.
 */
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import JSZip from "jszip";
import { imageSize } from "./images";

const NS = {
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
};
const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const EMU = 914400;

type El = Element;
const kids = (e: El | Document, ns: string, name: string): El[] => {
  const out: El[] = [];
  for (let n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (n as El).localName === name && (n as El).namespaceURI === ns) out.push(n as El);
  return out;
};
const kid = (e: El | Document | undefined, ns: string, name: string) => (e ? kids(e, ns, name)[0] : undefined);
const all = (e: El | Document, ns: string, name: string) => Array.from(e.getElementsByTagNameNS(ns, name)) as El[];
const path = (e: El | undefined, ...steps: [string, string][]) => steps.reduce<El | undefined>((cur, [ns, n]) => kid(cur, ns, n), e);

export type ShapeInfo = {
  id: string;
  name: string;
  kind: "text" | "picture" | "table" | "chart" | "group" | "other";
  placeholder?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  text?: string;
  rows?: string[][];
  image?: { width: number; height: number; file: string };
};
export type SlideInfo = { number: number; layout: string; shapes: ShapeInfo[]; notes?: string };

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml") as unknown as Document;
const serialize = (doc: Document) => new XMLSerializer().serializeToString(doc as never);
const round = (v: number) => Math.round(v * 100) / 100;

function textOf(body: El | undefined) {
  if (!body) return "";
  return kids(body, NS.a, "p")
    .map((p) => all(p, NS.a, "t").map((t) => t.textContent ?? "").join(""))
    .join("\n");
}

export class Pptx {
  private docs = new Map<string, Document>();
  private constructor(private zip: JSZip) {}

  static async load(data: Uint8Array) {
    return new Pptx(await JSZip.loadAsync(data));
  }

  private async doc(file: string) {
    if (!this.docs.has(file)) {
      const text = await this.zip.file(file)?.async("string");
      if (text === undefined) throw new Error(`Missing part ${file}`);
      this.docs.set(file, parse(text));
    }
    return this.docs.get(file)!;
  }
  private relsPath = (file: string) => file.replace(/([^/]+)$/, "_rels/$1.rels");
  private async rels(file: string) {
    const p = this.relsPath(file);
    if (!this.zip.file(p)) this.zip.file(p, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"/>`);
    return this.doc(p);
  }
  private resolve(from: string, target: string) {
    const parts = from.split("/").slice(0, -1);
    for (const seg of target.split("/")) seg === ".." ? parts.pop() : seg !== "." && parts.push(seg);
    return parts.join("/");
  }

  /** Slide files in show order. */
  async slideFiles() {
    const pres = await this.doc("ppt/presentation.xml");
    const rels = await this.rels("ppt/presentation.xml");
    const byId = new Map(all(rels, NS.rel, "Relationship").map((r) => [r.getAttribute("Id")!, r.getAttribute("Target")!]));
    return all(pres, NS.p, "sldId").map((s) => this.resolve("ppt/presentation.xml", byId.get(s.getAttributeNS(NS.r, "id")!)!));
  }

  private async slideFile(number: number) {
    const files = await this.slideFiles();
    const f = files[number - 1];
    if (!f) throw new Error(`No slide ${number} (the deck has ${files.length}).`);
    return f;
  }

  private async relTarget(file: string, id: string) {
    const r = all(await this.rels(file), NS.rel, "Relationship").find((x) => x.getAttribute("Id") === id);
    return r ? this.resolve(file, r.getAttribute("Target")!) : undefined;
  }

  private async shapeInfo(file: string, e: El): Promise<ShapeInfo | undefined> {
    const kind = e.localName;
    const nv = kid(e, NS.p, kind === "sp" ? "nvSpPr" : kind === "pic" ? "nvPicPr" : kind === "graphicFrame" ? "nvGraphicFramePr" : kind === "grpSp" ? "nvGrpSpPr" : "nvCxnSpPr");
    const c = kid(nv, NS.p, "cNvPr");
    if (!c) return undefined;
    const xfrm = kind === "graphicFrame" ? kid(e, NS.p, "xfrm") : path(e, [NS.p, kind === "grpSp" ? "grpSpPr" : "spPr"], [NS.a, "xfrm"]);
    const off = kid(xfrm, NS.a, "off");
    const ext = kid(xfrm, NS.a, "ext");
    const ph = path(nv, [NS.p, "nvPr"], [NS.p, "ph"]);
    const info: ShapeInfo = {
      id: c.getAttribute("id")!,
      name: c.getAttribute("name") ?? "",
      kind: "other",
      placeholder: ph ? (ph.getAttribute("type") ?? "body") : undefined,
      x: round(Number(off?.getAttribute("x") ?? 0) / EMU),
      y: round(Number(off?.getAttribute("y") ?? 0) / EMU),
      w: round(Number(ext?.getAttribute("cx") ?? 0) / EMU),
      h: round(Number(ext?.getAttribute("cy") ?? 0) / EMU),
    };
    if (kind === "sp") {
      const text = textOf(kid(e, NS.p, "txBody"));
      info.kind = kid(e, NS.p, "txBody") ? "text" : "other";
      if (text) info.text = text;
    } else if (kind === "pic") {
      info.kind = "picture";
      const blip = all(e, NS.a, "blip")[0];
      const target = blip && (await this.relTarget(file, blip.getAttributeNS(NS.r, "embed")!));
      const bytes = target ? await this.zip.file(target)?.async("uint8array") : undefined;
      const size = bytes && imageSize(bytes);
      if (target) info.image = { file: target.split("/").pop()!, width: size?.width ?? 0, height: size?.height ?? 0 };
    } else if (kind === "graphicFrame") {
      const tbl = all(e, NS.a, "tbl")[0];
      if (tbl) {
        info.kind = "table";
        info.rows = kids(tbl, NS.a, "tr").map((tr) => kids(tr, NS.a, "tc").map((tc) => textOf(kid(tc, NS.a, "txBody"))));
      } else if (all(e, "http://schemas.openxmlformats.org/drawingml/2006/chart", "chart").length) info.kind = "chart";
    } else if (kind === "grpSp") info.kind = "group";
    return info;
  }

  /** Every slide: layout name, shapes (groups opened), notes. */
  async structure(): Promise<SlideInfo[]> {
    const files = await this.slideFiles();
    const out: SlideInfo[] = [];
    for (const [i, file] of files.entries()) {
      const doc = await this.doc(file);
      const tree = all(doc, NS.p, "spTree")[0];
      const shapes: ShapeInfo[] = [];
      const walk = async (parent: El) => {
        for (let n = parent.firstChild; n; n = n.nextSibling) {
          if (n.nodeType !== 1) continue;
          const e = n as El;
          if (!["sp", "pic", "graphicFrame", "grpSp", "cxnSp"].includes(e.localName!)) continue;
          if (e.localName === "cxnSp") continue;
          const info = await this.shapeInfo(file, e);
          if (info) shapes.push(info);
          if (e.localName === "grpSp") await walk(e);
        }
      };
      if (tree) await walk(tree);
      // Layout name and notes.
      const rels = all(await this.rels(file), NS.rel, "Relationship");
      const layoutRel = rels.find((r) => /slideLayout$/.test(r.getAttribute("Type")!));
      let layout = "";
      if (layoutRel) {
        const ld = await this.doc(this.resolve(file, layoutRel.getAttribute("Target")!)).catch(() => undefined);
        layout = ld ? (all(ld, NS.p, "cSld")[0]?.getAttribute("name") ?? "") : "";
      }
      const notesRel = rels.find((r) => /notesSlide$/.test(r.getAttribute("Type")!));
      let notes: string | undefined;
      if (notesRel) {
        const nd = await this.doc(this.resolve(file, notesRel.getAttribute("Target")!)).catch(() => undefined);
        notes = nd ? all(nd, NS.p, "sp").filter((sp) => path(sp, [NS.p, "nvSpPr"], [NS.p, "nvPr"], [NS.p, "ph"])?.getAttribute("type") === "body").map((sp) => textOf(kid(sp, NS.p, "txBody"))).join("\n") || undefined : undefined;
      }
      out.push({ number: i + 1, layout, shapes, notes });
    }
    return out;
  }

  private async shape(number: number, id: string) {
    const file = await this.slideFile(number);
    const doc = await this.doc(file);
    const c = all(doc, NS.p, "cNvPr").find((x) => x.getAttribute("id") === String(id));
    if (!c) throw new Error(`No shape id ${id} on slide ${number}: read the structure for the ids.`);
    return { file, doc, shape: c.parentNode!.parentNode as El };
  }

  /** The text of a shape (a text box, a placeholder or a table cell `row,col` 0-based), its first formatting kept. Lines → paragraphs. */
  async setText(number: number, id: string, text: string, cell?: { row: number; col: number }) {
    const { doc, shape } = await this.shape(number, id);
    let body: El | undefined;
    if (cell) {
      const tbl = all(shape, NS.a, "tbl")[0];
      if (!tbl) throw new Error(`Shape ${id} is not a table.`);
      const tc = kids(kids(tbl, NS.a, "tr")[cell.row] ?? ({} as El), NS.a, "tc")[cell.col];
      if (!tc) throw new Error(`No cell ${cell.row},${cell.col} in the table.`);
      body = kid(tc, NS.a, "txBody");
    } else body = kid(shape, NS.p, "txBody");
    if (!body) throw new Error(`Shape ${id} holds no text (a picture or a chart?).`);
    const paragraphs = kids(body, NS.a, "p");
    const first = paragraphs[0];
    const pPr = first && kid(first, NS.a, "pPr");
    const rPr = (first && (all(first, NS.a, "rPr")[0] ?? kid(first, NS.a, "endParaRPr"))) as El | undefined;
    for (const p of paragraphs) body.removeChild(p);
    for (const line of text.split("\n")) {
      const p = doc.createElementNS(NS.a, "a:p");
      if (pPr) p.appendChild(pPr.cloneNode(true));
      if (line) {
        const r = doc.createElementNS(NS.a, "a:r");
        if (rPr) {
          const props = rPr.cloneNode(true) as El;
          if (props.localName === "endParaRPr") {
            const renamed = doc.createElementNS(NS.a, "a:rPr");
            for (const at of Array.from(props.attributes)) renamed.setAttribute(at.name, at.value);
            for (let n = props.firstChild; n; n = n.nextSibling) renamed.appendChild(n.cloneNode(true));
            r.appendChild(renamed);
          } else r.appendChild(props);
        }
        const t = doc.createElementNS(NS.a, "a:t");
        t.appendChild(doc.createTextNode(line));
        r.appendChild(t);
        p.appendChild(r);
      }
      body.appendChild(p);
    }
  }

  private async addMedia(file: string, data: Uint8Array) {
    const ext = data[0] === 0x89 ? "png" : data[0] === 0xff ? "jpeg" : "gif";
    const name = `ppt/media/hub-${crypto.randomUUID().slice(0, 8)}.${ext}`;
    this.zip.file(name, data);
    const ct = await this.doc("[Content_Types].xml");
    if (!all(ct, NS.ct, "Default").some((d) => d.getAttribute("Extension") === ext)) {
      const d = ct.createElementNS(NS.ct, "Default");
      d.setAttribute("Extension", ext);
      d.setAttribute("ContentType", `image/${ext}`);
      ct.documentElement!.insertBefore(d, ct.documentElement!.firstChild);
    }
    const rels = await this.rels(file);
    const used = new Set(all(rels, NS.rel, "Relationship").map((r) => r.getAttribute("Id")));
    let n = used.size + 1;
    while (used.has(`rId${n}`)) n++;
    const rel = rels.createElementNS(NS.rel, "Relationship");
    rel.setAttribute("Id", `rId${n}`);
    rel.setAttribute("Type", REL_IMAGE);
    rel.setAttribute("Target", `../media/${name.split("/").pop()}`);
    rels.documentElement!.appendChild(rel);
    return `rId${n}`;
  }

  /**
   * A picture's content replaced, its frame kept: "cover" crops the new picture to the frame (default),
   * "contain" resizes the frame to the picture's proportions, centered in the old one.
   */
  async replacePicture(number: number, id: string, data: Uint8Array, fit: "cover" | "contain" = "cover") {
    const { file, shape } = await this.shape(number, id);
    const blip = all(shape, NS.a, "blip")[0];
    if (!blip) throw new Error(`Shape ${id} is not a picture.`);
    const size = imageSize(data);
    if (!size) throw new Error("The new picture must be a PNG, JPEG or GIF.");
    blip.setAttributeNS(NS.r, "r:embed", await this.addMedia(file, data));
    const fill = blip.parentNode as El;
    const src = kid(fill, NS.a, "srcRect");
    if (src) fill.removeChild(src);
    const ext = path(shape, [NS.p, "spPr"], [NS.a, "xfrm"], [NS.a, "ext"]);
    const off = path(shape, [NS.p, "spPr"], [NS.a, "xfrm"], [NS.a, "off"]);
    if (!ext || !off) return;
    const cx = Number(ext.getAttribute("cx")), cy = Number(ext.getAttribute("cy"));
    const frame = cx / cy, pic = size.width / size.height;
    if (fit === "contain") {
      let w = cx, h = cx / pic;
      if (h > cy) (h = cy), (w = cy * pic);
      off.setAttribute("x", String(Math.round(Number(off.getAttribute("x")) + (cx - w) / 2)));
      off.setAttribute("y", String(Math.round(Number(off.getAttribute("y")) + (cy - h) / 2)));
      ext.setAttribute("cx", String(Math.round(w)));
      ext.setAttribute("cy", String(Math.round(h)));
    } else if (Math.abs(frame - pic) > 0.01) {
      const rect = fill.ownerDocument!.createElementNS(NS.a, "a:srcRect");
      // Crop (in 1/1000 %) the sides that overflow the frame.
      if (pic > frame) {
        const cut = Math.round(((1 - frame / pic) / 2) * 100000);
        rect.setAttribute("l", String(cut));
        rect.setAttribute("r", String(cut));
      } else {
        const cut = Math.round(((1 - pic / frame) / 2) * 100000);
        rect.setAttribute("t", String(cut));
        rect.setAttribute("b", String(cut));
      }
      fill.insertBefore(rect, kid(fill, NS.a, "stretch") ?? null);
    }
  }

  private async nextShapeId(doc: Document) {
    return Math.max(0, ...all(doc, NS.p, "cNvPr").map((c) => Number(c.getAttribute("id")) || 0)) + 1;
  }

  /** A new picture on a slide (inches). */
  async addPicture(number: number, data: Uint8Array, box: { x: number; y: number; w: number; h?: number }) {
    const file = await this.slideFile(number);
    const doc = await this.doc(file);
    const size = imageSize(data);
    if (!size) throw new Error("The picture must be a PNG, JPEG or GIF.");
    const h = box.h ?? (box.w * size.height) / size.width;
    const rid = await this.addMedia(file, data);
    const id = await this.nextShapeId(doc);
    const xml = `<p:pic xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round(box.x * EMU)}" y="${Math.round(box.y * EMU)}"/><a:ext cx="${Math.round(box.w * EMU)}" cy="${Math.round(h * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    all(doc, NS.p, "spTree")[0].appendChild(doc.importNode(parse(xml).documentElement!, true));
    return String(id);
  }

  /** A new text box on a slide (inches, font size in points). */
  async addText(number: number, text: string, box: { x: number; y: number; w: number; h: number }, o: { size?: number; bold?: boolean; color?: string } = {}) {
    const file = await this.slideFile(number);
    const doc = await this.doc(file);
    const id = await this.nextShapeId(doc);
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const rPr = `<a:rPr lang="en-US" sz="${Math.round((o.size ?? 14) * 100)}"${o.bold ? ' b="1"' : ""} dirty="0">${o.color ? `<a:solidFill><a:srgbClr val="${o.color.replace("#", "")}"/></a:solidFill>` : ""}</a:rPr>`;
    const paras = text.split("\n").map((l) => `<a:p>${l ? `<a:r>${rPr}<a:t>${esc(l)}</a:t></a:r>` : ""}</a:p>`).join("");
    const xml = `<p:sp xmlns:p="${NS.p}" xmlns:a="${NS.a}"><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(box.x * EMU)}" y="${Math.round(box.y * EMU)}"/><a:ext cx="${Math.round(box.w * EMU)}" cy="${Math.round(box.h * EMU)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody></p:sp>`;
    all(doc, NS.p, "spTree")[0].appendChild(doc.importNode(parse(xml).documentElement!, true));
    return String(id);
  }

  /** Moves or resizes a shape (inches; what is not given stays). */
  async moveShape(number: number, id: string, box: { x?: number; y?: number; w?: number; h?: number }) {
    const { shape } = await this.shape(number, id);
    const xfrm = shape.localName === "graphicFrame" ? kid(shape, NS.p, "xfrm") : path(shape, [NS.p, shape.localName === "grpSp" ? "grpSpPr" : "spPr"], [NS.a, "xfrm"]);
    const off = kid(xfrm, NS.a, "off"), ext = kid(xfrm, NS.a, "ext");
    if (!off || !ext) throw new Error(`Shape ${id} has no position of its own (it follows its layout placeholder).`);
    if (box.x !== undefined) off.setAttribute("x", String(Math.round(box.x * EMU)));
    if (box.y !== undefined) off.setAttribute("y", String(Math.round(box.y * EMU)));
    if (box.w !== undefined) ext.setAttribute("cx", String(Math.round(box.w * EMU)));
    if (box.h !== undefined) ext.setAttribute("cy", String(Math.round(box.h * EMU)));
  }

  async deleteShape(number: number, id: string) {
    const { shape } = await this.shape(number, id);
    shape.parentNode!.removeChild(shape);
  }

  /** A copy of a slide, placed right after it (its notes are not copied). Returns the new slide's number. */
  async duplicateSlide(number: number) {
    const file = await this.slideFile(number);
    const files = Object.keys(this.zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    const next = Math.max(...files.map((f) => Number(f.match(/(\d+)\.xml$/)![1]))) + 1;
    const target = `ppt/slides/slide${next}.xml`;
    this.zip.file(target, serialize(await this.doc(file)));
    const rels = parse(serialize(await this.rels(file)));
    for (const r of all(rels, NS.rel, "Relationship")) if (/notesSlide$/.test(r.getAttribute("Type")!)) r.parentNode!.removeChild(r);
    this.zip.file(this.relsPath(target), serialize(rels));
    const ct = await this.doc("[Content_Types].xml");
    const o = ct.createElementNS(NS.ct, "Override");
    o.setAttribute("PartName", `/${target}`);
    o.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
    ct.documentElement!.appendChild(o);
    const presRels = await this.rels("ppt/presentation.xml");
    const ids = new Set(all(presRels, NS.rel, "Relationship").map((r) => r.getAttribute("Id")));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    const rel = presRels.createElementNS(NS.rel, "Relationship");
    rel.setAttribute("Id", `rId${n}`);
    rel.setAttribute("Type", REL_SLIDE);
    rel.setAttribute("Target", `slides/slide${next}.xml`);
    presRels.documentElement!.appendChild(rel);
    const pres = await this.doc("ppt/presentation.xml");
    const list = all(pres, NS.p, "sldIdLst")[0];
    const items = kids(list, NS.p, "sldId");
    const s = pres.createElementNS(NS.p, "p:sldId");
    s.setAttribute("id", String(Math.max(...items.map((i) => Number(i.getAttribute("id")))) + 1));
    s.setAttributeNS(NS.r, "r:id", `rId${n}`);
    list.insertBefore(s, items[number] ?? null);
    return number + 1;
  }

  async deleteSlide(number: number) {
    const pres = await this.doc("ppt/presentation.xml");
    const list = all(pres, NS.p, "sldIdLst")[0];
    const item = kids(list, NS.p, "sldId")[number - 1];
    if (!item) throw new Error(`No slide ${number}.`);
    if (kids(list, NS.p, "sldId").length === 1) throw new Error("A deck keeps at least one slide.");
    list.removeChild(item);
  }

  async moveSlide(from: number, to: number) {
    const pres = await this.doc("ppt/presentation.xml");
    const list = all(pres, NS.p, "sldIdLst")[0];
    const items = kids(list, NS.p, "sldId");
    const item = items[from - 1];
    if (!item || to < 1 || to > items.length) throw new Error(`Slides are numbered 1 to ${items.length}.`);
    list.removeChild(item);
    const rest = kids(list, NS.p, "sldId");
    list.insertBefore(item, rest[to - 1] ?? null);
  }

  /** The bytes of a picture of the deck (by its file name in the structure). */
  async media(name: string) {
    const f = this.zip.file(`ppt/media/${name.replace(/[^\w.-]/g, "")}`);
    if (!f) throw new Error(`No picture "${name}" in the deck.`);
    return f.async("uint8array");
  }

  async save(): Promise<Uint8Array> {
    for (const [file, doc] of this.docs) this.zip.file(file, serialize(doc));
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
}
