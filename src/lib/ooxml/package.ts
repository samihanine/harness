/**
 * Office Open XML files (.pptx, .xlsx) opened in the browser: a zip of XML parts linked by
 * relationships. One small engine for the PowerPoint decks (features/presenter/pptx) and the Excel
 * tables' validation lists and colors (features/excel): parts read and written as XML documents,
 * relationships and content types kept consistent. Everything the code does not touch is written back
 * byte for byte (themes, masters, animations, charts…).
 */
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export type XmlDoc = Document;
export type XmlEl = Element;
export type Rel = { id: string; type: string; target: string; external?: boolean };

export const NS = {
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  types: "http://schemas.openxmlformats.org/package/2006/content-types",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  x: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
};
export const REL = {
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  notesSlide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  table: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
};

/** XML text → document; `strict`: any error throws (XML written by the AI), else the parser's best (files as they come). */
export function parseXml(text: string, strict = false): XmlDoc {
  const fail = (m: string) => {
    throw new Error(`Invalid XML: ${m.split("\n")[0]}`);
  };
  const quiet = () => undefined;
  return new DOMParser({ errorHandler: { warning: quiet, error: strict ? fail : quiet, fatalError: fail } }).parseFromString(text, "application/xml") as unknown as XmlDoc;
}
export const serializeXml = (doc: Node) => new XMLSerializer().serializeToString(doc as never);

/** Child elements (by qualified name, e.g. "p:sp"; none: all). */
export const kids = (el: XmlEl | XmlDoc, name?: string): XmlEl[] => {
  const out: XmlEl[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && (!name || (n as XmlEl).nodeName === name)) out.push(n as XmlEl);
  return out;
};
export const kid = (el: XmlEl | XmlDoc, name: string): XmlEl | undefined => kids(el, name)[0];
/** Descendants by qualified name, in document order. */
export const all = (el: XmlEl | XmlDoc, name: string): XmlEl[] => Array.from(el.getElementsByTagName(name) as ArrayLike<XmlEl>);
/** A path of children ("p:spPr/a:xfrm/a:off"). */
export const path = (el: XmlEl | undefined, steps: string): XmlEl | undefined => steps.split("/").reduce<XmlEl | undefined>((e, s) => (e ? kid(e, s) : undefined), el);
/** A new element in the namespace of its prefix (a, p, r, x…). */
export function element(doc: XmlDoc, name: string, attrs: Record<string, string | number | undefined> = {}, children: (XmlEl | string)[] = []) {
  const prefix = name.includes(":") ? name.split(":")[0] : "";
  const el = doc.createElementNS((NS as Record<string, string>)[prefix] ?? doc.documentElement?.namespaceURI ?? null, name) as XmlEl;
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined) k.startsWith("r:") ? el.setAttributeNS(NS.r, k, String(v)) : el.setAttribute(k, String(v));
  for (const c of children) el.appendChild(typeof c === "string" ? doc.createTextNode(c) : c);
  return el;
}

/** The directory of a part and a target relative to it, as a part name ("ppt/slides/../media/a.png" → "ppt/media/a.png"). */
export function resolvePart(base: string, target: string) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const seg of target.split("/")) seg === ".." ? parts.pop() : seg !== "." && parts.push(seg);
  return parts.join("/");
}
/** A part's name relative to another part's directory (for a relationship's target). */
export function relativePart(base: string, part: string) {
  const from = base.split("/").slice(0, -1);
  const to = part.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...from.slice(i).map(() => ".."), ...to.slice(i)].join("/");
}
const relsPath = (part: string) => {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
};

export class OoxmlPackage {
  private docs = new Map<string, XmlDoc>();
  private constructor(private zip: JSZip) {}

  static async open(data: ArrayBuffer | Uint8Array | Blob) {
    return new OoxmlPackage(await JSZip.loadAsync(data));
  }

  has(part: string) {
    return !!this.zip.file(part) || this.docs.has(part);
  }
  /** Part names (in the zip). */
  parts(prefix = "") {
    return Object.keys(this.zip.files).filter((n) => !this.zip.files[n].dir && n.startsWith(prefix));
  }
  async xml(part: string): Promise<XmlDoc> {
    if (!this.docs.has(part)) {
      const file = this.zip.file(part);
      if (!file) throw new Error(`The file has no part ${part}.`);
      this.docs.set(part, parseXml(await file.async("string")));
    }
    return this.docs.get(part)!;
  }
  /** A part's XML as text (parts changed here are serialized). */
  async text(part: string) {
    return this.docs.has(part) ? serializeXml(this.docs.get(part)!) : ((await this.zip.file(part)?.async("string")) ?? "");
  }
  setXml(part: string, doc: XmlDoc | string) {
    const d = typeof doc === "string" ? parseXml(doc, true) : doc;
    const error = d.getElementsByTagName("parsererror")[0];
    if (error || !d.documentElement) throw new Error(`Invalid XML for ${part}: ${error?.textContent?.slice(0, 200) ?? "empty"}`);
    this.docs.set(part, d);
  }
  async binary(part: string) {
    const file = this.zip.file(part);
    if (!file) throw new Error(`The file has no part ${part}.`);
    return file.async("uint8array");
  }
  setBinary(part: string, data: Uint8Array | ArrayBuffer | Blob) {
    this.docs.delete(part);
    this.zip.file(part, data);
  }
  remove(part: string) {
    this.docs.delete(part);
    this.zip.remove(part);
    this.zip.remove(relsPath(part));
    this.docs.delete(relsPath(part));
  }
  /** A free part name like "ppt/slides/slide12.xml". */
  freeName(prefix: string, ext: string) {
    for (let i = 1; ; i++) if (!this.has(`${prefix}${i}.${ext}`)) return `${prefix}${i}.${ext}`;
  }

  /* ------------------------------- Relationships ------------------------------ */

  async rels(part: string): Promise<Rel[]> {
    const p = relsPath(part);
    if (!this.has(p)) return [];
    return kids((await this.xml(p)).documentElement as XmlEl, "Relationship").map((r) => ({ id: r.getAttribute("Id")!, type: r.getAttribute("Type")!, target: r.getAttribute("Target")!, external: r.getAttribute("TargetMode") === "External" }));
  }
  /** The part a relationship points to (by id). */
  async target(part: string, id: string) {
    const rel = (await this.rels(part)).find((r) => r.id === id);
    return rel && !rel.external ? resolvePart(part, rel.target) : undefined;
  }
  /** Adds a relationship from `part` to `target` (a part name); returns its id. */
  async addRel(part: string, type: string, target: string, external = false) {
    const p = relsPath(part);
    if (!this.has(p)) this.setXml(p, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rels}"/>`);
    const doc = await this.xml(p);
    const ids = new Set(kids(doc.documentElement as XmlEl, "Relationship").map((r) => r.getAttribute("Id")));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    const rel = doc.createElementNS(NS.rels, "Relationship");
    rel.setAttribute("Id", `rId${n}`);
    rel.setAttribute("Type", type);
    rel.setAttribute("Target", external ? target : relativePart(part, target));
    if (external) rel.setAttribute("TargetMode", "External");
    doc.documentElement!.appendChild(rel);
    return `rId${n}`;
  }
  async removeRel(part: string, id: string) {
    const p = relsPath(part);
    if (!this.has(p)) return;
    const doc = await this.xml(p);
    for (const r of kids(doc.documentElement as XmlEl, "Relationship")) if (r.getAttribute("Id") === id) r.parentNode!.removeChild(r);
  }

  /* ------------------------------- Content types ------------------------------ */

  async setContentType(part: string, type: string) {
    const doc = await this.xml("[Content_Types].xml");
    const root = doc.documentElement as XmlEl;
    for (const o of kids(root, "Override")) if (o.getAttribute("PartName") === `/${part}`) root.removeChild(o);
    const o = doc.createElementNS(NS.types, "Override");
    o.setAttribute("PartName", `/${part}`);
    o.setAttribute("ContentType", type);
    root.appendChild(o);
  }
  async removeContentType(part: string) {
    const doc = await this.xml("[Content_Types].xml");
    for (const o of kids(doc.documentElement as XmlEl, "Override")) if (o.getAttribute("PartName") === `/${part}`) o.parentNode!.removeChild(o);
  }
  /** A file extension known to the package (e.g. png images). */
  async ensureDefault(ext: string, type: string) {
    const doc = await this.xml("[Content_Types].xml");
    const root = doc.documentElement as XmlEl;
    if (kids(root, "Default").some((d) => d.getAttribute("Extension")?.toLowerCase() === ext)) return;
    const d = doc.createElementNS(NS.types, "Default");
    d.setAttribute("Extension", ext);
    d.setAttribute("ContentType", type);
    root.insertBefore(d, root.firstChild);
  }
  async contentType(part: string) {
    const doc = await this.xml("[Content_Types].xml");
    const root = doc.documentElement as XmlEl;
    const o = kids(root, "Override").find((x) => x.getAttribute("PartName") === `/${part}`);
    if (o) return o.getAttribute("ContentType") ?? undefined;
    const ext = part.split(".").pop()?.toLowerCase();
    return kids(root, "Default").find((x) => x.getAttribute("Extension")?.toLowerCase() === ext)?.getAttribute("ContentType") ?? undefined;
  }

  /** The file, every changed part serialized. */
  async save(type = "application/octet-stream") {
    for (const [part, doc] of this.docs) this.zip.file(part, serializeXml(doc));
    return this.zip.generateAsync({ type: "blob", mimeType: type, compression: "DEFLATE" });
  }
  async bytes() {
    for (const [part, doc] of this.docs) this.zip.file(part, serializeXml(doc));
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
}
