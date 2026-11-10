/**
 * The text of a shape, read and written as plain lines the AI can handle: one paragraph per line,
 * two spaces (or a tab) before a line for one level deeper, "- " for a bullet, **bold** inside a
 * line. Written back, each line takes the look of the paragraph it replaces (font, size, color,
 * spacing), so a text changed by the AI keeps the slide's style.
 */
import { element, kid, kids, type XmlDoc, type XmlEl } from "@/lib/ooxml/package";

export type Paragraph = { text: string; marked: string; level: number; bullet: boolean; size?: number; bold?: boolean; color?: string; align?: string; font?: string };
/** The look of a new text (a shape made by the app): points, #rrggbb, a font name. */
export type TextStyle = { size?: number; bold?: boolean; italic?: boolean; color?: string; font?: string; align?: "l" | "ctr" | "r" };

const num = (v: string | null | undefined) => (v ? Number(v) : 0);
export const colorOf = (el: XmlEl | undefined) => {
  const c = el && kid(kid(el, "a:solidFill") ?? el, "a:srgbClr");
  return c ? `#${c.getAttribute("val")}` : undefined;
};
const BULLETS = ["a:buChar", "a:buAutoNum", "a:buBlip"];
const hasBullet = (pPr: XmlEl | undefined) => !!pPr && BULLETS.some((b) => kid(pPr, b));
const isBold = (r: XmlEl) => kid(r, "a:rPr")?.getAttribute("b") === "1";

export function paragraphsOf(txBody: XmlEl | undefined): Paragraph[] {
  if (!txBody) return [];
  return kids(txBody, "a:p").map((p) => {
    const runs = kids(p).filter((c) => c.nodeName === "a:r" || c.nodeName === "a:fld" || c.nodeName === "a:br");
    const textOf = (r: XmlEl) => (r.nodeName === "a:br" ? " " : (kid(r, "a:t")?.textContent ?? ""));
    const text = runs.map(textOf).join("");
    // Bold marked only where it is a part of the line (a wholly bold line is its style, kept as it is).
    const written = runs.filter((r) => textOf(r).trim());
    const mixed = written.some(isBold) && !written.every(isBold);
    const marked = mixed ? runs.map((r) => (isBold(r) && textOf(r).trim() ? textOf(r).replace(/^(\s*)(.*?)(\s*)$/, "$1**$2**$3") : textOf(r))).join("").replace(/\*\*\*\*/g, "") : text;
    const rPr = runs.map((r) => kid(r, "a:rPr")).find(Boolean) ?? kid(p, "a:endParaRPr");
    const pPr = kid(p, "a:pPr");
    return {
      text,
      marked,
      level: num(pPr?.getAttribute("lvl")),
      bullet: hasBullet(pPr),
      size: rPr?.getAttribute("sz") ? num(rPr.getAttribute("sz")) / 100 : undefined,
      bold: (written.length > 0 && written.every(isBold)) || undefined,
      color: colorOf(rPr),
      align: pPr?.getAttribute("algn") || undefined,
      font: (rPr && kid(rPr, "a:latin")?.getAttribute("typeface")) || undefined,
    };
  });
}

/** Paragraphs as the lines the AI reads and writes back. */
export const linesOf = (paragraphs: Paragraph[]) => paragraphs.map((p) => `${"  ".repeat(p.level)}${p.bullet ? "- " : ""}${p.marked}`);

/** The children of a:pPr / a:rPr come in a fixed order: a new one goes before the first that must follow it. */
const P_ORDER = ["a:lnSpc", "a:spcBef", "a:spcAft", "a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct", "a:buSzPts", "a:buFontTx", "a:buFont", "a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst"];
const R_ORDER = ["a:ln", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill", "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"];
function insertOrdered(parent: XmlEl, child: XmlEl, order: string[]) {
  const rank = order.indexOf(child.nodeName);
  parent.insertBefore(child, kids(parent).find((c) => order.indexOf(c.nodeName) > rank) ?? null);
}

/** Applies a style to run properties (a:rPr, a:endParaRPr, a:defRPr). */
export function styleRun(doc: XmlDoc, rPr: XmlEl, style: TextStyle) {
  if (style.size) rPr.setAttribute("sz", String(Math.round(style.size * 100)));
  if (style.bold !== undefined) rPr.setAttribute("b", style.bold ? "1" : "0");
  if (style.italic !== undefined) rPr.setAttribute("i", style.italic ? "1" : "0");
  if (style.color) {
    for (const f of ["a:noFill", "a:solidFill", "a:gradFill"]) for (const c of kids(rPr, f)) rPr.removeChild(c);
    insertOrdered(rPr, element(doc, "a:solidFill", {}, [element(doc, "a:srgbClr", { val: style.color.replace("#", "").toUpperCase() })]), R_ORDER);
  }
  if (style.font) {
    for (const c of kids(rPr, "a:latin")) rPr.removeChild(c);
    insertOrdered(rPr, element(doc, "a:latin", { typeface: style.font }), R_ORDER);
  }
}

/** Run properties copied from a model (an a:endParaRPr becomes an a:rPr). */
function runProps(doc: XmlDoc, from: XmlEl | undefined) {
  const rPr = element(doc, "a:rPr", from ? {} : { lang: "en-US", dirty: 0 });
  if (from) {
    for (const a of Array.from(from.attributes)) rPr.setAttribute(a.name, a.value);
    for (const c of Array.from(from.childNodes)) rPr.appendChild(c.cloneNode(true));
    rPr.removeAttribute("err");
  }
  return rPr;
}

function setBullet(doc: XmlDoc, pPr: XmlEl, on: boolean, level: number) {
  for (const b of [...BULLETS, "a:buNone"]) for (const c of kids(pPr, b)) pPr.removeChild(c);
  if (on) {
    if (!pPr.getAttribute("marL")) (pPr.setAttribute("marL", String(285_750 * (level + 1))), pPr.setAttribute("indent", "-285750"));
    if (!kid(pPr, "a:buFont")) insertOrdered(pPr, element(doc, "a:buFont", { typeface: "Arial", panose: "020B0604020202020204", pitchFamily: 34, charset: 0 }), P_ORDER);
    insertOrdered(pPr, element(doc, "a:buChar", { char: "•" }), P_ORDER);
  } else {
    pPr.removeAttribute("marL");
    pPr.removeAttribute("indent");
    for (const c of kids(pPr, "a:buFont")) pPr.removeChild(c);
    insertOrdered(pPr, element(doc, "a:buNone"), P_ORDER);
  }
}

/**
 * Replaces the paragraphs of a text body by lines. Line n takes the look of paragraph n it replaces
 * (the last one for added lines); `style` is for a body the app just made (nothing to copy from).
 */
export function writeParagraphs(doc: XmlDoc, body: XmlEl, text: string, style?: TextStyle) {
  const old = kids(body, "a:p");
  const withRuns = old.filter((p) => kids(p, "a:r").length);
  const lines = text.replace(/\r/g, "").replace(/\n+$/, "").split("\n");
  for (const p of old) body.removeChild(p);
  lines.forEach((line, i) => {
    const m = line.match(/^(\t+|(?: {2})+)?\s?([-•*]\s+)?(.*)$/)!;
    const indent = m[1] ?? "";
    const level = Math.min(8, indent.includes("\t") ? indent.length : indent.length / 2);
    const bullet = !!m[2];
    const from = (old[i] && kids(old[i], "a:r").length ? old[i] : undefined) ?? withRuns[Math.min(i, withRuns.length - 1)] ?? old[i] ?? old[0];
    const p = element(doc, "a:p");
    const fromPr = from && kid(from, "a:pPr");
    const pPr = fromPr ? (fromPr.cloneNode(true) as XmlEl) : element(doc, "a:pPr");
    if (level) pPr.setAttribute("lvl", String(level));
    else pPr.removeAttribute("lvl");
    if (style?.align) pPr.setAttribute("algn", style.align);
    // Bullets as the line says, when it differs from the paragraph it replaces (a placeholder's own bullets are left to it).
    if (bullet !== hasBullet(pPr) && (bullet || hasBullet(fromPr))) setBullet(doc, pPr, bullet, level);
    if (pPr.attributes.length || pPr.childNodes.length) p.appendChild(pPr);
    const modelRun = from && (kids(from, "a:r").find((r) => !isBold(r)) ?? kids(from, "a:r")[0]);
    const model = modelRun ? kid(modelRun, "a:rPr") : from && kid(from, "a:endParaRPr");
    // **bold** parts as their own runs; a line without any keeps the model's weight.
    const parts = m[3].split(/\*\*(.+?)\*\*/g);
    const marked = parts.length > 1;
    parts.forEach((part, k) => {
      if (!part) return;
      const rPr = runProps(doc, model);
      if (style) styleRun(doc, rPr, style);
      if (marked) rPr.setAttribute("b", k % 2 ? "1" : "0");
      p.appendChild(element(doc, "a:r", {}, [rPr, element(doc, "a:t", {}, [part])]));
    });
    const end = from && kid(from, "a:endParaRPr");
    if (end) p.appendChild(end.cloneNode(true));
    else if (style) {
      const e = element(doc, "a:endParaRPr", { lang: "en-US", dirty: 0 });
      styleRun(doc, e, { size: style.size });
      p.appendChild(e);
    }
    body.appendChild(p);
  });
}
