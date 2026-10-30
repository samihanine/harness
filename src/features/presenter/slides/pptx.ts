/**
 * The deck as a native, editable PowerPoint file (shapes and text; live reports become links). Images
 * are read first (PowerPoint gets their data, in their proportions); what React draws — the title
 * pages' flowers, the mascot, icons — becomes pictures drawn here in the deck's colors.
 */
import { createElement } from "react";
import JSZip from "jszip";
import PptxGenJS from "pptxgenjs";
import { toast } from "sonner";
import { Mascot, type Outfit } from "@/features/mascot";
import { preparePptxImages, setPptxPicture } from "./images";
import { brandOf, hex, plain } from "./kit";
import { checkSlide, deckInfo, type Deck } from ".";
import { iconOf } from "./picture";
import { svgPicture } from "./raster";
import { themeStyle } from "./stage";
import { ToileFlowers } from "./toile";
import { loadReportLinks } from "./powerbi";

/** The mascots and icons a deck shows (keys "mascot" and "icon" of its slides, at any depth). */
function drawnPictures(deck: Deck) {
  const mascots = new Set<string>();
  const icons = new Set<string>();
  const walk = (v: unknown, key = "") => {
    if (typeof v === "string" && key === "mascot") mascots.add(v);
    else if (typeof v === "string" && key === "icon" && v.trim()) icons.add(v.trim());
    else if (Array.isArray(v)) v.forEach((x) => walk(x));
    else if (v && typeof v === "object") Object.entries(v).forEach(([k, x]) => walk(x, k));
  };
  deck.slides.forEach((s) => walk(s));
  // Icons drawn by templates themselves.
  if (deck.slides.some((s) => s.template === "beware")) icons.add("badge-check");
  // Default icons of the people and boxes drawn without one.
  if (deck.slides.some((s) => ["usecase", "conversation", "responsibilities", "architecture"].includes(s.template))) ["user-round", "box", "sparkles"].forEach((i) => icons.add(i));
  if (deck.slides.some((s) => s.template === "features")) for (const s of deck.slides.filter((x) => x.template === "features")) for (const it of (s.items as { icon?: string }[]) ?? []) icons.add(it.icon?.trim() || "sparkles");
  return { mascots: [...mascots], icons: [...icons] };
}

/** Draws the flowers, mascots and icons of a deck as pictures for the export; returns the flowers' data. */
async function drawPictures(deck: Deck) {
  const theme = themeStyle(deck.theme) as Record<string, string>;
  const { mascots, icons } = drawnPictures(deck);
  for (const outfit of mascots) setPptxPicture(`mascot:${outfit}`, await svgPicture(createElement(Mascot, { outfit: outfit as Outfit, still: true, size: 176 }), { width: 176, height: 215, scale: 3 }));
  for (const name of icons) {
    setPptxPicture(`icon:${name}`, await svgPicture(createElement(iconOf(name), { width: 96, height: 96, strokeWidth: 1.6, color: deck.theme.primary }), { width: 96, height: 96, scale: 4 }));
    // White too: on the accent (avatars, headers).
    setPptxPicture(`icon-white:${name}`, await svgPicture(createElement(iconOf(name), { width: 96, height: 96, strokeWidth: 1.8, color: "#ffffff" }), { width: 96, height: 96, scale: 4 }));
  }
  if (!deck.slides.some((s) => s.template === "title")) return undefined;
  return (await svgPicture(createElement(ToileFlowers, { className: "size-full" }), { width: 1600, height: 900, className: "slide-still", style: theme }))?.data;
}

/**
 * Content pages: the blocks under the title (placed from 1.9 in down) are moved up, the title block
 * sitting higher too (pptxTitle): less room lost between the header and the content.
 */
function liftContent(slide: PptxGenJS.Slide, header: boolean) {
  // The title in the header: the content moves up further.
  const LIFT = header ? 0.6 : 0.4;
  const up = <O,>(o: O): O => {
    const y = (o as unknown as { y?: unknown } | undefined)?.y;
    return typeof y === "number" && y >= 1.9 ? ({ ...(o as object), y: y - LIFT } as O) : o;
  };
  const { addText, addShape, addImage, addTable } = slide;
  slide.addText = (text, o) => addText.call(slide, text, up(o));
  slide.addShape = (shape, o) => addShape.call(slide, shape, up(o));
  slide.addImage = (o) => addImage.call(slide, up(o));
  slide.addTable = (rows, o) => addTable.call(slide, rows, up(o));
}

export async function downloadPptx(deck: Deck, name: string) {
  await loadReportLinks();
  const missing = await preparePptxImages(deck);
  if (missing.length) toast.warning(`${missing.length} image(s) could not be read (the site does not allow it) and are left out of the PowerPoint: ${missing.slice(0, 3).join(", ")}${missing.length > 3 ? "…" : ""}`);
  const toile = await drawPictures(deck);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = deck.title;
  const infos = deckInfo(deck);
  deck.slides.forEach((s, i) => {
    const { template, args } = checkSlide(s);
    const slide = pptx.addSlide();
    slide.background = { color: hex(deck.theme.background) };
    const content = !template.bare && (template.kind ?? "content") === "content";
    if (content) liftContent(slide, deck.theme.header !== "none");
    template.pptx(slide, args, { theme: deck.theme, deck: infos[i], pptx, toile });
    // The header of content pages: where the page is (deck › part › subpart).
    const info = infos[i];
    if (content && deck.theme.header !== "none") {
      const trail = [deck.title, info.sections[info.section], info.subsection >= 0 ? info.subsections[info.subsection] : undefined].filter(Boolean).join("  ›  ");
      // As on screen: the brand letter, the page title on the left, where the page is (or the brand) on the right.
      const title = plain(String(args.title ?? ""));
      const number = deck.theme.header === "brand" && info.group.count > 1 ? ` (${info.group.index + 1}/${info.group.count})` : "";
      slide.addText(brandOf(deck.theme).charAt(0), { x: 0.3, y: 0.12, w: 0.5, h: 0.5, fontSize: 26, fontFace: "Bodoni MT", color: hex(deck.theme.primary), align: "center", valign: "middle" });
      slide.addText(`${title}${number}`, { x: 0.85, y: 0.12, w: 7.6, h: 0.5, fontSize: 22, bold: true, fontFace: deck.theme.fontTitle, color: hex(deck.theme.text), valign: "middle" });
      slide.addText(deck.theme.header === "brand" ? brandOf(deck.theme) : trail.toUpperCase(), { x: 8.5, y: 0.2, w: 4.4, h: 0.34, fontSize: deck.theme.header === "brand" ? 12 : 8, italic: deck.theme.header === "brand", charSpacing: deck.theme.header === "brand" ? 0 : 2, align: "right", color: hex(deck.theme.primary), fontFace: deck.theme.fontText });
      if (deck.theme.header !== "breadcrumb") slide.addShape("line", { x: 0.4, y: 0.72, w: 12.5, h: 0, line: { color: hex(deck.theme.text), width: 0.5, transparency: 75 } });
    }
    if (!template.bare) slide.addText(`${i + 1}`, { x: 12.3, y: 6.9, w: 0.6, h: 0.3, fontSize: 9, align: "right", color: hex(deck.theme.text), transparency: 50, fontFace: deck.theme.fontText });
  });
  const file = await withMorph((await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer);
  const url = URL.createObjectURL(new Blob([file], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
  Object.assign(document.createElement("a"), { href: url, download: `${name.replace(/[^\p{L}\p{N} _-]+/gu, "-")}.pptx` }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** The Morph transition on every slide (pptxgenjs has no transitions: written into each slide's XML; fade where Morph is unknown). */
const MORPH = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1200"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;
export async function withMorph(file: ArrayBuffer) {
  const zip = await JSZip.loadAsync(file);
  for (const path of Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))) {
    const xml = await zip.file(path)!.async("string");
    // The transition comes right after the slide's content (and its color map), before its timing.
    const at = xml.includes("</p:clrMapOvr>") ? xml.indexOf("</p:clrMapOvr>") + "</p:clrMapOvr>".length : xml.indexOf("</p:cSld>") + "</p:cSld>".length;
    zip.file(path, xml.slice(0, at) + MORPH + xml.slice(at));
  }
  return zip.generateAsync({ type: "arraybuffer" });
}
