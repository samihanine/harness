/**
 * A markdown deck (shared/deck.ts) as a native PowerPoint file, in the template's style (Dior: white,
 * black text, gold accents, Century Gothic): cover, agenda, section and end pages with their photos, and
 * content layouts laid out by code (text stays editable; tables, charts native; pictures fitted). The
 * template's pictures are in assets/template (cover.jpg, section.jpg, agenda.jpg, end.jpg, logo.png when
 * a logo picture is given; else the brand is written).
 */
import PptxGenJS from "pptxgenjs";
import { parseDeck, plain, type Block, type Cell, type Deck, type Inline, type Slide } from "@shared/deck";
import { loadPicture, type Picture } from "./images";
import { staticFile } from "../static";

const W = 13.333;
const H = 7.5;
const FONT = "Century Gothic";
const C = { text: "000000", muted: "595959", gold: "B39251", navy: "363D4E", light: "F7F3EA", line: "EBE0CB", white: "FFFFFF", palette: ["B39251", "1C4358", "5DA6CC", "CCA462", "363D4E", "1EA185", "9E813E", "91C3DC"] };
const M = 0.6; // side margin
const TOP = 1.45; // content starts under the header

type Box = { x: number; y: number; w: number; h: number };
type Pptx = InstanceType<typeof PptxGenJS>;
type PSlide = ReturnType<Pptx["addSlide"]>;

/** Text runs of inline markdown, for pptxgenjs. */
const runs = (content: Inline[], base: Record<string, unknown> = {}) =>
  content.map((c) => ({ text: c.text, options: { ...base, bold: c.bold || (base.bold as boolean) || undefined, italic: c.italic || undefined, fontFace: c.code ? "Consolas" : (base.fontFace as string) ?? FONT, ...(c.link ? { hyperlink: { url: c.link }, color: C.gold } : {}) } }));

/** Lines a text takes in a width at a font size (an estimate: average character width ~0.5 em). */
const linesOf = (text: string, width: number, size: number) => Math.max(1, Math.ceil((text.length * size * 0.0075) / Math.max(0.5, width)));

const template = (name: string) => (staticFile(`assets/template/${name}`) ? `template:${name}` : undefined);

async function picture(src: string | undefined): Promise<Picture | undefined> {
  if (!src) return undefined;
  return loadPicture(src).catch(() => undefined);
}

/** A picture fitted inside a box (contain) or filling it (cover: cropped). */
function addPicture(slide: PSlide, pic: Picture, box: Box, mode: "contain" | "cover" = "contain") {
  const ratio = pic.width / pic.height;
  if (mode === "cover") {
    slide.addImage({ data: pic.data, x: box.x, y: box.y, w: box.w, h: box.h, sizing: { type: "cover", w: box.w, h: box.h } as never });
    return;
  }
  let w = box.w;
  let h = w / ratio;
  if (h > box.h) (h = box.h), (w = h * ratio);
  slide.addImage({ data: pic.data, x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h });
}

function missingPicture(slide: PSlide, box: Box, src: string) {
  slide.addShape("rect", { ...box, fill: { color: C.light }, line: { color: C.line } });
  slide.addText(`Picture not found\n${src.slice(0, 80)}`, { ...box, fontFace: FONT, fontSize: 10, color: C.muted, align: "center", valign: "middle" });
}

function brand(slide: PSlide, deck: Deck, box: Box, color = C.text, size = 16) {
  const logo = template("logo.png") && logoData;
  if (logo) slide.addImage({ data: logo, ...box, sizing: { type: "contain", w: box.w, h: box.h } as never });
  else slide.addText(deck.head.brand.split("").join(" "), { ...box, fontFace: "Didot", fontSize: size, color, bold: false, charSpacing: 4, margin: 0, valign: "middle" });
}

/** The template's logo picture (assets/template/logo.png), read once per export. */
let logoData: string | undefined;

/* ------------------------------ Content blocks ------------------------------ */

/** Height a block needs in a width, at a scale of the font sizes. */
function blockHeight(b: Block, w: number, k: number): number {
  switch (b.type) {
    case "heading":
      return linesOf(plain(b.content), w, 18 * k) * 0.34 * k + 0.08;
    case "paragraph":
      return linesOf(plain(b.content), w, 15 * k) * 0.29 * k + 0.12;
    case "list":
      return b.items.reduce((n, it) => n + linesOf(plain(it.content), w - 0.35 * (it.level + 1), (it.level ? 13 : 15) * k) * 0.29 * k + 0.05, 0) + 0.1;
    case "quote":
      return 0.3 + (b.title ? 0.35 * k : 0) + b.content.reduce((n, l) => n + linesOf(plain(l), w - 0.5, 14 * k) * 0.28 * k, 0);
    case "table":
      return (1 + b.rows.length) * rowHeight(b) * k + 0.1;
    case "image":
      return Math.min(4.2, w * 0.56);
    case "chart":
      return Math.min(4.2, w * 0.6);
    case "code":
      return b.text.split("\n").length * 0.22 * k + 0.25;
  }
}

const rowHeight = (b: Extract<Block, { type: "table" }>) => (b.rows.some((r) => r.some((c) => c.image)) || b.header.some((c) => c.image) ? 1.0 : 0.38);

async function drawBlock(slide: PSlide, b: Block, box: Box, k: number) {
  switch (b.type) {
    case "heading":
      slide.addText(runs(b.content, { bold: true, color: C.navy }), { ...box, fontFace: FONT, fontSize: 18 * k, valign: "top", margin: 0 });
      return;
    case "paragraph":
      slide.addText(runs(b.content), { ...box, fontFace: FONT, fontSize: 15 * k, color: C.text, valign: "top", margin: 0, paraSpaceAfter: 4 });
      return;
    case "list": {
      const items = b.items.flatMap((it, n) => {
        const parts = runs(it.content, { fontSize: (it.level ? 13 : 15) * k, color: it.level ? C.muted : C.text });
        parts[0] = { ...parts[0], options: { ...parts[0].options, bullet: b.ordered && !it.level ? { type: "number" } : { code: it.level ? "2013" : "25AA", color: C.gold } as never, indentLevel: it.level, paraSpaceAfter: 5 } as never };
        if (n < b.items.length - 1) parts[parts.length - 1] = { ...parts[parts.length - 1], options: { ...parts[parts.length - 1].options, breakLine: true } as never };
        return parts;
      });
      slide.addText(items as never, { ...box, fontFace: FONT, valign: "top", margin: 0 });
      return;
    }
    case "quote": {
      slide.addShape("rect", { ...box, fill: { color: C.light }, line: { color: C.light } });
      slide.addShape("rect", { x: box.x, y: box.y, w: 0.06, h: box.h, fill: { color: C.gold }, line: { color: C.gold } });
      const parts: any[] = [];
      if (b.title) parts.push({ text: b.title, options: { bold: true, color: C.gold, fontSize: 14 * k, breakLine: true } });
      b.content.forEach((l, n) => parts.push(...runs(l, { fontSize: 14 * k, color: C.text }).map((r, m, all) => (m === all.length - 1 && n < b.content.length - 1 ? { ...r, options: { ...r.options, breakLine: true } } : r))));
      slide.addText(parts, { x: box.x + 0.25, y: box.y + 0.1, w: box.w - 0.4, h: box.h - 0.2, fontFace: FONT, valign: "middle", margin: 0 });
      return;
    }
    case "table":
      await drawTable(slide, b, box, k);
      return;
    case "image": {
      const pic = await picture(b.src);
      pic ? addPicture(slide, pic, box) : missingPicture(slide, box, b.src);
      if (b.alt && pic) slide.addText(b.alt, { x: box.x, y: box.y + box.h - 0.02, w: box.w, h: 0.25, fontFace: FONT, fontSize: 9, color: C.muted, align: "center", margin: 0 });
      return;
    }
    case "chart":
      drawChart(slide, b.spec, box);
      return;
    case "code":
      slide.addShape("rect", { ...box, fill: { color: "F4F4F5" }, line: { color: "E4E4E7" } });
      slide.addText(b.text, { x: box.x + 0.15, y: box.y + 0.1, w: box.w - 0.3, h: box.h - 0.2, fontFace: "Consolas", fontSize: 11 * k, color: C.text, valign: "top", margin: 0 });
      return;
  }
}

async function drawTable(slide: PSlide, b: Extract<Block, { type: "table" }>, box: Box, k: number) {
  const cols = Math.max(b.header.length, ...b.rows.map((r) => r.length));
  const rowH = rowHeight(b) * k;
  const colW = Array.from({ length: cols }, () => box.w / cols);
  const cellOf = (c: Cell | undefined, header: boolean) => ({
    text: c && !c.image ? (runs(c.content, { fontSize: (header ? 12 : 11) * k, bold: header || undefined, color: header ? C.white : C.text }) as never) : "",
    options: { fill: { color: header ? C.navy : C.white }, valign: "middle" as const, margin: 0.06 },
  });
  const rows = [b.header, ...b.rows].map((r, n) => Array.from({ length: cols }, (_, i) => cellOf(r[i], n === 0)));
  slide.addTable(rows as never, { x: box.x, y: box.y, w: box.w, colW, rowH: Array(rows.length).fill(rowH), fontFace: FONT, border: { type: "solid", pt: 0.5, color: C.line }, autoPage: false });
  // Pictures in cells: drawn over the table (a table cell holds no picture in PowerPoint).
  const all = [b.header, ...b.rows];
  for (let r = 0; r < all.length; r++)
    for (let c = 0; c < cols; c++) {
      const img = all[r][c]?.image;
      if (!img) continue;
      const cell = { x: box.x + c * colW[c] + 0.05, y: box.y + r * rowH + 0.05, w: colW[c] - 0.1, h: rowH - 0.1 };
      const pic = await picture(img.src);
      pic ? addPicture(slide, pic, cell) : missingPicture(slide, cell, img.src);
    }
}

function drawChart(slide: PSlide, spec: Extract<Block, { type: "chart" }>["spec"], box: Box) {
  const types: Record<string, string> = { bar: "bar", column: "bar", line: "line", pie: "pie", doughnut: "doughnut", area: "area" };
  const data = (spec.series ?? []).map((s) => ({ name: s.name, labels: spec.labels, values: s.values.map(Number) }));
  if (!data.length) return missingPicture(slide, box, "chart without series");
  slide.addChart((types[spec.type] ?? "bar") as never, data as never, {
    ...box,
    barDir: spec.type === "bar" ? "bar" : "col",
    barGrouping: spec.stacked ? "stacked" : "clustered",
    // One series: one color (PowerPoint would color each bar).
    chartColors: data.length === 1 && spec.type !== "pie" && spec.type !== "doughnut" ? [C.palette[0]] : C.palette,
    showLegend: data.length > 1 || spec.type === "pie" || spec.type === "doughnut",
    legendPos: "b",
    legendFontFace: FONT,
    catAxisLabelFontFace: FONT,
    valAxisLabelFontFace: FONT,
    catAxisLabelFontSize: 10,
    valAxisLabelFontSize: 10,
    showTitle: !!spec.title,
    title: spec.title,
    titleFontFace: FONT,
    titleFontSize: 13,
    showValue: spec.type === "pie" || spec.type === "doughnut",
    dataLabelFontSize: 9,
    valGridLine: { color: "E5E5E5", size: 0.5 },
    catGridLine: { style: "none" },
  } as never);
}

/** Blocks stacked in a box, the font made smaller (down to 70 %) when they would not fit. */
async function drawBlocks(slide: PSlide, blocks: Block[], box: Box) {
  if (!blocks.length) return;
  const fixed = (b: Block) => b.type === "image" || b.type === "chart";
  let k = 1;
  const total = (scale: number) => blocks.reduce((n, b) => n + blockHeight(b, box.w, scale) + 0.12, 0);
  while (k > 0.7 && total(k) > box.h) k -= 0.05;
  // Little content: bigger text, so the slide does not look empty (pictures and charts take the rest anyway).
  if (!blocks.some(fixed)) while (k < 1.6 && total(k + 0.05) < box.h * 0.8) k += 0.05;
  // Pictures and charts share the room left by the text.
  const flexible = blocks.filter(fixed).length;
  const textH = blocks.filter((b) => !fixed(b)).reduce((n, b) => n + blockHeight(b, box.w, k) + 0.12, 0);
  const share = flexible ? Math.max(1.2, (box.h - textH) / flexible - 0.12) : 0;
  let y = box.y;
  for (const b of blocks) {
    const h = fixed(b) ? share : blockHeight(b, box.w, k);
    await drawBlock(slide, b, { x: box.x, y, w: box.w, h: Math.min(h, box.y + box.h - y) }, k);
    y += h + 0.12;
  }
}

/* --------------------------------- Layouts --------------------------------- */

function chrome(slide: PSlide, deck: Deck, s: Slide, number: number) {
  brand(slide, deck, { x: M, y: 0.32, w: 1.6, h: 0.32 }, C.text, 13);
  slide.addShape("line", { x: M, y: 1.22, w: W - 2 * M, h: 0, line: { color: C.gold, width: 0.75 } });
  if (s.title) slide.addText(runs(s.title), { x: M, y: 0.62, w: W - 2 * M, h: 0.56, fontFace: FONT, fontSize: 24, color: C.text, valign: "bottom", margin: 0, fit: "shrink" });
  slide.addText(String(number), { x: W - M - 0.6, y: H - 0.45, w: 0.6, h: 0.3, fontFace: FONT, fontSize: 9, color: C.muted, align: "right", margin: 0 });
  if (deck.head.footer) slide.addText(deck.head.footer, { x: M, y: H - 0.45, w: 8, h: 0.3, fontFace: FONT, fontSize: 9, color: C.muted, margin: 0 });
}

const body: Box = { x: M, y: TOP, w: W - 2 * M, h: H - TOP - 0.6 };

async function heroPhoto(slide: PSlide, s: Slide, file: string, box: Box) {
  const pic = (await picture(s.meta.image)) ?? (await picture(template(file)));
  if (pic) addPicture(slide, pic, box, "cover");
  else slide.addShape("rect", { ...box, fill: { color: C.navy }, line: { color: C.navy } });
}

async function drawSlide(pptx: Pptx, deck: Deck, s: Slide) {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  if (s.notes) slide.addNotes(s.notes);
  const n = s.index + 1;
  const title = s.title ? plain(s.title) : "";
  switch (s.layout) {
    case "cover": {
      await heroPhoto(slide, s, "cover.jpg", { x: W / 2, y: 0, w: W / 2, h: H });
      brand(slide, deck, { x: 0.8, y: 0.7, w: 2.4, h: 0.5 }, C.text, 22);
      slide.addShape("line", { x: 0.8, y: 3.0, w: 1.2, h: 0, line: { color: C.gold, width: 1.5 } });
      slide.addText(title, { x: 0.8, y: 3.15, w: W / 2 - 1.4, h: 1.5, fontFace: FONT, fontSize: 34, color: C.text, valign: "top", margin: 0, fit: "shrink" });
      const sub = s.blocks.filter((b) => b.type === "paragraph").map((b) => plain((b as { content: Inline[] }).content));
      if (sub.length) slide.addText(sub.join("\n"), { x: 0.8, y: 4.8, w: W / 2 - 1.4, h: 1.6, fontFace: FONT, fontSize: 14, color: C.muted, valign: "top", margin: 0, paraSpaceAfter: 6 });
      return;
    }
    case "section": {
      const number = deck.slides.filter((x) => x.layout === "section" && x.index <= s.index).length;
      await heroPhoto(slide, s, "section.jpg", { x: 0, y: 0, w: W, h: H });
      slide.addShape("rect", { x: 0.8, y: 2.3, w: 6.2, h: 2.9, fill: { color: C.white }, line: { color: C.white } });
      slide.addText(String(number).padStart(2, "0"), { x: 1.2, y: 2.55, w: 2, h: 0.9, fontFace: FONT, fontSize: 44, color: C.gold, margin: 0 });
      slide.addText(title, { x: 1.2, y: 3.5, w: 5.4, h: 1.4, fontFace: FONT, fontSize: 28, color: C.text, valign: "top", margin: 0, fit: "shrink" });
      return;
    }
    case "agenda": {
      await heroPhoto(slide, s, "agenda.jpg", { x: 0, y: 0, w: 4.4, h: H });
      brand(slide, deck, { x: 5.2, y: 0.5, w: 1.6, h: 0.32 }, C.text, 13);
      slide.addText(title || "Agenda", { x: 5.2, y: 1.0, w: 7, h: 0.8, fontFace: FONT, fontSize: 30, color: C.text, margin: 0 });
      const list = s.blocks.find((b) => b.type === "list") as Extract<Block, { type: "list" }> | undefined;
      const items = (list?.items ?? []).filter((i) => i.level === 0);
      const active = Number(s.meta.active) || 0;
      const step = Math.min(0.75, 4.8 / Math.max(1, items.length));
      items.forEach((it, i) => {
        const y = 2.2 + i * step;
        const on = active === i + 1;
        if (on) slide.addShape("rect", { x: 5.2, y: y + 0.08, w: 0.06, h: step - 0.2, fill: { color: C.gold }, line: { color: C.gold } });
        slide.addText(String(i + 1).padStart(2, "0"), { x: 5.45, y, w: 0.8, h: step - 0.05, fontFace: FONT, fontSize: 18, color: C.gold, margin: 0, valign: "middle" });
        slide.addText(plain(it.content), { x: 6.3, y, w: 6.4, h: step - 0.05, fontFace: FONT, fontSize: 18, bold: on, color: active && !on ? C.muted : C.text, margin: 0, valign: "middle" });
      });
      return;
    }
    case "end": {
      await heroPhoto(slide, s, "end.jpg", { x: 0, y: 0, w: W, h: H });
      slide.addShape("rect", { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 55 }, line: { color: "000000", transparency: 100 } });
      brand(slide, deck, { x: W / 2 - 1.5, y: 2.6, w: 3, h: 0.7 }, C.white, 30);
      const text = [title, ...s.blocks.filter((b) => b.type === "paragraph").map((b) => plain((b as { content: Inline[] }).content))].filter(Boolean).join("\n");
      if (text) slide.addText(text, { x: 1, y: 3.6, w: W - 2, h: 1.5, fontFace: FONT, fontSize: 18, color: C.white, align: "center", margin: 0 });
      return;
    }
    case "full-image": {
      const pic = await picture(s.meta.image ?? (s.blocks.find((b) => b.type === "image") as { src: string } | undefined)?.src);
      if (pic) addPicture(slide, pic, { x: 0, y: 0, w: W, h: H }, s.meta.fit === "contain" ? "contain" : "cover");
      else missingPicture(slide, { x: 0, y: 0, w: W, h: H }, s.meta.image ?? "");
      if (title) {
        slide.addShape("rect", { x: 0, y: H - 1.3, w: W, h: 1.3, fill: { color: "000000", transparency: 35 }, line: { color: "000000", transparency: 100 } });
        slide.addText(title, { x: M, y: H - 1.2, w: W - 2 * M, h: 1.1, fontFace: FONT, fontSize: 26, color: C.white, valign: "middle", margin: 0 });
      }
      return;
    }
    case "image-left":
    case "image-right": {
      const left = s.layout === "image-left";
      const half = { x: left ? 0 : W / 2, y: 0, w: W / 2, h: H };
      const pic = await picture(s.meta.image);
      pic ? addPicture(slide, pic, half, s.meta.fit === "contain" ? "contain" : "cover") : missingPicture(slide, half, s.meta.image ?? "no image:");
      const x = left ? W / 2 + 0.5 : M;
      brand(slide, deck, { x, y: 0.32, w: 1.6, h: 0.32 }, C.text, 13);
      if (title) slide.addText(title, { x, y: 0.75, w: W / 2 - 1.1, h: 0.9, fontFace: FONT, fontSize: 24, color: C.text, valign: "bottom", margin: 0, fit: "shrink" });
      slide.addShape("line", { x, y: 1.7, w: 1.2, h: 0, line: { color: C.gold, width: 1 } });
      await drawBlocks(slide, s.blocks, { x, y: 1.95, w: W / 2 - 1.1, h: H - 2.5 });
      slide.addText(String(n), { x: left ? W - M - 0.6 : W / 2 - 1.1, y: H - 0.45, w: 0.6, h: 0.3, fontFace: FONT, fontSize: 9, color: C.muted, align: "right", margin: 0 });
      return;
    }
    case "kpis": {
      chrome(slide, deck, s, n);
      const list = s.blocks.find((b) => b.type === "list") as Extract<Block, { type: "list" }> | undefined;
      const items = (list?.items ?? []).filter((i) => i.level === 0).slice(0, 8);
      const rest = s.blocks.filter((b) => b !== list);
      const perRow = items.length <= 4 ? items.length : Math.ceil(items.length / 2);
      const rowsN = Math.ceil(items.length / Math.max(1, perRow));
      const gap = 0.3;
      const cardW = (body.w - gap * (perRow - 1)) / Math.max(1, perRow);
      const cardH = Math.min(rest.length ? 2.2 : 2.8, (body.h - (rest.length ? 1.4 : 0) - gap * (rowsN - 1)) / Math.max(1, rowsN));
      // Alone on the slide: the cards in the middle of the page.
      const top = rest.length ? body.y + 0.1 : body.y + Math.max(0.1, (body.h - rowsN * cardH - (rowsN - 1) * gap) / 2.5);
      items.forEach((it, i) => {
        const x = body.x + (i % perRow) * (cardW + gap);
        const y = top + Math.floor(i / perRow) * (cardH + gap);
        const value = it.content.find((c) => c.bold)?.text ?? plain(it.content).split(" ")[0];
        const label = plain(it.content.filter((c) => !c.bold)).trim();
        const [main, detail] = label.split(/\s+[—–-]\s+/, 2);
        slide.addShape("rect", { x, y, w: cardW, h: cardH, fill: { color: C.white }, line: { color: C.line, width: 1 } });
        slide.addShape("rect", { x, y, w: cardW, h: 0.06, fill: { color: C.gold }, line: { color: C.gold } });
        slide.addText(value, { x: x + 0.2, y: y + 0.25, w: cardW - 0.4, h: cardH * 0.42, fontFace: FONT, fontSize: cardH > 2.4 ? 36 : 30, color: C.navy, bold: true, margin: 0, fit: "shrink", valign: "middle" });
        slide.addText([{ text: main ?? "", options: { fontSize: cardH > 2.4 ? 15 : 13, color: C.text, breakLine: !!detail } }, ...(detail ? [{ text: detail, options: { fontSize: cardH > 2.4 ? 12 : 11, color: C.muted } }] : [])], { x: x + 0.2, y: y + 0.25 + cardH * 0.42, w: cardW - 0.4, h: cardH * 0.45, fontFace: FONT, margin: 0, valign: "top" });
      });
      if (rest.length) await drawBlocks(slide, rest, { x: body.x, y: body.y + rowsN * (cardH + gap) + 0.1, w: body.w, h: body.h - rowsN * (cardH + gap) - 0.1 });
      return;
    }
    case "quote": {
      chrome(slide, deck, { ...s, title: undefined }, n);
      const q = s.blocks.find((b) => b.type === "quote") as Extract<Block, { type: "quote" }> | undefined;
      const author = s.blocks.find((b) => b.type === "paragraph" && /^[—–-]/.test(plain(b.content))) as { content: Inline[] } | undefined;
      slide.addText("“", { x: 1.2, y: 1.2, w: 1.2, h: 1.4, fontFace: "Georgia", fontSize: 96, color: C.gold, margin: 0 });
      const text = q ? q.content.map(plain).join("\n") : title;
      slide.addText(text, { x: 2.0, y: 2.0, w: W - 4, h: 3.0, fontFace: FONT, fontSize: 26, italic: true, color: C.text, valign: "middle", margin: 0, fit: "shrink" });
      if (author) slide.addText(plain(author.content), { x: 2.0, y: 5.1, w: W - 4, h: 0.5, fontFace: FONT, fontSize: 14, color: C.muted, margin: 0 });
      if (title && q) slide.addText(title, { x: M, y: 0.62, w: W - 2 * M, h: 0.56, fontFace: FONT, fontSize: 24, color: C.text, valign: "bottom", margin: 0 });
      return;
    }
    case "two-cols": {
      chrome(slide, deck, s, n);
      const w = (body.w - 0.5) / 2;
      await drawBlocks(slide, s.blocks, { ...body, w });
      await drawBlocks(slide, s.right, { ...body, x: body.x + w + 0.5, w });
      return;
    }
    case "columns": {
      chrome(slide, deck, s, n);
      const cols = s.columns.length ? s.columns : [s.blocks];
      const intro = s.columns.length ? s.blocks : [];
      const introH = intro.length ? intro.reduce((h, b) => h + blockHeight(b, body.w, 1) + 0.12, 0) : 0;
      if (intro.length) await drawBlocks(slide, intro, { ...body, h: introH });
      const gap = 0.35;
      const w = (body.w - gap * (cols.length - 1)) / cols.length;
      for (let i = 0; i < cols.length; i++) {
        const x = body.x + i * (w + gap);
        slide.addShape("rect", { x, y: body.y + introH, w, h: 0.05, fill: { color: C.gold }, line: { color: C.gold } });
        await drawBlocks(slide, cols[i], { x, y: body.y + introH + 0.2, w, h: body.h - introH - 0.2 });
      }
      return;
    }
    default:
      chrome(slide, deck, s, n);
      await drawBlocks(slide, [...s.blocks, ...s.right], body);
  }
}

/** The deck as a .pptx file. */
export async function deckToPptx(source: string): Promise<Buffer> {
  const deck = parseDeck(source);
  logoData = template("logo.png") ? (await loadPicture("template:logo.png").catch(() => undefined))?.data : undefined;
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = deck.head.title;
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
  for (const s of deck.slides) await drawSlide(pptx, deck, s);
  if (!deck.slides.length) pptx.addSlide();
  return (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
}
