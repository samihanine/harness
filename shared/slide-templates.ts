/**
 * The templates a new browser starts with: the company's look (Century Gothic, gold, navy; its photos and logo
 * in assets/template). Copied into the browser's templates, then free to change, duplicate or replace.
 */
import type { Element, Template } from "./slides";

const GOLD = "#B39251";
const NAVY = "#1C4358";
const INK = "#000000";
const MUTED = "#595959";
const SAND = "#F7F3EA";

/** The header of content slides: the logo on the left, the title on the right, on one line; a gold rule under. */
const header: Element[] = [
  { type: "image", src: "asset:logo.svg", x: 0.6, y: 0.42, w: 1.1, h: 0.42, fit: "contain" },
  { type: "text", slot: "title", x: 2.0, y: 0.3, w: 10.73, h: 0.66, style: { size: 24, color: INK, align: "right", valign: "middle" } },
  { type: "shape", shape: "line", x: 0.6, y: 1.1, w: 12.13, h: 0, line: { color: GOLD, width: 1 } },
];
const footer: Element = { type: "text", slot: "footer", x: 0.6, y: 7.0, w: 9, h: 0.3, style: { size: 9, color: MUTED } };

export const DEFAULT_TEMPLATES: Template[] = [
  {
    id: "cover",
    name: "Cover",
    description: "First slide: the cover photo on the left, the logo, the title, a subtitle and a date (field date:).",
    builtin: true,
    elements: [
      { type: "image", src: "asset:cover.jpg", x: 0, y: 0, w: 6.2, h: 7.5 },
      { type: "image", src: "asset:logo.svg", x: 7.0, y: 1.0, w: 1.6, h: 0.6, fit: "contain" },
      { type: "text", slot: "title", x: 7.0, y: 2.4, w: 5.7, h: 1.9, style: { size: 34, color: INK, valign: "bottom" } },
      { type: "shape", shape: "line", x: 7.0, y: 4.5, w: 1.2, h: 0, line: { color: GOLD, width: 2 } },
      { type: "text", slot: "subtitle", x: 7.0, y: 4.7, w: 5.7, h: 1.0, style: { size: 16, color: MUTED } },
      { type: "text", slot: "date", x: 7.0, y: 6.5, w: 5.7, h: 0.4, style: { size: 11, color: GOLD, upper: true, spacing: 1 } },
    ],
  },
  {
    id: "agenda",
    name: "Agenda",
    description: "The plan: the boutique photo on the left, numbered lines (the slide's list items).",
    builtin: true,
    elements: [
      { type: "image", src: "asset:agenda.jpg", x: 0, y: 0, w: 4.4, h: 7.5 },
      { type: "image", src: "asset:logo.svg", x: 5.2, y: 0.55, w: 1.1, h: 0.42, fit: "contain" },
      { type: "text", slot: "title", x: 5.2, y: 1.2, w: 7.5, h: 0.8, style: { size: 28, color: INK } },
      {
        type: "items",
        x: 5.2,
        y: 2.3,
        w: 7.5,
        h: 4.6,
        direction: "column",
        gap: 0.1,
        max: 7,
        item: [
          { type: "text", slot: "index", x: 0, y: 0, w: 0.12, h: 1, style: { size: 20, color: GOLD, valign: "middle" } },
          { type: "text", slot: "text", x: 0.15, y: 0, w: 0.85, h: 1, style: { size: 18, color: INK, valign: "middle" } },
        ],
      },
    ],
  },
  {
    id: "section",
    name: "Section",
    description: "A new part of the presentation: the section photo, the title on a navy band.",
    builtin: true,
    elements: [
      { type: "image", src: "asset:section.jpg", x: 0, y: 0, w: 13.333, h: 7.5 },
      { type: "shape", x: 0, y: 5.0, w: 13.333, h: 2.5, fill: NAVY },
      { type: "text", slot: "title", x: 0.8, y: 5.3, w: 11.7, h: 1.1, style: { size: 32, color: "#FFFFFF", valign: "middle" } },
      { type: "text", slot: "subtitle", x: 0.8, y: 6.4, w: 11.7, h: 0.6, style: { size: 15, color: GOLD } },
    ],
  },
  {
    id: "content",
    name: "Text",
    description: "A message and its points: the header (logo left, title right), the body text and bullets.",
    builtin: true,
    elements: [...header, { type: "text", slot: "subtitle", x: 0.6, y: 1.3, w: 12.13, h: 0.5, style: { size: 16, color: GOLD } }, { type: "text", slot: "body", x: 0.6, y: 1.95, w: 12.13, h: 4.9, style: { size: 16, color: INK } }, footer],
  },
  {
    id: "text-picture",
    name: "Text and picture",
    description: "The body on the left, a picture (a public link) on the right.",
    builtin: true,
    elements: [...header, { type: "text", slot: "body", x: 0.6, y: 1.45, w: 6.0, h: 5.4, style: { size: 16, color: INK } }, { type: "image", slot: "image", x: 7.0, y: 1.45, w: 5.73, h: 5.4 }, footer],
  },
  {
    id: "kpis",
    name: "Key figures",
    description: "Up to 4 cards: each list item `**value** label`; a short text under them.",
    builtin: true,
    elements: [
      ...header,
      {
        type: "items",
        x: 0.6,
        y: 1.6,
        w: 12.13,
        h: 2.6,
        direction: "row",
        gap: 0.3,
        max: 4,
        item: [
          { type: "shape", x: 0, y: 0, w: 1, h: 1, fill: SAND },
          { type: "shape", x: 0, y: 0, w: 1, h: 0.04, fill: GOLD },
          { type: "text", slot: "value", x: 0.08, y: 0.18, w: 0.84, h: 0.4, style: { size: 30, color: NAVY, bold: true, valign: "middle" } },
          { type: "text", slot: "label", x: 0.08, y: 0.6, w: 0.84, h: 0.32, style: { size: 13, color: MUTED } },
        ],
      },
      { type: "text", slot: "body", x: 0.6, y: 4.5, w: 12.13, h: 2.3, style: { size: 15, color: INK } },
      footer,
    ],
  },
  {
    id: "powerbi",
    name: "Power BI view and message",
    description: "A Power BI visual or page (its link) on the left, live in the presentation; the message on the right.",
    builtin: true,
    elements: [...header, { type: "powerbi", slot: "powerbi", x: 0.6, y: 1.4, w: 8.4, h: 5.45 }, { type: "shape", x: 9.3, y: 1.4, w: 0.05, h: 5.45, fill: GOLD }, { type: "text", slot: "body", x: 9.55, y: 1.4, w: 3.18, h: 5.45, style: { size: 14, color: INK } }, footer],
  },
  {
    id: "powerbi-full",
    name: "Power BI full width",
    description: "A Power BI page or visual on the whole slide under the header.",
    builtin: true,
    elements: [...header, { type: "powerbi", slot: "powerbi", x: 0.6, y: 1.35, w: 12.13, h: 5.6 }, footer],
  },
  {
    id: "compare",
    name: "Two Power BI views",
    description: "Two Power BI links side by side (e.g. the same visual with two filters), a caption under each (fields caption, caption2).",
    builtin: true,
    elements: [
      ...header,
      { type: "powerbi", slot: "powerbi", x: 0.6, y: 1.4, w: 5.9, h: 4.9 },
      { type: "powerbi", slot: "powerbi2", x: 6.83, y: 1.4, w: 5.9, h: 4.9 },
      { type: "text", slot: "caption", x: 0.6, y: 6.4, w: 5.9, h: 0.45, style: { size: 13, color: GOLD, align: "center" } },
      { type: "text", slot: "caption2", x: 6.83, y: 6.4, w: 5.9, h: 0.45, style: { size: 13, color: GOLD, align: "center" } },
    ],
  },
  {
    id: "quote",
    name: "Quote",
    description: "One strong sentence (the body), its author (subtitle).",
    builtin: true,
    elements: [
      { type: "shape", x: 0, y: 0, w: 13.333, h: 7.5, fill: SAND },
      { type: "shape", x: 1.2, y: 2.1, w: 0.08, h: 2.6, fill: GOLD },
      { type: "text", slot: "body", x: 1.6, y: 1.9, w: 10.5, h: 3.0, style: { size: 28, color: NAVY, italic: true, valign: "middle" } },
      { type: "text", slot: "subtitle", x: 1.6, y: 5.1, w: 10.5, h: 0.6, style: { size: 14, color: GOLD } },
    ],
  },
  {
    id: "end",
    name: "End",
    description: "Last slide: the closing photo and a word (title).",
    builtin: true,
    elements: [
      { type: "image", src: "asset:end.jpg", x: 0, y: 0, w: 13.333, h: 7.5 },
      { type: "text", slot: "title", x: 0.8, y: 3.0, w: 11.7, h: 1.4, style: { size: 40, color: "#FFFFFF", align: "center", valign: "middle" } },
    ],
  },
];

/** An example of content for a template (its preview in the Visual tab). */
export function exampleFor(t: Template) {
  const slots = JSON.stringify(t.elements);
  const lines = [`template: ${t.id}`, ...(slots.includes('"date"') ? ["date: November 2026"] : []), ...(slots.includes('"caption"') ? ["caption: Left view", "caption2: Right view"] : []), "---", `# ${t.name}`];
  if (slots.includes('"subtitle"')) lines.push("## A subtitle that frames the message");
  if (slots.includes('"items"')) lines.push(...(slots.includes('"value"') ? ["- **1.97 bn$** Revenue", "- **458** Opportunities", "- **23 %** Conversion"] : ["- Key figures", "- Monthly trend", "- Regions and stages", "- Next steps"]));
  if (slots.includes('"body"')) lines.push("", "The message of the slide in one sentence.", "", "- A first point with **a figure**", "- A second point", "- What to do next");
  if (slots.includes('"image"')) lines.push("![](asset:section.jpg)");
  return lines.join("\n");
}
