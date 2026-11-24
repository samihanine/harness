/**
 * A deck written in markdown (Slidev-like): slides separated by `---` lines; a slide may start with a
 * frontmatter (`layout: …`, `image: …`) between two `---`. The first block of the file is the deck's
 * headmatter (title, footer, brand). Parsed the same way by the browser (outline, editor) and the server
 * (PowerPoint export with the template's layouts, server/pptx/deck.ts).
 */

export type Inline = { text: string; bold?: boolean; italic?: boolean; link?: string; code?: boolean };
export type ListItem = { content: Inline[]; level: number };
export type Block =
  | { type: "heading"; depth: number; content: Inline[] }
  | { type: "paragraph"; content: Inline[] }
  | { type: "list"; ordered: boolean; items: ListItem[] }
  | { type: "image"; src: string; alt: string }
  | { type: "quote"; title?: string; content: Inline[][] }
  | { type: "table"; header: Cell[]; rows: Cell[][] }
  | { type: "chart"; spec: ChartSpec }
  | { type: "code"; lang: string; text: string };

/** A table cell: text, or a picture (an image alone in the cell: `![alt](url)`). */
export type Cell = { content: Inline[]; image?: { src: string; alt: string } };

export type ChartSpec = { type: "bar" | "column" | "line" | "pie" | "doughnut" | "area"; title?: string; labels: string[]; series: { name: string; values: number[] }[]; stacked?: boolean };

/** Layouts of the template (Dior) and free layouts for any content. */
export const LAYOUTS = {
  cover: "First page: `# Title`, then paragraphs (subtitle, date). `image:` replaces the photo.",
  agenda: "Table of contents: `# Agenda` and a numbered list of the parts; `active: n` highlights part n.",
  section: "Section divider, numbered automatically: `# Part title`. `image:` replaces the photo.",
  default: "Content page: `# Title`, then any blocks: paragraphs, bullet lists (nested with 2 spaces), tables, a picture, a callout (`> **Title**` lines), a chart (```chart JSON```).",
  "two-cols": "Two columns: `# Title`, blocks for the left column, a `::right::` line, blocks for the right column.",
  columns: "2 to 4 columns: `# Title`, then each column starts with a `::col::` line (a `### heading` then blocks).",
  "image-right": "Text on the left, the picture of `image:` on the right (half the slide).",
  "image-left": "The picture of `image:` on the left, text on the right.",
  "full-image": "One picture filling the slide (`image:`), with the `# Title` over a dark band (optional).",
  kpis: "Key figures: `# Title`, then a list `- **1.97 bn** Revenue — +12 % vs 2025` (value in bold, label, optional detail after —). 2 to 6 items.",
  table: "A big table: `# Title`, a markdown table (a cell may hold only a picture: `![](url)`), an optional paragraph under it.",
  quote: "One statement: `# Title` (optional), a blockquote, an optional author line `— Name`.",
  end: "Last page: optional short text (thanks, contact).",
} as const;
export type Layout = keyof typeof LAYOUTS;

export type Slide = { index: number; line: number; layout: Layout; meta: Record<string, string>; title?: Inline[]; blocks: Block[]; right: Block[]; columns: Block[][]; notes: string };
export type Headmatter = { title: string; footer?: string; brand: string };
export type Deck = { head: Headmatter; slides: Slide[] };

const KEY = /^([A-Za-z_][\w-]*)\s*:\s*(.*)$/;

/** Inline markdown: **bold**, *italic*, `code`, [links](url). */
export function inline(text: string): Inline[] {
  const out: Inline[] = [];
  const re = /(\*\*([^*]+)\*\*|__([^_]+)__|\*([^*]+)\*|_([^_\s][^_]*)_|`([^`]+)`|\[([^\]]+)\]\(([^)\s]+)\))/g;
  let last = 0;
  for (const m of text.matchAll(re)) {
    if (m.index! > last) out.push({ text: text.slice(last, m.index) });
    if (m[2] ?? m[3]) out.push({ text: (m[2] ?? m[3])!, bold: true });
    else if (m[4] ?? m[5]) out.push({ text: (m[4] ?? m[5])!, italic: true });
    else if (m[6]) out.push({ text: m[6], code: true });
    else if (m[7]) out.push({ text: m[7], link: m[8] });
    last = m.index! + m[0].length;
  }
  if (last < text.length) out.push({ text: text.slice(last) });
  return out.filter((p) => p.text);
}

export const plain = (content: Inline[] = []) => content.map((c) => c.text).join("");

const IMAGE = /^!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)$/;

function cell(text: string): Cell {
  const img = text.trim().match(IMAGE);
  return img ? { content: [], image: { alt: img[1], src: img[2] } } : { content: inline(text.trim()) };
}

const splitRow = (line: string) =>
  line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split(/(?<!\\)\|/)
    .map((c) => c.replace(/\\\|/g, "|"));

/** Markdown blocks of a slide's body. */
export function parseBlocks(lines: string[]): Block[] {
  const blocks: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const t = line.trim();
    if (!t) {
      i++;
      continue;
    }
    const fence = t.match(/^(```|~~~)\s*(\w*)/);
    if (fence) {
      const body: string[] = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith(fence[1])) body.push(lines[i++]);
      i++;
      const text = body.join("\n");
      if (fence[2] === "chart") {
        try {
          blocks.push({ type: "chart", spec: JSON.parse(text) as ChartSpec });
        } catch {
          blocks.push({ type: "code", lang: "chart", text });
        }
      } else blocks.push({ type: "code", lang: fence[2], text });
      continue;
    }
    const heading = t.match(/^(#{1,6})\s+(.*)$/);
    if (heading) {
      blocks.push({ type: "heading", depth: heading[1].length, content: inline(heading[2]) });
      i++;
      continue;
    }
    const image = t.match(IMAGE);
    if (image) {
      blocks.push({ type: "image", alt: image[1], src: image[2] });
      i++;
      continue;
    }
    if (t.startsWith("|")) {
      const rows: string[] = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) rows.push(lines[i++]);
      const cells = rows.filter((r) => !/^\|?\s*:?-{2,}/.test(r.trim())).map(splitRow);
      if (cells.length) blocks.push({ type: "table", header: cells[0].map(cell), rows: cells.slice(1).map((r) => r.map(cell)) });
      continue;
    }
    if (t.startsWith(">")) {
      const quoted: string[] = [];
      while (i < lines.length && lines[i].trim().startsWith(">")) quoted.push(lines[i++].trim().replace(/^>\s?/, ""));
      const title = quoted[0]?.match(/^\*\*(.+)\*\*$/)?.[1] ?? quoted[0]?.match(/^\[!(\w+)\]\s*(.*)$/)?.[2];
      const body = (title !== undefined ? quoted.slice(1) : quoted).filter((l) => l.trim());
      blocks.push({ type: "quote", title: title || undefined, content: body.map(inline) });
      continue;
    }
    const item = (l: string) => l.match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/);
    if (item(line)) {
      const ordered = /\d/.test(item(line)![2]);
      const items: ListItem[] = [];
      while (i < lines.length && (item(lines[i]) || (lines[i].trim() && /^\s{2,}/.test(lines[i]) && items.length))) {
        const m = item(lines[i]);
        if (m) items.push({ level: Math.floor(m[1].replace(/\t/g, "  ").length / 2), content: inline(m[3]) });
        else items[items.length - 1].content.push({ text: ` ${lines[i].trim()}` });
        i++;
      }
      blocks.push({ type: "list", ordered, items });
      continue;
    }
    const para: string[] = [];
    while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|>|\||```|~~~|\s*([-*+]|\d+[.)])\s)/.test(lines[i].trim()) && !IMAGE.test(lines[i].trim())) para.push(lines[i++].trim());
    if (para.length) blocks.push({ type: "paragraph", content: inline(para.join(" ")) });
    else i++;
  }
  return blocks;
}

/** Splits on `---` lines outside code fences. */
function chunks(source: string) {
  const lines = source.replace(/\r/g, "").split("\n");
  const out: { line: number; lines: string[] }[] = [];
  let cur = { line: 0, lines: [] as string[] };
  let fence = false;
  lines.forEach((l, n) => {
    if (/^\s*(```|~~~)/.test(l)) fence = !fence;
    if (!fence && /^---\s*$/.test(l)) {
      out.push(cur);
      cur = { line: n + 1, lines: [] };
    } else cur.lines.push(l);
  });
  out.push(cur);
  return out;
}

const isMeta = (lines: string[]) => {
  const filled = lines.filter((l) => l.trim());
  return filled.length > 0 && filled.every((l) => KEY.test(l.trim()) && !/^https?:/.test(l.trim()));
};
const metaOf = (lines: string[]) => Object.fromEntries(lines.filter((l) => l.trim()).map((l) => { const m = l.trim().match(KEY)!; return [m[1], m[2].replace(/^["']|["']$/g, "")]; }));

export function parseDeck(source: string): Deck {
  const parts = chunks(source);
  // The file starts with `---`: the first chunk is empty, the second the headmatter.
  if (parts[0] && !parts[0].lines.some((l) => l.trim())) parts.shift();
  let head: Record<string, string> = {};
  if (parts[0] && isMeta(parts[0].lines)) head = metaOf(parts.shift()!.lines);
  const slides: Slide[] = [];
  for (let k = 0; k < parts.length; k++) {
    let meta: Record<string, string> = {};
    let part = parts[k];
    if (isMeta(part.lines) && parts[k + 1]) {
      meta = metaOf(part.lines);
      part = parts[++k];
    }
    if (!part.lines.some((l) => l.trim()) && !Object.keys(meta).length) continue;
    // Speaker notes: HTML comments.
    const text = part.lines.join("\n");
    const notes = [...text.matchAll(/<!--([\s\S]*?)-->/g)].map((m) => m[1].trim()).join("\n");
    const body = text.replace(/<!--[\s\S]*?-->/g, "").split("\n");
    const sections: string[][] = [[]];
    const kinds: string[] = ["main"];
    for (const l of body) {
      const marker = l.trim().match(/^::(right|col)::$/);
      if (marker) {
        sections.push([]);
        kinds.push(marker[1]);
      } else sections[sections.length - 1].push(l);
    }
    const main = parseBlocks(sections[0]);
    const titleAt = main.findIndex((b) => b.type === "heading" && b.depth <= 2);
    const title = titleAt >= 0 ? (main[titleAt] as Extract<Block, { type: "heading" }>).content : undefined;
    if (titleAt >= 0) main.splice(titleAt, 1);
    const first = slides.length === 0;
    const layout = ((meta.layout as Layout) in LAYOUTS ? meta.layout : first ? "cover" : "default") as Layout;
    slides.push({
      index: slides.length,
      line: part.line,
      layout,
      meta,
      title,
      blocks: main,
      right: kinds.includes("right") ? parseBlocks(sections[kinds.indexOf("right")]) : [],
      columns: sections.filter((_, n) => kinds[n] === "col").map(parseBlocks),
      notes,
    });
  }
  return { head: { title: head.title ?? (slides[0]?.title ? plain(slides[0].title) : "Presentation"), footer: head.footer, brand: head.brand ?? "DIOR" }, slides };
}

/** The markdown of a new deck: cover, agenda and one example content page. */
export const newDeck = (title: string) => `---
title: ${title}
footer: ${title}
---

# ${title}

Subtitle

${new Date().toLocaleDateString("en-GB", { month: "long", year: "numeric" })}

---
layout: agenda
---

# Agenda

1. Context
2. Analysis
3. Next steps

---
layout: section
---

# Context

---

# Example page

- A key message, in one line
- A second point
  - with a detail

> **To remember**
> One sentence the audience should keep.
`;

/** Splits a deck source into the text of each slide (with its frontmatter), and joins them back. */
export function slideSources(source: string) {
  const parts = chunks(source);
  const head: string[] = [];
  if (parts[0] && !parts[0].lines.some((l) => l.trim())) parts.shift();
  if (parts[0] && isMeta(parts[0].lines)) head.push(parts.shift()!.lines.join("\n"));
  const slides: string[] = [];
  for (let k = 0; k < parts.length; k++) {
    if (isMeta(parts[k].lines) && parts[k + 1]) {
      slides.push(`${parts[k].lines.join("\n").trim()}\n---\n${parts[k + 1].lines.join("\n").trim()}`);
      k++;
    } else if (parts[k].lines.some((l) => l.trim())) slides.push(parts[k].lines.join("\n").trim());
  }
  return { head: head[0] ?? "", slides };
}

export function joinSlides(head: string, slides: string[]) {
  const withMeta = (s: string) => (/^[A-Za-z_][\w-]*\s*:/.test(s.split("\n")[0]) && s.includes("\n---\n") ? s : `\n${s}`);
  return `---\n${head.trim()}\n---\n${slides.map((s, i) => (i === 0 ? `\n${s}` : withMeta(s))).join("\n\n---\n")}\n`;
}
