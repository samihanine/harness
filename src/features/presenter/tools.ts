/**
 * The Presenter's agent: tools on a real deck — texts, new text boxes and tables, slides from the
 * template or the deck's layouts, Power BI pictures, look, order — with the places computed by code,
 * and the raw XML for anything else. Arguments are read tolerantly (a shape as 6, "6" or "#6"; a text
 * as lines or a list). Each change is kept in this browser at once (with undo); the user saves the deck.
 */
import { z } from "zod";
import { datasetText, withContexts } from "@/agent/context";
import { tool, type Agent, type Tool } from "@/agent/loop";
import { daxTool, queryTool } from "@/agent/pbi-tools";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { describeReport } from "@/services/report-schema";
import { presenterInstructions } from "@/prompts/presenter";
import { captureVisual } from "./capture";
import type { Deck, Where } from "./pptx/deck";
import { typeOf, typesText, type SlideType } from "./pptx/template";

export type PresenterSession = {
  deck: () => Deck | undefined;
  template: () => { deck: Deck; types: SlideType[] } | undefined;
  /** Keeps the deck after a change (a new version here, the previous one for undo) and shows the slide. */
  changed: (slide?: number) => Promise<void>;
  undo: () => Promise<boolean>;
  show: (slide: number) => void;
  current: () => number;
  sources: () => { datasets: DatasetEntry[]; reports: ReportEntry[] };
};

/* --------------------------------- Arguments --------------------------------- */

const slide = z.coerce.number().int().min(1).describe("Slide number (1 = first)");
/** A shape: its id as read (#12, "12" or 12), its name, or "title" / "body" for a placeholder. */
const shape = z.preprocess((v) => (typeof v === "number" ? String(v) : v), z.string().min(1)).describe('Shape id as read ("#12"), or its name');
/** A text as lines (a list of lines is joined). */
const lines = z.preprocess((v) => (Array.isArray(v) ? v.join("\n") : v), z.string()).describe('Lines: one paragraph per line; two spaces before: one level deeper; "- " before: a bullet; **bold**');
const box = z.object({ x: z.coerce.number(), y: z.coerce.number(), w: z.coerce.number(), h: z.coerce.number() }).describe("In % of the slide (only when the user gives a position)");
const where = z
  .preprocess((v) => (typeof v === "string" ? { inside: v } : v), z.object({ inside: shape.optional(), below: shape.optional(), right_of: shape.optional(), box: box.optional() }))
  .optional()
  .describe('Where it goes: {"inside": "#12"} (a frame or a shape), {"below": "#25"}, {"right_of": "#25"}; nothing: the largest free area of the slide');
const color = z.string().regex(/^#?[0-9a-fA-F]{6}$|^none$/).describe("#rrggbb");
const rows = z.array(z.array(z.preprocess((v) => (v === null || v === undefined ? "" : String(v)), z.string()))).min(1).describe("Rows of cells; the first row is the header");
const textStyle = { size: z.coerce.number().optional().describe("Points"), bold: z.boolean().optional(), italic: z.boolean().optional(), color: color.optional(), font: z.string().optional(), align: z.enum(["l", "ctr", "r"]).optional().describe("l: left, ctr: centered, r: right") };

const hex = (c?: string) => (c && c !== "none" ? `#${c.replace("#", "")}` : c);

/** The reports' pages and visuals, for insert_visual. */
async function reportsText(reports: ReportEntry[]) {
  const parts = await Promise.all(
    reports.map(async (r) => {
      const schema = r.schema ?? (await describeReport(r).catch(() => undefined));
      if (!schema) return `Report "${r.name}" (report "${r.id}"): its pages could not be read.`;
      return `Report "${r.name}" — report "${r.id}"\n${schema.pages
        .filter((p) => !p.hidden)
        .map((p) => `  - page "${p.displayName}" — page "${p.name}"\n${schema.visuals.filter((v) => v.page === p.name && !v.hidden && !["shape", "basicShape", "image", "textbox", "actionButton"].includes(v.type)).map((v) => `      · ${v.type} "${v.title ?? ""}" — visual "${v.name}" (${Math.round(v.width)}×${Math.round(v.height)})`).join("\n")}`)
        .join("\n")}`;
    }),
  );
  return parts.length ? `Reports (insert_visual: give report, page and visual — none: the whole page):\n${parts.join("\n")}` : "";
}

export function presenterAgent(s: PresenterSession): Agent {
  const deck = () => {
    const d = s.deck();
    if (!d) throw new Error("No deck is open.");
    return d;
  };
  const chosen = s.sources();
  const models = () => s.sources().datasets.map((d) => ({ id: d.id, groupId: d.groupId, name: d.name, info: d.info }));
  const sourceTools: Tool[] = chosen.datasets.length ? [daxTool(() => undefined, models), queryTool(() => undefined, models)] : [];
  const template = !!s.template();
  /**
   * A shape after a change, as the AI would read it (its id, place and room), with what the AI cannot
   * see: a text too long for its box, a shape now over another one or out of the slide.
   */
  const after = async (n: number, id: string) => {
    const info = await deck().slide(n);
    const sh = info.shapes.find((x) => x.id === id.replace(/^#/, ""));
    if (!sh) return "";
    const length = sh.paragraphs.reduce((sum, p) => sum + p.text.length, 0);
    const warnings: string[] = [];
    if (sh.kind === "text" && sh.budget && length > sh.budget * 1.15) warnings.push(`${length} characters for a room of about ${sh.budget}: it overflows — shorten it, or lower its size (set_style)`);
    const b = sh.box;
    if (b) {
      const { w, h } = deck().size;
      if (b.x < -w * 0.01 || b.y < -h * 0.01 || b.x + b.w > w * 1.01 || b.y + b.h > h * 1.01) warnings.push("it goes out of the slide");
      const area = (o: NonNullable<typeof b>) => Math.max(0, Math.min(b.x + b.w, o.x + o.w) - Math.max(b.x, o.x)) * Math.max(0, Math.min(b.y + b.h, o.y + o.h) - Math.max(b.y, o.y));
      const over = info.shapes.filter((o) => o !== sh && !o.role && o.box && area(o.box) > Math.min(b.w * b.h, o.box.w * o.box.h) * 0.08);
      if (over.length) warnings.push(`it now covers ${over.map((o) => `#${o.id}`).join(", ")} — move or shrink it (set_box, fewer rows or lines, a smaller size)`);
    }
    return ` #${sh.id} (${deck().pct(sh.box)}${sh.budget ? `, room ≈ ${sh.budget} chars` : ""}).${warnings.length ? ` WARNING: ${warnings.join("; ")}.` : ""}`;
  };
  /** Fills shapes of a slide: key (id, name, "title" / "body") → lines; what could not be filled is said, not thrown. */
  const fill = async (n: number, texts: Record<string, string> | undefined) => {
    const missed: string[] = [];
    for (const [key, text] of Object.entries(texts ?? {})) await deck().setText(n, key, text).catch(() => missed.push(key));
    return missed.length ? ` Not filled (no such shape on the new slide): ${missed.join(", ")} — read_slide ${n}, then set_text with its ids.` : "";
  };

  return {
    instructions: presenterInstructions({ datasets: chosen.datasets.length, reports: chosen.reports.length, template }),
    context: async () => {
      const src = s.sources();
      const t = s.template();
      const d = s.deck();
      return withContexts(
        [
          t ? `Slide types of the template deck (add_slide "type"):\n${typesText(t.types, t.deck)}` : "",
          d ? `Layouts of this deck (add_slide "layout"): ${(await d.layouts()).map((l) => `"${l.name}"${l.placeholders.length ? ` (${l.placeholders.join(", ")})` : " (empty)"}`).join("; ")}` : "",
          ...(await Promise.all(src.datasets.map((x) => datasetText(x)))),
          await reportsText(src.reports),
        ],
        [...src.datasets.map((x) => x.contextIds), ...src.reports.map((r) => r.contextIds)],
      );
    },
    state: async () => {
      const d = s.deck();
      return d ? `The user sees slide ${s.current()}. Deck (${await d.count()} slides):\n${await d.outline()}` : "No deck is open.";
    },
    tools: [
      ...sourceTools,
      tool({
        name: "read_slide",
        description: "Reads a slide whole: every shape with its text (lines), place, size and room; tables as rows; frames; the notes. all: also the decorative shapes.",
        args: z.object({ slide, all: z.boolean().optional() }),
        readOnly: true,
        run: async ({ slide: n, all }) => deck().describe(await deck().slide(n), { full: true, all }),
      }),
      tool({
        name: "set_text",
        description: "Replaces the text of an EXISTING text shape (its look is kept). Not for adding content elsewhere: add_text.",
        args: z.object({ slide, shape, text: lines }),
        run: async ({ slide: n, shape: id, text }) => {
          await deck().setText(n, id, text);
          await s.changed(n);
          const sh = (await deck().slide(n)).shapes.find((x) => x.id === id.replace(/^#/, "") || x.name.toLowerCase() === id.toLowerCase());
          return `Text set on slide ${n}:${sh ? await after(n, sh.id) : ` ${id}.`}`;
        },
      }),
      tool({
        name: "add_text",
        description: "Creates a NEW text box on a slide, placed by the app (where) in the look of the slide's text (or of the shape given in like).",
        args: z.object({ slide, text: lines, where, like: shape.optional().describe("A text shape whose look is copied"), ...textStyle }),
        run: async ({ slide: n, text, where: at, like, ...style }) => {
          const id = await deck().addText(n, text, (at ?? {}) as Where, { like, style: { ...style, color: hex(style.color) } });
          await s.changed(n);
          return `Text box added on slide ${n}:${await after(n, id)}`;
        },
      }),
      tool({
        name: "replace_text",
        description: "Replaces a text everywhere in the deck (or on the given slides), e.g. a figure or a name.",
        args: z.object({ find: z.string().min(1), replace: z.string(), slides: z.array(slide).optional() }),
        run: async ({ find, replace, slides }) => {
          const n = await deck().replaceText(find, replace, slides);
          if (!n) throw new Error(`"${find}" is not in the ${slides ? "given slides" : "deck"} as written (a text split by formatting may differ: read the slide and use set_text).`);
          await s.changed();
          return `${n} occurrence(s) replaced.`;
        },
      }),
      tool({
        name: "set_table",
        description: "Replaces the content of an EXISTING table (rows added or removed as needed, its look kept).",
        args: z.object({ slide, shape, rows }),
        run: async ({ slide: n, shape: id, rows: cells }) => {
          await deck().setTable(n, id, cells);
          await s.changed(n);
          const table = (await deck().slide(n)).shapes.find((x) => x.id === id.replace(/^#/, "") || x.name.toLowerCase() === id.toLowerCase());
          return `Table on slide ${n}, ${cells.length} rows:${table ? await after(n, table.id) : ""}`;
        },
      }),
      tool({
        name: "add_table",
        description: "Creates a NEW table on a slide, placed by the app (where), in the look of the slide's own table when it has one (or of the table given in like).",
        args: z.object({ slide, rows, where, like: shape.optional() }),
        run: async ({ slide: n, rows: cells, where: at, like }) => {
          const id = await deck().addTable(n, cells, (at ?? {}) as Where, like);
          await s.changed(n);
          return `Table added on slide ${n}:${await after(n, id)}`;
        },
      }),
      tool({
        name: "add_slide",
        description: `Adds a slide at a position (number of the new slide; none: the end): ${template ? 'from a "type" of the template deck, or ' : ""}from a "layout" of this deck (none: its title-and-content layout). texts fills it: shape id, name, or "title" / "body" → lines.`,
        args: z.object({ type: z.string().optional().describe(template ? "A slide type of the template deck" : "Not available: no template deck is set"), layout: z.string().optional().describe("A layout of this deck"), at: z.coerce.number().int().optional(), texts: z.record(z.string(), lines).optional() }),
        run: async ({ type, layout, at, texts }) => {
          const count = await deck().count();
          const n = at ? Math.max(1, Math.min(at, count + 1)) : count + 1;
          const t = s.template();
          const kind = type && t ? typeOf(t.types, type) : undefined;
          let made: string;
          if (kind) made = `type "${kind.name}"${await deck().copySlide(t!.deck, kind.number, n).then((note) => (note ? ` (${note})` : ""))}`;
          else {
            // A name given as a type may be a layout of this deck (and the reverse): tried before refusing.
            const names = (await deck().layouts()).map((l) => l.name);
            const asked = layout ?? type;
            if (asked && !names.some((x) => x.toLowerCase() === asked.toLowerCase())) throw new Error(`No ${t ? `slide type or ` : ""}layout "${asked}". ${t ? `Types: ${t.types.map((x) => `"${x.name}"`).join(", ")}. ` : ""}Layouts: ${names.map((x) => `"${x}"`).join(", ")}.`);
            made = `layout "${await deck().addSlideFromLayout(asked, n)}"`;
          }
          const note = await fill(n, texts);
          await s.changed(n);
          return `Slide ${n} added (${made}).${note}\n${deck().describe(await deck().slide(n))}`;
        },
      }),
      tool({
        name: "duplicate_slide",
        description: "Copies a slide of this deck to a position (none: right after it), to rework the copy.",
        args: z.object({ slide, at: z.coerce.number().int().optional() }),
        run: async ({ slide: n, at }) => {
          const to = at ?? n + 1;
          await deck().copySlide(deck(), n, to);
          await s.changed(to);
          return `Slide ${n} copied as slide ${to} (same shape ids).`;
        },
      }),
      tool({
        name: "delete_slides",
        description: "Deletes slides by number.",
        args: z.object({ slides: z.array(slide).min(1) }),
        run: async ({ slides }) => {
          for (const n of [...new Set(slides)].sort((a, b) => b - a)) await deck().deleteSlide(n);
          await s.changed();
          return `${slides.length} slide(s) deleted.`;
        },
      }),
      tool({
        name: "move_slide",
        description: "Moves a slide to another position.",
        args: z.object({ slide, to: z.coerce.number().int().min(1) }),
        run: async ({ slide: n, to }) => {
          await deck().moveSlide(n, to);
          await s.changed(to);
          return `Slide ${n} moved to ${to}.`;
        },
      }),
      tool({
        name: "insert_visual",
        description: "Puts a picture of a Power BI visual (or of a whole page: no visual) on a slide, in its proportions: at a place (where), or in place of a shape (replace: a picture or a placeholder).",
        args: z.object({ slide, report: z.string(), page: z.string(), visual: z.string().optional(), where, replace: shape.optional().describe("A shape the picture takes the place of") }),
        run: async ({ slide: n, report, page, visual, where: at, replace }) => {
          const r = s.sources().reports.find((x) => x.id === report || x.name.toLowerCase() === report.toLowerCase());
          if (!r) throw new Error(`Report "${report}" is not among the deck's reports (Reports tab): ${s.sources().reports.map((x) => `"${x.name}"`).join(", ") || "none is chosen"}.`);
          const schema = r.schema;
          const p = schema?.pages.find((x) => x.name === page || x.displayName.toLowerCase() === page.toLowerCase());
          const v = visual ? schema?.visuals.find((x) => (x.name === visual || x.title?.toLowerCase() === visual.toLowerCase()) && (!p || x.page === p.name)) : undefined;
          if (schema && !p) throw new Error(`No page "${page}" in "${r.name}": ${schema.pages.map((x) => `"${x.displayName}"`).join(", ")}.`);
          if (visual && schema && !v) throw new Error(`No visual "${visual}" on that page: ${schema.visuals.filter((x) => x.page === p?.name).map((x) => `"${x.title ?? x.name}"`).join(", ")}.`);
          const aspect = v ? v.width / v.height : p ? p.width / p.height : 16 / 9;
          // The place first: refused before the capture when there is no room.
          if (!replace) await deck().placeFor(n, (at ?? {}) as Where, aspect);
          s.show(n);
          const picture = await captureVisual({ id: r.id, groupId: r.groupId }, p?.name ?? page, v?.name ?? visual, aspect);
          const id = await deck().putPicture(n, picture.png, picture, { ...((at ?? {}) as Where), replace }, v?.title || p?.displayName || "Power BI");
          await s.changed(n);
          return `Picture of ${v ? `the visual "${v.title ?? v.name}"` : `the page "${p?.displayName ?? page}"`} put on slide ${n}:${await after(n, id)}`;
        },
      }),
      tool({
        name: "set_style",
        description: "Changes the look of a shape: its text (size, bold, italic, color, font, align) and / or its fill and outline (#rrggbb or \"none\").",
        args: z.object({ slide, shape, ...textStyle, fill: color.optional(), line: color.optional() }),
        run: async ({ slide: n, shape: id, ...style }) => {
          if (!Object.values(style).some((v) => v !== undefined)) throw new Error("Say what to change: size, bold, italic, color, font, align, fill or line.");
          await deck().setStyle(n, id, { ...style, color: hex(style.color), fill: hex(style.fill), line: hex(style.line) });
          await s.changed(n);
          return `Look of #${id.replace(/^#/, "")} changed.`;
        },
      }),
      tool({
        name: "set_box",
        description: "Moves / resizes a shape (box in % of the slide).",
        args: z.object({ slide, shape, box }),
        run: async ({ slide: n, shape: id, box: b }) => (await deck().setBox(n, id, b), await s.changed(n), `Shape moved:${await after(n, id)}`),
      }),
      tool({
        name: "add_box",
        description: "Adds a plain rectangle under the content (a frame, a colored band): fill and outline as #rrggbb.",
        args: z.object({ slide, where, fill: color.optional(), line: color.optional() }),
        run: async ({ slide: n, where: at, fill: f, line: l }) => {
          const id = await deck().addBox(n, (at ?? {}) as Where, { fill: hex(f), line: hex(l) });
          await s.changed(n);
          return `Rectangle added:${await after(n, id)}`;
        },
      }),
      tool({
        name: "delete_shape",
        description: "Deletes a shape of a slide.",
        args: z.object({ slide, shape }),
        run: async ({ slide: n, shape: id }) => (await deck().deleteShape(n, id), await s.changed(n), "Shape deleted."),
      }),
      tool({
        name: "read_xml",
        description: 'Reads the raw XML of a slide (by number) or of any part of the deck (part: its name, e.g. ppt/slideLayouts/slideLayout2.xml; "list": every part).',
        args: z.object({ slide: slide.optional(), part: z.string().optional() }),
        readOnly: true,
        run: async ({ slide: n, part }) => (part === "list" ? deck().parts().join("\n") : deck().xml(n ?? part ?? s.current())),
      }),
      tool({
        name: "write_xml",
        description: "Replaces the raw XML of a slide or a part (the whole part, valid PresentationML; relationship ids unchanged). Only for what no other tool does.",
        args: z.object({ slide: slide.optional(), part: z.string().optional(), xml: z.string().min(20) }),
        run: async ({ slide: n, part, xml }) => {
          await deck().setXml(n ?? part ?? s.current(), xml);
          await s.changed(n);
          return "XML written.";
        },
      }),
      tool({
        name: "show_slide",
        description: "Shows a slide to the user.",
        args: z.object({ slide }),
        readOnly: true,
        run: async ({ slide: n }) => (await deck().partOf(n), s.show(n), `Slide ${n} shown.`),
      }),
      tool({
        name: "undo_last_change",
        description: "Puts the deck back as it was before the last change (only when the user asks).",
        args: z.object({}),
        run: async () => ((await s.undo()) ? "Previous version restored." : "Nothing to undo."),
      }),
    ],
  };
}
