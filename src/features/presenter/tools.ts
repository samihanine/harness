/**
 * The Presenter's agent: small tools on a real deck (texts, slides from the template, Power BI
 * pictures, order) and the raw XML for anything else. Each change is kept in this browser at once
 * (with undo); the user saves the deck to its file.
 */
import { z } from "zod";
import { datasetText, withContexts } from "@/agent/context";
import { tool, type Agent, type Tool } from "@/agent/loop";
import { daxTool, queryTool } from "@/agent/pbi-tools";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { describeReport } from "@/services/report-schema";
import { presenterInstructions } from "@/prompts/presenter";
import { captureVisual } from "./capture";
import type { Deck } from "./pptx/deck";
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

const box = z.object({ x: z.number(), y: z.number(), w: z.number(), h: z.number() }).describe("In % of the slide");

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
  return {
    instructions: presenterInstructions({ datasets: chosen.datasets.length, reports: chosen.reports.length, template: !!s.template() }),
    context: async () => {
      const src = s.sources();
      const t = s.template();
      return withContexts(
        [
          t ? `Slide types of the template deck (add_slide):\n${typesText(t.types, t.deck)}` : "No template deck (Settings › Slides).",
          ...(await Promise.all(src.datasets.map((d) => datasetText(d)))),
          await reportsText(src.reports),
        ],
        [...src.datasets.map((d) => d.contextIds), ...src.reports.map((r) => r.contextIds)],
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
        description: "Reads a slide whole: every shape with its full text, place, size and room, and the notes.",
        args: z.object({ slide: z.number() }),
        readOnly: true,
        run: async ({ slide }) => {
          const info = await deck().slide(slide);
          return [`Slide ${slide} — layout "${info.layout}"`, ...info.shapes.map((sh) => `#${sh.id} ${sh.kind}${sh.placeholder ? `/${sh.placeholder}` : ""} "${sh.name}" (${deck().pct(sh.box)}, ${sh.size}pt${sh.budget ? `, room ≈ ${sh.budget} chars` : ""})${sh.paragraphs.length ? `:\n${sh.paragraphs.map((p) => `    ${"  ".repeat(p.level)}${p.text}`).join("\n")}` : ""}`), info.notes ? `Notes: ${info.notes}` : ""].filter(Boolean).join("\n");
        },
      }),
      tool({
        name: "set_text",
        description: "Replaces the text of a shape (one paragraph per line; '- ' or two spaces before a line: one level deeper).",
        args: z.object({ slide: z.number(), shape: z.string().describe("Its id (#12 or 12) or name"), text: z.string() }),
        run: async ({ slide, shape, text }) => {
          await deck().setText(slide, shape, text);
          await s.changed(slide);
          const sh = (await deck().slide(slide)).shapes.find((x) => x.id === shape.replace(/^#/, "") || x.name === shape);
          const over = sh?.budget && text.length > sh.budget * 1.1 ? ` Warning: ${text.length} characters for a room of about ${sh.budget}: shorten it.` : "";
          return `Text of #${shape.replace(/^#/, "")} on slide ${slide} set.${over}`;
        },
      }),
      tool({
        name: "replace_text",
        description: "Replaces a text everywhere in the deck (or on the given slides), e.g. a figure or a name.",
        args: z.object({ find: z.string().min(1), replace: z.string(), slides: z.array(z.number()).optional() }),
        run: async ({ find, replace, slides }) => {
          const n = await deck().replaceText(find, replace, slides);
          if (!n) throw new Error(`"${find}" is not in the ${slides ? "given slides" : "deck"} (as written).`);
          await s.changed();
          return `${n} occurrence(s) replaced.`;
        },
      }),
      tool({
        name: "add_slide",
        description: "Adds a slide of a type of the template deck at a position (number of the new slide; none: the end), its slots filled (shape id of the TYPE → text).",
        args: z.object({ type: z.string(), at: z.number().optional(), texts: z.record(z.string(), z.string()).optional() }),
        run: async ({ type, at, texts }) => {
          const t = s.template();
          if (!t) throw new Error("No template deck is set (Settings › Slides): use duplicate_slide instead.");
          const kind = typeOf(t.types, type);
          if (!kind) throw new Error(`No slide type "${type}": ${t.types.map((x) => `"${x.name}"`).join(", ")}.`);
          const count = await deck().count();
          const number = at ? Math.max(1, Math.min(at, count + 1)) : count + 1;
          const note = await deck().copySlide(t.deck, kind.number, number);
          // The copy keeps the template's shape ids.
          for (const [shape, text] of Object.entries(texts ?? {})) await deck().setText(number, shape, text);
          await s.changed(number);
          return `Slide ${number} added ("${kind.name}").${note ? ` ${note}` : ""} Read it to check the texts that are left from the template.`;
        },
      }),
      tool({
        name: "duplicate_slide",
        description: "Copies a slide of this deck to a position (none: right after it).",
        args: z.object({ slide: z.number(), at: z.number().optional() }),
        run: async ({ slide, at }) => {
          const number = at ?? slide + 1;
          await deck().copySlide(deck(), slide, number);
          await s.changed(number);
          return `Slide ${slide} copied as slide ${number}.`;
        },
      }),
      tool({
        name: "delete_slides",
        description: "Deletes slides by number.",
        args: z.object({ slides: z.array(z.number()).min(1) }),
        run: async ({ slides }) => {
          for (const n of [...slides].sort((a, b) => b - a)) await deck().deleteSlide(n);
          await s.changed();
          return `${slides.length} slide(s) deleted.`;
        },
      }),
      tool({
        name: "move_slide",
        description: "Moves a slide to another position.",
        args: z.object({ slide: z.number(), to: z.number() }),
        run: async ({ slide, to }) => {
          await deck().moveSlide(slide, to);
          await s.changed(to);
          return `Slide ${slide} moved to ${to}.`;
        },
      }),
      tool({
        name: "insert_visual",
        description: "Puts a picture of a Power BI visual (or of a whole page: no visual) on a slide: in place of a shape, or at a box (% of the slide); none: a large free area.",
        args: z.object({ slide: z.number(), report: z.string(), page: z.string(), visual: z.string().optional(), shape: z.string().optional(), box: box.optional() }),
        run: async ({ slide, report, page, visual, shape, box: at }) => {
          const r = s.sources().reports.find((x) => x.id === report || x.name.toLowerCase() === report.toLowerCase());
          if (!r) throw new Error(`Report "${report}" is not among the deck's reports (Reports tab).`);
          const schema = r.schema;
          const p = schema?.pages.find((x) => x.name === page || x.displayName === page);
          const v = visual ? schema?.visuals.find((x) => (x.name === visual || x.title === visual) && (!p || x.page === p.name)) : undefined;
          const aspect = v ? v.width / v.height : p ? p.width / p.height : 16 / 9;
          s.show(slide);
          const picture = await captureVisual({ id: r.id, groupId: r.groupId }, p?.name ?? page, v?.name ?? visual, aspect);
          await deck().putPicture(slide, picture.png, picture, { shape, box: at }, v?.title || p?.displayName || "Power BI");
          await s.changed(slide);
          return `Picture of ${v ? `the visual "${v.title ?? v.name}"` : `the page "${p?.displayName ?? page}"`} put on slide ${slide}.`;
        },
      }),
      tool({
        name: "set_box",
        description: "Moves / resizes a shape (box in % of the slide).",
        args: z.object({ slide: z.number(), shape: z.string(), box }),
        run: async ({ slide, shape, box: b }) => (await deck().setBox(slide, shape, b), await s.changed(slide), "Shape moved."),
      }),
      tool({
        name: "delete_shape",
        description: "Deletes a shape of a slide.",
        args: z.object({ slide: z.number(), shape: z.string() }),
        run: async ({ slide, shape }) => (await deck().deleteShape(slide, shape), await s.changed(slide), "Shape deleted."),
      }),
      tool({
        name: "read_xml",
        description: "Reads the raw XML of a slide (by number) or of any part of the deck (name, e.g. ppt/slideLayouts/slideLayout2.xml, ppt/theme/theme1.xml); part: \"list\" lists them.",
        args: z.object({ slide: z.number().optional(), part: z.string().optional() }),
        readOnly: true,
        run: async ({ slide, part }) => (part === "list" ? deck().parts().join("\n") : deck().xml(slide ?? part ?? 1)),
      }),
      tool({
        name: "write_xml",
        description: "Replaces the raw XML of a slide or a part (the whole part, valid PresentationML; relationships ids unchanged).",
        args: z.object({ slide: z.number().optional(), part: z.string().optional(), xml: z.string().min(20) }),
        run: async ({ slide, part, xml }) => {
          await deck().setXml(slide ?? part ?? 1, xml);
          await s.changed(slide);
          return "XML written.";
        },
      }),
      tool({
        name: "show_slide",
        description: "Shows a slide to the user.",
        args: z.object({ slide: z.number() }),
        readOnly: true,
        run: async ({ slide }) => (s.show(slide), `Slide ${slide} shown.`),
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
