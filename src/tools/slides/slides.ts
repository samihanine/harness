/**
 * The Presenter's tools: the deck's content (markdown, slide by slide), the visual templates (read, created,
 * changed — the user's own, favorites first), pictures of slides and templates as the app draws them (a second,
 * no Microsoft call), and an attached PowerPoint read to recreate its look as templates.
 */
import { z } from "zod";
import { joinBlocks, parseSlides, slideBlocks, slideContent, templateOf, templateSlots, type Template } from "@shared/slides";
import { exampleFor } from "@shared/slide-templates";
import { readStoredFile, saveFile } from "@server/files";
import { Pptx } from "@server/pptx/ooxml";
import { slidePictures as pptxPictures } from "@server/pptx/render";
import { slidePictures } from "@server/slides/render";
import { base64, tool, type ToolContext } from "../tool";

type DeckDoc = { title?: string; markdown: string };

const deckOf = (ctx: ToolContext) => {
  const d = ctx.scope.doc as DeckDoc | undefined;
  if (!d) throw new Error("No presentation is open.");
  return d;
};
const setMarkdown = (ctx: ToolContext, markdown: string) => {
  const d = { ...deckOf(ctx), markdown };
  ctx.scope.doc = d;
  ctx.effect("deck", markdown);
};
export const templatesOf = (ctx: ToolContext) => (ctx.scope.templates ?? []) as Template[];
const setTemplates = (ctx: ToolContext, list: Template[]) => {
  ctx.scope.templates = list;
  ctx.effect("templates", list);
};

/** Slides whose template does not exist, or with an empty slot the template needs. */
function problems(markdown: string, templates: Template[]) {
  return parseSlides(markdown).flatMap((c, i) => {
    const t = templates.find((x) => x.id === c.template || x.name.toLowerCase() === c.template.toLowerCase());
    if (!t) return [`slide ${i + 1}: no template "${c.template}" (drawn with "${templateOf(templates, c.template)?.id}")`];
    const needs = JSON.stringify(t.elements);
    const out: string[] = [];
    if (needs.includes('"slot":"powerbi"') && !c.powerbi.length) out.push(`slide ${i + 1}: template ${t.id} shows a Power BI view: give its link`);
    if (needs.includes('"slot":"image"') && !c.images.length) out.push(`slide ${i + 1}: template ${t.id} shows a picture: give ![](https://…)`);
    return out;
  });
}

const pictures = async (ctx: ToolContext, markdown: string, slides?: number[]) => {
  ctx.status("Drawing the slides…");
  const r = await slidePictures(markdown, templatesOf(ctx), slides?.length ? slides : undefined);
  return r.pictures.slice(0, 8).map((png) => ({ mediaType: "image/png", data: base64(png) }));
};

export const readDeck = tool({
  name: "read_deck",
  description: "The presentation's markdown slide by slide (numbered), with each slide's template.",
  args: z.object({}),
  async run(_, ctx) {
    const blocks = slideBlocks(deckOf(ctx).markdown);
    if (!blocks.length) return "The presentation is empty.";
    return blocks.map((b, i) => `## Slide ${i + 1} (template ${slideContent(b).template})\n${b}`).join("\n\n");
  },
});

export const writeDeck = tool({
  name: "write_deck",
  description: "Replace the whole content: slides separated by --- lines, each starting with `template: <id>` (and other `key: value` fields), then `# Title`, `## Subtitle`, text and bullets, `![](https://…)` pictures, Power BI links, list items `**Value** label` for repeated items.",
  args: z.object({ markdown: z.string() }),
  changes: true,
  async run(a, ctx) {
    const md = joinBlocks(slideBlocks(a.markdown));
    setMarkdown(ctx, md);
    const p = problems(md, templatesOf(ctx));
    return `Content written: ${slideBlocks(md).length} slides.${p.length ? ` To fix: ${p.join("; ")}.` : ""} Check them with look_at_slides.`;
  },
});

export const editSlides = tool({
  name: "edit_slides",
  description: "Change slides one by one: replace slide n (its whole markdown block, fields included), insert a new one after slide n (0 = first), or delete slide n. The edits apply in order (numbers shift after an insert or a delete).",
  args: z.object({ edits: z.array(z.object({ action: z.enum(["replace", "insert", "delete"]), slide: z.number().int().min(0), markdown: z.string().optional() })).min(1) }),
  changes: true,
  async run(a, ctx) {
    const blocks = slideBlocks(deckOf(ctx).markdown);
    for (const e of a.edits) {
      if (e.action === "delete") {
        if (!blocks[e.slide - 1]) throw new Error(`No slide ${e.slide}.`);
        blocks.splice(e.slide - 1, 1);
      } else {
        const md = slideBlocks(e.markdown ?? "").join("\n");
        if (!md) throw new Error("Give the slide's markdown.");
        if (e.action === "replace") {
          if (!blocks[e.slide - 1]) throw new Error(`No slide ${e.slide}.`);
          blocks[e.slide - 1] = md;
        } else blocks.splice(e.slide, 0, md);
      }
    }
    const md = joinBlocks(blocks);
    setMarkdown(ctx, md);
    const p = problems(md, templatesOf(ctx));
    return `Done: ${blocks.length} slides.${p.length ? ` To fix: ${p.join("; ")}.` : ""}`;
  },
});

export const lookAtSlides = tool({
  name: "look_at_slides",
  description: "Pictures of slides as the user sees them (up to 8; all when none given): check the look after a change.",
  args: z.object({ slides: z.array(z.number().int().min(1)).max(8).optional() }),
  async run(a, ctx) {
    const images = await pictures(ctx, deckOf(ctx).markdown, a.slides);
    return { text: `${images.length} slide(s).`, images };
  },
});

export const listTemplates = tool({
  name: "list_templates",
  description: "The visual templates (the user's favorites first): id, name, what each is for, the slots its content fills.",
  args: z.object({}),
  async run(_, ctx) {
    const list = [...templatesOf(ctx)].sort((a, b) => Number(!!b.favorite) - Number(!!a.favorite));
    return list.map((t) => `- ${t.favorite ? "★ " : ""}${t.id} "${t.name}": ${t.description ?? ""} Slots: ${templateSlots(t).join(", ") || "none"}.`).join("\n");
  },
});

export const readTemplate = tool({
  name: "read_template",
  description: "A template's definition (JSON): its elements in inches on the 13.333 × 7.5 in slide.",
  args: z.object({ id: z.string() }),
  async run(a, ctx) {
    const t = templatesOf(ctx).find((x) => x.id === a.id || x.name.toLowerCase() === a.id.toLowerCase());
    if (!t) throw new Error(`No template "${a.id}" (list_templates).`);
    return JSON.stringify(t, null, 1);
  },
});

const style = z.object({ font: z.string().optional(), size: z.number().optional(), color: z.string().optional(), bold: z.boolean().optional(), italic: z.boolean().optional(), align: z.enum(["left", "center", "right"]).optional(), valign: z.enum(["top", "middle", "bottom"]).optional(), upper: z.boolean().optional(), spacing: z.number().optional(), lineHeight: z.number().optional() });
const box = { x: z.number(), y: z.number(), w: z.number(), h: z.number() };
const leaf = z.discriminatedUnion("type", [
  z.object({ type: z.literal("text"), ...box, slot: z.string().optional(), text: z.string().optional(), style: style.optional(), fill: z.string().optional() }),
  z.object({ type: z.literal("shape"), ...box, shape: z.enum(["rect", "roundRect", "ellipse", "line"]).optional(), fill: z.string().optional(), line: z.object({ color: z.string().optional(), width: z.number().optional() }).optional(), radius: z.number().optional() }),
  z.object({ type: z.literal("image"), ...box, slot: z.string().optional(), src: z.string().optional(), fit: z.enum(["cover", "contain"]).optional() }),
  z.object({ type: z.literal("powerbi"), ...box, slot: z.string().optional() }),
]);
const element = z.union([leaf, z.object({ type: z.literal("items"), ...box, slot: z.string().optional(), direction: z.enum(["row", "column", "grid"]).optional(), columns: z.number().optional(), gap: z.number().optional(), max: z.number().optional(), item: z.array(leaf) })]);

export const saveTemplate = tool({
  name: "save_template",
  description:
    'Create or change a visual template (same id: replaced). Elements in inches on the 13.333 × 7.5 in slide, drawn in order: text (slot: title, subtitle, body, footer, a field name… or fixed text; style: font, size in pt, color, bold, align, valign, upper), shape (rect, roundRect, ellipse, line; fill, line), image (slot "image", "image2"… or src: an https link, asset:cover.jpg / agenda.jpg / section.jpg / end.jpg / logo.svg, file:<id>), powerbi (slot "powerbi", "powerbi2"), items (repeated list items, direction row / column / grid; its `item` elements are placed in a 1 × 1 cell: x, y, w, h as fractions; their slots: value, label, text, index). Returns a picture of the template with example content.',
  args: z.object({ id: z.string().regex(/^[\w-]+$/), name: z.string(), description: z.string(), background: z.string().optional(), favorite: z.boolean().optional(), elements: z.array(element).min(1) }),
  changes: true,
  async run(a, ctx) {
    const list = templatesOf(ctx);
    const old = list.find((t) => t.id === a.id);
    const t: Template = { ...a, favorite: a.favorite ?? old?.favorite, elements: a.elements as Template["elements"] };
    setTemplates(ctx, [...list.filter((x) => x.id !== a.id), t]);
    const images = await pictures(ctx, exampleFor(t));
    return { text: `Template ${a.id} ${old ? "changed" : "created"} (example content below). Slots: ${templateSlots(t).join(", ")}.`, images };
  },
});

export const deleteTemplate = tool({
  name: "delete_template",
  description: "Remove a template the user no longer wants (slides using it are drawn with the Text template).",
  args: z.object({ id: z.string() }),
  changes: true,
  async run(a, ctx) {
    setTemplates(ctx, templatesOf(ctx).filter((t) => t.id !== a.id));
    return `Template ${a.id} removed.`;
  },
});

export const readPowerPoint = tool({
  name: "read_powerpoint",
  description: "An attached PowerPoint (file id): each slide's shapes (kind, place and size in inches, texts with their size, pictures) and pictures of the slides as PowerPoint draws them. Its pictures are given as file:<id> to reuse in templates. To recreate its look as templates, then its content as markdown.",
  args: z.object({ file: z.string(), slides: z.array(z.number().int().min(1)).max(8).optional() }),
  async run(a, ctx) {
    const f = await readStoredFile(a.file.replace(/^file:/, ""));
    const pptx = await Pptx.load(f.data);
    const all = await pptx.structure();
    const wanted = a.slides?.length ? a.slides : all.slice(0, 6).map((s) => s.number);
    ctx.status("Reading the PowerPoint…");
    // Its pictures as files of the app: templates can use them (logos, photos).
    const media = new Map<string, string>();
    for (const s of all.filter((x) => wanted.includes(x.number)))
      for (const sh of s.shapes)
        if (sh.image && !media.has(sh.image.file)) {
          const bytes = await pptx.media(sh.image.file).catch(() => undefined);
          if (bytes) media.set(sh.image.file, (await saveFile(Buffer.from(bytes), sh.image.file, bytes[0] === 0x89 ? "image/png" : "image/jpeg")).id);
        }
    const text = all
      .filter((s) => wanted.includes(s.number))
      .map((s) => [`## Slide ${s.number}${s.layout ? ` (layout "${s.layout}")` : ""}`, ...s.shapes.map((x) => `- ${x.kind}${x.placeholder ? ` [${x.placeholder}]` : ""} at ${x.x},${x.y} size ${x.w}×${x.h}${x.text ? `: ${JSON.stringify(x.text.slice(0, 300))}` : ""}${x.image ? ` picture file:${media.get(x.image.file)}` : ""}${x.rows ? ` table ${JSON.stringify(x.rows).slice(0, 300)}` : ""}`)].join("\n"))
      .join("\n\n");
    const pics = await pptxPictures(f.data, wanted, 1000).catch(() => [] as Buffer[]);
    return { text: `${all.length} slides.\n${text}`, images: pics.map((p) => ({ mediaType: "image/png", data: base64(p) })) };
  },
});

export const SLIDE_TOOLS = [readDeck, writeDeck, editSlides, lookAtSlides, listTemplates, readTemplate, saveTemplate, deleteTemplate, readPowerPoint];
