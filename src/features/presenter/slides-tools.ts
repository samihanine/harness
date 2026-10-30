/**
 * The Slides screen's agent: it writes the deck's JSON (theme + slides, each a template and its args),
 * checked against the templates before anything is saved.
 */
import { z } from "zod";
import { datasetText, withContexts } from "@/agent/context";
import { tool, type Agent, type Tool } from "@/agent/loop";
import { DAX_RULES, daxTool, queryTool } from "@/agent/pbi-tools";
import { describeReport } from "@/services/report-schema";
import { patch, type DatasetEntry, type DeckEntry, type ReportEntry } from "@/lib/store";
import { checkSlide, templatesText, type Deck, type SlideJson } from "@/features/presenter/slides";
import { themeSchema } from "@/features/presenter/slides/kit";
import { slidesInstructions } from "@/prompts/presenter";

const titleOf = (s: SlideJson) => String(s.title ?? s.text ?? "").replace(/\*\*|==/g, "").slice(0, 60);

/** Decks before the AI's last change (one step of undo per deck). */
const previous = new Map<string, Deck>();

/** What the deck is written from (chosen in the Datasets / Reports tabs; none by default). */
export type DeckSources = { datasets: DatasetEntry[]; reports: ReportEntry[] };

/** The sources' tools: DAX on the datasets (nothing when none is chosen). */
function sourceTools(sources: () => DeckSources): Tool[] {
  if (!sources().datasets.length) return [];
  return [daxTool(() => undefined, () => sources().datasets.map((d) => ({ id: d.id, groupId: d.groupId, name: d.name }))), queryTool(() => undefined, () => sources().datasets.map((d) => ({ id: d.id, groupId: d.groupId, name: d.name, info: d.info })))];
}

/**
 * The sources for the AI: models (structure and examples), reports (their kept schema: page ids,
 * visual ids, titles, fields and zones, for the report / visual / compare templates).
 */
async function sourcesText(s: DeckSources) {
  const reports = await Promise.all(
    s.reports.map(async (r) => {
      // The schema kept with the report (Reports screen), made now when missing (embed if needed).
      const schema = r.schema ?? (await describeReport(r).catch(() => undefined));
      if (!schema) return `Report "${r.name}" (report: "${r.id}"): its pages could not be read.`;
      const pages = schema.pages.map((p) => {
        const visuals = schema.visuals.filter((v) => v.page === p.name && !v.hidden && !["shape", "basicShape", "image", "textbox", "actionButton"].includes(v.type));
        const scale = 1280 / (p.width || 1280);
        return `  - page "${p.displayName}" — page: "${p.name}" (${p.width}×${p.height})\n${visuals.map((v) => `      · ${v.type} "${v.title ?? ""}" visual "${v.name}" — crop {"x": ${Math.round(v.x * scale)}, "y": ${Math.round(v.y * scale)}, "w": ${Math.round(v.width * scale)}, "h": ${Math.round(v.height * scale)}}${Object.keys(v.fields).length ? ` — ${Object.entries(v.fields).map(([role, f]) => `${role}: ${f.join(", ")}`).join("; ")}` : ""}`).join("\n")}`;
      });
      return `Report "${r.name}" — report: "${r.id}"\n${pages.join("\n")}`;
    }),
  );
  return [
    ...(await Promise.all(s.datasets.map((d) => datasetText(d)))),
    reports.length ? `Reports for the report / visual / compare templates. Give "report" (its id), "page" (a page's name) and, for one visual, its "crop" (already on a 1280×720 page):\n${reports.join("\n")}` : "",
    "Images (image args): a public https link (e.g. Unsplash) or an inline \"data:image/png;base64,…\".",
  ].filter(Boolean);
}

export function slidesAgent(get: () => DeckEntry, show: (index: number) => void, current: () => number, sources: () => DeckSources = () => ({ datasets: [], reports: [] })): Agent {
  const chosen = sources();
  const save = async (deck: Deck) => {
    previous.set(get().id, get().deck);
    await patch("decks", get().id, { deck, updatedAt: Date.now() });
  };
  const slideAt = (n: number) => {
    const s = get().deck.slides[n - 1];
    if (!s) throw new Error(`No slide ${n}: the deck has ${get().deck.slides.length} slides.`);
    return s;
  };
  const slideArg = z.looseObject({ template: z.string() }).describe("{template, ...its args}");
  return {
    instructions: slidesInstructions(chosen),
    context: async () => {
      const s = sources();
      const parts = await sourcesText(s);
      return withContexts(parts.length ? parts : ["No data source chosen for this deck (Datasets / Reports tabs): write from the user's request only."], [...s.datasets.map((d) => d.contextIds), ...s.reports.map((r) => r.contextIds)]);
    },
    state: async () => {
      const { deck } = get();
      return `Deck "${deck.title}" — theme ${JSON.stringify(deck.theme)}\nSlides (${deck.slides.length}), the user sees slide ${current() + 1}:\n${deck.slides.map((s, i) => `${i + 1}. ${s.template} — ${titleOf(s)}`).join("\n") || "none"}`;
    },
    tools: [
      ...sourceTools(sources),
      tool({
        name: "read_slides",
        description: "Reads the JSON of slides (by number; none: all).",
        args: z.object({ numbers: z.array(z.number()).optional() }),
        readOnly: true,
        run: async ({ numbers }) => JSON.stringify((numbers ?? get().deck.slides.map((_, i) => i + 1)).map((n) => ({ number: n, ...slideAt(n) }))),
      }),
      tool({
        name: "add_slides",
        description: "Adds slides (each {template, ...args}) at a position (number of the first new slide; default: the end).",
        args: z.object({ slides: z.array(slideArg).min(1), at: z.number().optional() }),
        run: async ({ slides, at }) => {
          slides.forEach((s) => checkSlide(s as SlideJson));
          const list = [...get().deck.slides];
          const i = at ? Math.max(0, Math.min(at - 1, list.length)) : list.length;
          list.splice(i, 0, ...(slides as SlideJson[]));
          await save({ ...get().deck, slides: list });
          show(i);
          return `${slides.length} slide(s) added from number ${i + 1}.`;
        },
      }),
      tool({
        name: "update_slide",
        description: "Changes a slide: the given args replace those of the slide (template: another template; then give all its args).",
        args: z.object({ number: z.number(), args: z.record(z.string(), z.unknown()).default({}), template: z.string().optional() }),
        run: async ({ number, args, template }) => {
          const old = slideAt(number);
          const next = { ...(template && template !== old.template ? {} : old), ...args, template: template ?? old.template } as SlideJson;
          checkSlide(next);
          await save({ ...get().deck, slides: get().deck.slides.map((s, i) => (i === number - 1 ? next : s)) });
          show(number - 1);
          return `Slide ${number} updated.`;
        },
      }),
      tool({
        name: "delete_slides",
        description: "Deletes slides by number.",
        args: z.object({ numbers: z.array(z.number()).min(1) }),
        run: async ({ numbers }) => {
          numbers.forEach(slideAt);
          await save({ ...get().deck, slides: get().deck.slides.filter((_, i) => !numbers.includes(i + 1)) });
          return `${numbers.length} slide(s) deleted.`;
        },
      }),
      tool({
        name: "move_slide",
        description: "Moves a slide to another number.",
        args: z.object({ number: z.number(), to: z.number() }),
        run: async ({ number, to }) => {
          const s = slideAt(number);
          const list = get().deck.slides.filter((x) => x !== s);
          list.splice(Math.max(0, Math.min(to - 1, list.length)), 0, s);
          await save({ ...get().deck, slides: list });
          show(to - 1);
          return `Slide ${number} moved to ${to}.`;
        },
      }),
      tool({
        name: "set_deck",
        description: "Changes the deck's title and/or theme (primary, background, text: hex colors; fontTitle, fontText: Google font names).",
        args: z.object({ title: z.string().optional(), theme: themeSchema.partial().optional() }),
        run: async ({ title, theme }) => {
          const deck = get().deck;
          await save({ ...deck, title: title ?? deck.title, theme: { ...deck.theme, ...theme } });
          return "Deck updated.";
        },
      }),
      tool({
        name: "show_slide",
        description: "Shows a slide to the user.",
        args: z.object({ number: z.number() }),
        readOnly: true,
        run: async ({ number }) => (slideAt(number), show(number - 1), `Slide ${number} shown.`),
      }),
      tool({
        name: "undo_last_change",
        description: "Puts the deck back as it was before your last change (only when the user asks).",
        args: z.object({}),
        run: async () => {
          const before = previous.get(get().id);
          if (!before) throw new Error("Nothing to undo.");
          previous.delete(get().id);
          await patch("decks", get().id, { deck: before, updatedAt: Date.now() });
          return "Previous version restored.";
        },
      }),
    ],
  };
}

