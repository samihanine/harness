/**
 * Decks: a short JSON — the theme and the slides, each a template id and its args — checked here
 * against each template's zod schema. Content and visuals stay apart: the templates draw (screen,
 * PDF, PowerPoint), the JSON only says what to show.
 */
import { z } from "zod";
import { themeSchema, type DeckInfo, type Template } from "./kit";
import { cards } from "./cards";
import { compare } from "./compare";
import { end } from "./end";
import { hero } from "./hero";
import { kpis } from "./kpis";
import { list } from "./list";
import { markdown } from "./markdown";
import { quote } from "./quote";
import { section } from "./section";
import { visual } from "./visual";

export const TEMPLATES = [hero, section, list, cards, kpis, quote, markdown, visual, compare, end] as unknown as Template<z.ZodObject>[];
export const templateOf = (id: string) => TEMPLATES.find((t) => t.id === id);

export type SlideJson = { template: string } & Record<string, unknown>;
export type Deck = { title: string; theme: z.infer<typeof themeSchema>; slides: SlideJson[] };

const deckShape = z.object({ title: z.string().default("Presentation"), theme: themeSchema.default(themeSchema.parse({})), slides: z.array(z.looseObject({ template: z.string() })).default([]) });

const issue = (e: z.ZodError) => e.issues.map((i) => `${i.path.join(".") || "value"}: ${i.message}`).join("; ");

/** One slide checked against its template (its args with the defaults filled). */
export function checkSlide(slide: SlideJson) {
  const t = templateOf(slide.template);
  if (!t) throw new Error(`Unknown template "${slide.template}". Templates: ${TEMPLATES.map((x) => x.id).join(", ")}.`);
  const { template: _, ...args } = slide;
  const parsed = t.args.safeParse(args);
  if (!parsed.success) throw new Error(`Slide "${slide.template}": ${issue(parsed.error)}.`);
  return { template: t, args: parsed.data as Record<string, unknown> };
}

/** The whole deck checked: the problems of each slide (by number), or the deck. */
export function parseDeck(value: unknown): { deck?: Deck; errors: string[] } {
  const shape = deckShape.safeParse(value);
  if (!shape.success) return { errors: [issue(shape.error)] };
  const errors = shape.data.slides.flatMap((s, i) => {
    try {
      checkSlide(s as SlideJson);
      return [];
    } catch (e) {
      return [`Slide ${i + 1}: ${e instanceof Error ? e.message : e}`];
    }
  });
  return { deck: shape.data as Deck, errors };
}

/** Where each slide is: its part (the last "section" slide before it) and the deck's outline. */
export function deckInfo(deck: Deck): DeckInfo[] {
  const sections = deck.slides.filter((s) => templateOf(s.template)?.section).map((s) => String(s.title ?? ""));
  let current = -1;
  return deck.slides.map((s, index) => {
    if (templateOf(s.template)?.section) current++;
    return { title: deck.title, index, count: deck.slides.length, section: Math.max(current, 0), sections };
  });
}

/** For the AI: each template with its args (from the zod schemas). */
export const templatesText = () =>
  TEMPLATES.map((t) => {
    const shape = z.toJSONSchema(t.args) as { properties?: Record<string, { type?: string; description?: string }>; required?: string[] };
    const args = Object.entries(shape.properties ?? {}).map(([k, v]) => `${k}${shape.required?.includes(k) ? "" : "?"}: ${v.type ?? "object"}${v.description ? ` (${v.description})` : ""}`);
    return `- ${t.id}: ${t.description} Args: ${args.join(", ")}`;
  }).join("\n");

export const EXAMPLE_DECK: Deck = {
  title: "Quarterly review",
  theme: themeSchema.parse({}),
  slides: [
    { template: "hero", eyebrow: "Sales · Q3", title: "Quarterly **review**", subtitle: "Results, trends and ==next steps==", footer: "Team · October" },
    { template: "section", title: "Results" },
    { template: "kpis", title: "Key figures", items: [{ value: "12.4 M€", label: "Revenue", note: "+8 % vs Q2" }, { value: "38 %", label: "Margin", note: "+2 pts" }, { value: "1 240", label: "Orders" }] },
    { template: "list", title: "What changed", lead: "Growth came from **two regions**.", items: [{ title: "North", text: "New distributor since July." }, { title: "West", text: "Price update accepted." }] },
    { template: "section", title: "Next steps" },
    { template: "cards", title: "Priorities", items: [{ title: "Stock", text: "Reduce lead times." }, { title: "Pricing", text: "Review the discount grid." }, { title: "Data", text: "One dashboard for all." }] },
    { template: "end", title: "Questions?", text: "sales@company.com" },
  ],
};
