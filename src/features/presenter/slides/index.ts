/**
 * Decks: a short JSON — the theme and the slides, each a template id and its args — checked here
 * against each template's zod schema. Content and visuals stay apart: the templates draw (screen,
 * PDF, PowerPoint), the JSON only says what to show.
 */
import { z } from "zod";
import { themeSchema, type DeckInfo, type Template } from "./kit";
import { agenda } from "./agenda";
import { beware } from "./beware";
import { cards } from "./cards";
import { cover } from "./cover";
import { features } from "./features";
import { tutorial } from "./tutorial";
import { compare } from "./compare";
import { end } from "./end";
import { hero } from "./hero";
import { image } from "./image";
import { causality } from "./causality";
import { context } from "./context";
import { conversation } from "./conversation";
import { architecture } from "./architecture";
import { responsibilities } from "./responsibilities";
import { pyramid } from "./pyramid";
import { upgradeSlide } from "./upgrade";
import { lists } from "./lists";
import { roadmap } from "./roadmap";
import { usecase } from "./usecase";
import { kpis } from "./kpis";
import { list } from "./list";
import { markdown } from "./markdown";
import { quote } from "./quote";
import { report } from "./report";
import { subtitle } from "./subtitle";
import { summary } from "./summary";
import { table } from "./table";
import { title } from "./title";
import { visual } from "./visual";

export const TEMPLATES = [title, cover, hero, summary, agenda, subtitle, list, cards, features, kpis, table, lists, context, causality, responsibilities, pyramid, architecture, roadmap, usecase, conversation, beware, quote, markdown, image, report, visual, tutorial, compare, end] as unknown as Template<z.ZodObject>[];
export const templateOf = (id: string) => TEMPLATES.find((t) => t.id === id);

export type SlideJson = { template: string } & Record<string, unknown>;
export type Deck = { title: string; theme: z.infer<typeof themeSchema>; slides: SlideJson[] };

const deckShape = z.object({ title: z.string().default("Presentation"), theme: themeSchema.default(() => themeSchema.parse({})), slides: z.array(z.preprocess(upgradeSlide, z.looseObject({ template: z.string() }))).default([]) });

const issue = (e: z.ZodError) => e.issues.map((i) => `${i.path.join(".") || "value"}: ${i.message}`).join("; ");

/** One slide checked against its template (its args with the defaults filled). */
export function checkSlide(given: SlideJson) {
  const slide = upgradeSlide(given) as SlideJson;
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

/** A slide's kind (content when its template says nothing). */
export const kindOf = (s: SlideJson) => {
  const t = templateOf(s.template);
  return t?.kindOf?.(s) ?? t?.kind ?? "content";
};

/**
 * Where each slide is: its part (the last part page before it: an agenda naming its part), its subpart (the last subsection
 * page since), and the deck's outline (summary pages, headers).
 */
export function deckInfo(deck: Deck): DeckInfo[] {
  const outline: DeckInfo["outline"] = [];
  for (const s of deck.slides) {
    if (kindOf(s) === "section") outline.push({ title: String(s.part ?? s.title ?? ""), subsections: [] });
    if (kindOf(s) !== "subsection") continue;
    if (!outline.length) outline.push({ title: "", subsections: [] });
    outline.at(-1)!.subsections.push(String(s.title ?? ""));
  }
  // Content pages grouped by the title page above them (a subpart, else a part): "(2/6)".
  const groups: number[] = [];
  let group = 0;
  for (const s of deck.slides) {
    if (["cover", "section", "subsection", "summary", "end"].includes(kindOf(s))) group++;
    groups.push(group);
  }
  const contentOf = (g: number) => deck.slides.map((s, i) => ({ s, i })).filter(({ s, i }) => groups[i] === g && kindOf(s) === "content").map(({ i }) => i);
  let section = -1;
  let subsection = -1;
  return deck.slides.map((s, index) => {
    if (kindOf(s) === "section") (section++, (subsection = -1));
    if (kindOf(s) === "subsection") subsection++;
    const part = Math.max(section, 0);
    const mates = contentOf(groups[index]);
    return {
      title: deck.title,
      index,
      count: deck.slides.length,
      section: part,
      sections: outline.map((o) => o.title),
      subsection,
      subsections: outline[part]?.subsections ?? [],
      outline,
      next: section + 1,
      group: { index: Math.max(0, mates.indexOf(index)), count: mates.length },
    };
  });
}

const KIND_LABELS: Record<NonNullable<Template<z.ZodObject>["kind"]>, string> = { cover: "cover page", section: "part title page", subsection: "subpart title page", summary: "summary page", content: "content page", end: "closing page" };

type JsonShape = { type?: string | string[]; properties?: Record<string, JsonShape>; required?: string[]; items?: JsonShape; enum?: unknown[] };

/** An args shape in short: {title, items[{text, type?}], colors?} — names and nesting, no types. */
function signature(shape: JsonShape): string {
  if (shape.items) return `[${signature(shape.items)}]`;
  if (!shape.properties) return shape.enum ? `(${shape.enum.join("|")})` : "";
  const names = Object.keys(shape.properties);
  const ordered = [...names.filter((k) => shape.required?.includes(k)), ...names.filter((k) => !shape.required?.includes(k))];
  return `{${ordered.map((k) => `${k}${shape.required?.includes(k) ? "" : "?"}${signature(shape.properties![k]).replace(/^\(.*\)$/, "")}`).join(", ")}}`;
}

/**
 * For the AI: each template in one short line — its kind, what it is for, its args (required first,
 * optional with "?"). Short on purpose: a long list made the model return empty replies.
 */
export const templatesText = () =>
  TEMPLATES.map((t) => {
    return `- ${t.id} (${KIND_LABELS[t.kind ?? "content"]}): ${t.description.length > 240 ? `${t.description.slice(0, 240)}…` : t.description} ${signature(z.toJSONSchema(t.args) as JsonShape)}`;
  }).join("\n");

export const EXAMPLE_DECK: Deck = {
  title: "Quarterly review",
  theme: themeSchema.parse({ header: "rule" }),
  slides: [
    { template: "title", eyebrow: "Sales · Q3", title: "Quarterly review", subtitle: "Results, trends and next steps", context: "Reporting · October", footer: "Sales team — October" },
    { template: "summary", title: "Summary" },
    { template: "agenda", part: "Results" },
    { template: "subtitle", title: "Key figures", description: "What the quarter **delivered**." },
    { template: "kpis", title: "Key figures", items: [{ value: "12.4 M€", label: "Revenue", note: "+8 % vs Q2" }, { value: "38 %", label: "Margin", note: "+2 pts" }, { value: "1 240", label: "Orders" }] },
    { template: "subtitle", title: "By region", description: "Where the growth came from." },
    { template: "table", title: "Revenue by region", columns: ["Region", "Q2", "Q3", "Change"], rows: [["North", "3.1 M€", "3.6 M€", "+16 %"], ["West", "2.8 M€", "3.0 M€", "+7 %"], ["South", "2.6 M€", "2.5 M€", "−4 %"], ["East", "1.9 M€", "2.0 M€", "+5 %"]], cellColors: [{ row: 3, column: 4, color: "#FEE2E2" }], rowColors: [{ row: 1, color: "#ECFDF5" }] },
    { template: "agenda", part: "Next steps" },
    { template: "cards", title: "Priorities", items: [{ title: "Stock", text: "Reduce lead times." }, { title: "Pricing", text: "Review the discount grid." }, { title: "Data", text: "One dashboard for all." }] },
    { template: "title", eyebrow: "Thank you", title: "Questions?", footer: "sales@company.com" },
  ],
};

