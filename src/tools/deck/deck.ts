/** The deck written in markdown (create mode): read, written whole or slide by slide. */
import { z } from "zod";
import { joinSlides, LAYOUTS, parseDeck, plain, slideSources } from "@shared/deck";
import { docOf, setDoc } from "../helpers";
import { tool, type ToolContext } from "../tool";

export const deckOf = (ctx: ToolContext) => {
  const d = docOf<string>(ctx);
  if (typeof d !== "string") throw new Error("No deck is open: ask the user to create one.");
  return d;
};

/** Problems the export would show: unknown layouts, images without source, too much text. */
export function deckProblems(md: string) {
  const deck = parseDeck(md);
  const out: string[] = [];
  for (const s of deck.slides) {
    if (s.meta.layout && !(s.meta.layout in LAYOUTS)) out.push(`slide ${s.index + 1}: unknown layout "${s.meta.layout}" (shown as default)`);
    if (["image-left", "image-right", "full-image"].includes(s.layout) && !s.meta.image && !s.blocks.some((b) => b.type === "image")) out.push(`slide ${s.index + 1}: layout ${s.layout} needs image:`);
    const words = [...s.blocks, ...s.right, ...s.columns.flat()].map((b) => ("content" in b && Array.isArray(b.content) ? plain(b.content as never) : "")).join(" ").split(/\s+/).length;
    if (words > 140) out.push(`slide ${s.index + 1}: ${words} words, too much text for one slide (split it)`);
  }
  return out;
}

export const readDeck = tool({
  name: "read_deck",
  description: "The deck's markdown, slide by slide (numbered), with the layout of each slide.",
  args: z.object({}),
  async run(_, ctx) {
    const md = deckOf(ctx);
    const { head, slides } = slideSources(md);
    const deck = parseDeck(md);
    return [`Headmatter:\n${head}`, ...slides.map((s, i) => `## Slide ${i + 1} (${deck.slides[i]?.layout ?? "?"})\n${s}`)].join("\n\n");
  },
});

export const writeDeck = tool({
  name: "write_deck",
  description: "Replace the whole deck with new markdown (headmatter, then slides separated by --- lines, each optionally starting with its frontmatter `layout: …`).",
  args: z.object({ markdown: z.string() }),
  changes: true,
  async run(a, ctx) {
    setDoc(ctx, a.markdown);
    const problems = deckProblems(a.markdown);
    return `Deck written: ${parseDeck(a.markdown).slides.length} slides.${problems.length ? `\nTo fix:\n${problems.join("\n")}` : ""} Check it with look_at_deck.`;
  },
});

export const editSlides = tool({
  name: "edit_slides",
  description:
    "Change slides without rewriting the deck: replace slide n with new markdown (its frontmatter included, e.g. \"layout: kpis\\n---\\n# Title\\n…\"), insert new slides after slide n (0 = at the start), delete slide n, or move slide n to position `to`.",
  args: z.object({
    action: z.enum(["replace", "insert_after", "delete", "move"]),
    slide: z.number().int().min(0),
    markdown: z.string().optional().describe("replace / insert_after: the slide(s); several slides separated by --- lines."),
    to: z.number().int().optional(),
  }),
  changes: true,
  async run(a, ctx) {
    const { head, slides } = slideSources(deckOf(ctx));
    const given = a.markdown ? slideSources(`---\nx: 1\n---\n${a.markdown}`).slides : [];
    const i = a.slide - 1;
    if (a.action !== "insert_after" && (i < 0 || i >= slides.length)) throw new Error(`Slides are numbered 1 to ${slides.length}.`);
    if (a.action === "replace") slides.splice(i, 1, ...given);
    else if (a.action === "insert_after") slides.splice(a.slide, 0, ...given);
    else if (a.action === "delete") slides.splice(i, 1);
    else {
      const [s] = slides.splice(i, 1);
      slides.splice(Math.max(0, Math.min(slides.length, (a.to ?? 1) - 1)), 0, s);
    }
    const md = joinSlides(head, slides);
    setDoc(ctx, md);
    const problems = deckProblems(md);
    return `Done: the deck has ${slides.length} slides.${problems.length ? `\nTo fix:\n${problems.join("\n")}` : ""}`;
  },
});
