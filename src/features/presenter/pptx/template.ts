/**
 * The template deck (Settings › Slides): an ordinary PowerPoint where each slide is a type of slide
 * (cover, agenda, two columns, a chart page…). Its name comes from the slide's notes ("type: Two
 * columns"), else its title, else its layout. The AI adds slides of these types to any deck: the slide
 * is copied with its look, then its texts are filled (deck.ts › copySlide).
 */
import type { Deck, Shape } from "./deck";

export type SlideType = { number: number; name: string; layout: string; slots: Shape[] };

export async function slideTypes(deck: Deck): Promise<SlideType[]> {
  const seen = new Map<string, number>();
  return (await deck.slides()).map((s) => {
    const base = s.notes.match(/^\s*(?:type|layout|name)\s*:\s*(.+)$/im)?.[1]?.trim() || s.title || s.layout || `Slide ${s.number}`;
    const n = (seen.get(base) ?? 0) + 1;
    seen.set(base, n);
    return { number: s.number, name: n > 1 ? `${base} ${n}` : base, layout: s.layout, slots: s.shapes.filter((x) => x.kind === "text" || x.kind === "picture" || x.placeholder === "pic") };
  });
}

export const typeOf = (types: SlideType[], name: string) => types.find((t) => t.name === name) ?? types.find((t) => t.name.toLowerCase() === name.trim().toLowerCase()) ?? types.find((t) => String(t.number) === name.trim());

/** The types for the AI: each with its slots (shape ids it may fill, their place and room). */
export function typesText(types: SlideType[], deck: Deck) {
  return types
    .map((t) => `- "${t.name}" (layout ${t.layout}): ${t.slots.map((s) => `#${s.id} ${s.kind}${s.placeholder ? `/${s.placeholder}` : ""} (${deck.pct(s.box)}${s.budget ? `, ≈ ${s.budget} chars` : ""})${s.paragraphs.length ? ` "${s.paragraphs.map((p) => p.text).join(" / ").slice(0, 60)}"` : ""}`).join("; ")}`)
    .join("\n");
}
