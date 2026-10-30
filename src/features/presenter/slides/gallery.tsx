/**
 * The slide templates, one at a time and large (Settings › Slides): drawn with example content (lorem
 * ipsum, Unsplash); on the right, a category and its templates, the header tried, arrows going
 * through all of them.
 */
import { useMemo, useState } from "react";
import { ChevronLeftIcon, ChevronRightIcon } from "lucide-react";
import type { z } from "zod";
import { Button } from "@/components/ui/button";
import { EXAMPLE_DECK, TEMPLATES, deckInfo, type Deck } from "./index";
import { exampleSlide } from "./examples";
import { HEADERS, appPrimary, type Template } from "./kit";
import { SlideView, useFonts } from "./stage";

/** The templates by category (Templates tab); a template in none goes in the last one. */
const CATEGORIES: { label: string; ids: string[] }[] = [
  { label: "Covers and endings", ids: ["title", "cover", "hero", "end"] },
  { label: "Structure", ids: ["summary", "agenda", "subtitle"] },
  { label: "Text", ids: ["list", "markdown", "quote", "beware", "lists", "context", "causality"] },
  { label: "Cards and figures", ids: ["cards", "features", "kpis", "table", "roadmap", "image"] },
  { label: "People and use cases", ids: ["usecase", "conversation", "responsibilities"] },
  { label: "Diagrams", ids: ["architecture", "pyramid"] },
  { label: "Power BI", ids: ["report", "visual", "tutorial", "compare"] },
];
const categoryOf = (id: string) => Math.max(0, CATEGORIES.findIndex((c) => c.ids.includes(id)));

export function TemplateGallery({ report }: { report?: string }) {
  const theme = useMemo(() => ({ ...EXAMPLE_DECK.theme, primary: appPrimary() }), []);
  const [i, setI] = useState(0);
  // The header of content pages, tried here without changing the deck.
  const [header, setHeader] = useState(theme.header);
  const template = TEMPLATES[i] as Template<z.ZodObject>;
  const slide = useMemo(() => exampleSlide(template, report), [template, report]);
  const deck: Deck = { title: "Lorem ipsum", theme: { ...theme, header }, slides: [{ template: "agenda", part: "Lorem ipsum" }, slide] };
  useFonts(theme);
  const info = deckInfo(deck)[1];
  const go = (n: number) => setI((n + TEMPLATES.length) % TEMPLATES.length);
  const category = categoryOf(template.id);
  const inCategory = TEMPLATES.map((t, n) => ({ t, n })).filter(({ t }) => categoryOf(t.id) === category);
  return (
    <div className="flex min-h-0 flex-1 bg-muted/40">
      <div className="flex min-w-0 flex-1 flex-col gap-3 overflow-y-auto p-6">
        <div className="flex flex-col">
          <span className="font-mono text-[15px] font-semibold text-primary">{template.id}</span>
          <span className="text-[12.5px] text-muted-foreground">{template.description}</span>
        </div>
        <div className="mx-auto w-full max-w-[min(100%,calc((100vh-16rem)*16/9))] overflow-hidden rounded-xl shadow-lg ring-1 ring-border">
          <SlideView key={template.id} slide={slide} deck={deck} info={info} />
        </div>
      </div>
      <aside className="flex w-60 shrink-0 flex-col gap-3 overflow-y-auto border-l bg-card p-4">
        <div className="flex items-center gap-2">
          <Button size="icon" variant="outline" className="size-8" aria-label="Previous template" onClick={() => go(i - 1)}>
            <ChevronLeftIcon />
          </Button>
          <span className="flex-1 text-center text-[12px] text-muted-foreground tabular-nums">{i + 1} / {TEMPLATES.length}</span>
          <Button size="icon" variant="outline" className="size-8" aria-label="Next template" onClick={() => go(i + 1)}>
            <ChevronRightIcon />
          </Button>
        </div>
        <select
          className="h-9 rounded-md border border-input bg-background px-2 text-[13px]"
          aria-label="Category"
          value={category}
          onChange={(e) => setI(TEMPLATES.findIndex((t) => categoryOf(t.id) === Number(e.target.value)))}
        >
          {CATEGORIES.map((c, n) => (
            <option key={c.label} value={n}>
              {c.label}
            </option>
          ))}
        </select>
        <label className="flex flex-col gap-1 text-[12px] text-muted-foreground">
          Header
          <select className="h-9 rounded-md border border-input bg-background px-2 text-[13px] text-foreground" value={header} onChange={(e) => setHeader(e.target.value as Deck["theme"]["header"])}>
            {HEADERS.map((h) => (
              <option key={h} value={h}>
                {h}
              </option>
            ))}
          </select>
        </label>
        <div className="flex flex-col gap-0.5 border-t pt-3">
          {inCategory.map(({ t, n }) => (
            <button key={t.id} type="button" onClick={() => setI(n)} className={`cursor-pointer rounded-md px-2.5 py-1.5 text-left font-mono text-[12.5px] transition-colors ${n === i ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}>
              {t.id}
            </button>
          ))}
        </div>
      </aside>
    </div>
  );
}
