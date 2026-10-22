import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import CodeMirror from "@uiw/react-codemirror";
import { json as jsonLanguage } from "@codemirror/lang-json";
import { AnimatePresence, motion } from "motion/react";
import { BracesIcon, DatabaseIcon, FileChartColumnIcon, FileSpreadsheetIcon, ChevronLeftIcon, ChevronRightIcon, DownloadIcon, FileTextIcon, MaximizeIcon, PlusIcon, PresentationIcon, SaveIcon } from "lucide-react";
import { EntityMenu, withToast } from "@/components/entity";
import { Button } from "@/components/ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Switch } from "@/components/ui/switch";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { slidesAgent } from "@/agent/slides-tools";
import { useLast } from "@/lib/last";
import { errorText } from "@/lib/ms";
import { patch, put, remove, useCollection, type DeckEntry } from "@/lib/store";
import { EXAMPLE_DECK, TEMPLATES, deckInfo, parseDeck, type Deck } from "@/slides";
import { downloadPptx } from "@/slides/pptx";
import { SlideView, useFonts } from "@/slides/stage";

export const Route = createFileRoute("/slides")({
  validateSearch: (s: Record<string, unknown>): { deck?: string; tab?: string } => ({ deck: (s.deck as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Slides,
});

/**
 * Slides: a presentation written as JSON (theme + slides, each a template and its args), drawn by the
 * templates of src/slides: animated preview (fullscreen), the JSON itself, PDF (print) and PowerPoint.
 */
function Slides() {
  const { deck: deckId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/slides" });
  useLast("slides:deck", deckId, (deck) => navigate({ search: { deck }, replace: true }));
  const decks = useCollection("decks");
  const entry = decks.find((d) => d.id === deckId);
  const [index, setIndex] = useState(0);
  const live = useRef({ entry, index });
  live.current = { entry, index };
  useEffect(() => setIndex(0), [deckId]);
  const count = entry?.deck.slides.length ?? 0;
  useEffect(() => {
    if (index >= count && count) setIndex(count - 1);
  }, [count, index]);
  const datasets = useCollection("datasets");
  const reports = useCollection("reports");
  const schemas = useCollection("schemas").filter((x) => x.excel);
  const sources = entry?.sources ?? { datasetIds: [], reportIds: [], schemaIds: [] };
  const chosen = { datasets: datasets.filter((d) => sources.datasetIds.includes(d.id)), reports: reports.filter((r) => sources.reportIds.includes(r.id)), schemas: schemas.filter((x) => sources.schemaIds.includes(x.id)) };
  const liveSources = useRef(chosen);
  liveSources.current = chosen;
  const agent = entry && slidesAgent(() => live.current.entry!, setIndex, () => live.current.index, () => liveSources.current);
  const toggle = (kind: keyof typeof sources, id: string, on: boolean) => entry && void patch("decks", entry.id, { sources: { ...sources, [kind]: on ? [...sources[kind], id] : sources[kind].filter((x) => x !== id) } });
  const create = async () => {
    const name = prompt("Name of the presentation", "New presentation")?.trim();
    if (name) navigate({ search: { deck: (await put("decks", { id: crypto.randomUUID(), name, deck: { ...EXAMPLE_DECK, title: name }, updatedAt: Date.now() })).id } });
  };

  return (
    <Workspace
      title="Slides"
      icon={PresentationIcon}
      picker={<EntityPicker items={decks.map((d) => ({ id: d.id, label: d.name }))} value={deckId} onChange={(deck) => navigate({ search: { deck } })} placeholder={decks.length ? "Choose a presentation" : "New presentation"} onAdd={() => void create()} addLabel="New presentation…" />}
      actions={
        <>
          {entry && (
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button size="sm" variant="outline">
                  <DownloadIcon /> Export
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-60">
                <DropdownMenuItem onSelect={() => setTimeout(() => window.print(), 100)}>
                  <FileTextIcon />
                  <span className="flex flex-col">
                    PDF
                    <span className="text-[11px] text-muted-foreground">Print › Save as PDF (landscape, no margins)</span>
                  </span>
                </DropdownMenuItem>
                <DropdownMenuItem onSelect={() => void withToast("Writing the PowerPoint…", downloadPptx(entry.deck, entry.name), "Downloaded.")}>
                  <PresentationIcon />
                  <span className="flex flex-col">
                    PowerPoint (.pptx)
                    <span className="text-[11px] text-muted-foreground">Editable; Power BI slides become links</span>
                  </span>
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          )}
          <Button size="sm" variant="outline" onClick={() => void create()}>
            <PlusIcon /> New
          </Button>
          <EntityMenu
            kind="decks"
            entry={entry}
            onRename={entry && (() => { const name = prompt("Name of the presentation", entry.name)?.trim(); if (name) void patch("decks", entry.id, { name }); })}
            onRemove={entry && (() => confirm(`Delete the presentation "${entry.name}"?`) && void remove("decks", entry.id).then(() => navigate({ search: {} })))}
          />
        </>
      }
      ai={{ scope: `slides:${entry?.id}`, agent, placeholder: "Ask to write a deck about…, add a slide, change the colors…", empty: "Choose a presentation, or start a new one." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        { id: "preview", label: "Preview", icon: PresentationIcon, count: entry ? count : undefined, content: entry ? <Preview deck={entry.deck} index={index} onIndex={setIndex} active={(tab ?? "preview") === "preview"} /> : <Empty icon={PresentationIcon} action={<Button onClick={() => void create()}><PlusIcon /> New presentation</Button>}>A presentation is a short JSON (a theme and slides from templates) written with the assistant or by hand.</Empty> },
        { id: "content", label: "Content", icon: BracesIcon, content: entry ? <Content key={entry.id} entry={entry} /> : <Empty>Choose a presentation.</Empty> },
        {
          id: "datasets",
          label: "Datasets",
          icon: DatabaseIcon,
          count: entry ? sources.datasetIds.length : undefined,
          content: entry ? <SourceList title="Semantic models" description="The assistant computes the deck's figures on them with DAX." items={datasets.map((d) => ({ id: d.id, label: d.name, hint: d.access === "view" ? "no DAX access" : undefined, disabled: d.access === "view" }))} chosen={sources.datasetIds} onToggle={(id, on) => toggle("datasetIds", id, on)} /> : <Empty>Choose a presentation.</Empty>,
        },
        {
          id: "reports",
          label: "Reports",
          icon: FileChartColumnIcon,
          count: entry ? sources.reportIds.length : undefined,
          content: entry ? <SourceList title="Power BI reports" description="Their pages and visuals can be shown in the slides (visual / compare)." items={reports.map((r) => ({ id: r.id, label: r.name }))} chosen={sources.reportIds} onToggle={(id, on) => toggle("reportIds", id, on)} /> : <Empty>Choose a presentation.</Empty>,
        },
        {
          id: "excel",
          label: "Excel",
          icon: FileSpreadsheetIcon,
          count: entry ? sources.schemaIds.length : undefined,
          content: entry ? <SourceList title="Excel tables" description="The assistant reads their rows to write the slides." items={schemas.map((x) => ({ id: x.id, label: x.title }))} chosen={sources.schemaIds} onToggle={(id, on) => toggle("schemaIds", id, on)} /> : <Empty>Choose a presentation.</Empty>,
        },
      ]}
    />
  );
}

/* ---------------------------------- Sources ---------------------------------- */

/** The items a deck is written from (none by default: the assistant then gets no tool for them). */
function SourceList({ title, description, items, chosen, onToggle }: { title: string; description: string; items: { id: string; label: string; hint?: string; disabled?: boolean }[]; chosen: string[]; onToggle: (id: string, on: boolean) => void }) {
  return (
    <div className="mx-auto flex w-full max-w-3xl flex-col gap-4 overflow-y-auto p-6">
      <Section title={title} description={description}>
        {!items.length && <p className="text-[13px] text-muted-foreground">None in the app yet: add them in their screen.</p>}
        <div className="flex flex-col divide-y">
          {items.map((i) => (
            <label key={i.id} className="flex cursor-pointer items-center gap-3 py-2.5 text-[13.5px]">
              <Switch checked={chosen.includes(i.id)} disabled={i.disabled} onCheckedChange={(on) => onToggle(i.id, on)} />
              <span className="min-w-0 flex-1 truncate font-medium">{i.label}</span>
              {i.hint && <span className="text-[12px] text-muted-foreground">{i.hint}</span>}
            </label>
          ))}
        </div>
      </Section>
    </div>
  );
}

/* ---------------------------------- Preview ---------------------------------- */

function Preview({ deck, index, onIndex, active }: { deck: Deck; index: number; onIndex: (i: number) => void; active: boolean }) {
  useFonts(deck.theme);
  const infos = useMemo(() => deckInfo(deck), [deck]);
  const stage = useRef<HTMLDivElement>(null);
  const [direction, setDirection] = useState(1);
  const go = (i: number) => {
    const next = Math.max(0, Math.min(i, deck.slides.length - 1));
    setDirection(next >= index ? 1 : -1);
    onIndex(next);
  };
  // Arrows / space / page keys change the slide (not while typing).
  useEffect(() => {
    if (!active) return;
    const key = (e: KeyboardEvent) => {
      if (e.target instanceof Element && e.target.closest("input, textarea, [contenteditable=true], .cm-editor")) return;
      if (["ArrowRight", "PageDown", " "].includes(e.key)) (e.preventDefault(), go(index + 1));
      if (["ArrowLeft", "PageUp"].includes(e.key)) (e.preventDefault(), go(index - 1));
      if (e.key === "Home") go(0);
      if (e.key === "End") go(deck.slides.length - 1);
    };
    window.addEventListener("keydown", key);
    return () => window.removeEventListener("keydown", key);
  });
  if (!deck.slides.length) return <Empty icon={PresentationIcon}>No slide yet: ask the assistant, or write the JSON in the Content tab.</Empty>;
  const slide = deck.slides[index] ?? deck.slides[0];
  return (
    <div className="flex min-h-0 flex-1 flex-col bg-muted/40">
      <div className="flex min-h-0 flex-1 items-center justify-center p-6">
        <div ref={stage} className="group relative w-full max-w-[min(100%,calc((100vh-15rem)*16/9))] overflow-hidden rounded-xl bg-black shadow-lg ring-1 ring-border [&:fullscreen]:flex [&:fullscreen]:max-w-none [&:fullscreen]:items-center [&:fullscreen]:rounded-none">
          <AnimatePresence mode="popLayout" initial={false} custom={direction}>
            <motion.div
              key={`${index}:${JSON.stringify(slide).length}`}
              custom={direction}
              initial={{ opacity: 0, x: `${direction * 4}%` }}
              animate={{ opacity: 1, x: 0, transition: { duration: 0.5, ease: [0.22, 1, 0.36, 1] } }}
              exit={{ opacity: 0, x: `${-direction * 4}%`, transition: { duration: 0.35 } }}
              className="w-full"
            >
              <SlideView slide={slide} deck={deck} info={infos[index] ?? infos[0]} />
            </motion.div>
          </AnimatePresence>
          <div className="absolute right-3 bottom-3 flex gap-1.5 opacity-0 transition-opacity group-hover:opacity-100">
            <Button size="icon" variant="secondary" className="size-8" aria-label="Previous slide" onClick={() => go(index - 1)}>
              <ChevronLeftIcon />
            </Button>
            <Button size="icon" variant="secondary" className="size-8" aria-label="Next slide" onClick={() => go(index + 1)}>
              <ChevronRightIcon />
            </Button>
            <Button size="icon" variant="secondary" className="size-8" aria-label="Fullscreen" onClick={() => void stage.current?.requestFullscreen()}>
              <MaximizeIcon />
            </Button>
          </div>
        </div>
      </div>
      <div className="flex shrink-0 gap-2.5 overflow-x-auto border-t bg-card px-5 py-3">
        {deck.slides.map((s, i) => (
          <button key={i} type="button" onClick={() => go(i)} className={`relative w-36 shrink-0 cursor-pointer overflow-hidden rounded-md ring-offset-2 transition-shadow ${i === index ? "ring-2 ring-primary" : "ring-1 ring-border hover:ring-primary/50"}`} aria-label={`Slide ${i + 1}`}>
            <div className="pointer-events-none">
              <SlideView slide={s} deck={deck} info={infos[i]} still />
            </div>
            <span className="absolute top-1 left-1 rounded bg-black/60 px-1 text-[10px] text-white tabular-nums">{i + 1}</span>
          </button>
        ))}
      </div>
      {/* Printed instead of the app (PDF export): every slide, one per page, without animation. */}
      <div className="print-deck">
        {deck.slides.map((s, i) => (
          <div key={i} className="print-slide">
            <SlideView slide={s} deck={deck} info={infos[i]} still />
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------- Content ---------------------------------- */

/** The deck as JSON: checked as it is typed (each slide against its template), saved on demand. */
function Content({ entry }: { entry: DeckEntry }) {
  const saved = JSON.stringify(entry.deck, null, 2);
  const [text, setText] = useState(saved);
  const [error, setError] = useState<string>();
  const last = useRef(saved);
  useEffect(() => {
    if (text === last.current) setText(saved);
    last.current = saved;
  }, [saved]); // eslint-disable-line react-hooks/exhaustive-deps
  const problems = useMemo(() => {
    try {
      return parseDeck(JSON.parse(text)).errors;
    } catch (e) {
      return [`Invalid JSON: ${errorText(e)}`];
    }
  }, [text]);
  const changed = text !== saved;
  const save = () => {
    const { deck } = parseDeck(JSON.parse(text));
    if (deck) patch("decks", entry.id, { deck, updatedAt: Date.now() }).catch((e) => setError(errorText(e)));
  };
  return (
    <div className="flex min-h-0 flex-1 gap-4 p-6">
      <div className="flex min-w-0 flex-1 flex-col gap-3">
        <div className="flex items-center gap-2">
          <span className="text-[13px] text-muted-foreground">{"{ title, theme: { primary, background, text, fontTitle, fontText }, slides: [ { template, …args } ] }"}</span>
          <span className="ml-auto" />
          {changed && (
            <>
              <Button size="sm" variant="ghost" onClick={() => setText(saved)}>Discard</Button>
              <Button size="sm" disabled={problems.length > 0} onClick={save}>
                <SaveIcon /> Save
              </Button>
            </>
          )}
        </div>
        {problems.length > 0 && <p className="rounded-lg bg-destructive/10 px-3 py-2 text-[12.5px] text-destructive">{problems.join(" · ")}</p>}
        {error && <p className="text-[12.5px] text-destructive">{error}</p>}
        <div className="min-h-0 flex-1 overflow-hidden rounded-xl border bg-card">
          <CodeMirror value={text} height="100%" className="h-full text-[13px]" extensions={[jsonLanguage()]} onChange={setText} basicSetup={{ foldGutter: true, highlightActiveLine: false }} />
        </div>
      </div>
      <aside className="hidden w-80 shrink-0 flex-col gap-2 overflow-y-auto rounded-xl border bg-card p-4 xl:flex">
        <span className="text-[13px] font-semibold">Templates</span>
        {TEMPLATES.map((t) => (
          <div key={t.id} className="flex flex-col gap-0.5 border-b pb-2 last:border-b-0">
            <code className="text-[12px] font-semibold text-primary">{t.id}</code>
            <span className="text-[12px] text-muted-foreground">{t.description}</span>
            <span className="font-mono text-[11px] text-muted-foreground/80">{Object.keys(t.args.shape).join(", ")}</span>
          </div>
        ))}
      </aside>
    </div>
  );
}
