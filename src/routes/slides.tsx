import { Fragment, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import CodeMirror from "@uiw/react-codemirror";
import { json as jsonLanguage } from "@codemirror/lang-json";
import { AnimatePresence, motion } from "motion/react";
import { BracesIcon, ClipboardCopyIcon, DatabaseIcon, FileChartColumnIcon, ChevronLeftIcon, ChevronRightIcon, DownloadIcon, FileTextIcon, MaximizeIcon, PlusIcon, PresentationIcon, SaveIcon } from "lucide-react";
import { SCREEN_LABELS } from "@/config";
import { EntityMenu, withToast } from "@/components/entity";
import { Button } from "@/components/ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Switch } from "@/components/ui/switch";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { slidesAgent } from "@/features/presenter/slides-tools";
import { useLast } from "@/lib/last";
import { errorText } from "@/services/ms";
import { patch, put, remove, useCollection, type DeckEntry } from "@/lib/store";
import { toast } from "sonner";
import { capturePowerBi, powerBiSlides } from "@/features/presenter/slides/capture";
import { argRows } from "@/features/presenter/slides/examples";
import { appPrimary } from "@/features/presenter/slides/kit";
import { EXAMPLE_DECK, TEMPLATES, checkSlide, deckInfo, parseDeck, templateOf, type Deck, type SlideJson } from "@/features/presenter/slides";
import { downloadPptx } from "@/features/presenter/slides/pptx";
import { SlideView, useFonts } from "@/features/presenter/slides/stage";
import { deckPrompt } from "@/prompts/presenter";

export const Route = createFileRoute("/slides")({
  validateSearch: (s: Record<string, unknown>): { deck?: string; tab?: string } => ({ deck: (s.deck as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Slides,
});

/**
 * Slides: a presentation written as JSON (theme + slides, each a template and its args), drawn by the
 * templates of src/features/presenter/slides: animated preview (fullscreen), the JSON itself, PDF (print) and PowerPoint.
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
  const sources = { datasetIds: entry?.sources?.datasetIds ?? [], reportIds: entry?.sources?.reportIds ?? [] };
  const chosen = { datasets: datasets.filter((d) => sources.datasetIds.includes(d.id)), reports: reports.filter((r) => sources.reportIds.includes(r.id)) };
  const shown = entry?.deck;
  const liveSources = useRef(chosen);
  liveSources.current = chosen;
  const agent = entry && slidesAgent(() => live.current.entry!, setIndex, () => live.current.index, () => liveSources.current);
  const toggle = (kind: keyof typeof sources, id: string, on: boolean) => entry && void patch("decks", entry.id, { sources: { ...entry.sources, ...sources, [kind]: on ? [...sources[kind], id] : sources[kind].filter((x) => x !== id) } });
  /**
   * Before an export, the Power BI views are captured from the screen (each slide shown in turn) so the
   * PDF and the PowerPoint hold their picture; declined, they keep a placeholder.
   */
  const capture = async () => {
    if (!shown || !powerBiSlides(shown).length) return;
    navigate({ search: (s) => ({ ...s, tab: "preview" }), replace: true });
    const back = live.current.index;
    try {
      await capturePowerBi(shown, setIndex);
    } catch (e) {
      toast.warning(`The Power BI views were not captured (${errorText(e)}): they show as placeholders.`);
    } finally {
      setIndex(back);
    }
  };
  const create = async () => {
    const name = prompt("Name of the presentation", "New presentation")?.trim();
    if (name) navigate({ search: { deck: (await put("decks", { id: crypto.randomUUID(), name, deck: { ...EXAMPLE_DECK, title: name, theme: { ...EXAMPLE_DECK.theme, primary: appPrimary() } }, updatedAt: Date.now() })).id } });
  };

  return (
    <Workspace
      title={SCREEN_LABELS.slides}
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
                <DropdownMenuItem onSelect={() => void capture().then(() => setTimeout(() => window.print(), 300))}>
                  <FileTextIcon />
                  <span className="flex flex-col">
                    PDF
                    <span className="text-[11px] text-muted-foreground">Print › Save as PDF (landscape, no margins); share this tab to capture Power BI</span>
                  </span>
                </DropdownMenuItem>
                <DropdownMenuItem onSelect={() => void withToast("Writing the PowerPoint…", capture().then(() => downloadPptx(shown!, entry.name)), "Downloaded.")}>
                  <PresentationIcon />
                  <span className="flex flex-col">
                    PowerPoint (.pptx)
                    <span className="text-[11px] text-muted-foreground">Editable; Power BI views as pictures linking to the report</span>
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
        { id: "preview", label: "Preview", icon: PresentationIcon, count: entry ? count : undefined, content: entry ? <Preview deck={shown!} index={index} onIndex={setIndex} active={(tab ?? "preview") === "preview"} /> : <Empty icon={PresentationIcon} action={<Button onClick={() => void create()}><PlusIcon /> New presentation</Button>}>A presentation is a short JSON (a theme and slides from templates) written with the assistant or by hand.</Empty> },
        { id: "content", label: "Edit", icon: BracesIcon, content: entry ? <Edit key={entry.id} entry={entry} index={index} onIndex={setIndex} /> : <Empty>Choose a presentation.</Empty> },
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
          {/* "wait": the slide leaving goes first, so the next one reuses its Power BI embed. */}
          <AnimatePresence mode="wait" initial={false} custom={direction}>
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

/* ----------------------------------- Edit ----------------------------------- */

type EditMode = "slide" | "deck";

/** The deck's JSON: slide by slide (its preview, its JSON, its template), or the whole deck at once. */
function Edit({ entry, index, onIndex }: { entry: DeckEntry; index: number; onIndex: (i: number) => void }) {
  const [mode, setMode] = useState<EditMode>(() => (localStorage.getItem("slides:edit") as EditMode) || "slide");
  const pick = (m: EditMode) => (setMode(m), localStorage.setItem("slides:edit", m));
  const slides = entry.deck.slides.length;
  const toggle = (
    <ToggleGroup type="single" variant="outline" size="sm" value={mode} onValueChange={(v) => v && pick(v as EditMode)}>
      <ToggleGroupItem value="slide" className="px-3">Slide by slide</ToggleGroupItem>
      <ToggleGroupItem value="deck" className="px-3">Whole deck</ToggleGroupItem>
    </ToggleGroup>
  );
  if (mode === "deck" || !slides) return <DeckJson entry={entry} toggle={toggle} />;
  const i = Math.min(index, slides - 1);
  return <SlideJson key={`${i}:${JSON.stringify(entry.deck.slides[i])}`} entry={entry} index={i} onIndex={onIndex} toggle={toggle} />;
}

/** A JSON text checked as it is typed, saved on demand (Discard / Save). */
function useJsonDraft(saved: string, check: (value: unknown) => string[]) {
  const [text, setText] = useState(saved);
  const last = useRef(saved);
  useEffect(() => {
    if (text === last.current) setText(saved);
    last.current = saved;
  }, [saved]); // eslint-disable-line react-hooks/exhaustive-deps
  const problems = useMemo(() => {
    try {
      return check(JSON.parse(text));
    } catch (e) {
      return [`Invalid JSON: ${errorText(e)}`];
    }
  }, [text]); // eslint-disable-line react-hooks/exhaustive-deps
  return { text, setText, problems, changed: text !== saved, discard: () => setText(saved) };
}

function SaveBar({ draft, onSave, children }: { draft: ReturnType<typeof useJsonDraft>; onSave: () => void; children?: ReactNode }) {
  return (
    <div className="flex items-center gap-2">
      {children}
      <span className="ml-auto" />
      {draft.changed && (
        <>
          <Button size="sm" variant="ghost" onClick={draft.discard}>Discard</Button>
          <Button size="sm" disabled={draft.problems.length > 0} onClick={onSave}>
            <SaveIcon /> Save
          </Button>
        </>
      )}
    </div>
  );
}

const editor = (draft: ReturnType<typeof useJsonDraft>) => (
  <div className="min-h-0 flex-1 overflow-hidden rounded-xl border bg-card">
    <CodeMirror value={draft.text} height="100%" className="h-full text-[13px]" extensions={[jsonLanguage()]} onChange={draft.setText} basicSetup={{ foldGutter: true, highlightActiveLine: false }} />
  </div>
);

/** The whole deck as JSON (theme and every slide). */
function DeckJson({ entry, toggle }: { entry: DeckEntry; toggle: ReactNode }) {
  const draft = useJsonDraft(JSON.stringify(entry.deck, null, 2), (v) => parseDeck(v).errors);
  const save = () => {
    const { deck } = parseDeck(JSON.parse(draft.text));
    if (deck) void withToast("Saving…", patch("decks", entry.id, { deck, updatedAt: Date.now() }), "Saved.");
  };
  return (
    <div className="flex min-h-0 flex-1 flex-col gap-3 p-6">
      <SaveBar draft={draft} onSave={save}>
        {toggle}
        <Button size="sm" variant="outline" onClick={() => void navigator.clipboard.writeText(promptFor(draft.text)).then(() => toast.success("Prompt copied: paste it into ChatGPT (or another assistant) with your request, then paste its JSON here."))}>
          <ClipboardCopyIcon /> Copy the prompt
        </Button>
      </SaveBar>
      {draft.problems.length > 0 && <p className="rounded-lg bg-destructive/10 px-3 py-2 text-[12.5px] text-destructive">{draft.problems.join(" · ")}</p>}
      {editor(draft)}
    </div>
  );
}

/** One slide: its preview on the left (drawn from the JSON as it is typed), its JSON and its template's arguments on the right. */
function SlideJson({ entry, index, onIndex, toggle }: { entry: DeckEntry; index: number; onIndex: (i: number) => void; toggle: ReactNode }) {
  const deck = entry.deck;
  useFonts(deck.theme);
  const draft = useJsonDraft(JSON.stringify(deck.slides[index], null, 2), (v) => {
    try {
      checkSlide(v as SlideJson);
      return [];
    } catch (e) {
      return [errorText(e)];
    }
  });
  // The slide as typed when it is valid, else as saved.
  const shown = useMemo(() => {
    const slides = [...deck.slides];
    if (!draft.problems.length) slides[index] = JSON.parse(draft.text) as SlideJson;
    return { ...deck, slides };
  }, [deck, index, draft.text, draft.problems.length]);
  const info = deckInfo(shown)[index];
  const template = templateOf(String((shown.slides[index] as SlideJson).template));
  const save = () => void withToast("Saving…", patch("decks", entry.id, { deck: shown, updatedAt: Date.now() }), "Saved.");
  return (
    <div className="flex min-h-0 flex-1 flex-col gap-3 p-6">
      <SaveBar draft={draft} onSave={save}>{toggle}</SaveBar>
      <div className="grid min-h-0 flex-1 grid-cols-[minmax(0,2fr)_minmax(0,3fr)] gap-5">
        <div className="flex min-h-0 flex-col gap-3">
          <div className="flex items-center gap-2">
            <Button size="icon" variant="outline" className="size-8" aria-label="Previous slide" disabled={index === 0} onClick={() => onIndex(index - 1)}>
              <ChevronLeftIcon />
            </Button>
            <select className="h-8 min-w-0 flex-1 rounded-md border border-input bg-background px-2 text-[13px]" aria-label="Slide" value={index} onChange={(e) => onIndex(Number(e.target.value))}>
              {deck.slides.map((s, i) => (
                <option key={i} value={i}>
                  {i + 1}. {s.template}
                  {s.title || s.part ? ` — ${String(s.title ?? s.part)}` : ""}
                </option>
              ))}
            </select>
            <Button size="icon" variant="outline" className="size-8" aria-label="Next slide" disabled={index === deck.slides.length - 1} onClick={() => onIndex(index + 1)}>
              <ChevronRightIcon />
            </Button>
          </div>
          <div className="overflow-hidden rounded-lg shadow-md ring-1 ring-border">
            <SlideView slide={shown.slides[index]} deck={shown} info={info} still />
          </div>
          {draft.problems.length > 0 && <p className="rounded-lg bg-destructive/10 px-3 py-2 text-[12.5px] text-destructive">{draft.problems.join(" · ")} — the preview shows the saved slide.</p>}
        </div>
        <div className="flex min-h-0 flex-col gap-3">
          {editor(draft)}
          {template && (
            <details className="max-h-[40%] shrink-0 overflow-y-auto rounded-xl border bg-card px-4 py-3 text-[12.5px]" open>
              <summary className="cursor-pointer font-medium">
                Template <span className="font-mono text-primary">{template.id}</span>
              </summary>
              <p className="mt-1.5 text-muted-foreground">{template.description}</p>
              <dl className="mt-2 grid grid-cols-[auto_minmax(0,1fr)] gap-x-3 gap-y-1">
                {argRows(template).map((a) => (
                  <Fragment key={a.name}>
                    <dt className="font-mono">
                      {a.name}
                      {a.required ? "" : "?"}
                    </dt>
                    <dd className="text-muted-foreground">
                      <span className="font-mono text-foreground/70">{a.type}</span>
                      {a.description && ` — ${a.description}`}
                    </dd>
                  </Fragment>
                ))}
              </dl>
            </details>
          )}
        </div>
      </div>
    </div>
  );
}

/**
 * The prompt for an outside assistant (ChatGPT…): the deck's format, every template with its
 * arguments, the current JSON; it answers with the whole deck as JSON only, in the user's language.
 */
function promptFor(current: string) {
  const templates = TEMPLATES.map((t) => `### ${t.id} (${t.kind ?? "content"})\n${t.description}\n${argRows(t).map((a) => `- ${a.name}${a.required ? "" : "?"}: ${a.type}${a.description ? ` — ${a.description}` : ""}`).join("\n")}`).join("\n\n");
  return deckPrompt(templates, current);
}
