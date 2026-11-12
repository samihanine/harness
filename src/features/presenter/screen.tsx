/**
 * Presenter: a real PowerPoint deck (from this computer or SharePoint) changed with the assistant —
 * texts, new slides from the template deck (Settings › Slides), Power BI visuals as pictures — then
 * saved back to its file. The slides are drawn from their XML (preview.tsx); the deck's sources
 * (models for figures, reports for visuals) are chosen in their tabs.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { Link, useNavigate, useSearch } from "@tanstack/react-router";
import { ChevronLeftIcon, ChevronRightIcon, CloudIcon, DatabaseIcon, DownloadIcon, ExternalLinkIcon, FileChartColumnIcon, FolderOpenIcon, LayoutTemplateIcon, ListTreeIcon, PresentationIcon, SaveIcon } from "lucide-react";
import { AddDialog, EntityMenu, withToast } from "@/components/entity";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { EntityPicker, Empty, ScrollPage, Section, Workspace } from "@/components/workspace";
import { useLast } from "@/lib/last";
import { patch, put, remove, useCollection } from "@/lib/store";
import { cn } from "@/lib/utils";
import { errorText } from "@/services/ms";
import { warmOffice } from "@/services/office";
import { stopCapture } from "./capture";
import { download, forgetDeck, keepDeckFile, keepHandle, pickLocalFile, readDeckFile, readSharePoint, readTemplate, saveLocal, saveSharePoint, undoDeckFile, type PresentationEntry } from "./files";
import { presenter } from "./module";
import { useSlides } from "./use-slides";
import { useRendered, type Rendered } from "./rendered";
import { Deck } from "./pptx/deck";
import { slideTypes, type SlideType } from "./pptx/template";
import { presenterAgent } from "./tools";

/** The template deck and its slide types (Settings › Slides), read once. */
function useTemplate() {
  const [template, setTemplate] = useState<{ deck: Deck; types: SlideType[] }>();
  useEffect(() => {
    void readTemplate().then(async (t) => {
      if (!t) return;
      const deck = await Deck.open(t.bytes);
      setTemplate({ deck, types: await slideTypes(deck) });
    }, () => undefined);
  }, []);
  return template;
}

export default function Presenter() {
  const { deck: deckId, tab } = useSearch({ from: "/slides" });
  const navigate = useNavigate({ from: "/slides" });
  useLast("slides:deck", deckId, (deck) => navigate({ search: { deck }, replace: true }));
  const entries = useCollection("presentations");
  const entry = entries.find((d) => d.id === deckId);
  const [deck, setDeck] = useState<Deck>();
  const [error, setError] = useState<string>();
  const [version, setVersion] = useState(0);
  const [index, setIndex] = useState(1);
  const [adding, setAdding] = useState(false);
  const template = useTemplate();
  const slides = useSlides(deck, version);
  const rendered = useRendered(deck, version, index, entry?.id);

  const load = useCallback(async () => {
    setDeck(undefined);
    setError(undefined);
    if (!deckId) return;
    try {
      const bytes = await readDeckFile(deckId);
      if (!bytes) throw new Error("This deck's file is not in this browser any more: open it again.");
      setDeck(await Deck.open(bytes));
      setIndex(1);
    } catch (e) {
      setError(errorText(e));
    }
  }, [deckId]);
  useEffect(() => void load(), [load]);
  useEffect(() => () => stopCapture(), []);
  // The preview engine starts while the user chooses a deck.
  useEffect(() => warmOffice(), []);

  const datasets = useCollection("datasets");
  const reports = useCollection("reports");
  const sources = { datasetIds: entry?.sources?.datasetIds ?? [], reportIds: entry?.sources?.reportIds ?? [] };
  const chosen = { datasets: datasets.filter((d) => sources.datasetIds.includes(d.id)), reports: reports.filter((r) => sources.reportIds.includes(r.id)) };
  const live = useRef({ deck, entry, index, chosen, template });
  live.current = { deck, entry, index, chosen, template };

  const agent =
    entry &&
    presenterAgent({
      deck: () => live.current.deck,
      template: () => live.current.template,
      changed: async (slide) => {
        const { deck: d, entry: e } = live.current;
        if (!d || !e) return;
        await keepDeckFile(e.id, await d.bytes());
        await patch("presentations", e.id, { unsaved: true, updatedAt: Date.now() });
        setVersion((v) => v + 1);
        if (slide) setIndex(slide);
      },
      undo: async () => {
        const e = live.current.entry;
        if (!e || !(await undoDeckFile(e.id))) return false;
        const d = await Deck.open((await readDeckFile(e.id))!);
        live.current.deck = d;
        setDeck(d);
        setVersion((v) => v + 1);
        return true;
      },
      show: setIndex,
      current: () => live.current.index,
      sources: () => live.current.chosen,
    });

  /** A new deck in the library from its file's content. */
  const create = async (name: string, source: PresentationEntry["source"], bytes: Uint8Array, handle?: unknown) => {
    await Deck.open(bytes); // refused here when it is not a deck
    const id = crypto.randomUUID();
    await keepDeckFile(id, bytes);
    if (handle) await keepHandle(id, handle);
    await put("presentations", { id, name, source, updatedAt: Date.now(), contextIds: [] });
    navigate({ search: { deck: id } });
  };
  const openLocal = async () => {
    const picked = await pickLocalFile();
    if (picked) await create(picked.file.name.replace(/\.pptx$/i, ""), { kind: "local", fileName: picked.file.name }, new Uint8Array(await picked.file.arrayBuffer()), picked.handle);
  };
  const fromTemplate = async () => {
    const t = await readTemplate();
    if (!t) throw new Error("No template deck: set one in Settings › Slides.");
    const name = prompt("Name of the new deck", "New deck")?.trim();
    if (name) await create(name, { kind: "local", fileName: `${name}.pptx` }, t.bytes);
  };
  const save = async () => {
    const { deck: d, entry: e } = live.current;
    if (!d || !e) return "";
    const blob = await d.blob();
    let done: string;
    if (e.source.kind === "sharepoint") {
      const eTag = await saveSharePoint(e.source, blob);
      await patch("presentations", e.id, { source: { ...e.source, eTag } });
      done = "Saved to SharePoint.";
    } else done = await saveLocal(e.id, e.source.fileName, blob);
    await patch("presentations", e.id, { unsaved: false });
    return done;
  };
  const toggle = (kind: keyof typeof sources, id: string, on: boolean) => entry && void patch("presentations", entry.id, { sources: { ...sources, [kind]: on ? [...sources[kind], id] : sources[kind].filter((x) => x !== id) } });
  const none = <Empty icon={PresentationIcon}>Open a deck.</Empty>;

  return (
    <Workspace
      title={presenter.label}
      icon={PresentationIcon}
      picker={
        <EntityPicker
          items={entries.map((d) => ({ id: d.id, label: d.name, hint: d.source.kind === "sharepoint" ? "SharePoint" : undefined }))}
          value={deckId}
          onChange={(deck) => navigate({ search: { deck } })}
          placeholder={entries.length ? "Choose a deck" : "Open a deck"}
          onAdd={() => void withToast("Opening…", openLocal())}
          addLabel="Open a .pptx of this computer…"
          more={[
            { label: "Open from SharePoint / OneDrive…", onSelect: () => setAdding(true) },
            { label: "New deck from the template…", onSelect: () => void withToast("Creating…", fromTemplate()) },
          ]}
        />
      }
      actions={
        <>
          {entry && (
            <Button size="sm" variant={entry.unsaved ? "default" : "outline"} onClick={() => void withToast("Saving…", save())}>
              <SaveIcon /> {entry.unsaved ? "Save" : "Saved"}
            </Button>
          )}
          {deck && entry && (
            <Button size="sm" variant="outline" onClick={() => void deck.blob().then((b) => download(b, `${entry.name}.pptx`))}>
              <DownloadIcon /> .pptx
            </Button>
          )}
          {entry?.source.kind === "sharepoint" && entry.source.webUrl && (
            <Button size="sm" variant="outline" asChild>
              <a href={entry.source.webUrl} target="_blank" rel="noreferrer">
                <ExternalLinkIcon /> PowerPoint
              </a>
            </Button>
          )}
          <EntityMenu
            kind="presentations"
            entry={entry}
            onRename={entry && (() => { const name = prompt("Name of the deck", entry.name)?.trim(); if (name) void patch("presentations", entry.id, { name }); })}
            onRefresh={entry?.source.kind === "sharepoint" ? () => void withToast("Reading the file again…", readSharePoint(entry.source.kind === "sharepoint" ? entry.source.link : "").then(async (r) => (await keepDeckFile(entry.id, new Uint8Array(r.data)), await patch("presentations", entry.id, { source: r.source, unsaved: false }), await load())), "Read again.") : undefined}
            onRemove={entry && (() => confirm(`Remove "${entry.name}" from the app? (its file is kept)`) && void Promise.all([remove("presentations", entry.id), forgetDeck(entry.id)]).then(() => navigate({ search: {} })))}
          />
          <AddDialog
            open={adding}
            onOpenChange={setAdding}
            title="Open a deck from SharePoint / OneDrive"
            description="Paste the sharing link of a .pptx: it is edited here and saved back to the same file."
            fields={[{ name: "link", label: "PowerPoint file link", placeholder: "https://…sharepoint.com/…/deck.pptx" }]}
            onSubmit={async ({ link }) => {
              const r = await readSharePoint(link);
              await create(r.name, r.source, new Uint8Array(r.data));
            }}
          />
        </>
      }
      ai={{ scope: `slides:${entry?.id}`, agent, placeholder: "Ask to update the figures, add a slide, put a Power BI visual…", empty: "Open a deck to work on it with the assistant." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      notices={[
        ...(template ? [] : [{ id: "template", level: "info" as const, title: "No template deck", body: <span>New slides can be copied from a template deck: set one in <Link to="/settings" search={{ tab: "slides" }} className="underline">Settings › Slides</Link>. Without it, the assistant uses the deck's own layouts.</span> }]),
        ...(rendered.problem ? [{ id: "engine", level: "warning" as const, title: "The slides cannot be drawn", body: <span>{rendered.problem.message}{rendered.problem.missing ? " The local preview engine is not installed: run “bun install”, then start the app again." : ""} Your deck itself is not affected: the assistant and the list of shapes keep working. In Settings › Power BI, “Presenter” lets Microsoft Office draw the slides.</span> }] : []),
      ]}
      tabs={[
        {
          id: "slides",
          label: "Slides",
          icon: PresentationIcon,
          count: entry ? slides.length : undefined,
          content: error ? (
            <Empty>{error}</Empty>
          ) : deck && slides.length ? (
            <Slides deck={deck} slides={slides} index={index} onIndex={setIndex} rendered={rendered} />
          ) : entry ? (
            <Empty>Reading the deck…</Empty>
          ) : (
            <Empty
              icon={PresentationIcon}
              action={
                <span className="flex flex-wrap justify-center gap-2">
                  <Button onClick={() => void withToast("Opening…", openLocal())}>
                    <FolderOpenIcon /> Open a .pptx
                  </Button>
                  <Button variant="outline" onClick={() => setAdding(true)}>
                    <CloudIcon /> From SharePoint
                  </Button>
                  <Button variant="outline" onClick={() => void withToast("Creating…", fromTemplate())}>
                    <LayoutTemplateIcon /> From the template
                  </Button>
                </span>
              }
            >
              Any PowerPoint deck: the assistant changes its texts, adds slides of your template's types and puts Power BI visuals in it — then you save it back to its file.
            </Empty>
          ),
        },
        { id: "outline", label: "Outline", icon: ListTreeIcon, content: deck ? <Outline deck={deck} version={version} /> : none },
        {
          id: "datasets",
          label: "Datasets",
          icon: DatabaseIcon,
          count: entry ? sources.datasetIds.length : undefined,
          content: entry ? <SourceList title="Semantic models" description="The assistant computes the deck's figures on them with DAX." items={datasets.map((d) => ({ id: d.id, label: d.name, hint: d.access === "view" ? "no DAX access" : undefined, disabled: d.access === "view" }))} chosen={sources.datasetIds} onToggle={(id, on) => toggle("datasetIds", id, on)} /> : none,
        },
        {
          id: "reports",
          label: "Reports",
          icon: FileChartColumnIcon,
          count: entry ? sources.reportIds.length : undefined,
          content: entry ? <SourceList title="Power BI reports" description="Their visuals can be put on the slides as pictures (the tab is shared once to capture them)." items={reports.map((r) => ({ id: r.id, label: r.name }))} chosen={sources.reportIds} onToggle={(id, on) => toggle("reportIds", id, on)} /> : none,
        },
      ]}
    />
  );
}

/* ----------------------------------- Slides ----------------------------------- */

/**
 * The slides as the real thing draws them (Microsoft Office, else LibreOffice: rendered.ts), never
 * rebuilt by the app, with the shapes' places laid over the picture: pointing a shape — on the slide or
 * in the list under it — shows which is which, with the id the assistant uses. While a picture is not
 * there yet, an empty frame says so.
 */
function Pending({ deck, label, className }: { deck: Deck; label?: string; className?: string }) {
  return (
    <div style={{ aspectRatio: `${deck.size.w} / ${deck.size.h}` }} className={cn("flex items-center justify-center bg-muted/60 px-6 text-center text-[12.5px] text-muted-foreground ring-1 ring-border", !label && "animate-pulse", className)}>
      {label}
    </div>
  );
}

/** (see above) */
function Slides({ deck, slides, index, onIndex, rendered }: { deck: Deck; slides: ReturnType<typeof useSlides>; index: number; onIndex: (i: number) => void; rendered: Rendered }) {
  const slide = slides[Math.min(index, slides.length) - 1] ?? slides[0];
  const [shape, setShape] = useState<string>();
  const rail = useRef<HTMLDivElement>(null);
  useEffect(() => void rail.current?.querySelector(`[data-n="${slide.number}"]`)?.scrollIntoView({ block: "nearest" }), [slide.number]);
  const picture = rendered.large[slide.number] ?? rendered.small[slide.number];
  const listed = slide.shapes.filter((s) => s.role !== "decoration");
  return (
    <div className="flex min-h-0 flex-1">
      <div ref={rail} className="flex w-44 shrink-0 flex-col gap-3 overflow-y-auto border-r bg-muted/30 p-3">
        {slides.map((s) => (
          <button key={s.part} data-n={s.number} type="button" onClick={() => onIndex(s.number)} className="flex gap-2 text-left">
            <span className="w-4 pt-0.5 text-right text-[11px] text-muted-foreground tabular-nums">{s.number}</span>
            <span className={cn("block flex-1 rounded-sm", s.number === slide.number && "ring-2 ring-primary ring-offset-2")}>
              {rendered.small[s.number] ? <img src={rendered.small[s.number]} alt={`Slide ${s.number}`} className="block w-full shadow-sm ring-1 ring-border" /> : <Pending deck={deck} />}
            </span>
          </button>
        ))}
      </div>
      <div className="flex min-w-0 flex-1 flex-col gap-3 overflow-y-auto bg-muted/40 p-6">
        <div className="mx-auto w-full" style={{ maxWidth: `min(100%, calc((100vh - 16rem) * ${deck.size.w / deck.size.h}))` }}>
          {picture ? (
            <div className="relative overflow-hidden rounded-md shadow-sm ring-1 ring-border">
              <img src={picture} alt={`Slide ${slide.number}`} className={cn("block w-full transition-opacity", rendered.busy && "opacity-70")} />
              {listed.map((s) =>
                s.box ? (
                  <span
                    key={s.id}
                    title={`#${s.id} ${s.name}`}
                    onMouseEnter={() => setShape(s.id)}
                    onMouseLeave={() => setShape(undefined)}
                    style={{ left: `${(s.box.x / deck.size.w) * 100}%`, top: `${(s.box.y / deck.size.h) * 100}%`, width: `${(s.box.w / deck.size.w) * 100}%`, height: `${(s.box.h / deck.size.h) * 100}%` }}
                    className={cn("absolute rounded-xs", shape === s.id && "bg-primary/10 outline-2 outline-primary")}
                  >
                    {shape === s.id && <span className="absolute -top-5 left-0 rounded bg-primary px-1.5 text-[11px] font-medium whitespace-nowrap text-primary-foreground">#{s.id}</span>}
                  </span>
                ) : null,
              )}
            </div>
          ) : (
            <Pending deck={deck} className="rounded-md" label={rendered.problem ? rendered.problem.message : "Drawing the slide…"} />
          )}
        </div>
        <div className="mx-auto flex w-full max-w-3xl items-center gap-2 text-[12.5px] text-muted-foreground">
          <Button size="icon" variant="ghost" className="size-8" disabled={slide.number <= 1} onClick={() => onIndex(slide.number - 1)} aria-label="Previous slide">
            <ChevronLeftIcon />
          </Button>
          <span className="tabular-nums">
            {slide.number} / {slides.length}
          </span>
          <Button size="icon" variant="ghost" className="size-8" disabled={slide.number >= slides.length} onClick={() => onIndex(slide.number + 1)} aria-label="Next slide">
            <ChevronRightIcon />
          </Button>
          <span className="truncate">Layout "{slide.layout}"</span>
          {rendered.busy && <span className="ml-auto">Drawing…</span>}
          {rendered.source && !rendered.busy && <span className="ml-auto">Drawn by {rendered.source === "office" ? "Microsoft Office" : "LibreOffice (close to PowerPoint)"}</span>}
        </div>
        <div className="mx-auto flex w-full max-w-3xl flex-wrap gap-1.5">
          {listed.map((s) => (
            <button key={s.id} type="button" onMouseEnter={() => setShape(s.id)} onMouseLeave={() => setShape(undefined)} className={cn("chip h-7 text-[12px]", shape === s.id && "border-primary")}>
              <span className="text-muted-foreground">#{s.id}</span> {s.role === "frame" ? "frame" : s.name}
            </button>
          ))}
        </div>
        {slide.notes && <p className="mx-auto w-full max-w-3xl rounded-md border bg-card p-3 text-[12.5px] whitespace-pre-wrap text-muted-foreground">{slide.notes}</p>}
      </div>
    </div>
  );
}

function Outline({ deck, version }: { deck: Deck; version: number }) {
  const [text, setText] = useState("");
  useEffect(() => void deck.outline(true).then(setText), [deck, version]);
  return (
    <div className="min-h-0 flex-1 overflow-auto p-6">
      <p className="mb-3 text-[12.5px] text-muted-foreground">The deck as the assistant reads it: each shape with its id, place, size and room for text.</p>
      <pre className="rounded-lg border bg-card p-4 text-[12px] leading-relaxed whitespace-pre-wrap">{text}</pre>
    </div>
  );
}

/* ---------------------------------- Sources ---------------------------------- */

/** The items a deck is written from (none by default: the assistant then gets no tool for them). */
function SourceList({ title, description, items, chosen, onToggle }: { title: string; description: string; items: { id: string; label: string; hint?: string; disabled?: boolean }[]; chosen: string[]; onToggle: (id: string, on: boolean) => void }) {
  return (
    <ScrollPage>
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
    </ScrollPage>
  );
}
