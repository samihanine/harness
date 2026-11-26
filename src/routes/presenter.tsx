/**
 * Presenter: the look apart from the content. Content: the slides' markdown (template, title, text, pictures,
 * Power BI links, items), drawn at once. Visual: the visual templates of this browser (the company's to start
 * with), favorites first, changed by the assistant or by hand, exported and imported. Present: full screen with
 * the Power BI views live. Reports / Models: what the assistant may use. Downloads: PowerPoint (editable), PDF.
 * An existing PowerPoint or a picture is attached in the chat: the assistant recreates its look and content.
 */
import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, CopyIcon, DatabaseIcon, DownloadIcon, FileTextIcon, LayoutTemplateIcon, Loader2Icon, PlayIcon, PlusIcon, PresentationIcon, RotateCcwIcon, StarIcon, Trash2Icon, UploadIcon, ArrowDownIcon, ArrowUpIcon } from "lucide-react";
import { toast } from "sonner";
import { joinBlocks, parseSlides, slideBlocks, slideContent, templateOf, templateSlots, type Template } from "@shared/slides";
import { DEFAULT_TEMPLATES, exampleFor } from "@shared/slide-templates";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { CatalogChecklist } from "@/components/pickers";
import { DocPicker } from "@/components/doc-picker";
import { Empty, Pills } from "@/components/common";
import { Button } from "@/components/ui/button";
import { Present } from "@/features/presenter/present";
import { SlideView, useTemplates } from "@/features/presenter/slides";
import { useChat } from "@/lib/chat";
import { screenSearch, useItemParam } from "@/lib/params";
import { api, download, errorText } from "@/lib/api";
import { withStructure } from "@/lib/items";
import { datasetOf, openReport, useLibrary } from "@/lib/library";
import { remove, save, useItems, type Stored } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/presenter")({ component: Presenter, validateSearch: screenSearch });

type Deck = Stored["deck"];
type TabId = "content" | "visual" | "present" | "reports" | "models";

const START = (title: string) => `template: cover\ndate: ${new Date().toLocaleDateString(undefined, { month: "long", year: "numeric" })}\n---\n# ${title}\n## What this presentation is about`;

/** The Content tab: the slides on the left, the chosen one drawn large, its markdown under it (or the whole deck's). */
function Content({ deck, templates, onChange }: { deck: Deck; templates: Template[]; onChange: (markdown: string) => void }) {
  const [mode, setMode] = useState<"slide" | "deck">("slide");
  const [index, setIndex] = useState(0);
  const blocks = useMemo(() => slideBlocks(deck.markdown), [deck.markdown]);
  const slides = useMemo(() => blocks.map(slideContent), [blocks]);
  const i = Math.min(index, Math.max(0, blocks.length - 1));
  const set = (list: string[], next = i) => (onChange(joinBlocks(list)), setIndex(Math.max(0, Math.min(list.length - 1, next))));
  const c = slides[i];
  return (
    <div className="flex h-full flex-col gap-3 p-5">
      <div className="flex flex-wrap items-center gap-2">
        <Pills value={mode} onChange={setMode} options={[{ id: "slide", label: "Slide by slide" }, { id: "deck", label: "Whole content (markdown)" }]} />
        <div className="flex-1" />
        <Button variant="outline" size="sm" onClick={() => set([...blocks, "template: content\n---\n# New slide\n\n- A first point"], blocks.length)}><PlusIcon /> Slide</Button>
      </div>
      {mode === "deck" ? (
        <textarea value={deck.markdown} onChange={(e) => onChange(e.target.value)} spellCheck={false} className="min-h-0 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30" />
      ) : !c ? (
        <Empty icon={<PresentationIcon />} title="No slide yet">Ask the assistant, or add a slide.</Empty>
      ) : (
        <div className="flex min-h-0 flex-1 gap-4">
          <div className="flex w-48 shrink-0 flex-col gap-2 overflow-y-auto pr-1">
            {slides.map((s, n) => (
              <button key={n} onClick={() => setIndex(n)} className={cn("rounded-xl border-2 p-1 text-left", n === i ? "border-primary" : "border-transparent hover:border-border")}>
                <SlideView template={templateOf(templates, s.template)} content={s} className="rounded-lg border" />
                <div className="mt-1 truncate px-1 text-[11.5px] text-muted-foreground">{n + 1} · {s.title ?? s.template}</div>
              </button>
            ))}
          </div>
          <div className="flex min-w-0 flex-1 flex-col gap-3 overflow-y-auto">
            <SlideView template={templateOf(templates, c.template)} content={c} className="rounded-2xl border shadow-sm" />
            <div className="flex items-center gap-1">
              <span className="flex-1 text-[12.5px] text-muted-foreground">Slide {i + 1} · template <b>{templateOf(templates, c.template)?.name}</b> · slots: {templateSlots(templateOf(templates, c.template)).join(", ")}</span>
              <Button variant="ghost" size="icon-sm" title="Move up" disabled={i === 0} onClick={() => { const l = [...blocks]; [l[i - 1], l[i]] = [l[i], l[i - 1]]; set(l, i - 1); }}><ArrowUpIcon /></Button>
              <Button variant="ghost" size="icon-sm" title="Move down" disabled={i >= blocks.length - 1} onClick={() => { const l = [...blocks]; [l[i + 1], l[i]] = [l[i], l[i + 1]]; set(l, i + 1); }}><ArrowDownIcon /></Button>
              <Button variant="ghost" size="icon-sm" title="Duplicate" onClick={() => set([...blocks.slice(0, i + 1), blocks[i], ...blocks.slice(i + 1)], i + 1)}><CopyIcon /></Button>
              <Button variant="ghost" size="icon-sm" title="Delete" disabled={blocks.length < 2} onClick={() => set(blocks.filter((_, n) => n !== i))}><Trash2Icon /></Button>
            </div>
            <textarea value={blocks[i]} onChange={(e) => set(blocks.map((b, n) => (n === i ? e.target.value : b)))} spellCheck={false} className="min-h-44 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30" />
          </div>
        </div>
      )}
    </div>
  );
}

/** The Visual tab: the templates (favorites first), drawn with example content; favorite, edit, duplicate, export / import. */
function Visual({ templates, onUse }: { templates: Template[]; onUse: (t: Template) => void }) {
  const [open, setOpen] = useState<string>();
  const [text, setText] = useState("");
  const input = useRef<HTMLInputElement>(null);
  const exportAll = () => download(new Blob([JSON.stringify(templates, null, 2)], { type: "application/json" }), "templates.json");
  const importFile = async (f: File) => {
    try {
      const list = JSON.parse(await f.text()) as Template[];
      for (const t of Array.isArray(list) ? list : [list]) if (t?.id && Array.isArray(t.elements)) await save("template", t);
      toast.success("Templates imported");
    } catch (e) {
      toast.error(`Not a templates file: ${errorText(e)}`);
    }
  };
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex flex-wrap items-center gap-2">
        <div className="flex-1 text-[12.5px] text-muted-foreground">The look of the slides: ask the assistant for a new template or a change (also from an attached PowerPoint or picture), or edit one here. ★ favorites are the assistant's first choice.</div>
        <input ref={input} type="file" hidden accept=".json" onChange={(e) => e.target.files?.[0] && void importFile(e.target.files[0])} />
        <Button variant="outline" size="sm" onClick={() => input.current?.click()}><UploadIcon /> Import</Button>
        <Button variant="outline" size="sm" onClick={exportAll}><DownloadIcon /> Export</Button>
        <Button variant="ghost" size="sm" title="Put the company's templates back (yours are kept)" onClick={() => void Promise.all(DEFAULT_TEMPLATES.map((t) => save("template", { ...t, favorite: templates.find((x) => x.id === t.id)?.favorite }))).then(() => toast.success("Company templates restored"))}><RotateCcwIcon /> Company templates</Button>
      </div>
      <div className="grid gap-5 lg:grid-cols-2 2xl:grid-cols-3">
        {templates.map((t) => (
          <div key={t.id} className="flex flex-col overflow-hidden rounded-2xl border bg-card">
            <SlideView template={t} content={slideContent(exampleFor(t))} className="border-b" />
            <div className="flex flex-1 flex-col gap-2 p-4">
              <div className="flex items-center gap-2">
                <button title={t.favorite ? "Remove from favorites" : "Favorite: the assistant's first choice"} onClick={() => void save("template", { ...t, favorite: !t.favorite })}>
                  <StarIcon className={cn("size-4", t.favorite ? "fill-primary text-primary" : "text-muted-foreground")} />
                </button>
                <span className="flex-1 text-[14px] font-semibold">{t.name}</span>
                <span className="rounded-full bg-muted px-2 py-0.5 font-mono text-[11px]">{t.id}</span>
              </div>
              <div className="text-[12.5px] text-muted-foreground">{t.description}</div>
              {open === t.id && <textarea value={text} onChange={(e) => setText(e.target.value)} spellCheck={false} className="h-72 rounded-xl border bg-muted/40 p-3 font-mono text-[11px] outline-none" />}
              <div className="mt-auto flex flex-wrap gap-1.5 pt-1">
                {open === t.id ? (
                  <>
                    <Button variant="dark" size="sm" onClick={() => { try { const next = JSON.parse(text) as Template; void save("template", { ...next, id: t.id }); setOpen(undefined); } catch (e) { toast.error(errorText(e)); } }}>Save</Button>
                    <Button variant="ghost" size="sm" onClick={() => setOpen(undefined)}>Cancel</Button>
                  </>
                ) : (
                  <Button variant="ghost" size="sm" onClick={() => (setText(JSON.stringify(t, null, 2)), setOpen(t.id))}>Edit</Button>
                )}
                <Button variant="ghost" size="sm" onClick={() => void save("template", { ...t, id: `${t.id}-copy-${Math.random().toString(36).slice(2, 5)}`, name: `${t.name} (copy)`, builtin: false, favorite: false })}><CopyIcon /> Duplicate</Button>
                <Button variant="ghost" size="sm" onClick={() => confirm(`Remove the template "${t.name}"?`) && void remove("template", t.id)}><Trash2Icon /></Button>
                <div className="flex-1" />
                <Button variant="outline" size="sm" onClick={() => onUse(t)}><PlusIcon /> Add a slide</Button>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Presenter() {
  const decks = (useItems("deck") ?? []).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  const templates = useTemplates();
  const [id, setId] = useItemParam("presenter:current");
  const deck = decks.find((d) => d.id === id) ?? decks[0];
  const [tab, setTab] = useState<TabId>("content");
  const [busy, setBusy] = useState<string>();
  const deckRef = useRef(deck);
  deckRef.current = deck;
  const templatesRef = useRef(templates);
  templatesRef.current = templates;
  const { entries: reports } = useLibrary("report");
  const { entries: models } = useLibrary("dataset");
  const slides = useMemo(() => (deck ? parseSlides(deck.markdown) : []), [deck?.markdown]); // eslint-disable-line react-hooks/exhaustive-deps

  const update = (d: Deck) => {
    const next = { ...d, updatedAt: new Date().toISOString() };
    deckRef.current = next;
    return save("deck", next);
  };
  const create = async () => {
    const title = "New presentation";
    const d: Deck = { id: crypto.randomUUID(), title, markdown: START(title), updatedAt: new Date().toISOString() };
    await save("deck", d);
    setId(d.id);
    setTab("content");
  };
  const chat = useChat({
    screen: "presenter",
    item: deck?.id ?? "none",
    scope: async () => {
      const d = deckRef.current!;
      const refs = await Promise.all((reports ?? []).filter((e) => e.id && d.reportIds?.includes(e.id)).map(async (e) => (await withStructure(await openReport(e).catch(() => undefined))) ?? { id: e.id!, name: e.name }));
      return { doc: { title: d.title, markdown: d.markdown }, templates: templatesRef.current, references: refs, datasets: (models ?? []).filter((e) => e.id && d.datasetIds?.includes(e.id)).map(datasetOf) };
    },
    onEffect: (name, data) => {
      if (name === "deck" && deckRef.current) void update({ ...deckRef.current, markdown: data as string });
      if (name === "templates") {
        const next = data as Template[];
        for (const t of next) void save("template", t);
        for (const t of templatesRef.current) if (!next.some((x) => x.id === t.id)) void remove("template", t.id);
      }
    },
  });
  const exportAs = async (kind: "pptx" | "pdf") => {
    if (!deck) return;
    setBusy(kind);
    try {
      const blob = await api<Blob>(`/slides/${kind}`, { json: { markdown: deck.markdown, templates, title: deck.title } });
      download(blob, `${deck.title}.${kind}`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  return (
    <Workspace
      screen="presenter"
      title="Presenter"
      item={<DocPicker items={decks} current={deck} name={(d) => d.title} onSelect={setId} onCreate={create} onDelete={(x) => void remove("deck", x)} onRename={(d, title) => void update({ ...d, title })} createLabel="New presentation" empty="No presentation" />}
      actions={
        deck && (
          <>
            <Button variant="outline" size="sm" disabled={!!busy} onClick={() => void exportAs("pdf")}>{busy === "pdf" ? <Loader2Icon className="animate-spin" /> : <FileTextIcon />} PDF</Button>
            <Button variant="dark" size="sm" disabled={!!busy} onClick={() => void exportAs("pptx")}>{busy === "pptx" ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} PowerPoint</Button>
          </>
        )
      }
      tabs={[
        { id: "content", label: "Content", icon: <PresentationIcon />, count: slides.length || undefined },
        { id: "visual", label: "Visual", icon: <LayoutTemplateIcon />, count: templates.length || undefined },
        { id: "present", label: "Present", icon: <PlayIcon /> },
        { id: "reports", label: "Reports", icon: <BarChart3Icon />, count: deck?.reportIds?.length || undefined },
        { id: "models", label: "Models", icon: <DatabaseIcon />, count: deck?.datasetIds?.length || undefined },
      ]}
      tab={tab}
      onTab={setTab}
      scroll={tab !== "content"}
      chat={<ChatPanel chat={chat} screen="presenter" itemLabel={(x) => decks.find((d) => d.id === x)?.title} accept="image/*,.pptx" disabled={deck ? undefined : "Create a presentation first"} placeholder="Describe the slides, or attach a PowerPoint or a picture to redo…" suggestions={["A 6-slide review of the pipeline with key figures and Power BI visuals", "Redo this PowerPoint with our templates (attach it)", "Make a new template: a big figure on the left, three points on the right"]} />}
    >
      {!deck ? (
        <Empty icon={<PresentationIcon />} title="No presentation yet">
          Start empty, then ask the assistant — or attach an existing PowerPoint or a picture of a slide in the chat to redo it.
          <div><Button variant="dark" size="sm" className="mt-3" onClick={create}><PlusIcon /> New presentation</Button></div>
        </Empty>
      ) : tab === "content" ? (
        <Content deck={deck} templates={templates} onChange={(markdown) => void update({ ...deck, markdown })} />
      ) : tab === "visual" ? (
        <Visual templates={templates} onUse={(t) => (void update({ ...deck, markdown: joinBlocks([...slideBlocks(deck.markdown), exampleFor(t)]) }), setTab("content"), toast.success(`Slide "${t.name}" added at the end`))} />
      ) : tab === "present" ? (
        <Present slides={slides} templates={templates} />
      ) : tab === "reports" ? (
        <CatalogChecklist kind="report" ids={deck.reportIds ?? []} setIds={(reportIds) => void update({ ...deck, reportIds })} title="Reports" hint="The assistant takes its Power BI views (pages and visuals) from these reports." />
      ) : (
        <CatalogChecklist kind="dataset" ids={deck.datasetIds ?? []} setIds={(datasetIds) => void update({ ...deck, datasetIds })} title="Semantic models" hint="The assistant takes its figures from these models (DAX)." />
      )}
    </Workspace>
  );
}
