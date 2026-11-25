/**
 * Presenter: decks made from the company template (markdown, its layouts and Power BI views) and imported
 * PowerPoints changed in place (custom slides, plus new slides from the template). Tabs: Edit, Present
 * (full screen, Power BI live), Template, and the reports and models the assistant may use.
 */
import { useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, CheckIcon, DatabaseIcon, DownloadIcon, FilePenLineIcon, LayoutTemplateIcon, Loader2Icon, PencilLineIcon, PlayIcon, PlusIcon, PresentationIcon, SparklesIcon, Undo2Icon, UploadIcon } from "lucide-react";
import { toast } from "sonner";
import { joinSlides, newDeck, slideSources } from "@shared/deck";
import type { FileRef } from "@shared/types";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { PickerMenu, AddDialog } from "@/components/pickers";
import { Empty } from "@/components/common";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { useSlides } from "@/features/presenter/use-slides";
import { Present } from "@/features/presenter/present";
import { EditDeck, EditPptx } from "@/features/presenter/edit";
import { TemplateGallery } from "@/features/presenter/template-gallery";
import { useChat } from "@/lib/chat";
import { screenSearch, useItemParam } from "@/lib/params";
import { api, download, errorText, fileUrl, post } from "@/lib/api";
import { withStructure } from "@/lib/items";
import { keepFile, keepServerFile, remove, save, useItems, type Stored } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/presenter")({ component: Presenter, validateSearch: screenSearch });

type Deck = Stored["deck"];
type Pptx = Stored["pptx"];
type Doc = ({ kind: "deck" } & Deck) | ({ kind: "pptx" } & Pptx);
type TabId = "edit" | "present" | "template" | "reports" | "models";
const PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

/** "New presentation": from the template, or an imported PowerPoint. */
function NewDialog({ open, onOpenChange, onCreated }: { open: boolean; onOpenChange: (o: boolean) => void; onCreated: (key: string) => void }) {
  const [title, setTitle] = useState("New presentation");
  const [busy, setBusy] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const fromTemplate = async () => {
    const d: Deck = { id: crypto.randomUUID(), title, markdown: newDeck(title), updatedAt: new Date().toISOString() };
    await save("deck", d);
    onCreated(`deck:${d.id}`);
    onOpenChange(false);
  };
  const upload = async (f: File) => {
    setBusy(true);
    try {
      const ref = await keepFile(f, f.name);
      const d: Pptx = { id: crypto.randomUUID(), name: f.name.replace(/\.pptx$/i, ""), fileId: ref.id, history: [], updatedAt: new Date().toISOString() };
      await save("pptx", d);
      onCreated(`pptx:${d.id}`);
      onOpenChange(false);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-2xl rounded-3xl">
        <DialogHeader>
          <DialogTitle>New presentation</DialogTitle>
        </DialogHeader>
        <div className="grid gap-4 sm:grid-cols-2">
          <div className="flex flex-col gap-3 rounded-2xl border p-5">
            <SparklesIcon className="size-6 text-primary" />
            <div className="text-[15px] font-semibold">From the template</div>
            <div className="text-[13px] text-muted-foreground">A new deck with the company's layouts (cover, agenda, a section, an example page), Power BI views and pictures. The assistant writes it with you.</div>
            <input className="input" value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Title" />
            <Button variant="dark" onClick={fromTemplate} disabled={!title.trim()}><PlusIcon /> Create</Button>
          </div>
          <div className="flex flex-col gap-3 rounded-2xl border p-5">
            <FilePenLineIcon className="size-6 text-primary" />
            <div className="text-[15px] font-semibold">Import a PowerPoint</div>
            <div className="text-[13px] text-muted-foreground">An existing .pptx changed in place: texts, tables, pictures (Power BI screenshots refreshed), and new slides from the template. Its design stays.</div>
            <input ref={input} type="file" accept=".pptx" hidden onChange={(e) => e.target.files?.[0] && void upload(e.target.files[0])} />
            <div className="flex-1" />
            <Button variant="dark" onClick={() => input.current?.click()} disabled={busy}>{busy ? <Loader2Icon className="animate-spin" /> : <UploadIcon />} Choose a file</Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** The reports or models the assistant may use for this presentation (checked), with "Add". */
function Sources({ kind, ids, setIds }: { kind: "report" | "dataset"; ids: string[]; setIds: (v: string[]) => void }) {
  const items = (useItems(kind) ?? []) as { id: string; name: string }[];
  const [adding, setAdding] = useState(false);
  const Icon = kind === "report" ? BarChart3Icon : DatabaseIcon;
  return (
    <div className="mx-auto flex max-w-3xl flex-col gap-3 p-6">
      <div className="flex items-center gap-3">
        <div className="flex-1 text-[13px] text-muted-foreground">{kind === "report" ? "Reports whose pages and visuals the assistant may put in the slides (live in the presentation, pictures in the PowerPoint)." : "Semantic models the assistant may query (DAX) for the figures of the slides."}</div>
        <Button variant="outline" size="sm" onClick={() => setAdding(true)}><PlusIcon /> {kind === "report" ? "Add a report" : "Add a model"}</Button>
      </div>
      {items.map((r) => {
        const on = ids.includes(r.id);
        return (
          <button key={r.id} onClick={() => setIds(on ? ids.filter((x) => x !== r.id) : [...ids, r.id])} className={cn("flex items-center gap-3 rounded-2xl border px-4 py-3 text-left", on && "border-foreground/40 bg-muted/50")}>
            <span className={cn("flex size-5 items-center justify-center rounded-md border", on && "border-foreground bg-foreground text-background")}>{on && <CheckIcon className="size-3.5" />}</span>
            <Icon className="size-4 text-muted-foreground" />
            <span className="flex-1 text-[14px]">{r.name}</span>
          </button>
        );
      })}
      {!items.length && <Empty title={kind === "report" ? "No report yet" : "No model yet"} />}
      <AddDialog kind={kind} open={adding} onOpenChange={setAdding} onAdded={(id) => setIds([...new Set([...ids, id])])} />
    </div>
  );
}

function Presenter() {
  const decks = useItems("deck") ?? [];
  const files = useItems("pptx") ?? [];
  const reports = useItems("report") ?? [];
  const datasets = useItems("dataset") ?? [];
  const docs: Doc[] = [...decks.map((d) => ({ ...d, kind: "deck" as const })), ...files.map((d) => ({ ...d, kind: "pptx" as const }))].sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  const [key, setKey] = useItemParam("presenter:current");
  const doc = docs.find((d) => `${d.kind}:${d.id}` === key) ?? docs[0];
  const [tab, setTab] = useState<TabId>("edit");
  const [creating, setCreating] = useState(false);
  const [busy, setBusy] = useState(false);
  const docRef = useRef(doc);
  docRef.current = doc;
  const isDeck = doc?.kind === "deck";
  const slides = useSlides(doc ? (doc.kind === "deck" ? { markdown: doc.markdown } : { fileId: doc.fileId }) : undefined, doc ? (doc.kind === "deck" ? `md:${doc.id}:${doc.updatedAt}` : `file:${doc.fileId}`) : "none");
  const update = (d: Doc) => {
    const { kind, ...rest } = d;
    const next = { ...rest, updatedAt: new Date().toISOString() };
    docRef.current = { ...next, kind } as Doc;
    return save(kind, next as never);
  };
  const name = (d: Doc) => (d.kind === "deck" ? d.title : d.name);
  const chat = useChat({
    screen: "presenter",
    item: doc ? `${doc.kind}:${doc.id}` : "none",
    scope: async () => {
      const d = docRef.current!;
      const refs = await Promise.all(reports.filter((r) => d.reportIds?.includes(r.id)).map(async (r) => (await withStructure(r))!));
      const models = datasets.filter((x) => d.datasetIds?.includes(x.id));
      return d.kind === "deck"
        ? { mode: "create" as const, doc: d.markdown, references: refs, datasets: models }
        : { mode: "edit" as const, doc: { fileId: d.fileId, name: `${d.name}.pptx` }, files: [{ id: d.fileId, name: `${d.name}.pptx`, type: PPTX, size: 0 } as FileRef], references: refs, datasets: models };
    },
    onEffect: (effect, data) => {
      const d = docRef.current;
      if (effect !== "doc" || !d) return;
      if (d.kind === "deck" && typeof data === "string") void update({ ...d, markdown: data, title: data.match(/^title:\s*(.+)$/m)?.[1]?.trim() ?? d.title });
      if (d.kind === "pptx" && (data as { fileId?: string })?.fileId) {
        const fileId = (data as { fileId: string }).fileId;
        void update({ ...d, fileId, history: [...d.history, d.fileId] });
        void keepServerFile({ id: fileId, name: `${d.name}.pptx`, type: PPTX, size: 0 });
      }
    },
  });
  const exportPptx = async () => {
    if (!doc) return;
    if (doc.kind === "pptx") return void Object.assign(document.createElement("a"), { href: fileUrl(doc.fileId), download: `${doc.name}.pptx` }).click();
    setBusy(true);
    try {
      download(await api<Blob>("/deck/pptx", { json: { markdown: doc.markdown, name: doc.title } }), `${doc.title}.pptx`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  /** A template layout added at the end: to the deck's markdown, or as a slide of the imported PowerPoint. */
  const addTemplate = async (markdown: string) => {
    const d = docRef.current;
    if (!d) return;
    if (d.kind === "deck") {
      const { head, slides: list } = slideSources(d.markdown);
      await update({ ...d, markdown: joinSlides(head, [...list, markdown]) });
    } else {
      const ref = await post<FileRef>("/pptx/add-template", { fileId: d.fileId, name: `${d.name}.pptx`, markdown, after: slides.pics?.length ?? 999 });
      await update({ ...d, fileId: ref.id, history: [...d.history, d.fileId] });
    }
    setTab("edit");
  };
  return (
    <Workspace
      screen="presenter"
      title="Presenter"
      item={
        <PickerMenu
          items={docs}
          current={doc}
          name={(d) => `${name(d)}${d.kind === "pptx" ? " (.pptx)" : ""}`}
          onSelect={(id) => setKey(docs.find((d) => d.id === id) ? `${docs.find((d) => d.id === id)!.kind}:${id}` : "")}
          onRename={(d, n) => void update(d.kind === "deck" ? { ...d, title: n } : { ...d, name: n })}
          onDelete={(id) => void remove(docs.find((d) => d.id === id)?.kind ?? "deck", id)}
          addLabel="New presentation…"
          onAdd={() => setCreating(true)}
          empty="No presentation"
        />
      }
      actions={
        doc && (
          <>
            {doc.kind === "pptx" && <Button variant="outline" size="sm" onClick={() => void update({ ...doc, fileId: doc.history.at(-1)!, history: doc.history.slice(0, -1) })} disabled={!doc.history.length}><Undo2Icon /> Undo</Button>}
            <Button variant="dark" size="sm" onClick={exportPptx} disabled={busy}>{busy ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} PowerPoint</Button>
          </>
        )
      }
      tabs={[
        { id: "edit", label: "Edit", icon: <PencilLineIcon /> },
        { id: "present", label: "Present", icon: <PlayIcon /> },
        { id: "template", label: "Template", icon: <LayoutTemplateIcon /> },
        { id: "reports", label: "Reports", icon: <BarChart3Icon />, count: doc?.reportIds?.length || undefined },
        { id: "models", label: "Models", icon: <DatabaseIcon />, count: doc?.datasetIds?.length || undefined },
      ]}
      tab={tab}
      onTab={setTab}
      scroll={tab !== "edit"}
      chat={<ChatPanel chat={chat} screen="presenter" itemLabel={(id) => { const d = docs.find((x) => `${x.kind}:${x.id}` === id); return d && name(d); }} disabled={doc ? undefined : "Create or import a presentation first"} placeholder={isDeck ? "Describe the deck, or the change to make…" : "What should change in the PowerPoint?"} suggestions={isDeck ? ["Make a 6-slide review with the key figures and Power BI visuals", "Add a slide comparing two regions with the same visual", "Shorten every slide to one key message"] : ["Refresh every Power BI screenshot with the current data", "Add a key figures slide from the template after slide 2", "Fix the typos and shorten the long texts"]} />}
    >
      {!doc ? (
        <Empty icon={<PresentationIcon />} title="No presentation yet">
          Create a deck from the template, or import a PowerPoint to update.
          <div><Button variant="dark" size="sm" className="mt-3" onClick={() => setCreating(true)}><PlusIcon /> New presentation</Button></div>
        </Empty>
      ) : tab === "edit" ? (
        doc.kind === "deck" ? (
          // Typing keeps the version (no new drawing at each key): "Refresh the preview" draws it.
          <EditDeck markdown={doc.markdown} onChange={(markdown) => { const { kind: _kind, ...deck } = doc; void save("deck", { ...deck, markdown }); }} slides={slides} />
        ) : <EditPptx fileId={doc.fileId} slides={slides} />
      ) : tab === "present" ? (
        <Present slides={slides} markdown={doc.kind === "deck" ? doc.markdown : undefined} />
      ) : tab === "template" ? (
        <TemplateGallery onAdd={addTemplate} />
      ) : (
        <Sources kind={tab === "reports" ? "report" : "dataset"} ids={(tab === "reports" ? doc.reportIds : doc.datasetIds) ?? []} setIds={(ids) => void update({ ...doc, ...(tab === "reports" ? { reportIds: ids } : { datasetIds: ids }) })} />
      )}
      <NewDialog open={creating} onOpenChange={setCreating} onCreated={setKey} />
    </Workspace>
  );
}
