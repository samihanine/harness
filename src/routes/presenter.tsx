/**
 * Presenter: two separate ways. Create: a deck written in markdown, drawn with the template's layouts,
 * exported to PowerPoint. Edit: an uploaded PowerPoint changed in place (texts, pictures — e.g. Power BI
 * screenshots refreshed —, slides), each change a version the user can go back to.
 */
import { useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BookOpenTextIcon, DownloadIcon, FilePenLineIcon, ImagesIcon, Loader2Icon, PencilLineIcon, PresentationIcon, SparklesIcon, Undo2Icon, UploadIcon } from "lucide-react";
import { toast } from "sonner";
import { newDeck } from "@shared/deck";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { DocPicker } from "@/components/doc-picker";
import { ContextLinks } from "@/components/context-links";
import { Empty, Pills } from "@/components/common";
import { Button } from "@/components/ui/button";
import { SlidesPreview } from "@/features/presenter/slides-preview";
import { DeckEditor } from "@/features/presenter/deck-editor";
import { useChat } from "@/lib/chat";
import { api, download, errorText, fileUrl } from "@/lib/api";
import { contextsOf, keepFile, keepServerFile, remove, save, useItems, useLocal, type Stored } from "@/lib/store";

export const Route = createFileRoute("/presenter")({ component: Presenter });

function Presenter() {
  const [mode, setMode] = useLocal<"create" | "edit">("presenter:mode", "create");
  return mode === "create" ? <CreateMode setMode={setMode} /> : <EditMode setMode={setMode} />;
}

const ModeSwitch = ({ mode, setMode }: { mode: "create" | "edit"; setMode: (m: "create" | "edit") => void }) => (
  <Pills value={mode} onChange={setMode} options={[{ id: "create", label: "Create from template", icon: <SparklesIcon /> }, { id: "edit", label: "Edit a PowerPoint", icon: <FilePenLineIcon /> }]} />
);

function CreateMode({ setMode }: { setMode: (m: "create" | "edit") => void }) {
  const decks = useItems("deck") ?? [];
  const datasets = useItems("dataset") ?? [];
  const [id, setId] = useLocal("presenter:deck", "");
  const deck = decks.find((d) => d.id === id) ?? decks[0];
  const [tab, setTab] = useState<"preview" | "edit" | "context">("preview");
  const [busy, setBusy] = useState(false);
  const deckRef = useRef(deck);
  deckRef.current = deck;
  const create = async () => {
    const title = `Presentation ${decks.length + 1}`;
    const d: Stored["deck"] = { id: crypto.randomUUID(), title, markdown: newDeck(title), updatedAt: new Date().toISOString() };
    await save("deck", d);
    setId(d.id);
  };
  const itemKey = `deck:${deck?.id ?? "none"}`;
  const chat = useChat({
    screen: "presenter",
    item: `create:${deck?.id ?? "none"}`,
    scope: () => ({ mode: "create", doc: deckRef.current?.markdown, datasets, contexts: contextsOf(itemKey) }),
    onEffect: (name, data) => {
      if (name === "doc" && deckRef.current && typeof data === "string") {
        const next = { ...deckRef.current, markdown: data, title: data.match(/^title:\s*(.+)$/m)?.[1]?.trim() ?? deckRef.current.title, updatedAt: new Date().toISOString() };
        deckRef.current = next;
        void save("deck", next);
        setTab("preview");
      }
    },
  });
  const exportPptx = async () => {
    if (!deck) return;
    setBusy(true);
    try {
      download(await api<Blob>("/deck/pptx", { json: { markdown: deck.markdown, name: deck.title } }), `${deck.title}.pptx`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <Workspace
      screen="presenter"
      title="Presenter"
      item={<DocPicker items={decks} current={deck} name={(d) => d.title} onSelect={setId} onCreate={create} onDelete={(x) => void remove("deck", x)} createLabel="New deck (template)" empty="No deck" />}
      actions={
        <>
          {deck && (
            <Button variant="dark" size="sm" onClick={exportPptx} disabled={busy}>
              {busy ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} PowerPoint
            </Button>
          )}
        </>
      }
      tabs={[{ id: "preview", label: "Preview", icon: <ImagesIcon /> }, { id: "edit", label: "Edit", icon: <PencilLineIcon /> }, { id: "context", label: "Context", icon: <BookOpenTextIcon /> }]}
      tab={tab}
      onTab={setTab}
      aside={<ModeSwitch mode="create" setMode={setMode} />}
      scroll={tab !== "edit"}
      chat={<ChatPanel chat={chat} screen="presenter" disabled={deck ? undefined : "Create a deck first"} placeholder="Describe the deck, or the change to make…" suggestions={["Make a 6-slide review of the revenue with the key figures and Power BI visuals", "Add a slide comparing the regions in a table", "Shorten every slide to one key message"]} />}
    >
      {!deck ? (
        <Empty icon={<PresentationIcon />} title="No deck yet">
          A new deck starts from the template (cover, agenda, a section and an example page).
          <div><Button variant="dark" size="sm" className="mt-3" onClick={create}>New deck</Button></div>
        </Empty>
      ) : tab === "preview" ? (
        <SlidesPreview source={{ markdown: deck.markdown }} cacheKey={`md:${deck.id}:${deck.markdown.length}:${deck.updatedAt}`} />
      ) : tab === "edit" ? (
        <DeckEditor markdown={deck.markdown} onChange={(markdown) => void save("deck", { ...deck, markdown, updatedAt: new Date().toISOString() })} />
      ) : (
        <ContextLinks itemKey={itemKey} title={deck.title} />
      )}
    </Workspace>
  );
}

function EditMode({ setMode }: { setMode: (m: "create" | "edit") => void }) {
  const files = useItems("pptx") ?? [];
  const datasets = useItems("dataset") ?? [];
  const [id, setId] = useLocal("presenter:pptx", "");
  const doc = files.find((d) => d.id === id) ?? files[0];
  const [tab, setTab] = useState<"slides" | "context">("slides");
  const [uploading, setUploading] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const docRef = useRef(doc);
  docRef.current = doc;
  const upload = async (f: File) => {
    setUploading(true);
    try {
      const ref = await keepFile(f, f.name);
      const d: Stored["pptx"] = { id: crypto.randomUUID(), name: f.name.replace(/\.pptx$/i, ""), fileId: ref.id, history: [], updatedAt: new Date().toISOString() };
      await save("pptx", d);
      setId(d.id);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setUploading(false);
    }
  };
  const itemKey = `pptx:${doc?.id ?? "none"}`;
  const chat = useChat({
    screen: "presenter",
    item: `edit:${doc?.id ?? "none"}`,
    scope: () => ({ mode: "edit", doc: docRef.current && { fileId: docRef.current.fileId, name: `${docRef.current.name}.pptx` }, files: docRef.current ? [{ id: docRef.current.fileId, name: `${docRef.current.name}.pptx`, type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", size: 0 }] : [], datasets, contexts: contextsOf(itemKey) }),
    onEffect: (name, data) => {
      const d = docRef.current;
      if (name === "doc" && d && (data as { fileId?: string })?.fileId) {
        const fileId = (data as { fileId: string }).fileId;
        const next = { ...d, fileId, history: [...d.history, d.fileId], updatedAt: new Date().toISOString() };
        docRef.current = next;
        void save("pptx", next);
        void keepServerFile({ id: fileId, name: `${d.name}.pptx`, type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", size: 0 });
      }
    },
  });
  const undo = () => doc && doc.history.length && void save("pptx", { ...doc, fileId: doc.history.at(-1)!, history: doc.history.slice(0, -1), updatedAt: new Date().toISOString() });
  return (
    <Workspace
      screen="presenter"
      title="Presenter"
      item={<DocPicker items={files} current={doc} name={(d) => d.name} onSelect={setId} onCreate={() => input.current?.click()} onDelete={(x) => void remove("pptx", x)} createLabel="Upload a PowerPoint…" empty="Upload a PowerPoint" />}
      actions={
        <>
          <input ref={input} type="file" accept=".pptx" hidden onChange={(e) => e.target.files?.[0] && void upload(e.target.files[0])} />
          {doc && (
            <>
              <Button variant="outline" size="sm" onClick={undo} disabled={!doc.history.length}><Undo2Icon /> Undo</Button>
              <Button variant="dark" size="sm" asChild>
                <a href={fileUrl(doc.fileId)} download={`${doc.name}.pptx`}><DownloadIcon /> PowerPoint</a>
              </Button>
            </>
          )}
        </>
      }
      tabs={[{ id: "slides", label: "Slides", icon: <ImagesIcon /> }, { id: "context", label: "Context", icon: <BookOpenTextIcon /> }]}
      tab={tab}
      onTab={setTab}
      aside={<ModeSwitch mode="edit" setMode={setMode} />}
      chat={<ChatPanel chat={chat} screen="presenter" disabled={doc ? undefined : "Upload a PowerPoint first"} placeholder="What should change in the deck?" suggestions={["Refresh every Power BI screenshot with the current data", "Update the figures of slide 3 for this year", "Fix the typos and shorten the long texts"]} />}
    >
      {!doc ? (
        <Empty icon={<UploadIcon />} title="Upload a PowerPoint">
          The assistant changes it in place: texts, tables, pictures (Power BI screenshots refreshed from their report), slides. Its design stays as it is.
          <div><Button variant="dark" size="sm" className="mt-3" onClick={() => input.current?.click()} disabled={uploading}>{uploading ? <Loader2Icon className="animate-spin" /> : <UploadIcon />} Upload</Button></div>
        </Empty>
      ) : tab === "slides" ? (
        <SlidesPreview source={{ fileId: doc.fileId }} cacheKey={`file:${doc.fileId}`} />
      ) : (
        <ContextLinks itemKey={itemKey} title={doc.name} />
      )}
    </Workspace>
  );
}

