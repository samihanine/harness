/** Builder: the assistant builds and changes a Power BI report (its PBIR definition), shown live on the right. */
import { useEffect, useMemo, useRef, useState } from "react";
import type { ReportAccess } from "@shared/types";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, BlocksIcon, CheckIcon, DatabaseIcon, ExternalLinkIcon, FileJsonIcon, LibraryIcon, Loader2Icon, PencilRulerIcon, PlusIcon } from "lucide-react";
import { toast } from "sonner";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { AddButton, AddDialog, ItemPicker, addReport, usePicked } from "@/components/pickers";
import { Empty } from "@/components/common";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { ReportFrame } from "@/features/views/report-view";
import { ModelView } from "@/features/views/model-view";
import { StructureView } from "@/features/views/structure-view";
import { editorActions } from "@/tools/browser/editor-actions";
import { captureView, showView, useEmbed, type Report } from "@/features/powerbi/embed";
import { reportActions } from "@/tools/browser/report-actions";
import { useChat } from "@/lib/chat";
import { screenSearch } from "@/lib/params";
import { errorText, get, post } from "@/lib/api";
import { save, useItems, useLocal } from "@/lib/store";
import { withStructure } from "@/lib/items";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/builder")({ component: Builder, validateSearch: screenSearch });

type TabId = "report" | "structure" | "model" | "references";

function NewReport({ open, onOpenChange, onCreated }: { open: boolean; onOpenChange: (o: boolean) => void; onCreated: (id: string) => void }) {
  const [name, setName] = useState("New report");
  const [models, setModels] = useState<{ id: string; name: string; groupId?: string; workspace: string }[]>();
  const [model, setModel] = useState<string>();
  const [busy, setBusy] = useState(false);
  const [query, setQuery] = useState("");
  if (open && !models) void get("/powerbi/datasets").then(setModels, (e) => toast.error(errorText(e)));
  const create = async () => {
    const m = models?.find((x) => x.id === model);
    if (!m) return;
    setBusy(true);
    try {
      const r = await post<{ id: string; groupId?: string }>("/powerbi/new-report", { name, datasetId: m.id, groupId: m.groupId });
      await addReport(r);
      onCreated(r.id);
      onOpenChange(false);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-lg rounded-3xl">
        <DialogHeader>
          <DialogTitle>New report</DialogTitle>
        </DialogHeader>
        <label className="label">Name</label>
        <input className="input" value={name} onChange={(e) => setName(e.target.value)} />
        <label className="label">Semantic model</label>
        <input className="input" placeholder="Search" value={query} onChange={(e) => setQuery(e.target.value)} />
        <div className="max-h-60 overflow-y-auto rounded-xl border">
          {!models && <div className="flex items-center gap-2 p-3 text-[13px] text-muted-foreground"><Loader2Icon className="size-4 animate-spin" /> Loading…</div>}
          {models?.filter((m) => `${m.name} ${m.workspace}`.toLowerCase().includes(query.toLowerCase())).map((m) => (
            <button key={m.id} onClick={() => setModel(m.id)} className={cn("flex w-full items-center gap-2 px-3 py-2 text-left text-[13.5px] hover:bg-muted", model === m.id && "bg-muted")}>
              <CheckIcon className={cn("size-4", model === m.id ? "opacity-100" : "opacity-0")} />
              <span className="flex-1">{m.name}</span>
              <span className="text-[12px] text-muted-foreground">{m.workspace}</span>
            </button>
          ))}
        </div>
        <Button variant="dark" disabled={!model || !name.trim() || busy} onClick={create}>
          {busy ? <Loader2Icon className="animate-spin" /> : <PlusIcon />} Create
        </Button>
      </DialogContent>
    </Dialog>
  );
}

function References({ current, ids, setIds }: { current?: string; ids: string[]; setIds: (v: string[]) => void }) {
  const reports = (useItems("report") ?? []).filter((r) => r.id !== current);
  const [adding, setAdding] = useState(false);
  return (
    <div className="mx-auto flex max-w-3xl flex-col gap-4 p-6">
      <div className="flex items-center gap-3">
        <div className="flex-1">
          <div className="text-[15px] font-semibold">Reference reports</div>
          <div className="text-[13px] text-muted-foreground">The assistant takes them as inspiration (layout, visuals, style). Editable ones give it their whole definition; view-only ones, what they show.</div>
        </div>
        <Button variant="outline" size="sm" onClick={() => setAdding(true)}><PlusIcon /> Add a report</Button>
      </div>
      {reports.length === 0 && <Empty icon={<LibraryIcon />} title="No other report">Add reports to use them as references.</Empty>}
      {reports.map((r) => {
        const on = ids.includes(r.id);
        return (
          <button key={r.id} onClick={() => setIds(on ? ids.filter((x) => x !== r.id) : [...ids, r.id])} className={cn("flex items-center gap-3 rounded-2xl border px-4 py-3 text-left", on && "border-foreground/40 bg-muted/50")}>
            <span className={cn("flex size-5 items-center justify-center rounded-md border", on && "border-foreground bg-foreground text-background")}>{on && <CheckIcon className="size-3.5" />}</span>
            <span className="flex-1 text-[14px]">{r.name}</span>
            <span className="badge bg-muted text-muted-foreground">{r.access === "edit" ? "full definition" : "view only"}</span>
          </button>
        );
      })}
      <AddDialog kind="report" open={adding} onOpenChange={setAdding} onAdded={(id) => setIds([...ids, id])} />
    </div>
  );
}

function Builder() {
  const { items, current, select } = usePicked("report", "builder");
  const all = useItems("report") ?? [];
  const [tab, setTab] = useState<TabId>("report");
  const [reload, setReload] = useState(0);
  const [creating, setCreating] = useState(false);
  const [refIds, setRefIds] = useLocal<string[]>(`builder:refs:${current?.id}`, []);
  // Editor mode: the report opens in the Power BI editor (its definition cannot be written by the API).
  const editorMode = current?.access === "editor";
  const embed = useEmbed(current, reload, editorMode ? "edit" : "view");
  const reportRef = useRef<Report | undefined>(undefined);
  reportRef.current = embed.report;
  const references = all.filter((r) => refIds.includes(r.id));
  const chat = useChat({
    screen: "builder",
    item: current?.id ?? "none",
    scope: async () => ({ report: await withStructure(current), references: await Promise.all(references.map(async (r) => (await withStructure(r))!)) }),
    actions: () => ({ ...reportActions(() => reportRef.current, embed.ready), ...editorActions(() => reportRef.current, embed.ready) }),
    onEffect: (name) => name === "reload" && setReload((n) => n + 1),
    view: () => captureView(reportRef.current),
    restoreView: (state) => showView(() => reportRef.current, current!.id, embed.ready, state),
  });
  const model = useMemo(() => (current?.datasetId ? { id: current.datasetId, groupId: current.datasetGroupId } : undefined), [current?.datasetId, current?.datasetGroupId]);
  // A report kept as not editable is checked again (its rights, or what the Builder can write, may have changed).
  useEffect(() => {
    if (!current || current.access === "edit") return;
    void post<Record<string, ReportAccess>>("/powerbi/access", [{ id: current.id, groupId: current.groupId, appId: current.appId }]).then((a) => {
      if (a[current.id] === "edit") void save("report", { ...current, access: "edit" });
    }, () => undefined);
  }, [current?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Workspace
      screen="builder"
      title="Builder"
      item={<ItemPicker kind="report" items={items} current={current} onSelect={select} editable extra={(r) => (r.access === "view" ? "view only" : r.access === "editor" ? "editor mode" : undefined)} />}
      actions={
        <>
          <Button variant="dark" size="sm" onClick={() => setCreating(true)}><PlusIcon /> New report</Button>
          {current?.webUrl && (
            <Button variant="outline" size="sm" asChild>
              <a href={current.webUrl} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Power BI</a>
            </Button>
          )}
        </>
      }
      tabs={[{ id: "report", label: "Report", icon: <BarChart3Icon /> }, { id: "structure", label: "Structure", icon: <FileJsonIcon /> }, { id: "model", label: "Semantic model", icon: <DatabaseIcon /> }, { id: "references", label: "References", icon: <LibraryIcon />, count: references.length || undefined }]}
      tab={tab}
      onTab={setTab}
      scroll={false}
      chat={<ChatPanel chat={chat} screen="builder" itemLabel={(id) => all.find((r) => r.id === id)?.name} disabled={!current ? "Choose or create a report first" : current.access === "view" ? "This report is view only: choose one you can edit, or try the editor mode" : undefined} placeholder="Describe the page to build, or drop a screenshot to redo…" suggestions={["Build an overview page with the key figures", "Redo the page of this screenshot (attach it)", "Make this page look like my reference report"]} />}
    >
      {!current ? (
        <Empty icon={<BlocksIcon />} title="Choose or create a report">Pick a report you can edit, or create one on a semantic model: the assistant builds its pages.<div><AddButton kind="report" onAdded={select} editable /></div></Empty>
      ) : (
        <>
          <div className={cn("min-h-0 flex-1 flex-col", tab === "report" ? "flex" : "hidden")}>
            {current.access === "view" && (
              <div className="mx-3 mt-3 flex items-center gap-3 rounded-2xl border bg-muted/40 px-4 py-2.5 text-[13px]">
                <span className="flex-1">The API gives no definition of this report. If you can edit it in Power BI, the assistant can work in the Power BI editor instead.</span>
                <Button variant="dark" size="sm" onClick={() => void save("report", { ...current, access: "editor" })}><PencilRulerIcon /> Try the editor mode</Button>
              </div>
            )}
            {editorMode && (
              <div className="mx-3 mt-3 flex items-center gap-3 rounded-2xl bg-primary/10 px-4 py-2 text-[12.5px]">
                <span className="flex-1">{embed.error ? "Power BI refused to open this report in its editor: you cannot edit it." : "Editor mode: the assistant changes this report live in the Power BI editor, then saves it."}</span>
                {embed.error && <Button variant="outline" size="sm" onClick={() => void save("report", { ...current, access: "view" })}>Back to view only</Button>}
              </div>
            )}
            <ReportFrame embed={embed} refOf={current} />
          </div>
          {tab === "structure" && <div className="min-h-0 flex-1 overflow-auto"><StructureView report={current} /></div>}
          {tab === "model" && <div className="min-h-0 flex-1 overflow-auto"><ModelView refOf={model} name={current.name} /></div>}
          {tab === "references" && <div className="min-h-0 flex-1 overflow-auto"><References current={current.id} ids={refIds} setIds={setRefIds} /></div>}
        </>
      )}
      <NewReport open={creating} onOpenChange={setCreating} onCreated={select} />
    </Workspace>
  );
}
