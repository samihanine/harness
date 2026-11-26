/**
 * Builder: the user drops a .pbix; it is imported privately into their My workspace and shown embedded, as Power
 * BI draws it; the assistant changes its definition (every option of Power BI Desktop: pages, visuals,
 * formatting, report measures — never the model); the user downloads the .pbix with the changes, its
 * connection to the model kept.
 */
import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, BlocksIcon, DatabaseIcon, DownloadIcon, FileJsonIcon, LibraryIcon, Loader2Icon, UploadIcon } from "lucide-react";
import { toast } from "sonner";
import type { ReportRef } from "@shared/types";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { CatalogChecklist } from "@/components/pickers";
import { DocPicker } from "@/components/doc-picker";
import { Empty, ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { ReportFrame } from "@/features/views/report-view";
import { ModelView } from "@/features/views/model-view";
import { StructureView } from "@/features/views/structure-view";
import { captureView, showView, useEmbed, type Report } from "@/features/powerbi/embed";
import { reportActions } from "@/tools/browser/report-actions";
import { useChat } from "@/lib/chat";
import { screenSearch, useItemParam } from "@/lib/params";
import { download, errorText, fileUrl, post } from "@/lib/api";
import { item, keepFile, keepServerFile, remove, save, useItems, type Stored } from "@/lib/store";
import { reportStructure, withStructure } from "@/lib/items";
import { openReport, useLibrary } from "@/lib/library";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/builder")({ component: Builder, validateSearch: screenSearch });

type TabId = "report" | "structure" | "model" | "references";
type Doc = Stored["pbix"];
type Opened = ReportRef & { format?: string; live?: boolean };

/** Imports a .pbix (or its latest version) into Power BI: the copy the app shows and the assistant changes. */
async function openCopy(d: Doc) {
  const r = await post<Opened>("/pbix/open", { fileId: d.fileId, name: d.name, key: d.id });
  const next: Doc = { ...d, reportId: r.id, datasetId: r.datasetId, datasetGroupId: r.datasetGroupId, format: r.format, live: r.live, updatedAt: new Date().toISOString() };
  await save("pbix", next);
  // The copy as a report of this browser: its structure kept like the others'.
  await save("report", { id: r.id, name: d.name, datasetId: r.datasetId, datasetGroupId: r.datasetGroupId, access: "edit", addedAt: Date.now() });
  void reportStructure({ id: r.id, name: d.name }, true).catch(() => undefined);
  return next;
}

function Drop({ onFiles, busy }: { onFiles: (f: FileList | File[]) => void; busy: boolean }) {
  const input = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  return (
    <div
      onDragOver={(e) => (e.preventDefault(), setOver(true))}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => (e.preventDefault(), setOver(false), e.dataTransfer.files.length && onFiles(e.dataTransfer.files))}
      className={cn("m-6 flex flex-1 flex-col items-center justify-center gap-3 rounded-3xl border-2 border-dashed p-10 text-center", over ? "border-primary bg-primary/5" : "border-border")}
    >
      <BlocksIcon className="size-8 text-muted-foreground" />
      <div className="text-[16px] font-semibold">Drop a Power BI file (.pbix)</div>
      <div className="max-w-md text-[13px] text-muted-foreground">A report connected to a semantic model (the Build permission is enough). It is shown here as Power BI draws it; the assistant changes it with every option of Power BI Desktop; you download it with its connection kept.</div>
      <input ref={input} type="file" hidden accept=".pbix" onChange={(e) => e.target.files?.length && onFiles(e.target.files)} />
      <Button variant="dark" size="sm" disabled={busy} onClick={() => input.current?.click()}>{busy ? <Loader2Icon className="animate-spin" /> : <UploadIcon />} Choose a .pbix</Button>
    </div>
  );
}

function Builder() {
  const docs = (useItems("pbix") ?? []).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  const [id, setId] = useItemParam("builder:current");
  const doc = docs.find((d) => d.id === id) ?? docs[0];
  const [tab, setTab] = useState<TabId>("report");
  const [reload, setReload] = useState(0);
  const [busy, setBusy] = useState<string>();
  const [error, setError] = useState<string>();
  const report: ReportRef | undefined = useMemo(() => (doc?.reportId ? { id: doc.reportId, name: doc.name, datasetId: doc.datasetId, datasetGroupId: doc.datasetGroupId, access: "edit" } : undefined), [doc?.reportId, doc?.name, doc?.datasetId, doc?.datasetGroupId]);
  const embed = useEmbed(report, reload);
  const reportRef = useRef<Report | undefined>(undefined);
  reportRef.current = embed.report;
  const { entries } = useLibrary("report");
  const referenceIds = doc?.referenceIds ?? [];

  /** The open item's latest version saved as a file (after each save of the assistant: the .pbix to download). */
  const snapshot = async (d: Doc) => {
    if (!d.reportId) return d;
    const f = await post<{ id: string; name: string }>("/pbix/save", { fileId: d.fileId, reportId: d.reportId, name: d.name });
    await keepServerFile(f as never);
    const next = { ...((await item("pbix", d.id)) ?? d), fileId: f.id, updatedAt: new Date().toISOString() };
    await save("pbix", next);
    return next;
  };

  const add = async (files: FileList | File[]) => {
    const f = Array.from(files).find((x) => /\.pbix$/i.test(x.name));
    if (!f) return void toast.error("Drop a .pbix file.");
    setBusy("open");
    setError(undefined);
    try {
      const ref = await keepFile(f, f.name);
      const d: Doc = { id: crypto.randomUUID(), name: f.name.replace(/\.pbix$/i, ""), fileId: ref.id, updatedAt: new Date().toISOString() };
      await save("pbix", d);
      setId(d.id);
      await openCopy(d);
      setTab("report");
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };

  const chat = useChat({
    screen: "builder",
    item: doc?.id ?? "none",
    scope: async () => {
      const refs = await Promise.all((entries ?? []).filter((e) => e.id && referenceIds.includes(e.id)).map(async (e) => (await withStructure(await openReport(e)))!));
      return { report: await withStructure(report), references: refs };
    },
    actions: () => reportActions(() => reportRef.current, embed.ready),
    onEffect: (name) => {
      if (name !== "reload") return;
      setReload((n) => n + 1);
      // The .pbix to download follows the changes (in the background).
      if (doc) void snapshot(doc).catch(() => undefined);
    },
    view: () => captureView(reportRef.current),
    restoreView: (state) => showView(() => reportRef.current, report!.id, embed.ready, state),
  });

  const dl = async () => {
    if (!doc) return;
    setBusy("download");
    try {
      const d = await snapshot(doc);
      const blob = await (await fetch(fileUrl(d.fileId))).blob();
      download(blob, `${d.name}.pbix`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  // The copy may be gone (removed in Power BI): imported again from the latest version.
  const reopen = async () => {
    if (!doc) return;
    setBusy("open");
    setError(undefined);
    try {
      await openCopy(doc);
      setReload((n) => n + 1);
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  const removeDoc = async (x: string) => {
    const d = docs.find((y) => y.id === x);
    if (d?.reportId) void post("/pbix/forget", { reportId: d.reportId }).catch(() => undefined);
    await remove("pbix", x);
  };

  const model = useMemo(() => (doc?.datasetId ? { id: doc.datasetId, groupId: doc.datasetGroupId } : undefined), [doc?.datasetId, doc?.datasetGroupId]);
  return (
    <Workspace
      screen="builder"
      title="Builder"
      item={<DocPicker items={docs} current={doc} name={(d) => d.name} onSelect={setId} onCreate={() => setId("")} onDelete={(x) => void removeDoc(x)} onRename={(d, name) => void save("pbix", { ...d, name })} createLabel="Open another .pbix…" empty="No report" />}
      actions={
        doc && (
          <Button variant="dark" size="sm" disabled={!!busy || !doc.reportId} onClick={() => void dl()}>
            {busy === "download" ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} Download .pbix
          </Button>
        )
      }
      tabs={doc ? [{ id: "report", label: "Report", icon: <BarChart3Icon /> }, { id: "structure", label: "Structure", icon: <FileJsonIcon /> }, { id: "model", label: "Semantic model", icon: <DatabaseIcon /> }, { id: "references", label: "References", icon: <LibraryIcon />, count: referenceIds.length || undefined }] : undefined}
      tab={tab}
      onTab={setTab}
      scroll={false}
      chat={<ChatPanel chat={chat} screen="builder" itemLabel={(x) => docs.find((d) => d.id === x)?.name} disabled={!doc ? "Drop a .pbix first" : !doc.reportId ? "Opening the report…" : undefined} placeholder="Describe the page to build, or drop a screenshot to redo…" suggestions={["Build an overview page with the key figures", "Redo the page of this screenshot (attach it)", "Make this page look like my reference report"]} />}
    >
      {!doc || id === "" ? (
        <div className="flex min-h-0 flex-1 flex-col">
          {error && <div className="mx-6 mt-6"><ErrorBox>{error}</ErrorBox></div>}
          <Drop onFiles={(f) => void add(f)} busy={busy === "open"} />
        </div>
      ) : (
        <>
          <div className={cn("min-h-0 flex-1 flex-col", tab === "report" ? "flex" : "hidden")}>
            {doc.live === false && <div className="mx-3 mt-3 rounded-2xl bg-muted/50 px-4 py-2 text-[12.5px]">This file holds its own model (imported data): the report is changed, its model stays as it is.</div>}
            {(error || embed.error) && (
              <div className="mx-3 mt-3 flex items-center gap-3">
                <ErrorBox className="flex-1">{error ?? embed.error}</ErrorBox>
                <Button variant="outline" size="sm" disabled={!!busy} onClick={() => void reopen()}>{busy === "open" ? <Loader2Icon className="animate-spin" /> : null} Open it again</Button>
              </div>
            )}
            {!doc.reportId ? <Empty icon={<Loader2Icon className="animate-spin" />} title="Opening the report in Power BI…">The file is imported into your My workspace (private) to be shown here.</Empty> : <ReportFrame embed={embed} refOf={report} />}
          </div>
          {tab === "structure" && <div className="min-h-0 flex-1 overflow-auto"><StructureView report={report} /></div>}
          {tab === "model" && <div className="min-h-0 flex-1 overflow-auto"><ModelView refOf={model} name={doc.name} /></div>}
          {tab === "references" && <div className="min-h-0 flex-1 overflow-auto"><CatalogChecklist kind="report" ids={referenceIds} setIds={(ids) => void save("pbix", { ...doc, referenceIds: ids })} title="Reference reports" hint="The assistant takes them as inspiration (layout, visuals, style): their structure, and what they show." /></div>}
        </>
      )}
    </Workspace>
  );
}
