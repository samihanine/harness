/**
 * Data Scientist: an Excel workbook built by the assistant in the user's OneDrive — CUBE formulas connected to
 * the Power BI models, the uploaded Excel / CSV files as sheets, formulas, formats, conditional formats, lists,
 * tables, charts — computed by Excel itself, shown here sheet by sheet, downloaded working.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { CodeIcon, DatabaseIcon, DownloadIcon, ExternalLinkIcon, FileSpreadsheetIcon, FlaskConicalIcon, Loader2Icon, PlusIcon, RefreshCwIcon, SheetIcon, UploadIcon } from "lucide-react";
import { toast } from "sonner";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { DocPicker } from "@/components/doc-picker";
import { CatalogChecklist } from "@/components/pickers";
import { DaxConsole } from "@/features/views/dax-console";
import { Empty, ErrorBox, Pills } from "@/components/common";
import { Button } from "@/components/ui/button";
import { useChat } from "@/lib/chat";
import { screenSearch, useItemParam } from "@/lib/params";
import { download, errorText, get, post } from "@/lib/api";
import { datasetOf, useLibrary } from "@/lib/library";
import { keepFile, remove, save, useItems, type Stored } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/scientist")({ component: Scientist, validateSearch: screenSearch });

type Doc = Stored["analysis"];
type Preview = { sheets: string[]; sheet: string; html: string; total: { rows: number; cols: number }; shown: { rows: number; cols: number } };

/** The workbook's sheets as Excel computes them (first rows and columns), drawn here; refreshed after each change. */
function WorkbookView({ doc, version }: { doc: Doc; version: number }) {
  const [sheet, setSheet] = useState<string>();
  const [preview, setPreview] = useState<Preview>();
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(false);
  const load = (s = sheet) => {
    if (!doc.itemId) return;
    setBusy(true);
    setError(undefined);
    get<Preview>(`/workbook/preview?${new URLSearchParams({ itemId: doc.itemId, ...(s ? { sheet: s } : {}) })}`)
      .then((p) => (setPreview(p), setSheet(p.sheet)), (e) => setError(errorText(e)))
      .finally(() => setBusy(false));
  };
  useEffect(() => load(), [doc.itemId, version]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!doc.itemId) return <Empty icon={<Loader2Icon className="animate-spin" />} title="Creating the workbook in your OneDrive…" />;
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex flex-wrap items-center gap-2 px-5 py-2.5">
        {preview && <Pills value={preview.sheet} onChange={(s) => (setSheet(s), load(s))} options={preview.sheets.map((s) => ({ id: s, label: s }))} />}
        <div className="flex-1 text-right text-[12px] text-muted-foreground">{preview && (preview.total.rows > preview.shown.rows || preview.total.cols > preview.shown.cols) ? `${preview.total.rows.toLocaleString()} rows × ${preview.total.cols} columns: the first ${preview.shown.rows} × ${preview.shown.cols} shown` : ""}</div>
        <Button variant="ghost" size="sm" onClick={() => load()} disabled={busy}>{busy ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh</Button>
        {doc.webUrl && <Button variant="outline" size="sm" asChild><a href={doc.webUrl} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Open in Excel</a></Button>}
      </div>
      {error && <div className="px-5"><ErrorBox>{error}</ErrorBox></div>}
      <div className={cn("relative mx-3 mb-3 min-h-0 flex-1 overflow-hidden rounded-2xl border bg-white", busy && "opacity-70")}>
        {preview ? <iframe title="Workbook" srcDoc={preview.html} className="h-full w-full" /> : !error && <div className="flex h-full items-center justify-center gap-2 text-[13px] text-muted-foreground"><Loader2Icon className="size-4 animate-spin" /> Excel computes the workbook…</div>}
      </div>
    </div>
  );
}

/** The Excel / CSV files added to the workbook as sheets (Excel works on them with formulas). */
function Files({ doc, onAdded }: { doc: Doc; onAdded: () => void }) {
  const [busy, setBusy] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const add = async (files: FileList) => {
    if (!doc.itemId) return;
    setBusy(true);
    try {
      const refs = [];
      for (const f of Array.from(files)) {
        const ref = await keepFile(f, f.name);
        const r = await post<{ sheets: string[] }>("/workbook/add-file", { itemId: doc.itemId, fileId: ref.id });
        toast.success(`${f.name}: sheet${r.sheets.length > 1 ? "s" : ""} ${r.sheets.join(", ")} added`);
        refs.push(ref);
      }
      await save("analysis", { ...doc, files: [...doc.files, ...refs], updatedAt: new Date().toISOString() });
      onAdded();
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="mx-auto flex max-w-3xl flex-col gap-3 p-6">
      <div className="flex items-center gap-2">
        <div className="flex-1">
          <div className="text-[15px] font-semibold">Excel and CSV files</div>
          <div className="text-[13px] text-muted-foreground">Each file's sheets are added to the workbook: the assistant works on them with Excel formulas, even with hundreds of thousands of rows.</div>
        </div>
        <input ref={input} type="file" hidden multiple accept=".xlsx,.xlsm,.csv,.tsv,.txt" onChange={(e) => e.target.files && void add(e.target.files)} />
        <Button variant="dark" size="sm" onClick={() => input.current?.click()} disabled={busy || !doc.itemId}>{busy ? <Loader2Icon className="animate-spin" /> : <UploadIcon />} Add a file</Button>
      </div>
      {doc.files.map((f) => (
        <div key={f.id} className="flex items-center gap-3 rounded-2xl border px-4 py-3">
          <FileSpreadsheetIcon className="size-4 shrink-0 text-muted-foreground" />
          <span className="flex-1 truncate">{f.name}</span>
          <span className="text-[12px] text-muted-foreground">{(f.size / 1024).toFixed(0)} KB</span>
        </div>
      ))}
      {!doc.files.length && <div className="text-[13px] text-muted-foreground">No file added yet.</div>}
    </div>
  );
}

function Scientist() {
  const docs = (useItems("analysis") ?? []).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  const { entries } = useLibrary("dataset");
  const catalog = useMemo(() => (entries ?? []).filter((e) => e.id), [entries]);
  const [id, setId] = useItemParam("scientist:current");
  const doc = docs.find((d) => d.id === id) ?? docs[0];
  const [tab, setTab] = useState<"workbook" | "models" | "excel" | "dax">("workbook");
  const [version, setVersion] = useState(0);
  const [busy, setBusy] = useState<string>();
  const docRef = useRef(doc);
  docRef.current = doc;
  const models = useMemo(() => catalog.filter((e) => doc?.datasetIds.includes(e.id!)).map(datasetOf), [catalog, doc?.datasetIds]);

  // A new analysis gets its workbook in OneDrive.
  useEffect(() => {
    if (!doc || doc.itemId) return;
    post<{ itemId: string; webUrl?: string }>("/workbook/create", { key: doc.id, title: doc.title }).then((w) => save("analysis", { ...doc, ...w, updatedAt: new Date().toISOString() }), (e) => toast.error(errorText(e)));
  }, [doc?.id, doc?.itemId]); // eslint-disable-line react-hooks/exhaustive-deps
  // The checked models get their connection in the workbook (CUBE formulas name it).
  useEffect(() => {
    if (!doc?.itemId || !models.length) return;
    void post("/workbook/connect", { itemId: doc.itemId, models: models.map((m) => ({ id: m.id, name: m.name })) }).catch(() => undefined);
  }, [doc?.itemId, models.map((m) => m.id).join()]); // eslint-disable-line react-hooks/exhaustive-deps

  const create = async () => {
    const d: Doc = { id: crypto.randomUUID(), title: `Analysis ${docs.length + 1}`, datasetIds: catalog.length === 1 ? [catalog[0].id!] : [], files: [], updatedAt: new Date().toISOString() };
    await save("analysis", d);
    setId(d.id);
    setTab(catalog.length === 1 ? "workbook" : "models");
  };
  const chat = useChat({
    screen: "scientist",
    item: doc?.id ?? "none",
    scope: () => {
      const d = docRef.current!;
      return { datasets: models, files: d.files, doc: { itemId: d.itemId, title: d.title, webUrl: d.webUrl } };
    },
    // Each answer may have changed the workbook: drawn again.
    onEffect: () => undefined,
  });
  // The preview follows the assistant: drawn again when it stops answering.
  const wasRunning = useRef(false);
  useEffect(() => {
    if (wasRunning.current && !chat.running) setVersion((v) => v + 1);
    wasRunning.current = chat.running;
  }, [chat.running]);

  const dl = async () => {
    if (!doc?.itemId) return;
    setBusy("download");
    try {
      download(await get<Blob>(`/workbook/download?itemId=${doc.itemId}`), `${doc.title}.xlsx`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  return (
    <Workspace
      screen="scientist"
      title="Data Scientist"
      item={<DocPicker items={docs} current={doc} name={(d) => d.title} onSelect={setId} onCreate={create} onDelete={(x) => void remove("analysis", x)} onRename={(d, title) => void save("analysis", { ...d, title })} createLabel="New analysis" empty="No analysis" />}
      actions={doc && (
        <Button variant="dark" size="sm" disabled={!!busy || !doc.itemId} onClick={() => void dl()}>
          {busy === "download" ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} Excel
        </Button>
      )}
      tabs={[{ id: "workbook", label: "Workbook", icon: <SheetIcon /> }, { id: "models", label: "Semantic models", icon: <DatabaseIcon />, count: doc?.datasetIds.length || undefined }, { id: "excel", label: "Excel files", icon: <FileSpreadsheetIcon />, count: doc?.files.length || undefined }, { id: "dax", label: "DAX", icon: <CodeIcon /> }]}
      tab={tab}
      onTab={setTab}
      scroll={tab !== "workbook"}
      chat={<ChatPanel chat={chat} screen="scientist" itemLabel={(x) => docs.find((d) => d.id === x)?.title} accept="image/*,.xlsx,.xls,.csv" disabled={!doc ? "Create an analysis first" : !doc.itemId ? "Creating the workbook…" : undefined} suggestions={["What drives the revenue growth? Build a summary sheet", "Compare actuals with the targets of my Excel file", "A sheet per region with the monthly trend and red / green variances"]} />}
    >
      {!doc ? (
        <Empty icon={<FlaskConicalIcon />} title="No analysis yet">
          An analysis is an Excel workbook the assistant builds with you: connected to the Power BI models, with your files, formulas and formats.
          <div><Button variant="dark" size="sm" className="mt-3" onClick={create}><PlusIcon /> New analysis</Button></div>
        </Empty>
      ) : tab === "workbook" ? (
        <WorkbookView doc={doc} version={version} />
      ) : tab === "models" ? (
        <CatalogChecklist kind="dataset" ids={doc.datasetIds} setIds={(datasetIds) => void save("analysis", { ...doc, datasetIds, updatedAt: new Date().toISOString() })} title="Semantic models" hint="Check the models the assistant may query and connect in this workbook (CUBE formulas, refreshable)." />
      ) : tab === "excel" ? (
        <Files doc={doc} onAdded={() => setVersion((v) => v + 1)} />
      ) : (
        <DaxConsole models={models} />
      )}
    </Workspace>
  );
}
