/**
 * Data Scientist: an analysis built by the assistant from semantic models (DAX) and uploaded Excel / CSV
 * files, combined with SQL; shown as blocks (findings, key figures, charts, tables, pivots), exported to Excel.
 */
import { useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BookOpenTextIcon, DatabaseIcon, DownloadIcon, FileSpreadsheetIcon, FlaskConicalIcon, Loader2Icon, PlusIcon, RefreshCwIcon, SheetIcon, Trash2Icon, UploadIcon, XIcon } from "lucide-react";
import { toast } from "sonner";
import { emptyAnalysis, type Analysis } from "@shared/analysis";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { DocPicker } from "@/components/doc-picker";
import { AddDialog } from "@/components/pickers";
import { ContextLinks } from "@/components/context-links";
import { Empty } from "@/components/common";
import { Button } from "@/components/ui/button";
import { BlockView } from "@/features/scientist/blocks";
import { useChat } from "@/lib/chat";
import { download, errorText, post } from "@/lib/api";
import { contextsOf, keepFile, remove, save, useItems, useLocal, type Stored } from "@/lib/store";

export const Route = createFileRoute("/scientist")({ component: Scientist });

type Doc = Stored["analysis"];

function Sources({ doc, update }: { doc: Doc; update: (d: Doc) => void }) {
  const datasets = useItems("dataset") ?? [];
  const [adding, setAdding] = useState(false);
  const [uploading, setUploading] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const addFiles = async (files: FileList) => {
    setUploading(true);
    try {
      const refs = [];
      for (const f of Array.from(files)) refs.push(await keepFile(f, f.name));
      update({ ...doc, files: [...doc.files, ...refs] });
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setUploading(false);
    }
  };
  return (
    <div className="mx-auto flex max-w-4xl flex-col gap-6 p-6">
      <div>
        <div className="mb-2 flex items-center gap-2">
          <div className="flex-1 text-[15px] font-semibold">Semantic models</div>
          <Button variant="outline" size="sm" onClick={() => setAdding(true)}><PlusIcon /> Add a model</Button>
        </div>
        <div className="flex flex-col gap-2">
          {datasets.map((d) => {
            const on = doc.datasetIds.includes(d.id);
            return (
              <button key={d.id} onClick={() => update({ ...doc, datasetIds: on ? doc.datasetIds.filter((x) => x !== d.id) : [...doc.datasetIds, d.id] })} className={`flex items-center gap-3 rounded-2xl border px-4 py-3 text-left ${on ? "border-foreground/40 bg-muted/50" : ""}`}>
                <DatabaseIcon className="size-4 text-muted-foreground" />
                <span className="flex-1">{d.name}</span>
                <span className="text-[12px] text-muted-foreground">{on ? "used" : "not used"}</span>
              </button>
            );
          })}
          {!datasets.length && <div className="text-[13px] text-muted-foreground">No model yet.</div>}
        </div>
      </div>
      <div>
        <div className="mb-2 flex items-center gap-2">
          <div className="flex-1 text-[15px] font-semibold">Files (Excel, CSV)</div>
          <input ref={input} type="file" hidden multiple accept=".xlsx,.xls,.csv,.tsv,.txt" onChange={(e) => e.target.files && void addFiles(e.target.files)} />
          <Button variant="outline" size="sm" onClick={() => input.current?.click()} disabled={uploading}>{uploading ? <Loader2Icon className="animate-spin" /> : <UploadIcon />} Upload</Button>
        </div>
        <div className="flex flex-col gap-2">
          {doc.files.map((f) => (
            <div key={f.id} className="flex items-center gap-3 rounded-2xl border px-4 py-3">
              <FileSpreadsheetIcon className="size-4 text-muted-foreground" />
              <span className="flex-1">{f.name}</span>
              <span className="text-[12px] text-muted-foreground">{(f.size / 1024).toFixed(0)} KB</span>
              <Button variant="ghost" size="icon-sm" onClick={() => update({ ...doc, files: doc.files.filter((x) => x.id !== f.id) })} aria-label="Remove"><XIcon /></Button>
            </div>
          ))}
          {!doc.files.length && <div className="text-[13px] text-muted-foreground">Upload Excel or CSV files to combine them with the models (targets, mappings, exports…).</div>}
        </div>
      </div>
      {(doc.sources.length > 0 || doc.steps.length > 0) && (
        <div>
          <div className="mb-2 text-[15px] font-semibold">Loaded tables and steps</div>
          <div className="flex flex-col gap-2 text-[12.5px]">
            {doc.sources.map((s) => (
              <div key={s.name} className="rounded-xl border px-4 py-2">
                <span className="font-mono font-medium">{s.name}</span> <span className="text-muted-foreground">— {s.kind === "dax" ? `DAX on ${s.datasetName}` : `${s.fileName}${s.sheet ? ` / ${s.sheet}` : ""}`} · {s.rows ?? "?"} rows</span>
                {s.kind === "dax" && <pre className="mt-1 whitespace-pre-wrap font-mono text-[11.5px] text-muted-foreground">{s.query}</pre>}
              </div>
            ))}
            {doc.steps.map((s) => (
              <div key={s.name} className="rounded-xl border px-4 py-2">
                <span className="font-mono font-medium">{s.name}</span> <span className="text-muted-foreground">{s.description}</span>
                <pre className="mt-1 whitespace-pre-wrap font-mono text-[11.5px] text-muted-foreground">{s.sql}</pre>
              </div>
            ))}
          </div>
        </div>
      )}
      <AddDialog kind="dataset" open={adding} onOpenChange={setAdding} onAdded={(id) => update({ ...doc, datasetIds: [...new Set([...doc.datasetIds, id])] })} />
    </div>
  );
}

function Scientist() {
  const docs = useItems("analysis") ?? [];
  const datasets = useItems("dataset") ?? [];
  const [id, setId] = useLocal("scientist:current", "");
  const doc = docs.find((d) => d.id === id) ?? docs[0];
  const [tab, setTab] = useState<"analysis" | "sources" | "context">("analysis");
  const [busy, setBusy] = useState<string>();
  const docRef = useRef(doc);
  docRef.current = doc;
  const update = (d: Doc) => void save("analysis", d);
  const create = async () => {
    const d: Doc = { ...emptyAnalysis(`Analysis ${docs.length + 1}`), id: crypto.randomUUID(), datasetIds: datasets.length === 1 ? [datasets[0].id] : [], files: [] };
    await save("analysis", d);
    setId(d.id);
    setTab("sources");
  };
  const itemKey = `analysis:${doc?.id ?? "none"}`;
  const chat = useChat({
    screen: "scientist",
    item: doc?.id ?? "none",
    scope: () => {
      const d = docRef.current!;
      const { id: _i, datasetIds, files, ...analysis } = d;
      return { datasets: datasets.filter((x) => datasetIds.includes(x.id)), files, doc: analysis, contexts: contextsOf(itemKey) };
    },
    onEffect: (name, data) => {
      if (name === "doc" && docRef.current) {
        const next = { ...docRef.current, ...(data as Analysis) };
        docRef.current = next;
        void save("analysis", next);
        setTab("analysis");
      }
    },
  });
  const run = async (label: string, f: () => Promise<void>) => {
    setBusy(label);
    try {
      await f();
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  const analysisOf = (d: Doc): Analysis => ({ title: d.title, sources: d.sources, steps: d.steps, blocks: d.blocks, updatedAt: d.updatedAt });
  return (
    <Workspace
      screen="scientist"
      title="Data Scientist"
      item={<DocPicker items={docs} current={doc} name={(d) => d.title} onSelect={setId} onCreate={create} onDelete={(x) => void remove("analysis", x)} createLabel="New analysis" empty="New analysis" />}
      actions={doc && (
        <>
          <Button variant="outline" size="sm" disabled={!!busy || !doc.sources.length} onClick={() => run("refresh", async () => update({ ...doc, ...(await post<Analysis>("/analysis/refresh", analysisOf(doc))) }))}>
            {busy === "refresh" ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh data
          </Button>
          <Button variant="dark" size="sm" disabled={!!busy || !doc.blocks.length} onClick={() => run("export", async () => download(await post<Blob>("/analysis/export", analysisOf(doc)), `${doc.title}.xlsx`))}>
            {busy === "export" ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} Excel
          </Button>
        </>
      )}
      tabs={[{ id: "analysis", label: "Analysis", icon: <SheetIcon />, count: doc?.blocks.length || undefined }, { id: "sources", label: "Sources", icon: <DatabaseIcon />, count: (doc ? doc.datasetIds.length + doc.files.length : 0) || undefined }, { id: "context", label: "Context", icon: <BookOpenTextIcon /> }]}
      tab={tab}
      onTab={setTab}
      chat={<ChatPanel chat={chat} screen="scientist" accept="image/*,.xlsx,.xls,.csv" disabled={!doc ? "Create an analysis first" : !doc.datasetIds.length && !doc.files.length ? "Add a model or a file in Sources" : undefined} suggestions={["What drives the revenue growth?", "Compare actuals with the targets of my Excel file", "Find the anomalies by month and region"]} />}
    >
      {!doc ? (
        <Empty icon={<FlaskConicalIcon />} title="No analysis yet">
          <Button variant="dark" size="sm" className="mt-2" onClick={create}><PlusIcon /> New analysis</Button>
        </Empty>
      ) : tab === "analysis" ? (
        doc.blocks.length ? (
          <div className="mx-auto flex max-w-5xl flex-col gap-6 p-7">
            <input value={doc.title} onChange={(e) => update({ ...doc, title: e.target.value })} className="bg-transparent text-[22px] font-semibold outline-none" />
            {doc.blocks.map((b) => (
              <div key={b.id} className="group relative">
                <BlockView b={b} />
                <button onClick={() => update({ ...doc, blocks: doc.blocks.filter((x) => x.id !== b.id) })} className="absolute -top-2 -right-2 hidden size-7 items-center justify-center rounded-full border bg-card text-muted-foreground group-hover:flex" aria-label="Remove block"><Trash2Icon className="size-3.5" /></button>
              </div>
            ))}
          </div>
        ) : (
          <Empty icon={<SheetIcon />} title="Ask a question">The assistant explores the sources, asks you to confirm, then writes the analysis here: findings, key figures, charts, tables and pivots.</Empty>
        )
      ) : tab === "sources" ? (
        <Sources doc={doc} update={update} />
      ) : (
        <ContextLinks itemKey={itemKey} title={doc.title} />
      )}
    </Workspace>
  );
}
