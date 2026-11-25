/**
 * Data Scientist: an analysis built by the assistant from semantic models (DAX) and uploaded Excel / CSV
 * files, combined with SQL; shown as blocks (findings, key figures, charts, tables, pivots), exported to Excel.
 */
import { useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { CheckIcon, CodeIcon, DatabaseIcon, DownloadIcon, FileSpreadsheetIcon, FlaskConicalIcon, Loader2Icon, PencilIcon, PlusIcon, RefreshCwIcon, SheetIcon, Trash2Icon, UploadIcon, XIcon } from "lucide-react";
import { toast } from "sonner";
import { emptyAnalysis, type Analysis } from "@shared/analysis";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { DocPicker } from "@/components/doc-picker";
import { AddDialog } from "@/components/pickers";
import { DaxConsole } from "@/features/views/dax-console";
import { Empty } from "@/components/common";
import { Button } from "@/components/ui/button";
import { BlockView } from "@/features/scientist/blocks";
import { useChat } from "@/lib/chat";
import { screenSearch, useItemParam } from "@/lib/params";
import { download, errorText, post } from "@/lib/api";
import { keepFile, remove, save, useItems, type Stored } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/scientist")({ component: Scientist, validateSearch: screenSearch });

type Doc = Stored["analysis"];

/** A name changed in place (in this app only): the pencil turns it into a field. */
function Rename({ name, onRename }: { name: string; onRename: (n: string) => void }) {
  const [editing, setEditing] = useState(false);
  const [value, setValue] = useState(name);
  if (!editing)
    return (
      <span className="group flex min-w-0 flex-1 items-center gap-1.5">
        <span className="truncate">{name}</span>
        <button onClick={(e) => (e.stopPropagation(), setValue(name), setEditing(true))} className="opacity-0 group-hover:opacity-100" aria-label="Rename"><PencilIcon className="size-3.5 text-muted-foreground" /></button>
      </span>
    );
  const done = () => (setEditing(false), value.trim() && value.trim() !== name && onRename(value.trim()));
  return <input autoFocus value={value} onClick={(e) => e.stopPropagation()} onChange={(e) => setValue(e.target.value)} onBlur={done} onKeyDown={(e) => (e.key === "Enter" ? done() : e.key === "Escape" && setEditing(false))} className="input h-8 min-w-0 flex-1" />;
}

/** The semantic models the analysis uses (checked), renamable, and more added (also from a report's link). */
function Models({ doc, update }: { doc: Doc; update: (d: Doc) => void }) {
  const datasets = useItems("dataset") ?? [];
  const [adding, setAdding] = useState(false);
  return (
    <div className="mx-auto flex max-w-4xl flex-col gap-3 p-6">
      <div className="flex items-center gap-2">
        <div className="flex-1">
          <div className="text-[15px] font-semibold">Semantic models</div>
          <div className="text-[13px] text-muted-foreground">Check the ones the assistant may query. Add one from the list, or paste its link — or the link of one of its reports.</div>
        </div>
        <Button variant="outline" size="sm" onClick={() => setAdding(true)}><PlusIcon /> Add a model</Button>
      </div>
      {datasets.map((d) => {
        const on = doc.datasetIds.includes(d.id);
        return (
          <div key={d.id} role="button" tabIndex={0} onClick={() => update({ ...doc, datasetIds: on ? doc.datasetIds.filter((x) => x !== d.id) : [...doc.datasetIds, d.id] })} className={cn("flex cursor-pointer items-center gap-3 rounded-2xl border px-4 py-3 text-left", on && "border-foreground/40 bg-muted/50")}>
            <span className={cn("flex size-5 shrink-0 items-center justify-center rounded-md border", on && "border-foreground bg-foreground text-background")}>{on && <CheckIcon className="size-3.5" />}</span>
            <DatabaseIcon className="size-4 shrink-0 text-muted-foreground" />
            <Rename name={d.name} onRename={(name) => void save("dataset", { ...d, name })} />
            <Button variant="ghost" size="icon-sm" onClick={(e) => (e.stopPropagation(), confirm(`Remove "${d.name}" from the list?`) && void remove("dataset", d.id))} aria-label="Remove"><XIcon /></Button>
          </div>
        );
      })}
      {!datasets.length && <div className="text-[13px] text-muted-foreground">No model yet.</div>}
      <AddDialog kind="dataset" open={adding} onOpenChange={setAdding} onAdded={(id) => update({ ...doc, datasetIds: [...new Set([...doc.datasetIds, id])] })} />
    </div>
  );
}

/** The Excel / CSV files of the analysis (renamable), and the tables and steps loaded from the sources. */
function Files({ doc, update }: { doc: Doc; update: (d: Doc) => void }) {
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
      <div className="flex flex-col gap-3">
        <div className="flex items-center gap-2">
          <div className="flex-1">
            <div className="text-[15px] font-semibold">Excel and CSV files</div>
            <div className="text-[13px] text-muted-foreground">Targets, mappings, exports… the assistant combines them with the models.</div>
          </div>
          <input ref={input} type="file" hidden multiple accept=".xlsx,.xls,.csv,.tsv,.txt" onChange={(e) => e.target.files && void addFiles(e.target.files)} />
          <Button variant="outline" size="sm" onClick={() => input.current?.click()} disabled={uploading}>{uploading ? <Loader2Icon className="animate-spin" /> : <UploadIcon />} Upload</Button>
        </div>
        {doc.files.map((f) => (
          <div key={f.id} className="flex items-center gap-3 rounded-2xl border px-4 py-3">
            <FileSpreadsheetIcon className="size-4 shrink-0 text-muted-foreground" />
            <Rename name={f.name} onRename={(name) => update({ ...doc, files: doc.files.map((x) => (x.id === f.id ? { ...x, name } : x)) })} />
            <span className="text-[12px] text-muted-foreground">{(f.size / 1024).toFixed(0)} KB</span>
            <Button variant="ghost" size="icon-sm" onClick={() => update({ ...doc, files: doc.files.filter((x) => x.id !== f.id) })} aria-label="Remove"><XIcon /></Button>
          </div>
        ))}
        {!doc.files.length && <div className="text-[13px] text-muted-foreground">No file yet.</div>}
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
    </div>
  );
}

function Scientist() {
  const docs = useItems("analysis") ?? [];
  const datasets = useItems("dataset") ?? [];
  const [id, setId] = useItemParam("scientist:current");
  const doc = docs.find((d) => d.id === id) ?? docs[0];
  const [tab, setTab] = useState<"analysis" | "models" | "excel" | "dax">("analysis");
  const [busy, setBusy] = useState<string>();
  const docRef = useRef(doc);
  docRef.current = doc;
  const update = (d: Doc) => void save("analysis", d);
  const create = async () => {
    const d: Doc = { ...emptyAnalysis(`Analysis ${docs.length + 1}`), id: crypto.randomUUID(), datasetIds: datasets.length === 1 ? [datasets[0].id] : [], files: [] };
    await save("analysis", d);
    setId(d.id);
    setTab("models");
  };
  const chat = useChat({
    screen: "scientist",
    item: doc?.id ?? "none",
    scope: () => {
      const d = docRef.current!;
      const { id: _i, datasetIds, files, ...analysis } = d;
      return { datasets: datasets.filter((x) => datasetIds.includes(x.id)), files, doc: analysis };
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
      item={<DocPicker items={docs} current={doc} name={(d) => d.title} onSelect={setId} onCreate={create} onDelete={(x) => void remove("analysis", x)} onRename={(d, title) => void save("analysis", { ...d, title })} createLabel="New analysis" empty="No analysis" />}
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
      tabs={[{ id: "analysis", label: "Analysis", icon: <SheetIcon />, count: doc?.blocks.length || undefined }, { id: "models", label: "Semantic models", icon: <DatabaseIcon />, count: doc?.datasetIds.length || undefined }, { id: "excel", label: "Excel", icon: <FileSpreadsheetIcon />, count: doc?.files.length || undefined }, { id: "dax", label: "DAX", icon: <CodeIcon /> }]}
      tab={tab}
      onTab={setTab}
      chat={<ChatPanel chat={chat} screen="scientist" itemLabel={(id) => docs.find((d) => d.id === id)?.title} accept="image/*,.xlsx,.xls,.csv" disabled={!doc ? "Create an analysis first" : !doc.datasetIds.length && !doc.files.length ? "Add a semantic model or an Excel file first" : undefined} suggestions={["What drives the revenue growth?", "Compare actuals with the targets of my Excel file", "Find the anomalies by month and region"]} />}
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
      ) : tab === "models" ? (
        <Models doc={doc} update={update} />
      ) : tab === "excel" ? (
        <Files doc={doc} update={update} />
      ) : (
        <DaxConsole models={datasets.filter((d) => doc.datasetIds.includes(d.id))} />
      )}
    </Workspace>
  );
}
