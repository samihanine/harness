import { useEffect, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { DatabaseIcon, ExternalLinkIcon, FileDownIcon, FileUpIcon, FlaskConicalIcon, LayoutDashboardIcon, PlusIcon, RotateCcwIcon, SaveIcon, Table2Icon, TextIcon, Trash2Icon } from "lucide-react";
import { CsvDialog } from "@/components/csv-dialog";
import { DataTable } from "@/components/data-table";
import { EntityMenu, withToast } from "@/components/entity";
import { Markdown } from "@/components/markdown";
import { DatasetPicker } from "@/components/pickers";
import { downloadPbix } from "@/components/report-actions";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { publishReports, researchAgent } from "@/agent/research-tools";
import { useLast } from "@/lib/last";
import { embedVisual, resetEmbed, type Ref } from "@/lib/pbi";
import { errorText } from "@/lib/ms";
import { deleteResearchReports, newResearch, readResearch, type Block, type MatrixPage, type ResearchEntry } from "@/lib/research";
import { patch, put, remove, useCollection, type DatasetEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";

export const Route = createFileRoute("/research")({
  validateSearch: (s: Record<string, unknown>): { research?: string; tab?: string } => ({ research: (s.research as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Research,
});

const blockId = () => crypto.randomUUID().slice(0, 8);

/**
 * Research: questions answered from semantic models. Canvas: markdown and Power BI visuals (each the
 * visual of a page of the model's report); Report: one Power BI report per model, each page a short
 * text and one visual, a matrix by default (downloadable as .pbix);
 * Datasets: the models it reads (or a CSV file becoming a model).
 */
function Research() {
  const { research: researchId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/research" });
  useLast("research:research", researchId, (research) => navigate({ search: { research }, replace: true }));
  const researches = useCollection("researches").map(readResearch);
  const datasets = useCollection("datasets");
  const research = researches.find((r) => r.id === researchId);
  const live = useRef({ research, datasets });
  live.current = { research, datasets };
  const agent = research && researchAgent(() => readResearch(live.current.research!), () => live.current.datasets);
  const create = async () => {
    const name = prompt("Name of the research", "New research")?.trim();
    if (name) navigate({ search: { research: (await put("researches", newResearch(name))).id, tab: "datasets" } });
  };

  return (
    <Workspace
      title="Research"
      icon={FlaskConicalIcon}
      picker={<EntityPicker items={researches.map((r) => ({ id: r.id, label: r.name }))} value={researchId} onChange={(research) => navigate({ search: { research } })} placeholder={researches.length ? "Choose a research" : "New research"} onAdd={() => void create()} addLabel="New research…" />}
      actions={
        <>
          {research && (
            <Button
              size="sm"
              variant="outline"
              title="Empty the canvas and the reports (the models and the conversations are kept)"
              onClick={() =>
                confirm(`Reset "${research.name}"? The canvas is emptied and its Power BI reports are deleted.`) &&
                void withToast("Resetting…", deleteResearchReports(research).then(() => patch("researches", research.id, { blocks: [], reports: {}, updatedAt: Date.now() })), "Research reset.")
              }
            >
              <RotateCcwIcon /> Reset
            </Button>
          )}
          <Button size="sm" variant="outline" onClick={() => void create()}>
            <PlusIcon /> New
          </Button>
          <EntityMenu
            kind="researches"
            entry={research}
            onRename={research && (() => { const name = prompt("Name of the research", research.name)?.trim(); if (name) void patch("researches", research.id, { name }); })}
            onRemove={research && (() => confirm(`Delete the research "${research.name}" and its Power BI reports?`) && void deleteResearchReports(research).then(() => remove("researches", research.id)).then(() => navigate({ search: {} })))}
          />
        </>
      }
      ai={{ scope: `research:${research?.id}`, agent, placeholder: "Ask a question about your data…", empty: "Choose a research, or start a new one." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        { id: "canvas", label: "Canvas", icon: LayoutDashboardIcon, count: research?.blocks.length, content: research ? <Canvas research={research} datasets={datasets} /> : <Empty icon={FlaskConicalIcon} action={<Button onClick={() => void create()}><PlusIcon /> New research</Button>}>A research answers questions from your semantic models, with text and Power BI visuals (matrices first).</Empty> },
        { id: "report", label: "Report", icon: Table2Icon, content: research ? <Reports research={research} datasets={datasets} /> : <Empty>Choose a research.</Empty> },
        { id: "datasets", label: "Datasets", icon: DatabaseIcon, count: research?.datasetIds.length, content: research ? <Datasets research={research} datasets={datasets} /> : <Empty>Choose a research.</Empty> },
      ]}
    />
  );
}

/* ---------------------------------- Canvas ---------------------------------- */

function Canvas({ research, datasets }: { research: ResearchEntry; datasets: DatasetEntry[] }) {
  const setBlocks = (blocks: Block[]) => void patch("researches", research.id, { blocks, updatedAt: Date.now() });
  if (!research.blocks.length)
    return (
      <Empty icon={LayoutDashboardIcon} action={<Button variant="outline" onClick={() => setBlocks([{ id: blockId(), kind: "markdown", text: "## Notes\n" }])}><TextIcon /> Add a text</Button>}>
        {research.datasetIds.length ? "Ask the assistant a question: its answer comes here as text and Power BI visuals." : "Choose the semantic models of this research in the Datasets tab, then ask a question."}
      </Empty>
    );
  const remove = (b: Block) => setBlocks(research.blocks.filter((x) => x !== b));
  return (
    <div className="min-h-0 flex-1 overflow-y-auto bg-muted/30">
      <div className="mx-auto grid max-w-6xl gap-4 p-6 md:grid-cols-2">
        {research.blocks.map((b) =>
          b.kind === "markdown" ? (
            <TextBlock key={b.id} text={b.text} onChange={(text) => setBlocks(research.blocks.map((x) => (x.id === b.id ? { ...b, text } : x)))} onDelete={() => remove(b)} />
          ) : b.kind === "visual" ? (
            <VisualBlock key={b.id} research={research} block={b} model={datasets.find((d) => d.id === b.datasetId)?.name ?? "?"} onDelete={() => remove(b)} />
          ) : (
            // Older researches: DAX charts, no longer drawn (their figures are now Power BI visuals).
            <section key={b.id} className="flex flex-col gap-2 rounded-xl border bg-card p-4 text-[12.5px] shadow-xs">
              <div className="flex items-center gap-2">
                <h3 className="min-w-0 flex-1 truncate text-[14px] font-semibold">{b.title}</h3>
                <Badge variant="outline">older chart</Badge>
                <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Delete the chart" onClick={() => remove(b)}>
                  <Trash2Icon />
                </Button>
              </div>
              <p className="text-muted-foreground">Charts are now Power BI visuals: ask the assistant to show this again.</p>
              <pre className="overflow-x-auto rounded-md bg-muted p-2 font-mono text-[11.5px] whitespace-pre-wrap">{b.dax}</pre>
            </section>
          ),
        )}
      </div>
    </div>
  );
}

/** The visual of a research page on the canvas (embedded from the model's report once written), with how it is computed. */
function VisualBlock({ research, block, model, onDelete }: { research: ResearchEntry; block: Extract<Block, { kind: "visual" }>; model: string; onDelete: () => void }) {
  const rep = research.reports[block.datasetId];
  const page = rep?.pages.find((p) => p.title === block.page);
  const wide = !page?.chart || page.chart === "matrix" || page.chart === "table";
  return (
    <section className={`flex flex-col gap-3 rounded-xl border bg-card p-4 shadow-xs ${wide ? "md:col-span-2" : ""}`}>
      <div className="flex items-center gap-2">
        <h3 className="min-w-0 flex-1 truncate text-[14px] font-semibold">{block.page}</h3>
        <Badge variant="secondary">{model}</Badge>
        <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Remove from the canvas" onClick={onDelete}>
          <Trash2Icon />
        </Button>
      </div>
      {!page ? (
        <p className="rounded-md bg-muted p-3 text-[12.5px] text-muted-foreground">This page is no longer in the report.</p>
      ) : rep?.id && page.name && page.visualName ? (
        <EmbeddedVisual key={`${rep.id}:${rep.saved}:${page.visualName}`} reportRef={{ id: rep.id, groupId: rep.groupId }} page={page.name} visual={page.visualName} />
      ) : (
        <p className="rounded-md bg-muted p-3 text-[12.5px] text-muted-foreground">Shown once the report is written (at the end of the assistant's answer, or Report tab › “Create the report”).</p>
      )}
      {page && <PageFields page={page} />}
    </section>
  );
}

/** How a page's figures are obtained: its fields and filters. */
function PageFields({ page }: { page: MatrixPage }) {
  return (
    <details className="text-[12px] text-muted-foreground">
      <summary className="cursor-pointer">How it is computed</summary>
      <div className="mt-2 flex flex-col gap-0.5">
        <span>visual: {page.chart ?? "matrix"}</span>
        {(["rows", "columns", "values"] as const).map((k) => page[k].length > 0 && <span key={k}>{k}: {page[k].join(", ")}</span>)}
        {page.filters.map((f) => <span key={f.field}>filter: {f.field} = {f.values.join(", ")}</span>)}
      </div>
    </details>
  );
}

function EmbeddedVisual({ reportRef, page, visual }: { reportRef: Ref; page: string; visual: string }) {
  const element = useRef<HTMLDivElement>(null);
  const [error, setError] = useState<string>();
  useEffect(() => {
    const el = element.current;
    if (!el) return;
    embedVisual(el, reportRef, page, visual).catch((e) => setError(errorText(e)));
    return () => void resetEmbed(el);
  }, [reportRef.id, reportRef.groupId, page, visual]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <>
      {error && <p className="text-[12.5px] text-destructive">{error}</p>}
      <div ref={element} className="h-[420px] w-full" />
    </>
  );
}

function TextBlock({ text, onChange, onDelete }: { text: string; onChange: (t: string) => void; onDelete: () => void }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(text);
  useEffect(() => setDraft(text), [text]);
  return (
    <section className="group relative rounded-xl border bg-card p-5 shadow-xs md:col-span-2">
      {editing ? (
        <div className="flex flex-col gap-2">
          <Textarea rows={8} className="font-mono text-[13px]" value={draft} onChange={(e) => setDraft(e.target.value)} />
          <div className="flex gap-2">
            <Button size="sm" onClick={() => (onChange(draft), setEditing(false))}>Save</Button>
            <Button size="sm" variant="outline" onClick={() => (setDraft(text), setEditing(false))}>Cancel</Button>
          </div>
        </div>
      ) : (
        <div className="markdown cursor-text" onDoubleClick={() => setEditing(true)}>
          <Markdown>{text}</Markdown>
        </div>
      )}
      {!editing && (
        <div className="absolute top-3 right-3 flex gap-1 opacity-0 transition-opacity group-hover:opacity-100">
          <Button variant="outline" size="sm" className="h-7" onClick={() => setEditing(true)}>Edit</Button>
          <Button variant="outline" size="icon" className="size-7" aria-label="Delete the text" onClick={onDelete}>
            <Trash2Icon />
          </Button>
        </div>
      )}
    </section>
  );
}

/* ---------------------------------- Reports ---------------------------------- */

function Reports({ research, datasets }: { research: ResearchEntry; datasets: DatasetEntry[] }) {
  const models = datasets.filter((d) => research.datasetIds.includes(d.id));
  const [chosen, setChosen] = useState<string>();
  const model = models.find((d) => d.id === chosen) ?? models[0];
  if (!model) return <Empty icon={Table2Icon}>Choose the semantic models of this research in the Datasets tab: each gets its own report.</Empty>;
  const rep = research.reports[model.id];
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex shrink-0 flex-wrap items-center gap-2 border-b bg-card px-5 py-2.5">
        {models.map((d) => (
          <Button key={d.id} size="sm" variant={d.id === model.id ? "secondary" : "ghost"} className="h-8 rounded-full" onClick={() => setChosen(d.id)}>
            <DatabaseIcon /> {d.name}
          </Button>
        ))}
        <span className="ml-auto" />
        {rep?.dirty && (
          <Button size="sm" onClick={() => void withToast("Writing the report…", publishReports(research, datasets), "Report written.")}>
            <SaveIcon /> {rep.id ? "Save the report" : "Create the report"}
          </Button>
        )}
        {rep?.id && (
          <>
            <Button size="sm" variant="outline" onClick={() => void withToast("Downloading the .pbix…", downloadPbix({ id: rep.id!, groupId: rep.groupId, name: `${research.name} — ${model.name}` }), "Downloaded.")}>
              <FileDownIcon /> .pbix
            </Button>
            <Button size="sm" variant="outline" asChild>
              <a href={`https://app.powerbi.com/groups/${rep.groupId ?? "me"}/reports/${rep.id}`} target="_blank" rel="noreferrer">
                <ExternalLinkIcon /> Power BI
              </a>
            </Button>
          </>
        )}
      </div>
      <div className="flex min-h-0 flex-1">
        <aside className="flex w-72 shrink-0 flex-col gap-2 overflow-y-auto border-r bg-card p-4">
          <span className="text-[12px] font-medium text-muted-foreground">Pages ({rep?.pages.length ?? 0})</span>
          {!rep?.pages.length && <p className="text-[12.5px] text-muted-foreground">No page yet. Ask the assistant, e.g. “a matrix of sales by region and year”.</p>}
          {rep?.pages.map((p) => (
            <div key={p.title} className="flex flex-col gap-1 rounded-lg border p-3 text-[12px]">
              <span className="text-[13px] font-semibold">{p.title}</span>
              <span><span className="text-muted-foreground">visual:</span> {p.chart ?? "matrix"}</span>
              {(["rows", "columns", "values"] as const).map((k) => p[k].length > 0 && <span key={k}><span className="text-muted-foreground">{k}:</span> {p[k].join(", ")}</span>)}
              {p.filters.map((f) => <span key={f.field}><span className="text-muted-foreground">filter:</span> {f.field} = {f.values.join(", ")}</span>)}
            </div>
          ))}
        </aside>
        {rep?.id ? <Embedded key={`${rep.id}:${rep.saved}`} id={rep.id} groupId={rep.groupId} /> : <Empty icon={Table2Icon}>{rep?.pages.length ? "The report is created when the assistant's answer is done (or with “Create the report”)." : "The report is created with its first page."}</Empty>}
      </div>
    </div>
  );
}

function Embedded({ id, groupId }: { id: string; groupId?: string }) {
  const { element, error } = useEmbed({ id, groupId }, "view");
  return (
    <div className="flex min-w-0 flex-1 flex-col">
      {error && <p className="px-5 py-2 text-[13px] text-destructive">{error}</p>}
      <div ref={element} className="min-h-0 flex-1" />
    </div>
  );
}

/* ---------------------------------- Datasets ---------------------------------- */

function Datasets({ research, datasets }: { research: ResearchEntry; datasets: DatasetEntry[] }) {
  const [csv, setCsv] = useState(false);
  const toggle = (id: string, on: boolean) => void patch("researches", research.id, { datasetIds: on ? [...research.datasetIds, id] : research.datasetIds.filter((x) => x !== id) });
  return (
    <div className="mx-auto flex w-full max-w-4xl flex-col gap-4 overflow-y-auto p-6">
      <Section
        title="Semantic models of this research"
        description="The assistant reads these models (DAX). Each gets its own report, whose visuals show on the canvas."
        actions={
          <>
            <DatasetPicker value={undefined} onChange={(id) => toggle(id, true)} />
            <Button size="sm" variant="outline" onClick={() => setCsv(true)}>
              <FileUpIcon /> From a CSV
            </Button>
          </>
        }
      >
        {!datasets.length && <p className="text-[13px] text-muted-foreground">No semantic model in the app: add one from its link, or upload a CSV file.</p>}
        <div className="flex flex-col divide-y">
          {datasets.map((d) => (
            <label key={d.id} className="flex cursor-pointer items-center gap-3 py-2.5 text-[13.5px]">
              <Switch checked={research.datasetIds.includes(d.id)} onCheckedChange={(on) => toggle(d.id, on)} disabled={d.access === "view"} />
              <span className="min-w-0 flex-1 truncate font-medium">{d.name}</span>
              {d.info && <span className="text-[12px] text-muted-foreground">{d.info.tables.length} tables</span>}
              <Badge variant={d.access === "view" ? "outline" : "secondary"}>{d.access === "view" ? "no DAX access" : (d.access ?? "?")}</Badge>
            </label>
          ))}
        </div>
      </Section>
      <CsvDialog open={csv} onOpenChange={setCsv} onCreated={(id) => toggle(id, true)} />
    </div>
  );
}
