import { useEffect, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Legend, Line, LineChart, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { DatabaseIcon, ExternalLinkIcon, FileDownIcon, FileUpIcon, FlaskConicalIcon, LayoutDashboardIcon, PlusIcon, RefreshCwIcon, SaveIcon, Table2Icon, TextIcon, Trash2Icon } from "lucide-react";
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
import { computeBlock, newResearch, readResearch, type Block, type ResearchEntry } from "@/lib/research";
import { patch, put, remove, useCollection, type DatasetEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";

export const Route = createFileRoute("/research")({
  validateSearch: (s: Record<string, unknown>): { research?: string; tab?: string } => ({ research: (s.research as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Research,
});

type Chart = Extract<Block, { kind: "chart" }>;
const blockId = () => crypto.randomUUID().slice(0, 8);

/**
 * Research: questions answered from semantic models. Canvas: markdown and charts computed by DAX;
 * Report: one Power BI report per model, each page a full-page matrix (downloadable as .pbix);
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
          <Button size="sm" variant="outline" onClick={() => void create()}>
            <PlusIcon /> New
          </Button>
          <EntityMenu
            kind="researches"
            entry={research}
            onRename={research && (() => { const name = prompt("Name of the research", research.name)?.trim(); if (name) void patch("researches", research.id, { name }); })}
            onRemove={research && (() => confirm(`Delete the research "${research.name}"? (its reports stay in Power BI)`) && void remove("researches", research.id).then(() => navigate({ search: {} })))}
          />
        </>
      }
      ai={{ scope: `research:${research?.id}`, agent, placeholder: "Ask a question about your data…", empty: "Choose a research, or start a new one." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        { id: "canvas", label: "Canvas", icon: LayoutDashboardIcon, count: research?.blocks.length, content: research ? <Canvas research={research} datasets={datasets} /> : <Empty icon={FlaskConicalIcon} action={<Button onClick={() => void create()}><PlusIcon /> New research</Button>}>A research answers questions from your semantic models, with charts, text and matrix reports.</Empty> },
        { id: "report", label: "Report", icon: Table2Icon, content: research ? <Reports research={research} datasets={datasets} /> : <Empty>Choose a research.</Empty> },
        { id: "datasets", label: "Datasets", icon: DatabaseIcon, count: research?.datasetIds.length, content: research ? <Datasets research={research} datasets={datasets} /> : <Empty>Choose a research.</Empty> },
      ]}
    />
  );
}

/* ---------------------------------- Canvas ---------------------------------- */

const PALETTE = ["var(--chart-1)", "var(--chart-2)", "var(--chart-3)", "var(--chart-4)", "var(--chart-5)"];

function Canvas({ research, datasets }: { research: ResearchEntry; datasets: DatasetEntry[] }) {
  const setBlocks = (blocks: Block[]) => void patch("researches", research.id, { blocks, updatedAt: Date.now() });
  const refresh = async (b: Chart) => {
    const d = datasets.find((x) => x.id === b.datasetId);
    if (!d) return;
    const next = await computeBlock(b, d);
    setBlocks(research.blocks.map((x) => (x.id === b.id ? next : x)));
  };
  if (!research.blocks.length)
    return (
      <Empty icon={LayoutDashboardIcon} action={<Button variant="outline" onClick={() => setBlocks([{ id: blockId(), kind: "markdown", text: "## Notes\n" }])}><TextIcon /> Add a text</Button>}>
        {research.datasetIds.length ? "Ask the assistant a question: its answer comes here as charts and text." : "Choose the semantic models of this research in the Datasets tab, then ask a question."}
      </Empty>
    );
  return (
    <div className="min-h-0 flex-1 overflow-y-auto bg-muted/30">
      <div className="mx-auto grid max-w-6xl gap-4 p-6 md:grid-cols-2">
        {research.blocks.map((b) =>
          b.kind === "markdown" ? (
            <TextBlock key={b.id} text={b.text} onChange={(text) => setBlocks(research.blocks.map((x) => (x.id === b.id ? { ...b, text } : x)))} onDelete={() => setBlocks(research.blocks.filter((x) => x !== b))} />
          ) : (
            <section key={b.id} className={`flex flex-col gap-3 rounded-xl border bg-card p-4 shadow-xs ${b.chart === "table" ? "md:col-span-2" : ""}`}>
              <div className="flex items-center gap-2">
                <h3 className="min-w-0 flex-1 truncate text-[14px] font-semibold">{b.title}</h3>
                <Badge variant="secondary">{datasets.find((d) => d.id === b.datasetId)?.name ?? "?"}</Badge>
                <Button variant="ghost" size="icon" className="size-7" aria-label="Compute again" title={b.at ? `Computed ${new Date(b.at).toLocaleString()}` : "Compute"} onClick={() => void withToast("Computing…", refresh(b), "Up to date.")}>
                  <RefreshCwIcon />
                </Button>
                <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Delete the chart" onClick={() => setBlocks(research.blocks.filter((x) => x !== b))}>
                  <Trash2Icon />
                </Button>
              </div>
              {b.error ? <p className="rounded-md bg-destructive/10 p-3 text-[12.5px] text-destructive">{b.error}</p> : <ChartView block={b} />}
              <details className="text-[12px] text-muted-foreground">
                <summary className="cursor-pointer">DAX</summary>
                <pre className="mt-2 overflow-x-auto rounded-md bg-muted p-2 font-mono text-[11.5px] whitespace-pre-wrap">{b.dax}</pre>
              </details>
            </section>
          ),
        )}
      </div>
    </div>
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

const num = (v: unknown) => (typeof v === "number" ? v : Number(v));
const short = (v: number) => (Math.abs(v) >= 1e6 ? `${(v / 1e6).toFixed(1)}M` : Math.abs(v) >= 1e3 ? `${(v / 1e3).toFixed(1)}k` : Number.isInteger(v) ? String(v) : v.toFixed(2));
/** Columns named like a rate ("Margin %", "Growth pct", "Return Rate") are shown as percentages. */
const isRate = (k: string) => /%|pct|percent|rate|ratio|taux/i.test(k);
const pct = (v: number) => `${(v * 100).toFixed(Math.abs(v) < 0.1 ? 1 : 0)} %`;

/** A chart of the rows: first column = categories, the numeric columns after it = series. */
function ChartView({ block }: { block: Chart }) {
  const rows = block.rows ?? [];
  if (!rows.length) return <p className="text-[13px] text-muted-foreground">No data.</p>;
  const keys = Object.keys(rows[0]);
  const [category, ...rest] = keys;
  const series = rest.filter((k) => rows.some((r) => !Number.isNaN(num(r[k])) && r[k] !== null));
  const data = rows.map((r) => ({ ...r, [category]: String(r[category] ?? "—"), ...Object.fromEntries(series.map((k) => [k, num(r[k])])) }));
  if (block.chart === "table") return <DataTable head={keys} rows={rows.map((r) => keys.map((k) => String(r[k] ?? "")))} empty="No row." />;
  if (block.chart === "kpi")
    return (
      <div className="flex flex-wrap gap-6 py-2">
        {keys.map((k) => (
          <div key={k} className="flex flex-col">
            <span className="text-[12px] text-muted-foreground">{k}</span>
            <span className="text-[28px] font-semibold tracking-tight tabular-nums">{typeof rows[0][k] === "number" ? (isRate(k) ? pct(rows[0][k] as number) : (rows[0][k] as number).toLocaleString()) : String(rows[0][k] ?? "—")}</span>
          </div>
        ))}
      </div>
    );
  const axis = { tick: { fontSize: 11, fill: "var(--muted-foreground)" }, stroke: "var(--border)" };
  const rates = series.length > 0 && series.every(isRate);
  const tick = rates ? pct : short;
  const tip = { formatter: (v: unknown, name: unknown) => [typeof v === "number" ? (isRate(String(name)) ? pct(v) : v.toLocaleString()) : String(v), String(name)] as [string, string] };
  return (
    <div className="h-64">
      <ResponsiveContainer width="100%" height="100%">
        {block.chart === "pie" ? (
          <PieChart>
            <Pie data={data} dataKey={series[0]} nameKey={category} innerRadius="45%" outerRadius="80%" paddingAngle={2}>
              {data.map((_, i) => (
                <Cell key={i} fill={PALETTE[i % PALETTE.length]} />
              ))}
            </Pie>
            <Tooltip {...tip} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
          </PieChart>
        ) : block.chart === "line" ? (
          <LineChart data={data}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
            <XAxis dataKey={category} {...axis} />
            <YAxis tickFormatter={tick} {...axis} width={48} />
            <Tooltip {...tip} />
            {series.length > 1 && <Legend wrapperStyle={{ fontSize: 12 }} />}
            {series.map((k, i) => (
              <Line key={k} dataKey={k} stroke={PALETTE[i % PALETTE.length]} strokeWidth={2} dot={false} />
            ))}
          </LineChart>
        ) : block.chart === "area" ? (
          <AreaChart data={data}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
            <XAxis dataKey={category} {...axis} />
            <YAxis tickFormatter={tick} {...axis} width={48} />
            <Tooltip {...tip} />
            {series.map((k, i) => (
              <Area key={k} dataKey={k} stroke={PALETTE[i % PALETTE.length]} fill={PALETTE[i % PALETTE.length]} fillOpacity={0.2} />
            ))}
          </AreaChart>
        ) : (
          <BarChart data={data} layout={block.chart === "bar" ? "vertical" : "horizontal"}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
            {block.chart === "bar" ? <XAxis type="number" tickFormatter={tick} {...axis} /> : <XAxis dataKey={category} {...axis} />}
            {block.chart === "bar" ? <YAxis type="category" dataKey={category} width={110} {...axis} /> : <YAxis tickFormatter={tick} {...axis} width={48} />}
            <Tooltip cursor={{ fill: "var(--muted)" }} {...tip} />
            {series.length > 1 && <Legend wrapperStyle={{ fontSize: 12 }} />}
            {series.map((k, i) => (
              <Bar key={k} dataKey={k} fill={PALETTE[i % PALETTE.length]} radius={4} />
            ))}
          </BarChart>
        )}
      </ResponsiveContainer>
    </div>
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
        description="The assistant reads these models (DAX). Each gets its own matrix report."
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
