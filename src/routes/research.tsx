import { useEffect, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { z } from "zod";
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
import { withContexts } from "@/agent/context";
import { tool, type Agent } from "@/agent/loop";
import { DAX_RULES, daxTool } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { modelSummary } from "@/lib/model";
import { rowsText } from "@/lib/pbi";
import { CHARTS, checkPage, computeBlock, newResearch, readResearch, saveResearchReport, type Block, type MatrixPage, type ResearchEntry } from "@/lib/research";
import { patch, put, remove, useCollection, type DatasetEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";

export const Route = createFileRoute("/research")({
  validateSearch: (s: Record<string, unknown>): { research?: string; tab?: string } => ({ research: (s.research as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Research,
});

type Chart = Extract<Block, { kind: "chart" }>;
const blockId = () => crypto.randomUUID().slice(0, 8);
const fieldList = z.array(z.string()).default([]);

/** The research as the AI sees it: models, canvas blocks, report pages. */
function researchText(r: ResearchEntry, datasets: DatasetEntry[]) {
  const name = (id: string) => datasets.find((d) => d.id === id)?.name ?? id;
  return [
    `Models: ${r.datasetIds.map((id) => `"${name(id)}" (dataset_id ${id})`).join(", ") || "none (the user chooses them in the Datasets tab)"}.`,
    `Canvas (${r.blocks.length} blocks, in order):\n${r.blocks.map((b) => (b.kind === "markdown" ? `- [${b.id}] text: ${b.text.replace(/\s+/g, " ").slice(0, 90)}` : `- [${b.id}] ${b.chart} chart "${b.title}" on ${name(b.datasetId)}${b.error ? ` — ERROR: ${b.error.slice(0, 160)}` : ` — ${b.rows?.length ?? 0} rows`}`)).join("\n") || "empty"}`,
    ...r.datasetIds.map((id) => {
      const rep = r.reports[id];
      return `Report on "${name(id)}"${rep?.id ? "" : " (not created yet)"}${rep?.dirty ? " (unsaved changes)" : ""}: ${rep?.pages.length ? rep.pages.map((p) => `\n- page "${p.title}": rows ${p.rows.join(", ") || "—"}; columns ${p.columns.join(", ") || "—"}; values ${p.values.join(", ") || "—"}${p.filters.length ? `; filters ${p.filters.map((f) => `${f.field} in (${f.values.join(", ")})`).join("; ")}` : ""}`).join("") : "no page"}`;
    }),
  ].join("\n\n");
}

function researchAgent(get: () => ResearchEntry, datasets: () => DatasetEntry[]): Agent {
  const save = (change: Partial<ResearchEntry>) => patch("researches", get().id, { ...change, updatedAt: Date.now() });
  const models = () => datasets().filter((d) => get().datasetIds.includes(d.id));
  const model = (id?: string) => {
    const list = models();
    const d = id ? list.find((x) => x.id === id || x.name.toLowerCase() === id.toLowerCase()) : list.length === 1 ? list[0] : undefined;
    if (!d) throw new Error(list.length ? `Give dataset_id: one of ${list.map((x) => `${x.id} (${x.name})`).join(", ")}.` : "No model in this research: ask the user to choose one in the Datasets tab.");
    return d;
  };
  const block = (id: string) => {
    const b = get().blocks.find((x) => x.id === id);
    if (!b) throw new Error(`No block ${id}. Blocks: ${get().blocks.map((x) => x.id).join(", ") || "none"}.`);
    return b;
  };
  const insert = (b: Block, after?: string) => {
    const blocks = [...get().blocks];
    const i = after ? blocks.findIndex((x) => x.id === after) : -1;
    blocks.splice(i >= 0 ? i + 1 : blocks.length, 0, b);
    return blocks;
  };
  /** A chart is kept only when its query works (the error goes back to the AI to fix). */
  const computed = async (b: Chart) => {
    const d = model(b.datasetId);
    const done = await computeBlock({ ...b, datasetId: d.id }, d);
    if (done.error) throw new Error(`The DAX query failed: ${done.error}`);
    if (!done.rows?.length) throw new Error("The query returned no row: check the filters.");
    return done;
  };
  const pageArgs = z.object({
    dataset_id: z.string().optional(),
    title: z.string().describe("Page name (an existing title replaces that page)"),
    rows: fieldList.describe("Columns as matrix rows, e.g. ['Product'[Category]]"),
    columns: fieldList.describe("Columns as matrix columns (e.g. a year), often empty"),
    values: fieldList.describe("Measures, or Sum('T'[col]) / Avg / Count / Min / Max"),
    filters: z.array(z.object({ field: z.string(), values: z.array(z.union([z.string(), z.number()])) })).default([]).describe("Page filters: a column and the values kept"),
  });
  return {
    instructions: `Screen: RESEARCH. You answer business questions from the research's semantic models and present the results in two places:
- the CANVAS (Canvas tab): text blocks (markdown: findings, explanations) and charts, each computed by one DAX query. Chart kinds: ${CHARTS.join(", ")}. The query's first column is the category (x axis / slices), the next columns are the numeric series; kpi shows the first row's numbers; table shows all rows. Keep queries small (TOPN, SUMMARIZECOLUMNS, ≤ 30 rows for charts).
- the REPORTS (Report tab): one Power BI report per model; each page is ONE full-page matrix: rows, columns, values, page filters — no formatting, no other visual. set_page adds or replaces a page; the reports are written when your answer is done.
How to work: check numbers with run_dax first, then add the charts and a short markdown summary of what they show. Use only fields that exist in the models.
${DAX_RULES}`,
    context: async () => withContexts(models().map((d) => `Semantic model "${d.name}" (dataset_id ${d.id})\n${d.info ? modelSummary(d.info) : "Structure not read."}`), models().map((d) => d.contextIds)),
    state: async () => researchText(get(), datasets()),
    tools: [
      daxTool(() => undefined, () => models()),
      tool({
        name: "add_text",
        description: "Adds a markdown block to the canvas (after a block id, else at the end).",
        args: z.object({ text: z.string(), after: z.string().optional() }),
        run: async ({ text, after }) => {
          const b: Block = { id: blockId(), kind: "markdown", text };
          await save({ blocks: insert(b, after) });
          return `Text block ${b.id} added.`;
        },
      }),
      tool({
        name: "add_chart",
        description: "Adds a chart computed by a DAX query (EVALUATE …) to the canvas. Returns the first rows.",
        args: z.object({ title: z.string(), chart: z.enum(CHARTS), dax: z.string(), dataset_id: z.string().optional(), after: z.string().optional() }),
        run: async ({ title, chart, dax, dataset_id, after }) => {
          const b = await computed({ id: blockId(), kind: "chart", title, chart, dax, datasetId: model(dataset_id).id });
          await save({ blocks: insert(b, after) });
          return `Chart ${b.id} added (${b.rows!.length} rows):\n${rowsText(b.rows!, 8)}`;
        },
      }),
      tool({
        name: "update_block",
        description: "Changes a block: text (markdown block), or title / chart / dax (chart, recomputed).",
        args: z.object({ id: z.string(), text: z.string().optional(), title: z.string().optional(), chart: z.enum(CHARTS).optional(), dax: z.string().optional() }),
        run: async ({ id, text, title, chart, dax }) => {
          const b = block(id);
          const next: Block = b.kind === "markdown" ? { ...b, text: text ?? b.text } : dax && dax !== b.dax ? await computed({ ...b, title: title ?? b.title, chart: chart ?? b.chart, dax }) : { ...b, title: title ?? b.title, chart: chart ?? b.chart };
          await save({ blocks: get().blocks.map((x) => (x.id === id ? next : x)) });
          return `Block ${id} updated.`;
        },
      }),
      tool({
        name: "move_block",
        description: "Moves a block to a position (0 = first).",
        args: z.object({ id: z.string(), to: z.number() }),
        run: async ({ id, to }) => {
          const b = block(id);
          const blocks = get().blocks.filter((x) => x !== b);
          blocks.splice(Math.max(0, Math.min(to, blocks.length)), 0, b);
          await save({ blocks });
          return `Block ${id} moved.`;
        },
      }),
      tool({
        name: "delete_block",
        description: "Deletes a canvas block.",
        args: z.object({ id: z.string() }),
        run: async ({ id }) => {
          block(id);
          await save({ blocks: get().blocks.filter((x) => x.id !== id) });
          return `Block ${id} deleted.`;
        },
      }),
      tool({
        name: "set_page",
        description: "Adds or replaces a page of a model's report: one full-page matrix (rows, columns, values, page filters).",
        args: pageArgs,
        run: async ({ dataset_id, title, rows, columns, values, filters }) => {
          const d = model(dataset_id);
          const page: MatrixPage = { title: title.trim(), rows, columns, values, filters };
          checkPage(page, d.info);
          const rep = get().reports[d.id] ?? { pages: [] };
          const pages = rep.pages.some((p) => p.title === page.title) ? rep.pages.map((p) => (p.title === page.title ? page : p)) : [...rep.pages, page];
          await save({ reports: { ...get().reports, [d.id]: { ...rep, pages, dirty: true } } });
          return `Page "${page.title}" set (written to the report when your answer is done).`;
        },
      }),
      tool({
        name: "delete_page",
        description: "Deletes a page of a model's report.",
        args: z.object({ dataset_id: z.string().optional(), title: z.string() }),
        run: async ({ dataset_id, title }) => {
          const d = model(dataset_id);
          const rep = get().reports[d.id];
          if (!rep?.pages.some((p) => p.title === title)) throw new Error(`No page "${title}".`);
          await save({ reports: { ...get().reports, [d.id]: { ...rep, pages: rep.pages.filter((p) => p.title !== title), dirty: true } } });
          return `Page "${title}" deleted.`;
        },
      }),
    ],
    finish: async () => {
      const done = await publishReports(get(), datasets());
      return done.length ? `✓ Report${done.length > 1 ? "s" : ""} written: ${done.join(", ")}.` : undefined;
    },
  };
}

/** Writes the reports whose pages changed (created the first time). */
async function publishReports(r: ResearchEntry, datasets: DatasetEntry[]) {
  const done: string[] = [];
  for (const id of r.datasetIds) {
    const rep = r.reports[id];
    const d = datasets.find((x) => x.id === id);
    if (!rep?.dirty || !d || !rep.pages.length) continue;
    const saved = await saveResearchReport(`${r.name} — ${d.name}`, d, rep, rep.pages);
    await patch("researches", r.id, { reports: { ...r.reports, [id]: saved } });
    r = { ...r, reports: { ...r.reports, [id]: saved } };
    done.push(d.name);
  }
  return done;
}

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
            <span className="text-[28px] font-semibold tracking-tight tabular-nums">{typeof rows[0][k] === "number" ? (rows[0][k] as number).toLocaleString() : String(rows[0][k] ?? "—")}</span>
          </div>
        ))}
      </div>
    );
  const axis = { tick: { fontSize: 11, fill: "var(--muted-foreground)" }, stroke: "var(--border)" };
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
            <Tooltip />
            <Legend wrapperStyle={{ fontSize: 12 }} />
          </PieChart>
        ) : block.chart === "line" ? (
          <LineChart data={data}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
            <XAxis dataKey={category} {...axis} />
            <YAxis tickFormatter={short} {...axis} width={48} />
            <Tooltip />
            {series.length > 1 && <Legend wrapperStyle={{ fontSize: 12 }} />}
            {series.map((k, i) => (
              <Line key={k} dataKey={k} stroke={PALETTE[i % PALETTE.length]} strokeWidth={2} dot={false} />
            ))}
          </LineChart>
        ) : block.chart === "area" ? (
          <AreaChart data={data}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
            <XAxis dataKey={category} {...axis} />
            <YAxis tickFormatter={short} {...axis} width={48} />
            <Tooltip />
            {series.map((k, i) => (
              <Area key={k} dataKey={k} stroke={PALETTE[i % PALETTE.length]} fill={PALETTE[i % PALETTE.length]} fillOpacity={0.2} />
            ))}
          </AreaChart>
        ) : (
          <BarChart data={data} layout={block.chart === "bar" ? "vertical" : "horizontal"}>
            <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
            {block.chart === "bar" ? <XAxis type="number" tickFormatter={short} {...axis} /> : <XAxis dataKey={category} {...axis} />}
            {block.chart === "bar" ? <YAxis type="category" dataKey={category} width={110} {...axis} /> : <YAxis tickFormatter={short} {...axis} width={48} />}
            <Tooltip cursor={{ fill: "var(--muted)" }} />
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
