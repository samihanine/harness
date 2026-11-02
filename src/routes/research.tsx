import { useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Legend, Line, LineChart, Pie, PieChart, ResponsiveContainer, Scatter, ScatterChart, Tooltip, XAxis, YAxis } from "recharts";
import { BarChart3Icon, CalculatorIcon, CodeIcon, DatabaseIcon, FileSpreadsheetIcon, FlaskConicalIcon, LayoutDashboardIcon, PencilIcon, PlusIcon, RefreshCwIcon, SlidersHorizontalIcon, TableIcon, Trash2Icon } from "lucide-react";
import { EXPORT, SCREEN_LABELS } from "@/config";
import { withToast } from "@/components/entity";
import { powerBiToken } from "@/services/ms";
import { researchWorkbook, powerBiConnection, tenantOf, type PivotPart } from "@/features/data-scientist/pivot-excel";
import { columnName, computeWorkbook, type Result } from "@/features/data-scientist/formulas";
import { DatasetList } from "@/components/dataset-list";
import { EntityMenu } from "@/components/entity";
import { Markdown } from "@/components/markdown";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { EntityPicker, Empty, Workspace } from "@/components/workspace";
import { researchAgent } from "@/features/data-scientist/research-tools";
import { useLast } from "@/lib/last";
import { chartAxes, newResearch, pivot, readResearch, resultName, runSheet, sheetDax, type CalcSheet, type ChartSheet, type PivotSheet, type PivotSpec, type ResearchEntry, type Sheet } from "@/features/data-scientist/research";
import { patch, put, remove, useCollection, type DatasetEntry } from "@/lib/store";

export const Route = createFileRoute("/research")({
  validateSearch: (s: Record<string, unknown>): { research?: string; tab?: string } => ({ research: (s.research as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Research,
});

/**
 * Data Scientist: questions answered from semantic models — the analysis as markdown on top, and below it
 * a small workbook shown as a simplified Excel window (sheets as tabs at the bottom: pivots, charts of a
 * pivot, calculations with Excel formulas), downloadable with every pivot connected to its model.
 */
function Research() {
  const { research: researchId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/research" });
  useLast("research:research", researchId, (research) => navigate({ search: { research }, replace: true }));
  const researches = useCollection("researches").map(readResearch);
  const datasets = useCollection("datasets");
  const research = researches.find((r) => r.id === researchId);
  const agent = research && researchAgent(() => readResearch(researches.find((r) => r.id === research.id) ?? research), () => datasets);
  const create = async () => {
    const name = prompt("Name of the research", "New research")?.trim();
    if (name) navigate({ search: { research: (await put("researches", newResearch(name))).id, tab: "datasets" } });
  };
  const toggle = (id: string, on: boolean) => research && void patch("researches", research.id, { datasetIds: on ? [...research.datasetIds, id] : research.datasetIds.filter((x) => x !== id) });

  return (
    <Workspace
      title={SCREEN_LABELS.research}
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
            onRemove={research && (() => confirm(`Delete the research "${research.name}"?`) && void remove("researches", research.id).then(() => navigate({ search: {} })))}
          />
        </>
      }
      ai={{ scope: `research:${research?.id}`, agent, placeholder: "Ask a question about your data…", empty: "Choose a research, or start a new one." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        { id: "canvas", label: "Canvas", icon: LayoutDashboardIcon, count: research?.sheets.length || undefined, content: research ? <Canvas research={research} datasets={datasets} /> : <Empty icon={FlaskConicalIcon} action={<Button onClick={() => void create()}><PlusIcon /> New research</Button>}>A research answers questions from your semantic models: the analysis as text, the figures in a workbook of pivots, charts and calculations.</Empty> },
        {
          id: "datasets",
          label: "Datasets",
          icon: DatabaseIcon,
          count: research?.datasetIds.length,
          content: research ? <DatasetList title="Semantic models of this research" description="The assistant queries these models (DAX)." datasets={datasets} chosen={research.datasetIds} onToggle={toggle} /> : <Empty>Choose a research.</Empty>,
        },
      ]}
    />
  );
}

/* ---------------------------------- Canvas ---------------------------------- */

function Canvas({ research, datasets }: { research: ResearchEntry; datasets: DatasetEntry[] }) {
  const [editing, setEditing] = useState(false);
  const save = (change: Partial<ResearchEntry>) => void patch("researches", research.id, { ...change, updatedAt: Date.now() });
  if (!research.text && !research.sheets.length && !editing)
    return (
      <Empty icon={LayoutDashboardIcon} action={<Button variant="outline" onClick={() => setEditing(true)}><PencilIcon /> Write a note</Button>}>
        {research.datasetIds.length ? "Ask the assistant a question: it explores the models, proposes how to answer, then writes the analysis here with the pivots, charts and calculations behind it." : "Choose the semantic models of this research in the Datasets tab, then ask a question."}
      </Empty>
    );
  return (
    <div className="min-h-0 flex-1 overflow-y-auto bg-muted/30">
      <div className="mx-auto flex max-w-7xl flex-col gap-5 p-6">
        <section className="group relative rounded-xl border bg-card p-6 shadow-xs">
          {editing ? (
            <TextEditor text={research.text} onSave={(text) => (save({ text }), setEditing(false))} onCancel={() => setEditing(false)} />
          ) : (
            <>
              <div className="markdown text-[14px]">{research.text ? <Markdown>{research.text}</Markdown> : <p className="text-muted-foreground">No analysis yet.</p>}</div>
              <Button variant="ghost" size="icon" className="absolute top-3 right-3 size-7 text-muted-foreground opacity-0 group-hover:opacity-100" aria-label="Edit the analysis" onClick={() => setEditing(true)}>
                <PencilIcon />
              </Button>
            </>
          )}
        </section>
        {research.sheets.length > 0 && <Workbook research={research} datasets={datasets} onSheets={(sheets) => save({ sheets })} />}
      </div>
    </div>
  );
}

function TextEditor({ text, onSave, onCancel }: { text: string; onSave: (text: string) => void; onCancel: () => void }) {
  const [draft, setDraft] = useState(text);
  return (
    <div className="flex flex-col gap-2">
      <Textarea className="min-h-60 font-mono text-[13px]" value={draft} onChange={(e) => setDraft(e.target.value)} aria-label="Analysis (markdown)" />
      <div className="flex justify-end gap-2">
        <Button size="sm" variant="ghost" onClick={onCancel}>Cancel</Button>
        <Button size="sm" onClick={() => onSave(draft)}>Save</Button>
      </div>
    </div>
  );
}

/* ---------------------------------- Workbook ---------------------------------- */

const SHEET_ICONS = { pivot: TableIcon, chart: BarChart3Icon, calc: CalculatorIcon } as const;
/** A chart of an older canvas (a free DAX query) as a pivot for Excel: its x as rows, its y as values. */
const specOf = (s: PivotSheet | ChartSheet): PivotSpec => s.spec ?? (s.kind === "chart" ? { rows: s.x ? [s.x] : [], columns: [], values: s.y ?? [], filters: [] } : { rows: [], columns: [], values: [], filters: [] });

/** The workbook as an Excel file: every pivot (and chart) connected to its model, the calculations with their formulas. */
async function downloadWorkbook(research: ResearchEntry, datasets: DatasetEntry[]) {
  const tenant = tenantOf(await powerBiToken().catch(() => ""));
  const results = computeWorkbook(research.sheets);
  const sheets = research.sheets.flatMap((s): Parameters<typeof researchWorkbook>[0] => {
    if (s.kind === "calc") return [{ kind: "calc", sheet: s }];
    const d = datasets.find((x) => x.id === s.datasetId);
    if (!d || !s.rows) return [];
    const part: PivotPart = { name: s.name, spec: specOf(s), rows: s.rows, dax: sheetDax(s, d.info), connection: powerBiConnection(d.id, tenant), model: d.name, note: s.kind === "chart" ? `${s.chart} chart in the app` : s.note };
    return [{ kind: "pivot", part }];
  });
  const blob = await researchWorkbook(sheets, results);
  const url = URL.createObjectURL(blob);
  Object.assign(document.createElement("a"), { href: url, download: `${research.name.replace(/[^\p{L}\p{N} _-]+/gu, "-") || "research"}.xlsx` }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** The workbook as a simplified Excel window: a toolbar, the sheet shown, and the sheets as tabs at the bottom. */
function Workbook({ research, datasets, onSheets }: { research: ResearchEntry; datasets: DatasetEntry[]; onSheets: (sheets: Sheet[]) => void }) {
  const [shownId, setShown] = useState<string>();
  const [busy, setBusy] = useState(false);
  const [dialog, setDialog] = useState(false);
  const sheet = research.sheets.find((s) => s.id === shownId) ?? research.sheets[0];
  const results = computeWorkbook(research.sheets);
  const dataset = sheet.kind !== "calc" ? datasets.find((d) => d.id === sheet.datasetId) : undefined;
  const setSheet = (next: Sheet) => onSheets(research.sheets.map((s) => (s.id === next.id ? next : s)));
  const refresh = async (s: PivotSheet | ChartSheet = sheet as PivotSheet | ChartSheet) => {
    const d = datasets.find((x) => x.id === s.datasetId);
    if (!d) return;
    setBusy(true);
    setSheet(await runSheet(s, { id: d.id, groupId: d.groupId, info: d.info }));
    setBusy(false);
  };
  const refreshAll = async () => {
    setBusy(true);
    const next = await Promise.all(research.sheets.map(async (s) => {
      if (s.kind === "calc") return s;
      const d = datasets.find((x) => x.id === s.datasetId);
      return d ? runSheet(s, { id: d.id, groupId: d.groupId, info: d.info }) : s;
    }));
    onSheets(next);
    setBusy(false);
  };
  const Icon = SHEET_ICONS[sheet.kind];
  return (
    <section className="flex min-w-0 flex-col overflow-hidden rounded-xl border bg-card shadow-xs">
      <div className="flex items-center gap-1 border-b bg-muted/40 px-3 py-2">
        <FileSpreadsheetIcon className="size-4 shrink-0 text-emerald-600" />
        <span className="ml-1 truncate text-[13px] font-semibold">{sheet.name}</span>
        <span className="ml-2 flex min-w-0 items-center gap-1 truncate text-[11.5px] text-muted-foreground">
          <Icon className="size-3 shrink-0" /> {sheet.kind === "calc" ? "calculation" : sheet.kind === "chart" ? `${sheet.chart} chart` : "pivot"}
          {sheet.kind !== "calc" && <><span className="opacity-50">·</span><DatabaseIcon className="size-3 shrink-0" /> {dataset?.name ?? "model not in the app"}</>}
        </span>
        <span className="flex-1" />
        {sheet.kind !== "calc" && (
          <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label={sheet.spec ? "Edit the parameters" : "View the DAX"} title={sheet.spec ? "Edit the parameters" : "View the DAX"} onClick={() => setDialog(true)}>
            {sheet.spec ? <SlidersHorizontalIcon /> : <CodeIcon />}
          </Button>
        )}
        <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Refresh every sheet" title="Run every sheet's query again" disabled={busy} onClick={() => void refreshAll()}>
          <RefreshCwIcon className={busy ? "animate-spin" : ""} />
        </Button>
        {EXPORT.pivotExcel && (
          <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Download as Excel" title="Excel: every sheet, each pivot connected to its model (Data › Refresh All)" onClick={() => void withToast("Writing the Excel file…", downloadWorkbook(research, datasets), "Downloaded.")}>
            <FileSpreadsheetIcon />
          </Button>
        )}
        <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Delete the sheet" onClick={() => confirm(`Delete the sheet "${sheet.name}"?`) && onSheets(research.sheets.filter((s) => s.id !== sheet.id))}>
          <Trash2Icon />
        </Button>
      </div>
      <div className="min-h-64">
        {sheet.kind === "calc" ? (
          <CalcView sheet={sheet} results={results.get(sheet.name) ?? []} />
        ) : sheet.error ? (
          <p className="m-4 rounded-lg bg-destructive/10 px-3 py-2 text-[12.5px] whitespace-pre-wrap text-destructive">{sheet.error}</p>
        ) : !sheet.rows ? (
          <p className="m-4 text-[12.5px] text-muted-foreground">Not run yet: refresh.</p>
        ) : sheet.kind === "chart" ? (
          <div className="p-4"><ChartView sheet={sheet} /></div>
        ) : (
          <PivotView sheet={sheet} />
        )}
      </div>
      {sheet.kind !== "calc" && (
        <div className="border-t px-3 py-2">
          {sheet.truncated && <p className="mb-1 text-[11.5px] text-muted-foreground">Cut at the app's row limit ({sheet.rows?.length} rows).</p>}
          <SpecLine spec={specOf(sheet)} />
        </div>
      )}
      {sheet.note && <p className="border-t px-3 py-2 text-[12px] text-muted-foreground">{sheet.note}</p>}
      {/* The sheets as Excel's tabs, at the bottom. */}
      <div className="flex items-stretch gap-0.5 overflow-x-auto border-t bg-muted/60 px-2" role="tablist" aria-label="Sheets">
        {research.sheets.map((s) => {
          const I = SHEET_ICONS[s.kind];
          const on = s.id === sheet.id;
          return (
            <button key={s.id} type="button" role="tab" aria-selected={on} onClick={() => setShown(s.id)} className={`-mt-px flex shrink-0 cursor-pointer items-center gap-1.5 border-x border-b px-3 py-1.5 text-[12px] transition-colors ${on ? "border-border bg-card font-semibold text-emerald-700 dark:text-emerald-400" : "border-transparent text-muted-foreground hover:bg-card/60"}`}>
              <I className="size-3.5" /> {s.name}
              {s.kind !== "calc" && s.error && <span className="size-1.5 rounded-full bg-destructive" />}
            </button>
          );
        })}
      </div>
      {dialog && sheet.kind !== "calc" && !sheet.spec && <DaxDialog sheet={sheet} info={dataset?.info} onClose={() => setDialog(false)} />}
      {dialog && sheet.kind !== "calc" && sheet.spec && <SpecDialog title={sheet.name} spec={sheet.spec} onClose={() => setDialog(false)} onSave={(spec) => (setDialog(false), void refresh({ ...sheet, spec }))} />}
    </section>
  );
}

/** A figure in a grid cell, whole as in Excel (grouped, at most 4 decimals). */
const full = (v: number) => Intl.NumberFormat(undefined, { maximumFractionDigits: 4 }).format(v);

/** A grid as in Excel: lettered columns, numbered rows; `head` rows in bold; a cell selected shows its formula. */
function Grid({ cells, head = 1, formulas, errors }: { cells: unknown[][]; head?: number; formulas?: string[][]; errors?: (string | undefined)[][] }) {
  const [selected, setSelected] = useState<[number, number]>();
  const width = Math.max(1, ...cells.map((r) => r.length));
  const raw = selected && formulas?.[selected[0]]?.[selected[1]];
  const why = selected && errors?.[selected[0]]?.[selected[1]];
  return (
    <div className="flex flex-col">
      {formulas && (
        <div className="flex items-center gap-2 border-b px-3 py-1.5 font-mono text-[12px]">
          <span className="w-12 shrink-0 text-muted-foreground">{selected ? `${columnName(selected[1])}${selected[0] + 1}` : ""}</span>
          <span className="text-muted-foreground italic">fx</span>
          <span className="min-w-0 flex-1 truncate" title={raw ?? ""}>{raw ?? ""}</span>
          {why && <span className="truncate text-destructive" title={why}>{why}</span>}
        </div>
      )}
      <div className="max-h-[28rem] overflow-auto">
        <table className="border-collapse text-[12.5px]">
          <thead className="sticky top-0 z-10">
            <tr>
              <th className="sticky left-0 z-20 w-10 border-r border-b bg-muted" />
              {Array.from({ length: width }, (_, c) => <th key={c} className="min-w-24 border-r border-b bg-muted px-2 py-0.5 text-center text-[11px] font-medium text-muted-foreground">{columnName(c)}</th>)}
            </tr>
          </thead>
          <tbody>
            {cells.map((line, r) => (
              <tr key={r}>
                <td className="sticky left-0 border-r border-b bg-muted px-2 text-center text-[11px] text-muted-foreground tabular-nums">{r + 1}</td>
                {Array.from({ length: width }, (_, c) => {
                  const v = line[c];
                  const on = selected?.[0] === r && selected?.[1] === c;
                  const error = errors?.[r]?.[c];
                  return (
                    <td key={c} onClick={() => setSelected([r, c])} className={`cursor-cell border-r border-b px-2 py-1 whitespace-nowrap ${r < head ? "bg-muted/40 font-semibold" : ""} ${typeof v === "number" ? "text-right tabular-nums" : ""} ${error ? "text-destructive" : ""} ${on ? "outline-2 -outline-offset-2 outline-emerald-600" : ""}`}>
                      {typeof v === "number" ? full(v) : v === null || v === undefined ? "" : String(v)}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/** A calculation sheet: its computed cells, the formula of the selected one above (as Excel's formula bar). */
function CalcView({ sheet, results }: { sheet: CalcSheet; results: Result[][] }) {
  return <Grid cells={results.map((line) => line.map((c) => c.value))} formulas={sheet.cells} errors={results.map((line) => line.map((c) => (c.error ? `${c.error} ${c.why ?? ""}` : undefined)))} />;
}

/** A pivot sheet: its row fields on the left, a column per column value and value (as the PivotTable of the download). */
function PivotView({ sheet }: { sheet: PivotSheet }) {
  const p = pivot(sheet.spec, sheet.rows ?? []);
  if (!p.lines.length) return <p className="m-4 text-[12.5px] text-muted-foreground">No rows.</p>;
  const cells = [[...p.rowLabels, ...p.head.map((h) => h.label || "Value")], ...p.lines.map((l) => [...l.labels.map((x) => x || "(blank)"), ...p.head.map((h) => (l.cells.has(h.id) ? Number(l.cells.get(h.id)) : null))])];
  return <Grid cells={cells} />;
}

const COLORS = ["var(--chart-1)", "var(--chart-2)", "var(--chart-3)", "var(--chart-4)", "var(--chart-5)"];
const number = (v: unknown) => (typeof v === "number" ? v : Number(v));
const short = (v: unknown) => {
  const n = number(v);
  if (!Number.isFinite(n)) return String(v ?? "");
  return Intl.NumberFormat(undefined, { notation: Math.abs(n) >= 10_000 ? "compact" : "standard", maximumFractionDigits: 2 }).format(n);
};

/** A chart sheet: a pivot's rows as categories (x), its columns and values as series (an older canvas: its query's x and y). */
function ChartView({ sheet }: { sheet: ChartSheet }) {
  const axes0 = chartAxes(sheet);
  const card = { chart: sheet.chart, x: axes0.x, y: axes0.y };
  const data = (axes0.data ?? sheet.rows ?? []).map((r) => ({ ...r, [card.x]: r[card.x] ?? "(blank)", ...Object.fromEntries(card.y.map((y) => [y, number(r[y])])) }));
  const axes = [
    <CartesianGrid key="g" strokeDasharray="3 3" stroke="var(--border)" vertical={false} />,
    <Tooltip key="t" formatter={(v) => short(v)} contentStyle={{ background: "var(--popover)", border: "1px solid var(--border)", borderRadius: 8, fontSize: 12 }} />,
    ...(card.y.length > 1 ? [<Legend key="l" wrapperStyle={{ fontSize: 12 }} />] : []),
  ];
  const tick = { fontSize: 11, fill: "var(--muted-foreground)" };
  const chart = (() => {
    switch (card.chart) {
      case "pie":
        return (
          <PieChart>
            <Tooltip formatter={(v) => short(v)} />
            <Pie data={data} dataKey={card.y[0]} nameKey={card.x} innerRadius="45%" outerRadius="80%" paddingAngle={1} label={(p: { name?: string }) => String(p.name ?? "")} labelLine={false} fontSize={11}>
              {data.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
            </Pie>
          </PieChart>
        );
      case "line":
        return (
          <LineChart data={data}>
            {axes}
            <XAxis dataKey={card.x} tick={tick} />
            <YAxis tick={tick} tickFormatter={short} width={56} />
            {card.y.map((y, i) => <Line key={y} dataKey={y} stroke={COLORS[i % COLORS.length]} strokeWidth={2} dot={false} />)}
          </LineChart>
        );
      case "area":
        return (
          <AreaChart data={data}>
            {axes}
            <XAxis dataKey={card.x} tick={tick} />
            <YAxis tick={tick} tickFormatter={short} width={56} />
            {card.y.map((y, i) => <Area key={y} dataKey={y} stroke={COLORS[i % COLORS.length]} fill={COLORS[i % COLORS.length]} fillOpacity={0.15} />)}
          </AreaChart>
        );
      case "scatter":
        return (
          <ScatterChart>
            {axes}
            <XAxis dataKey={card.y[0]} name={card.y[0]} type="number" tick={tick} tickFormatter={short} />
            <YAxis dataKey={card.y[1] ?? card.y[0]} name={card.y[1] ?? card.y[0]} type="number" tick={tick} tickFormatter={short} width={56} />
            <Scatter data={data} fill={COLORS[0]} />
          </ScatterChart>
        );
      case "bar":
        return (
          <BarChart data={data} layout="vertical">
            {axes}
            <XAxis type="number" tick={tick} tickFormatter={short} />
            <YAxis type="category" dataKey={card.x} tick={tick} width={110} />
            {card.y.map((y, i) => <Bar key={y} dataKey={y} fill={COLORS[i % COLORS.length]} radius={[0, 4, 4, 0]} />)}
          </BarChart>
        );
      default:
        return (
          <BarChart data={data}>
            {axes}
            <XAxis dataKey={card.x} tick={tick} />
            <YAxis tick={tick} tickFormatter={short} width={56} />
            {card.y.map((y, i) => <Bar key={y} dataKey={y} fill={COLORS[i % COLORS.length]} radius={[4, 4, 0, 0]} />)}
          </BarChart>
        );
    }
  })();
  return (
    <div className="h-72 w-full">
      <ResponsiveContainer width="100%" height="100%">
        {chart}
      </ResponsiveContainer>
    </div>
  );
}

const filterText = (f: PivotSpec["filters"][number]) => `${f.field} ${f.op ?? "in"} ${f.values ? f.values.join(", ") : `${f.value ?? ""}${f.to !== undefined ? ` – ${f.to}` : ""}`}`;

/** A pivot's parameters, under it. */
function SpecLine({ spec }: { spec: PivotSpec }) {
  const parts: [string, string][] = [
    ["Rows", spec.rows.join(", ")],
    ["Columns", spec.columns.join(", ")],
    ["Values", spec.values.join(", ")],
    ["Filters", spec.filters.map(filterText).join("; ")],
  ];
  return (
    <dl className="grid grid-cols-[auto_1fr] gap-x-3 gap-y-0.5 border-t pt-2 font-mono text-[11px] text-muted-foreground">
      {parts.filter(([, v]) => v).map(([k, v]) => (
        <div key={k} className="contents">
          <dt className="font-sans font-medium">{k}</dt>
          <dd className="truncate" title={v}>{v}</dd>
        </div>
      ))}
    </dl>
  );
}

function DaxDialog({ sheet, info, onClose }: { sheet: PivotSheet | ChartSheet; info?: DatasetEntry["info"]; onClose: () => void }) {
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle>{sheet.name} — DAX</DialogTitle>
        </DialogHeader>
        <pre className="max-h-96 overflow-auto rounded-lg bg-muted p-3 font-mono text-[12px] whitespace-pre-wrap">{sheetDax(sheet, info)}</pre>
      </DialogContent>
    </Dialog>
  );
}

/** Edits a pivot's parameters (fields separated by commas; filters as JSON). */
function SpecDialog({ title, spec, onClose, onSave }: { title: string; spec: PivotSpec; onClose: () => void; onSave: (spec: PivotSpec) => void }) {
  const [draft, setDraft] = useState({ rows: spec.rows.join(", "), columns: spec.columns.join(", "), values: spec.values.join(", "), filters: JSON.stringify(spec.filters, null, 1) });
  const [error, setError] = useState<string>();
  const list = (s: string) => s.split(/,(?![^[]*\])/).map((x) => x.trim()).filter(Boolean);
  const save = () => {
    try {
      const filters = draft.filters.trim() ? JSON.parse(draft.filters) : [];
      if (!Array.isArray(filters)) throw new Error("Filters: a JSON list.");
      if (!list(draft.values).length) throw new Error("Give at least one value.");
      onSave({ rows: list(draft.rows), columns: list(draft.columns), values: list(draft.values), filters });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="sm:max-w-xl">
        <DialogHeader>
          <DialogTitle>{title} — parameters</DialogTitle>
        </DialogHeader>
        <div className="flex flex-col gap-3 text-[13px]">
          {(["rows", "columns", "values"] as const).map((k) => (
            <label key={k} className="flex flex-col gap-1">
              <span className="font-medium capitalize">{k}</span>
              <Input className="font-mono text-[12.5px]" value={draft[k]} placeholder={k === "values" ? "[Total Sales], Sum('Sales'[Qty])" : "'Date'[Year]"} onChange={(e) => setDraft({ ...draft, [k]: e.target.value })} />
            </label>
          ))}
          <label className="flex flex-col gap-1">
            <span className="font-medium">Filters</span>
            <Textarea className="min-h-24 font-mono text-[12px]" value={draft.filters} placeholder={'[{"field": "\'Date\'[Year]", "values": [2024]}]'} onChange={(e) => setDraft({ ...draft, filters: e.target.value })} />
          </label>
          {error && <p className="text-[12.5px] text-destructive">{error}</p>}
        </div>
        <DialogFooter>
          <Button variant="ghost" onClick={onClose}>Cancel</Button>
          <Button onClick={save}>Apply and refresh</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
