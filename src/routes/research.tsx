import { useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Legend, Line, LineChart, Pie, PieChart, ResponsiveContainer, Scatter, ScatterChart, Tooltip, XAxis, YAxis } from "recharts";
import { CodeIcon, DatabaseIcon, FlaskConicalIcon, LayoutDashboardIcon, PencilIcon, PlusIcon, RefreshCwIcon, SlidersHorizontalIcon, Trash2Icon } from "lucide-react";
import { SCREEN_LABELS } from "@/config";
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
import { cardDax, newResearch, pivot, readResearch, runCard, type Card, type ChartCard, type PivotCard, type PivotSpec, type ResearchEntry } from "@/features/data-scientist/research";
import { patch, put, remove, useCollection, type DatasetEntry } from "@/lib/store";

export const Route = createFileRoute("/research")({
  validateSearch: (s: Record<string, unknown>): { research?: string; tab?: string } => ({ research: (s.research as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Research,
});

/**
 * Data Scientist: questions answered from semantic models in a canvas — the answer as markdown, then cards
 * (a pivot table or a chart each, fed by DAX on one model); Datasets: the models the assistant reads.
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
        { id: "canvas", label: "Canvas", icon: LayoutDashboardIcon, count: research?.cards.length || undefined, content: research ? <Canvas research={research} datasets={datasets} /> : <Empty icon={FlaskConicalIcon} action={<Button onClick={() => void create()}><PlusIcon /> New research</Button>}>A research answers questions from your semantic models: the answer as text, the figures as pivot tables and charts.</Empty> },
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
  const setCard = (card: Card) => save({ cards: research.cards.map((c) => (c.id === card.id ? card : c)) });
  if (!research.text && !research.cards.length && !editing)
    return (
      <Empty icon={LayoutDashboardIcon} action={<Button variant="outline" onClick={() => setEditing(true)}><PencilIcon /> Write a note</Button>}>
        {research.datasetIds.length ? "Ask the assistant a question: its answer comes here, with the pivot tables and charts that justify it." : "Choose the semantic models of this research in the Datasets tab, then ask a question."}
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
              <div className="markdown text-[14px]">{research.text ? <Markdown>{research.text}</Markdown> : <p className="text-muted-foreground">No text yet.</p>}</div>
              <Button variant="ghost" size="icon" className="absolute top-3 right-3 size-7 text-muted-foreground opacity-0 group-hover:opacity-100" aria-label="Edit the text" onClick={() => setEditing(true)}>
                <PencilIcon />
              </Button>
            </>
          )}
        </section>
        {/* Two columns; a card alone on the last row takes the whole width (no hole). */}
        {research.cards.length > 0 && (
          <div className="grid gap-4 md:grid-cols-2">
            {research.cards.map((c, i) => (
              <CardView key={c.id} wide={i === research.cards.length - 1 && research.cards.length % 2 === 1} card={c} dataset={datasets.find((d) => d.id === c.datasetId)} onChange={setCard} onDelete={() => save({ cards: research.cards.filter((x) => x.id !== c.id) })} />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function TextEditor({ text, onSave, onCancel }: { text: string; onSave: (text: string) => void; onCancel: () => void }) {
  const [draft, setDraft] = useState(text);
  return (
    <div className="flex flex-col gap-2">
      <Textarea className="min-h-60 font-mono text-[13px]" value={draft} onChange={(e) => setDraft(e.target.value)} aria-label="Text of the canvas (markdown)" />
      <div className="flex justify-end gap-2">
        <Button size="sm" variant="ghost" onClick={onCancel}>Cancel</Button>
        <Button size="sm" onClick={() => onSave(draft)}>Save</Button>
      </div>
    </div>
  );
}

/* ----------------------------------- Cards ----------------------------------- */

/** A card: its title and model, its chart or pivot, and its buttons (DAX or parameters, refresh, delete). */
function CardView({ card, dataset, wide, onChange, onDelete }: { card: Card; dataset?: DatasetEntry; wide?: boolean; onChange: (card: Card) => void; onDelete: () => void }) {
  const [busy, setBusy] = useState(false);
  const [dialog, setDialog] = useState(false);
  const refresh = async (c: Card = card) => {
    if (!dataset) return;
    setBusy(true);
    onChange(await runCard(c, { id: dataset.id, groupId: dataset.groupId, info: dataset.info }));
    setBusy(false);
  };
  return (
    <section className={`flex min-w-0 flex-col gap-3 rounded-xl border bg-card p-4 shadow-xs ${wide ? "md:col-span-2" : ""}`}>
      <div className="flex items-start gap-1">
        <div className="flex min-w-0 flex-1 flex-col">
          <h3 className="truncate text-[14px] font-semibold">{card.title}</h3>
          <span className="flex items-center gap-1 truncate text-[11.5px] text-muted-foreground">
            <DatabaseIcon className="size-3 shrink-0" /> {dataset?.name ?? "model not in the app"}
          </span>
        </div>
        <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label={card.kind === "chart" ? "View the DAX" : "Edit the parameters"} title={card.kind === "chart" ? "View the DAX" : "Edit the parameters"} onClick={() => setDialog(true)}>
          {card.kind === "chart" ? <CodeIcon /> : <SlidersHorizontalIcon />}
        </Button>
        <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Refresh" title="Run the query again" disabled={busy || !dataset} onClick={() => void refresh()}>
          <RefreshCwIcon className={busy ? "animate-spin" : ""} />
        </Button>
        <Button variant="ghost" size="icon" className="size-7 text-muted-foreground" aria-label="Delete the card" onClick={() => confirm(`Delete "${card.title}"?`) && onDelete()}>
          <Trash2Icon />
        </Button>
      </div>
      {card.error ? (
        <p className="rounded-lg bg-destructive/10 px-3 py-2 text-[12.5px] whitespace-pre-wrap text-destructive">{card.error}</p>
      ) : !card.rows ? (
        <p className="text-[12.5px] text-muted-foreground">Not run yet: refresh.</p>
      ) : card.kind === "chart" ? (
        <ChartView card={card} />
      ) : (
        <PivotView card={card} />
      )}
      {card.truncated && <p className="text-[11.5px] text-muted-foreground">Cut at the app's row limit ({card.rows?.length} rows).</p>}
      {card.kind === "pivot" && <SpecLine spec={card.spec} />}
      {dialog && card.kind === "chart" && <DaxDialog card={card} info={dataset?.info} onClose={() => setDialog(false)} />}
      {dialog && card.kind === "pivot" && <SpecDialog card={card} onClose={() => setDialog(false)} onSave={(spec) => (setDialog(false), void refresh({ ...card, spec }))} />}
    </section>
  );
}

const COLORS = ["var(--chart-1)", "var(--chart-2)", "var(--chart-3)", "var(--chart-4)", "var(--chart-5)"];
const number = (v: unknown) => (typeof v === "number" ? v : Number(v));
const short = (v: unknown) => {
  const n = number(v);
  if (!Number.isFinite(n)) return String(v ?? "");
  return Intl.NumberFormat(undefined, { notation: Math.abs(n) >= 10_000 ? "compact" : "standard", maximumFractionDigits: 2 }).format(n);
};

/** A chart of the card's rows (x: categories, y: one series per value column). */
function ChartView({ card }: { card: ChartCard }) {
  const data = (card.rows ?? []).map((r) => ({ ...r, [card.x]: r[card.x] ?? "(blank)", ...Object.fromEntries(card.y.map((y) => [y, number(r[y])])) }));
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

/** The card's rows as a pivot table: its row fields on the left, a column per column value and value. */
function PivotView({ card }: { card: PivotCard }) {
  const p = pivot(card.spec, card.rows ?? []);
  if (!p.lines.length) return <p className="text-[12.5px] text-muted-foreground">No rows.</p>;
  return (
    <div className="max-h-96 overflow-auto rounded-lg border">
      <table className="w-full text-[12.5px]">
        <thead className="sticky top-0 bg-muted">
          <tr>
            {p.rowLabels.map((l) => <th key={l} className="px-2.5 py-1.5 text-left font-medium whitespace-nowrap">{l}</th>)}
            {p.head.map((h) => <th key={h.id} className="px-2.5 py-1.5 text-right font-medium whitespace-nowrap">{h.label}</th>)}
          </tr>
        </thead>
        <tbody className="divide-y">
          {p.lines.map((line, i) => (
            <tr key={i}>
              {line.labels.map((l, j) => <td key={j} className="px-2.5 py-1.5 whitespace-nowrap">{l}</td>)}
              {p.head.map((h) => <td key={h.id} className="px-2.5 py-1.5 text-right tabular-nums">{line.cells.has(h.id) ? short(line.cells.get(h.id)) : ""}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
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

function DaxDialog({ card, info, onClose }: { card: ChartCard; info?: DatasetEntry["info"]; onClose: () => void }) {
  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle>{card.title} — DAX</DialogTitle>
        </DialogHeader>
        <pre className="max-h-96 overflow-auto rounded-lg bg-muted p-3 font-mono text-[12px] whitespace-pre-wrap">{cardDax(card, info)}</pre>
        <p className="text-[12px] text-muted-foreground">x: {card.x} · y: {card.y.join(", ")} · {card.chart}</p>
      </DialogContent>
    </Dialog>
  );
}

/** Edits a pivot's parameters (fields separated by commas; filters as JSON). */
function SpecDialog({ card, onClose, onSave }: { card: PivotCard; onClose: () => void; onSave: (spec: PivotSpec) => void }) {
  const [draft, setDraft] = useState({ rows: card.spec.rows.join(", "), columns: card.spec.columns.join(", "), values: card.spec.values.join(", "), filters: JSON.stringify(card.spec.filters, null, 1) });
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
          <DialogTitle>{card.title} — parameters</DialogTitle>
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
