import { useEffect, useMemo, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { CodeIcon, PlusIcon, SigmaIcon, XIcon } from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures, useLatest } from "@/agent/server";
import { NeedsDb } from "@/components/guards";
import { IconButton } from "@/components/icon-button";
import { ModelTree } from "@/components/model-tree";
import type { PickedField } from "@/components/model-tree";
import { formatCell } from "@/components/result-table";
import { SimpleSelect } from "@/components/simple-select";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Switch } from "@/components/ui/switch";
import { useDb } from "@/db/db";
import { useSources } from "@/lib/use-source";
import { cn } from "@/lib/utils";
import { modelSummary } from "@/pbi/model";
import type { Source } from "@/query/engine";
import { runQuery, toCsv } from "@/query/engine";
import { AGGREGATES, compile, emptyPivot, layout, pivotSchema, runPivot } from "@/query/pivot";
import type { Pivot, PivotResult } from "@/query/pivot";

export const Route = createFileRoute("/research")({
  component: () => (
    <NeedsDb>
      <Research />
    </NeedsDb>
  ),
});

type Saved = { sourceIds: string[]; pivot: Pivot; allowMeasures: boolean };
const KEY = "pbi-agent:research";
const load = (): Saved => ({ sourceIds: [], pivot: emptyPivot(), allowMeasures: true, ...JSON.parse(localStorage.getItem(KEY) ?? "{}") });

const NUMERIC = /int|double|decimal|currency/i;

function Research() {
  const data = useDb();
  const [state, setState] = useState<Saved>(load);
  const { sourceIds, pivot, allowMeasures } = state;
  const { sources, error: sourceError } = useSources(sourceIds);
  const [result, setResult] = useState<PivotResult | null>(null);
  const [running, setRunning] = useState(false);
  const [showQueries, setShowQueries] = useState(false);
  const [measureOpen, setMeasureOpen] = useState(false);
  const update = (patch: Partial<Saved>) => setState((s) => ({ ...s, ...patch }));
  const setPivot = (next: Pivot) => update({ pivot: next });
  const latest = useLatest({ sources, pivot, result, allowMeasures, sourceIds });

  useEffect(() => localStorage.setItem(KEY, JSON.stringify(state)), [state]);

  // Re-run when the pivot or the sources change.
  const runKey = JSON.stringify(pivot) + sources.map((s) => s.id).join();
  useEffect(() => {
    if (!pivot.values.length || !sources.length) return setResult(null);
    let cancelled = false;
    setRunning(true);
    const timer = setTimeout(
      () =>
        void runPivot(pivot, sources)
          .then((r) => !cancelled && setResult(r))
          .catch((e: Error) => !cancelled && setResult({ rows: [], queries: [{ source: "", query: "", error: e.message }] }))
          .finally(() => !cancelled && setRunning(false)),
      300,
    );
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [runKey]);

  /** Adds a picked field to a zone; rows/columns are bound to the same column name in the other sources when it exists. */
  const add = (source: Source, field: PickedField, zone: "rows" | "columns" | "values" | "filters", aggregate?: (typeof AGGREGATES)[number]) => {
    if (zone === "values") {
      const label = uniqueLabel(pivot, `${field.name}${sources.length > 1 ? ` (${source.title})` : ""}`);
      return setPivot({ ...pivot, values: [...pivot.values, { label, source: source.id, field: field.ref, aggregate: field.kind === "measure" ? undefined : aggregate }] });
    }
    if (zone === "filters") return setPivot({ ...pivot, filters: [...pivot.filters, { source: source.id, field: field.ref, values: [] }] });
    const bindings: Record<string, string> = { [source.id]: field.ref };
    for (const other of sources)
      if (other.id !== source.id) {
        const match = other.model.tables.flatMap((t) => t.columns.map((c) => ({ t, c }))).find(({ c }) => c.name.toLowerCase() === field.name.toLowerCase());
        if (match) bindings[other.id] = `'${match.t.name}'[${match.c.name}]`;
      }
    setPivot({ ...pivot, [zone]: [...pivot[zone], { label: uniqueLabel(pivot, field.name), bindings }] });
  };

  useAgentFeatures(
    `Researcher: the user compares data from several sources with a pivot table. Sources (research://sources) are datasets of the database; remote ones speak DAX, local ones SQL.
The pivot (research://pivot) has rows/columns (labels bound to one column per source: "bindings": {sourceId: "'Table'[Column]"}), values (per source: a measure "[Name]" or a column with an aggregate, or a native expression${allowMeasures ? "; computed values combine other values by label, e.g. {Sales A} - {Sales B}, DIVIDE({x}, {y}), ABS(…)" : ""}) and filters (per source, list of values).
To compare two sources, bind each row label to the matching column of both sources and add one value per source${allowMeasures ? " plus a computed gap" : ""}.${allowMeasures ? "" : " New measures (expression / computed) are disabled by the user."}`,
    (server) => [
      server.registerResource(
        "sources",
        "research://sources",
        { title: "Sources and their models", mimeType: "text/plain", annotations: { priority: 1 } },
        async (uri) => ({
          contents: [
            {
              uri: uri.href,
              text:
                latest.current.sources.map((s) => `## Source ${s.id}: ${s.title} (${s.language})${s.context ? `\nContext: ${s.context}` : ""}\n${modelSummary(s.model, { expressions: false })}`).join("\n\n") ||
                "No source added yet.",
            },
          ],
        }),
      ),
      server.registerResource(
        "available",
        "research://datasets",
        { title: "Datasets that can be added as sources", mimeType: "application/jsonl", annotations: { priority: 0.5 } },
        async (uri) => ({ contents: [{ uri: uri.href, text: data.rows("datasets").map((d) => JSON.stringify({ id: d.id, title: d.title, source: d.source, context: d.context })).join("\n") }] }),
      ),
      server.registerResource(
        "pivot",
        "research://pivot",
        { title: "Current pivot and its result", mimeType: "text/plain", annotations: { priority: 1 } },
        async (uri) => {
          const { pivot: p, result: r } = latest.current;
          const csv = r ? toCsv({ columns: [...[...p.rows, ...p.columns].map((d) => d.label), ...p.values.map((v) => v.label)], rows: r.rows.map((row) => Object.fromEntries([...[...p.rows, ...p.columns].map((d, i) => [d.label, row[`d${i}`]]), ...p.values.map((v, i) => [v.label, row[`v${i}`]])])) }, 40) : "(no result)";
          const errors = r?.queries.filter((q) => q.error).map((q) => `${q.source}: ${q.error}`) ?? [];
          return { contents: [{ uri: uri.href, text: `Pivot:\n${JSON.stringify(p)}\n\nResult (${r?.rows.length ?? 0} rows):\n${csv}${errors.length ? `\n\nErrors:\n${errors.join("\n")}` : ""}` }] };
        },
      ),
      server.registerTool(
        "set_sources",
        { description: "Sets the sources (dataset ids from research://datasets).", inputSchema: z.object({ ids: z.array(z.string()) }) },
        async ({ ids }) => {
          const unknown = ids.filter((id) => !data.row("datasets", id));
          if (unknown.length) return errorResult(`Unknown dataset(s): ${unknown.join(", ")}`);
          const before = latest.current.sourceIds;
          update({ sourceIds: ids });
          return textResult("Sources set; read research://sources for their models.", { name: "set_sources", arguments: { ids: before } });
        },
      ),
      server.registerTool(
        "set_pivot",
        { description: "Replaces the pivot (rows, columns, values, filters). The result appears in research://pivot.", inputSchema: pivotSchema },
        async (next) => {
          if (!latest.current.allowMeasures && next.values.some((v) => v.expression || v.computed)) return errorResult("New measures are disabled: use model measures or aggregated columns only.");
          try {
            next.values.forEach((v) => v.computed && compile(v.computed, next));
          } catch (e) {
            return errorResult(e);
          }
          const before = latest.current.pivot;
          setPivot(next);
          const result = await runPivot(next, latest.current.sources);
          const errors = result.queries.filter((q) => q.error).map((q) => `${q.source}: ${q.error}`);
          return textResult(`${result.rows.length} rows.${errors.length ? ` Errors: ${errors.join(" | ")}` : ""} See research://pivot.`, { name: "set_pivot", arguments: { pivot: before } });
        },
      ),
      server.registerTool(
        "run_query",
        { description: "Runs a query on one source (DAX or SQL by its language) to explore data; returns CSV.", inputSchema: z.object({ source: z.string(), query: z.string() }), annotations: { readOnlyHint: true } },
        async ({ source, query }) => {
          const s = latest.current.sources.find((x) => x.id === source);
          if (!s) return errorResult(`Source ${source} is not added.`);
          try {
            const r = await runQuery(s, query);
            return textResult(`${r.rows.length} rows\n${toCsv(r, 200)}`);
          } catch (e) {
            return errorResult(e);
          }
        },
      ),
    ],
    [],
  );

  const datasets = data.rows("datasets");
  return (
    <div className="flex h-full">
      <div className="flex w-64 shrink-0 flex-col border-r">
        <div className="flex items-center gap-1 border-b px-2 py-1.5">
          <span className="text-[12px] font-medium">Sources</span>
          <Popover>
            <PopoverTrigger render={<IconButton label="Add sources" className="ml-auto" />}>
              <PlusIcon />
            </PopoverTrigger>
            <PopoverContent align="start" className="w-64 gap-1 p-1.5">
              {datasets.length === 0 && <p className="p-2 text-[12px] text-muted-foreground">No dataset in the database.</p>}
              {datasets.map((d) => (
                <label key={d.id} className="flex items-center gap-2 rounded-md px-2 py-1 hover:bg-muted">
                  <Checkbox
                    checked={sourceIds.includes(d.id)}
                    onCheckedChange={(on) => update({ sourceIds: on ? [...sourceIds, d.id] : sourceIds.filter((id) => id !== d.id) })}
                  />
                  <span className="truncate">{String(d.title)}</span>
                  <span className="ml-auto text-[10px] text-muted-foreground">{d.source === "excel" ? "SQL" : "DAX"}</span>
                </label>
              ))}
            </PopoverContent>
          </Popover>
        </div>
        {sourceError && <p className="px-2 py-1 text-[11px] text-destructive">{sourceError}</p>}
        {sources.length === 0 ? (
          <p className="p-3 text-[12px] text-muted-foreground">Add one or more datasets to explore and compare them.</p>
        ) : (
          <ModelTree
            sources={sources}
            onPick={(s, f) => add(s, f, f.kind === "measure" || NUMERIC.test(f.dataType ?? "") ? "values" : "rows", "sum")}
            action={(s, f) => (
              <DropdownMenu>
                <DropdownMenuTrigger render={<IconButton label="Add to…" size="icon-xs" />}>
                  <PlusIcon />
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" className="w-44">
                  <DropdownMenuItem onClick={() => add(s, f, "rows")}>Rows</DropdownMenuItem>
                  <DropdownMenuItem onClick={() => add(s, f, "columns")}>Columns</DropdownMenuItem>
                  {f.kind === "measure" ? (
                    <DropdownMenuItem onClick={() => add(s, f, "values")}>Values</DropdownMenuItem>
                  ) : (
                    AGGREGATES.map((agg) => (
                      <DropdownMenuItem key={agg} onClick={() => add(s, f, "values", agg)}>
                        Values ({agg})
                      </DropdownMenuItem>
                    ))
                  )}
                  {f.kind === "column" && <DropdownMenuItem onClick={() => add(s, f, "filters")}>Filter</DropdownMenuItem>}
                </DropdownMenuContent>
              </DropdownMenu>
            )}
          />
        )}
      </div>

      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex flex-col gap-1.5 border-b p-2.5">
          <Zone label="Rows" items={pivot.rows.map((d, i) => ({ key: i, label: d.label, hint: `${Object.keys(d.bindings).length}/${sources.length}`, warn: Object.keys(d.bindings).length < sources.filter((s) => pivot.values.some((v) => v.source === s.id)).length, edit: <DimEditor dim={d} sources={sources} onChange={(dim) => setPivot({ ...pivot, rows: pivot.rows.map((x, j) => (j === i ? dim : x)) })} /> }))} onRemove={(i) => setPivot({ ...pivot, rows: pivot.rows.filter((_, j) => j !== i) })} />
          <Zone label="Columns" items={pivot.columns.map((d, i) => ({ key: i, label: d.label, hint: `${Object.keys(d.bindings).length}/${sources.length}`, edit: <DimEditor dim={d} sources={sources} onChange={(dim) => setPivot({ ...pivot, columns: pivot.columns.map((x, j) => (j === i ? dim : x)) })} /> }))} onRemove={(i) => setPivot({ ...pivot, columns: pivot.columns.filter((_, j) => j !== i) })} />
          <Zone
            label="Values"
            items={pivot.values.map((v, i) => ({ key: i, label: v.label, hint: v.computed ? "ƒ" : v.aggregate, edit: <ValueEditor value={v} onChange={(value) => setPivot({ ...pivot, values: pivot.values.map((x, j) => (j === i ? value : x)) })} /> }))}
            onRemove={(i) => setPivot({ ...pivot, values: pivot.values.filter((_, j) => j !== i) })}
            extra={
              allowMeasures && (
                <Button variant="ghost" size="xs" className="text-muted-foreground" onClick={() => setMeasureOpen(true)}>
                  <SigmaIcon /> New measure
                </Button>
              )
            }
          />
          <Zone label="Filters" items={pivot.filters.map((f, i) => ({ key: i, label: `${f.field}${f.values.length ? ` = ${f.values.join(", ")}` : ""}`, edit: <FilterEditor filter={f} onChange={(filter) => setPivot({ ...pivot, filters: pivot.filters.map((x, j) => (j === i ? filter : x)) })} /> }))} onRemove={(i) => setPivot({ ...pivot, filters: pivot.filters.filter((_, j) => j !== i) })} />
          <div className="flex items-center gap-3 pt-0.5 text-[12px] text-muted-foreground">
            <label className="flex items-center gap-1.5">
              <Switch size="sm" checked={allowMeasures} onCheckedChange={(on) => update({ allowMeasures: on })} /> New measures
            </label>
            <button type="button" onClick={() => setShowQueries(!showQueries)} className={cn("flex items-center gap-1 hover:text-foreground", showQueries && "text-foreground")}>
              <CodeIcon className="size-3.5" /> Queries
            </button>
            {running && <span className="shimmer">Running…</span>}
            <button type="button" className="ml-auto hover:text-foreground" onClick={() => setPivot(emptyPivot())}>
              Clear
            </button>
          </div>
        </div>
        {showQueries && result && (
          <div className="max-h-56 shrink-0 overflow-auto border-b bg-muted/30 p-2">
            {result.queries.map((q, i) => (
              <div key={i} className="mb-2">
                <p className="text-[11px] text-muted-foreground">{sources.find((s) => s.id === q.source)?.title ?? q.source}</p>
                <pre className="font-mono text-[11px] whitespace-pre-wrap">{q.query}</pre>
              </div>
            ))}
          </div>
        )}
        {result?.queries.filter((q) => q.error).map((q, i) => (
          <p key={i} className="border-b bg-destructive/5 px-3 py-1.5 text-[12px] text-destructive">
            {sources.find((s) => s.id === q.source)?.title ?? "Pivot"}: {q.error}
          </p>
        ))}
        {result && pivot.values.length > 0 ? <PivotGrid pivot={pivot} result={result} /> : <p className="p-8 text-center text-[12px] text-muted-foreground">Add values (and rows) from the fields on the left, or ask the agent.</p>}
      </div>
      <MeasureDialog open={measureOpen} onClose={() => setMeasureOpen(false)} pivot={pivot} sources={sources} onAdd={(value) => setPivot({ ...pivot, values: [...pivot.values, value] })} />
    </div>
  );
}

const uniqueLabel = (pivot: Pivot, base: string) => {
  const used = new Set([...pivot.rows, ...pivot.columns, ...pivot.values].map((x) => x.label));
  let label = base;
  for (let i = 2; used.has(label); i++) label = `${base} ${i}`;
  return label;
};

function Zone({ label, items, onRemove, extra }: { label: string; items: { key: number; label: string; hint?: string; warn?: boolean; edit: React.ReactNode }[]; onRemove: (index: number) => void; extra?: React.ReactNode }) {
  return (
    <div className="flex min-h-7 flex-wrap items-center gap-1">
      <span className="w-16 shrink-0 text-[11px] text-muted-foreground uppercase">{label}</span>
      {items.map((item) => (
        <Popover key={item.key}>
          <span className={cn("inline-flex h-6 items-center rounded-md border bg-card text-[12px] shadow-soft", item.warn && "border-amber-500/50")}>
            <PopoverTrigger className="flex h-full items-center gap-1 pr-1 pl-2 hover:text-foreground">
              {item.label}
              {item.hint && <span className="text-[10px] text-muted-foreground">{item.hint}</span>}
            </PopoverTrigger>
            <button type="button" aria-label="Remove" onClick={() => onRemove(item.key)} className="px-1 text-muted-foreground hover:text-foreground">
              <XIcon className="size-3" />
            </button>
          </span>
          <PopoverContent align="start" className="w-80 gap-2 p-2.5">
            {item.edit}
          </PopoverContent>
        </Popover>
      ))}
      {extra}
    </div>
  );
}

/** Label of a row / column and its column in each source. */
function DimEditor({ dim, sources, onChange }: { dim: Pivot["rows"][number]; sources: Source[]; onChange: (dim: Pivot["rows"][number]) => void }) {
  return (
    <>
      <Input value={dim.label} onChange={(e) => onChange({ ...dim, label: e.target.value })} className="h-7" />
      {sources.map((s) => (
        <label key={s.id} className="flex flex-col gap-1">
          <span className="text-[11px] text-muted-foreground">{s.title}</span>
          <SimpleSelect
            size="sm"
            value={dim.bindings[s.id] ?? null}
            placeholder="Not mapped"
            onChange={(ref) => onChange({ ...dim, bindings: { ...dim.bindings, [s.id]: ref } })}
            options={s.model.tables.flatMap((t) => t.columns.map((c) => ({ value: `'${t.name}'[${c.name}]`, label: `${t.name} › ${c.name}` })))}
          />
        </label>
      ))}
    </>
  );
}

function ValueEditor({ value, onChange }: { value: Pivot["values"][number]; onChange: (value: Pivot["values"][number]) => void }) {
  return (
    <>
      <Input value={value.label} onChange={(e) => onChange({ ...value, label: e.target.value })} className="h-7" />
      {value.field && !value.field.startsWith("[") && (
        <SimpleSelect size="sm" value={value.aggregate ?? "sum"} onChange={(a) => onChange({ ...value, aggregate: a as (typeof AGGREGATES)[number] })} options={AGGREGATES.map((a) => ({ value: a, label: a }))} />
      )}
      {(value.expression !== undefined || value.computed !== undefined) && (
        <textarea
          value={value.computed ?? value.expression ?? ""}
          onChange={(e) => onChange(value.computed !== undefined ? { ...value, computed: e.target.value } : { ...value, expression: e.target.value })}
          className="field-sizing-content min-h-14 rounded-md border bg-transparent p-2 font-mono text-[11px] outline-none focus:border-ring"
        />
      )}
      <SimpleSelect size="sm" value={value.format ?? "number"} onChange={(f) => onChange({ ...value, format: f as "number" })} options={["number", "integer", "percent", "currency"].map((f) => ({ value: f, label: f }))} />
    </>
  );
}

function FilterEditor({ filter, onChange }: { filter: Pivot["filters"][number]; onChange: (filter: Pivot["filters"][number]) => void }) {
  const [draft, setDraft] = useState(filter.values.join("; "));
  return (
    <>
      <p className="font-mono text-[11px] text-muted-foreground">{filter.field}</p>
      <Input
        value={draft}
        placeholder="Values separated by ;"
        onChange={(e) => setDraft(e.target.value)}
        onBlur={() => onChange({ ...filter, values: draft.split(";").map((v) => v.trim()).filter(Boolean) })}
        className="h-7"
      />
    </>
  );
}

function MeasureDialog({ open, onClose, pivot, sources, onAdd }: { open: boolean; onClose: () => void; pivot: Pivot; sources: Source[]; onAdd: (value: Pivot["values"][number]) => void }) {
  const [label, setLabel] = useState("");
  const [kind, setKind] = useState<string>("computed");
  const [text, setText] = useState("");
  const [error, setError] = useState("");
  const source = sources.find((s) => s.id === kind);
  const submit = () => {
    try {
      const value = kind === "computed" ? { label, computed: text } : { label, source: kind, expression: text };
      if (kind === "computed") compile(text, pivot);
      onAdd(value);
      setLabel("");
      setText("");
      setError("");
      onClose();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };
  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="gap-3 sm:max-w-lg">
        <DialogTitle className="text-[13px] font-medium">New measure</DialogTitle>
        <Input autoFocus value={label} onChange={(e) => setLabel(e.target.value)} placeholder="Name" />
        <SimpleSelect value={kind} onChange={setKind} options={[{ value: "computed", label: "Across sources (from the values)" }, ...sources.map((s) => ({ value: s.id, label: `${s.title} (${s.language})` }))]} />
        <textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder={kind === "computed" ? `{${pivot.values[0]?.label ?? "Value A"}} - {${pivot.values[1]?.label ?? "Value B"}}` : source?.language === "SQL" ? "SUM([Table].[Amount]) / COUNT(*)" : "CALCULATE(SUM('Sales'[Amount]), 'Sales'[Channel] = \"Web\")"}
          className="field-sizing-content min-h-20 rounded-lg border bg-transparent p-2.5 font-mono text-[12px] outline-none focus:border-ring"
        />
        {kind === "computed" && (
          <p className="text-[11px] text-muted-foreground">
            Values by name in braces: {pivot.values.map((v) => `{${v.label}}`).join(", ") || "none yet"}. Functions: DIVIDE, ABS, COALESCE.
          </p>
        )}
        {error && <p className="text-[12px] text-destructive">{error}</p>}
        <div className="flex justify-end gap-2">
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button disabled={!label.trim() || !text.trim()} onClick={submit}>
            Add
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

function format(value: unknown, kind?: string) {
  if (typeof value !== "number") return formatCell(value);
  if (kind === "percent") return value.toLocaleString(undefined, { style: "percent", maximumFractionDigits: 1 });
  if (kind === "integer") return Math.round(value).toLocaleString();
  if (kind === "currency") return value.toLocaleString(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 2 });
  return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function PivotGrid({ pivot, result }: { pivot: Pivot; result: PivotResult }) {
  const grid = useMemo(() => layout(pivot, result), [pivot, result]);
  const multiValue = pivot.values.length > 1;
  return (
    <div className="min-h-0 flex-1 overflow-auto">
      <table className="w-max min-w-full border-separate border-spacing-0 text-[12px]">
        <thead className="sticky top-0 z-10 bg-background">
          <tr>
            {pivot.rows.map((d) => (
              <th key={d.label} rowSpan={multiValue && pivot.columns.length ? 2 : 1} className="border-r border-b px-2 py-1.5 text-left font-medium">
                {d.label}
              </th>
            ))}
            {grid.colKeys.map((key) => (
              <th key={JSON.stringify(key)} colSpan={pivot.columns.length ? pivot.values.length : 1} className="border-r border-b px-2 py-1.5 text-right font-medium">
                {pivot.columns.length ? key.map(formatCell).join(" · ") : pivot.values[0].label}
              </th>
            ))}
            {!pivot.columns.length && pivot.values.slice(1).map((v) => <th key={v.label} className="border-r border-b px-2 py-1.5 text-right font-medium">{v.label}</th>)}
          </tr>
          {multiValue && pivot.columns.length > 0 && (
            <tr>
              {grid.colKeys.flatMap((key) =>
                pivot.values.map((v) => (
                  <th key={JSON.stringify(key) + v.label} className="border-r border-b px-2 py-1 text-right text-[11px] font-normal text-muted-foreground">
                    {v.label}
                  </th>
                )),
              )}
            </tr>
          )}
        </thead>
        <tbody>
          {grid.rows.map((row) => (
            <tr key={JSON.stringify(row.labels)} className="hover:bg-muted/40">
              {row.labels.map((label, i) => (
                <td key={i} className="border-r border-b px-2 py-1 whitespace-nowrap">
                  {formatCell(label)}
                </td>
              ))}
              {grid.colKeys.flatMap((key) =>
                pivot.values.map((v, vi) => {
                  const value = grid.cell(row.cells, key, vi);
                  return (
                    <td key={JSON.stringify(key) + vi} className={cn("border-r border-b px-2 py-1 text-right whitespace-nowrap tabular-nums", v.computed && typeof value === "number" && value !== 0 && (value > 0 ? "text-emerald-600 dark:text-emerald-400" : "text-red-600 dark:text-red-400"))}>
                      {format(value, v.format)}
                    </td>
                  );
                }),
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
