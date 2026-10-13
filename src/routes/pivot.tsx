import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { z } from "zod";
import { ChevronRightIcon, CodeIcon, FilterIcon, PlayIcon, SearchIcon, SigmaIcon, XIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { Picker } from "@/components/picker";
import { Popover } from "@/components/popover";
import { Empty, SidePanel } from "@/components/side-panel";
import { PIVOT_TABS, visible } from "@/config";
import { tool, type Agent } from "@/agent/loop";
import { daxTool } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { addDataset } from "@/lib/library";
import { resolveField } from "@/lib/model";
import { errorText } from "@/lib/ms";
import { checkValues, columnValues, dax, rowsText, type Rows } from "@/lib/pbi";
import { AGGREGATIONS, EMPTY, OPERATORS, OPERATORS_OF, filterText, kindOf, labelOf, specText, toDax, type Aggregation, type Col, type Filter, type Operator, type Spec, type Value } from "@/lib/pivot";
import { last, useLast } from "@/lib/last";
import { useCollection, type DatasetEntry } from "@/lib/store";

export const Route = createFileRoute("/pivot")({
  validateSearch: (s: Record<string, unknown>) => ({ dataset: (s.dataset as string) || undefined }),
  component: PivotBuilder,
});

/* ----------------------------------- Spec ----------------------------------- */

/** The spec of each model, kept in this browser. */
const SPEC_KEY = (id?: string) => `pivot:spec:${id}`;
function useSpec(id?: string) {
  const [spec, setSpec] = useState<Spec>(EMPTY);
  useEffect(() => {
    try {
      setSpec({ ...EMPTY, ...JSON.parse(localStorage.getItem(SPEC_KEY(id)) ?? "{}") });
    } catch {
      setSpec(EMPTY);
    }
  }, [id]);
  const update = (next: Spec) => {
    setSpec(next);
    try {
      localStorage.setItem(SPEC_KEY(id), JSON.stringify(next));
    } catch {}
  };
  return [spec, update] as const;
}

/* ----------------------------------- Page ----------------------------------- */

function PivotBuilder() {
  const { dataset: datasetId } = Route.useSearch();
  const navigate = useNavigate({ from: "/pivot" });
  useLast("pivot:dataset", datasetId, (dataset) => navigate({ search: { dataset }, replace: true }));
  const datasets = useCollection("datasets");
  const dataset = datasets.find((d) => d.id === datasetId);
  const [spec, setSpec] = useSpec(datasetId);
  const query = useMemo(() => toDax(spec, dataset?.info), [spec, dataset?.info]);
  const [result, setResult] = useState<{ rows?: Rows; error?: string; ms?: number; running?: boolean }>({});

  // Run on each change (a pivot table updates as it is built).
  useEffect(() => {
    if (!dataset || !query) return setResult({});
    let live = true;
    const t0 = performance.now();
    setResult((r) => ({ ...r, running: true }));
    const timer = setTimeout(() => {
      dax(dataset, query).then(
        (rows) => live && setResult({ rows, ms: Math.round(performance.now() - t0) }),
        (e) => live && setResult({ error: errorText(e) }),
      );
    }, 300);
    return () => {
      live = false;
      clearTimeout(timer);
    };
  }, [dataset?.id, query]); // eslint-disable-line react-hooks/exhaustive-deps

  const openInDax = () => {
    last.set(`dax:query:${datasetId}`, query);
    void navigate({ to: "/dax", search: { dataset: datasetId } });
  };

  // The agent reads and changes the pivot as it is when called.
  const live = useRef({ spec, setSpec, result });
  live.current = { spec, setSpec, result };
  const agent = useMemo(() => (dataset ? pivotAgent(dataset, live) : undefined), [dataset]);

  return (
    <div className="flex h-full">
      <SidePanel
        width={320}
        tabs={visible([
          { id: "fields", label: "Fields", content: dataset ? <FieldList dataset={dataset} spec={spec} onChange={setSpec} /> : <Empty>Choose a semantic model.</Empty> },
          { id: "ai", label: "AI", content: agent ? <Chat scope={`pivot:${dataset!.id}`} agent={agent} placeholder="Describe the table you want…" /> : <Empty>Choose a semantic model.</Empty> },
        ], PIVOT_TABS)}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex h-12 shrink-0 items-center gap-2 border-b bg-card px-3">
          <Picker
            className="w-72"
            items={datasets.map((d) => ({ id: d.id, label: d.name }))}
            value={datasetId}
            onChange={(dataset) => navigate({ search: { dataset } })}
            placeholder="Choose a semantic model…"
            add={{ label: "Add a semantic model", placeholder: "Semantic model link or id", run: async (link) => (await addDataset(link)).id }}
          />
        </div>
        {!dataset ? (
          <Empty>Choose a semantic model (add some in the Library), then add fields from the list.</Empty>
        ) : (
          <>
            <Zones dataset={dataset} spec={spec} onChange={setSpec} />
            <div className="flex h-10 shrink-0 items-center gap-2 border-b bg-card px-3 text-[12px] text-muted-foreground">
              {result.running ? <span>Running…</span> : result.rows ? <span className="tabular-nums">{result.rows.length} rows · {result.ms} ms</span> : null}
              <span className="ml-auto" />
              {query && (
                <>
                  <Popover width={460} trigger={(open) => (
                    <button type="button" className="btn h-7" onClick={open}>
                      <CodeIcon /> DAX
                    </button>
                  )}>
                    {() => <pre className="max-h-80 overflow-auto rounded bg-muted p-2 font-mono text-[11px] whitespace-pre-wrap">{query}</pre>}
                  </Popover>
                  <button type="button" className="btn h-7" onClick={openInDax}>
                    <PlayIcon /> Open in the DAX builder
                  </button>
                </>
              )}
              {spec !== EMPTY && (
                <button type="button" className="btn h-7" onClick={() => setSpec(EMPTY)}>
                  <XIcon /> Clear
                </button>
              )}
            </div>
            {result.error && <p className="border-b bg-red-500/5 px-4 py-2 font-mono text-[12px] whitespace-pre-wrap text-destructive">{result.error}</p>}
            {!query ? <Empty>Add rows, columns or values from the field list.</Empty> : result.rows ? <PivotTable spec={spec} rows={result.rows} /> : null}
          </>
        )}
      </div>
    </div>
  );
}

/* -------------------------------- Field list -------------------------------- */

/** The model's tables (visible columns and measures), searchable; each field is added to a zone by its buttons. */
function FieldList({ dataset, spec, onChange }: { dataset: DatasetEntry; spec: Spec; onChange: (s: Spec) => void }) {
  const [search, setSearch] = useState("");
  const [open, setOpen] = useState<string[]>([]);
  const q = search.trim().toLowerCase();
  const tables = (dataset.info?.tables ?? []).filter((t) => !t.hidden);
  if (!dataset.info) return <Empty>The model structure is not read yet: use "Read again" on it in the Library.</Empty>;
  const add = (zone: "rows" | "columns" | "filters", c: Col) =>
    onChange({ ...spec, [zone]: [...spec[zone].filter((x) => !(x.table === c.table && x.column === c.column)), zone === "filters" ? { ...c, values: [] } : c] });
  const addValue = (v: Value) => onChange({ ...spec, values: [...spec.values, v] });
  const zoneBtn = (label: string, title: string, onClick: () => void) => (
    <button type="button" title={title} onClick={onClick} className="flex h-5 min-w-5 items-center justify-center rounded border bg-card px-1 text-[10px] font-medium text-muted-foreground hover:border-ring hover:text-foreground">
      {label}
    </button>
  );
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <label className="flex items-center gap-2 border-b px-3 py-2">
        <SearchIcon className="size-3.5 text-muted-foreground" />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search a field" className="h-6 flex-1 bg-transparent outline-none" />
      </label>
      <div className="min-h-0 flex-1 overflow-y-auto p-1.5 text-[12px]">
        {tables.map((t) => {
          const columns = t.columns.filter((c) => !c.hidden && (!q || c.name.toLowerCase().includes(q) || t.name.toLowerCase().includes(q)));
          const measures = t.measures.filter((m) => !m.hidden && (!q || m.name.toLowerCase().includes(q) || t.name.toLowerCase().includes(q)));
          if (!columns.length && !measures.length) return null;
          const expanded = !!q || open.includes(t.name);
          return (
            <div key={t.name}>
              <button type="button" onClick={() => setOpen(expanded ? open.filter((x) => x !== t.name) : [...open, t.name])} className="flex w-full items-center gap-1 rounded px-1.5 py-1 text-left font-medium hover:bg-muted">
                <ChevronRightIcon className={`size-3 shrink-0 transition-transform ${expanded ? "rotate-90" : ""}`} />
                <span className="truncate">{t.name}</span>
              </button>
              {expanded && (
                <ul className="ml-4">
                  {measures.map((m) => (
                    <li key={`m:${m.name}`} className="group flex items-center gap-1 rounded px-1.5 py-0.5 hover:bg-muted">
                      <SigmaIcon className="size-3 shrink-0 text-amber-600" />
                      <span className="min-w-0 flex-1 truncate" title={m.description}>{m.name}</span>
                      <span className="flex gap-0.5 opacity-0 group-hover:opacity-100 focus-within:opacity-100">{zoneBtn("V", "Add to values", () => addValue({ table: t.name, measure: m.name }))}</span>
                    </li>
                  ))}
                  {columns.map((c) => {
                    const col = { table: t.name, column: c.name };
                    const numeric = /int|double|decimal|number|currency/i.test(c.type ?? "");
                    return (
                      <li key={`c:${c.name}`} className="group flex items-center gap-1 rounded px-1.5 py-0.5 hover:bg-muted">
                        <span className="size-3 shrink-0" />
                        <span className="min-w-0 flex-1 truncate" title={c.description}>{c.name}</span>
                        <span className="flex gap-0.5 opacity-0 group-hover:opacity-100 focus-within:opacity-100">
                          {zoneBtn("R", "Add to rows", () => add("rows", col))}
                          {zoneBtn("C", "Add to columns", () => add("columns", col))}
                          {zoneBtn("V", numeric ? "Add to values (sum)" : "Add to values (count)", () => addValue({ ...col, aggregation: numeric ? "sum" : "count" }))}
                          {zoneBtn("F", "Add a filter", () => add("filters", col))}
                        </span>
                      </li>
                    );
                  })}
                </ul>
              )}
            </div>
          );
        })}
      </div>
      <p className="border-t px-3 py-2 text-[11px] text-muted-foreground">Hover a field: R rows · C columns · V values · F filter.</p>
    </div>
  );
}

/* ----------------------------------- Zones ---------------------------------- */

function Zones({ dataset, spec, onChange }: { dataset: DatasetEntry; spec: Spec; onChange: (s: Spec) => void }) {
  const chip = (key: string, text: string, onRemove: () => void, extra?: React.ReactNode) => (
    <span key={key} className="inline-flex h-6 max-w-64 items-center gap-1 rounded-md border bg-card pr-0.5 pl-2 text-[12px]">
      <span className="truncate">{text}</span>
      {extra}
      <button type="button" className="opacity-60 hover:opacity-100" title="Remove" onClick={onRemove}>
        <XIcon className="size-3" />
      </button>
    </span>
  );
  const zone = (title: string, items: React.ReactNode[]) => (
    <div className="flex min-h-8 items-center gap-1.5">
      <span className="label w-16 shrink-0">{title}</span>
      <div className="flex flex-wrap gap-1">{items.length ? items : <span className="text-[11px] text-muted-foreground/70">—</span>}</div>
    </div>
  );
  const without = <T,>(list: T[], i: number) => list.filter((_, k) => k !== i);
  return (
    <div className="flex shrink-0 flex-col gap-1 border-b bg-card px-3 py-2">
      {zone("Filters", spec.filters.map((f, i) => chip(`f${i}`, filterText(f), () => onChange({ ...spec, filters: without(spec.filters, i) }), (
        <FilterEditor dataset={dataset} filter={f} onChange={(next) => onChange({ ...spec, filters: spec.filters.map((x, k) => (k === i ? next : x)) })} />
      ))))}
      {zone("Rows", spec.rows.map((c, i) => chip(`r${i}`, c.column, () => onChange({ ...spec, rows: without(spec.rows, i) }))))}
      {zone("Columns", spec.columns.map((c, i) => chip(`c${i}`, c.column, () => onChange({ ...spec, columns: without(spec.columns, i) }))))}
      {zone("Values", spec.values.map((v, i) => chip(`v${i}`, v.measure ?? v.column ?? "", () => onChange({ ...spec, values: without(spec.values, i) }), v.column ? (
        <select className="h-5 rounded bg-muted px-0.5 text-[11px] outline-none" value={v.aggregation} onChange={(e) => onChange({ ...spec, values: spec.values.map((x, k) => (k === i ? { ...x, aggregation: e.target.value as Aggregation } : x)) })}>
          {Object.entries(AGGREGATIONS).map(([k, label]) => (
            <option key={k} value={k}>{label}</option>
          ))}
        </select>
      ) : undefined)))}
    </div>
  );
}

/**
 * A filter: an operator suited to the column (text: one of, contains, starts with…; number and date:
 * comparisons, between), then the values chosen in the column's real values, or the condition's value(s).
 */
function FilterEditor({ dataset, filter, onChange }: { dataset: DatasetEntry; filter: Filter; onChange: (f: Filter) => void }) {
  const [options, setOptions] = useState<{ values: unknown[]; more: boolean; error?: string }>();
  const [search, setSearch] = useState("");
  const kind = kindOf(dataset.info, filter);
  const op = filter.op ?? "in";
  const list = op === "in" || op === "notIn";
  const load = () => !options && columnValues(dataset, filter.table, filter.column).then(setOptions, (e) => setOptions({ values: [], more: false, error: errorText(e) }));
  const input = (key: "value" | "value2", placeholder: string) => (
    <input
      className="input h-7"
      type={kind === "number" ? "number" : kind === "date" ? "date" : "text"}
      placeholder={placeholder}
      value={filter[key] ?? ""}
      onChange={(e) => onChange({ ...filter, [key]: e.target.value })}
    />
  );
  return (
    <Popover width={280} trigger={(open) => (
      <button type="button" title="Edit the filter" className="opacity-60 hover:opacity-100" onClick={() => (open(), list && load())}>
        <FilterIcon className="size-3" />
      </button>
    )}>
      {() => (
        <div className="flex flex-col gap-1.5">
          <span className="label">{filter.column} ({kind})</span>
          <select
            className="input h-7"
            value={op}
            onChange={(e) => {
              const next = e.target.value as Operator;
              onChange({ ...filter, op: next });
              if (next === "in" || next === "notIn") load();
            }}
          >
            {OPERATORS_OF[kind].map((o) => (
              <option key={o} value={o}>{OPERATORS[o]}</option>
            ))}
          </select>
          {list ? (
            <>
              <input autoFocus className="input h-7" placeholder="Search" value={search} onChange={(e) => setSearch(e.target.value)} />
              {!options && <p className="text-[12px] text-muted-foreground">Reading the values…</p>}
              {options?.error && <p className="text-[11px] text-destructive">{options.error}</p>}
              <div className="flex max-h-56 flex-col overflow-y-auto">
                {options?.values.filter((v) => String(v).toLowerCase().includes(search.toLowerCase())).map((v) => {
                  const on = filter.values.some((x) => x === v);
                  return (
                    <label key={String(v)} className="flex items-center gap-2 py-0.5 text-[12px]">
                      <input type="checkbox" className="accent-foreground" checked={on} onChange={() => onChange({ ...filter, values: on ? filter.values.filter((x) => x !== v) : [...filter.values, v] })} />
                      <span className="truncate">{String(v)}</span>
                    </label>
                  );
                })}
              </div>
              {options?.more && <p className="text-[11px] text-muted-foreground">First 40 values only: use "contains" or a comparison for the others.</p>}
            </>
          ) : op === "between" ? (
            <div className="flex items-center gap-1.5">
              {input("value", "From")}
              <span className="text-muted-foreground">–</span>
              {input("value2", "To")}
            </div>
          ) : op === "blank" || op === "notBlank" ? null : (
            input("value", kind === "date" ? "YYYY-MM-DD" : "Value")
          )}
        </div>
      )}
    </Popover>
  );
}

/* ----------------------------------- Table ---------------------------------- */

const cellText = (v: unknown) => (v === null || v === undefined ? "" : typeof v === "number" ? v.toLocaleString(undefined, { maximumFractionDigits: 2 }) : String(v));

/** The query's rows laid out as a pivot table: row fields down, column fields across (× values). */
function PivotTable({ spec, rows }: { spec: Spec; rows: Rows }) {
  if (!rows.length) return <Empty>No data for this selection.</Empty>;
  const key = (r: Rows[number], cols: Col[]) => cols.map((c) => cellText(r[c.column])).join(" · ");
  const labels = spec.values.map(labelOf);
  const rowKeys = [...new Set(rows.map((r) => key(r, spec.rows)))];
  const colKeys = spec.columns.length ? [...new Set(rows.map((r) => key(r, spec.columns)))].slice(0, 60) : [""];
  const cells = new Map<string, Rows[number]>();
  for (const r of rows) cells.set(`${key(r, spec.rows)}|${key(r, spec.columns)}`, r);
  const first = new Map<string, Rows[number]>();
  for (const r of rows) if (!first.has(key(r, spec.rows))) first.set(key(r, spec.rows), r);
  // No value: rows only (an empty column per column key keeps the column fields readable).
  const measures = labels.length ? labels : spec.columns.length ? [""] : [];
  return (
    <div className="min-h-0 flex-1 overflow-auto bg-card">
      <table className="w-max min-w-full border-separate border-spacing-0 text-[12px]">
        <thead className="sticky top-0 z-10 bg-card">
          {spec.columns.length > 0 && (
            <tr>
              {spec.rows.map((c) => <th key={c.column} className="border-r border-b" />)}
              {colKeys.map((ck) => (
                <th key={ck} colSpan={measures.length} className="border-r border-b px-2 py-1 text-left font-medium">{ck}</th>
              ))}
            </tr>
          )}
          <tr>
            {spec.rows.map((c) => (
              <th key={c.column} className="border-r border-b px-2 py-1.5 text-left font-medium text-muted-foreground">{c.column}</th>
            ))}
            {colKeys.flatMap((ck) => measures.map((m) => (
              <th key={`${ck}|${m}`} className="border-r border-b px-2 py-1.5 text-right font-medium text-muted-foreground">{m}</th>
            )))}
          </tr>
        </thead>
        <tbody>
          {rowKeys.slice(0, 2000).map((rk) => (
            <tr key={rk} className="hover:bg-muted/40">
              {spec.rows.map((c) => (
                <td key={c.column} className="max-w-72 truncate border-r border-b px-2 py-1">{cellText(first.get(rk)?.[c.column])}</td>
              ))}
              {colKeys.flatMap((ck) => measures.map((m) => (
                <td key={`${ck}|${m}`} className="border-r border-b px-2 py-1 text-right tabular-nums">{m ? cellText(cells.get(`${rk}|${ck}`)?.[m]) : ""}</td>
              )))}
            </tr>
          ))}
        </tbody>
      </table>
      {(rowKeys.length > 2000 || colKeys.length >= 60) && <p className="p-2 text-[11px] text-muted-foreground">Large result: only the first rows / columns are shown.</p>}
    </div>
  );
}

/* ----------------------------------- AI ------------------------------------ */

const field = z.string().describe("Exact name from the model, best as 'Table'[Column] or 'Table'[Measure], e.g. 'Fact'[Revenue]");

/** The AI builds the pivot through one tool (names and filter values are checked against the model first). */
function pivotAgent(dataset: DatasetEntry, live: { current: { spec: Spec; setSpec: (s: Spec) => void; result: { rows?: Rows; error?: string } } }): Agent {
  const parse = (spec: string): Col => {
    const known = resolveField(dataset.info, spec);
    if (known.kind === "measure") throw new Error(`[${known.name}] is a measure: put it in values (rows, columns and filters take columns).`);
    return { table: known.table, column: known.name };
  };
  return {
    instructions: `You build a pivot table (TCD) for the user on one semantic model, with set_pivot: rows, columns, values (measures, or columns with an aggregation) and filters. set_pivot replaces the whole pivot and returns its first rows: then answer with a short reading of the key figures.
- Use exact names from the model (context, describe_table, search_model); prefer existing measures for values.
- Filters: op "in" / "notIn" with values (as they are in the data: read them with run_dax when unsure), or a condition: contains, notContains, startsWith (text); eq, ne, gt, ge, lt, le, between (numbers, dates as YYYY-MM-DD); blank, notBlank.
- To change one thing, send the whole pivot again with that change (the current pivot is in the state).
- Do the request in one go; ask only when it is really ambiguous.`,
    context: async () => `Semantic model "${dataset.name}"${dataset.context ? `\nNotes: ${dataset.context}` : ""}\n${dataset.model ?? "(structure not read: use the model tools)"}`,
    state: async () => {
      const { spec, result } = live.current;
      return [`Current pivot:\n${specText(spec)}`, result.error ? `Its query failed: ${result.error}` : result.rows ? `Its result: ${result.rows.length} rows\n${rowsText(result.rows, 10)}` : ""].filter(Boolean).join("\n\n");
    },
    tools: [
      tool({
        name: "set_pivot",
        description: "Replaces the pivot table (everything not given is emptied) and returns its first rows or its error.",
        args: z.object({
          rows: z.array(field).optional(),
          columns: z.array(field).optional(),
          values: z.array(z.object({ field, aggregation: z.enum(Object.keys(AGGREGATIONS) as [Aggregation, ...Aggregation[]]).optional().describe("For a column (not a measure): sum, count, distinct, average, min, max") })).optional(),
          filters: z
            .array(
              z.object({
                field,
                op: z.enum(Object.keys(OPERATORS) as [Operator, ...Operator[]]),
                values: z.array(z.union([z.string(), z.number(), z.boolean()])).optional().describe('For "in" / "notIn"'),
                value: z.union([z.string(), z.number()]).optional().describe("For a condition (between: the start)"),
                value2: z.union([z.string(), z.number()]).optional().describe("between: the end"),
              }),
            )
            .optional(),
        }),
        run: async ({ rows = [], columns = [], values = [], filters = [] }) => {
          const spec: Spec = {
            rows: rows.map(parse),
            columns: columns.map(parse),
            values: values.map((v): Value => {
              const known = resolveField(dataset.info, v.field);
              if (known.kind === "measure") return { table: known.table, measure: known.name };
              const numeric = kindOf(dataset.info, { table: known.table, column: known.name }) === "number";
              return { table: known.table, column: known.name, aggregation: v.aggregation ?? (numeric ? "sum" : "count") };
            }),
            filters: await Promise.all(
              filters.map(async (f): Promise<Filter> => {
                const c = parse(f.field);
                const list = f.op === "in" || f.op === "notIn";
                return { ...c, op: f.op, values: list ? await checkValues(dataset, c.table, c.column, f.values ?? []) : [], value: f.value === undefined ? undefined : String(f.value), value2: f.value2 === undefined ? undefined : String(f.value2) };
              }),
            ),
          };
          live.current.setSpec(spec);
          const query = toDax(spec, dataset.info);
          if (!query) return "Pivot emptied (give rows, columns or values).";
          const result = await dax(dataset, query);
          return `Pivot set: ${result.length} rows.\n${rowsText(result, 15)}`;
        },
      }),
      daxTool(() => dataset),
      ...modelTools(() => dataset),
    ],
  };
}
