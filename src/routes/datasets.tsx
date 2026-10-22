import { useEffect, useMemo, useRef, useState } from "react";
import { z } from "zod";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import CodeMirror from "@uiw/react-codemirror";
import { ArrowRightIcon, BookTextIcon, CalculatorIcon, DatabaseIcon, ExternalLinkIcon, FileSpreadsheetIcon, FileUpIcon, KeyRoundIcon, LayersIcon, PlayIcon, RefreshCwIcon, SigmaIcon, TerminalSquareIcon } from "lucide-react";
import { ContextLinks } from "@/components/context-links";
import { DataTable } from "@/components/data-table";
import { CsvDialog } from "@/components/csv-dialog";
import { AddDialog, EntityMenu, withToast } from "@/components/entity";
import { DatasetPicker } from "@/components/pickers";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { DropdownMenuItem } from "@/components/ui/dropdown-menu";
import { Empty, StatusText, SubTabs, Workspace } from "@/components/workspace";
import { datasetText, withContexts } from "@/agent/context";
import { tool, type Agent } from "@/agent/loop";
import { STEPS_TEXT, modelEditorTools, modelState, saveModelDraft } from "@/agent/model-editor-tools";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, daxTool } from "@/agent/pbi-tools";
import { last, useLast } from "@/lib/last";
import { createDatasetFromExcel, refreshDataset } from "@/lib/library";
import { fromTmdl, modelParameters, type ModelInfo } from "@/lib/model";
import { errorText } from "@/lib/ms";
import { REFRESH_QUOTA, dax, refreshesToday, rowsText, type Rows } from "@/lib/pbi";
import { patch, remove, useCollection, type DatasetEntry } from "@/lib/store";
import { readModelFiles } from "@/lib/tmdl";
import { ModelDraft } from "@/lib/tmdl-edit";
import { useTheme } from "@/lib/theme";
import { refreshModel } from "@/server/ms";

export const Route = createFileRoute("/datasets")({
  validateSearch: (s: Record<string, unknown>): { dataset?: string; tab?: string } => ({ dataset: (s.dataset as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Datasets,
});

type Model = { info: ModelInfo; steps: Record<string, string[]>; parameters: ReturnType<typeof modelParameters> };

/** The model as saved (its definition): structure, Power Query steps, parameters; read again after each save. */
function useModel(dataset?: DatasetEntry, version = 0) {
  const [model, setModel] = useState<Model>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    if (!dataset) return;
    // Without the edit right, the definition is not readable: the structure read by DAX is shown.
    if (dataset.access !== "edit") {
      setModel(dataset.info ? { info: dataset.info, steps: {}, parameters: [] } : undefined);
      setError(dataset.info ? undefined : dataset.error);
      return;
    }
    let live = true;
    setModel(undefined);
    setError(undefined);
    readModelFiles(dataset.id, dataset.groupId).then(
      (files) => {
        if (!live) return;
        const info = fromTmdl(files);
        const draft = new ModelDraft(files);
        const steps = Object.fromEntries(
          info.tables.map((t) => {
            try {
              return [t.name, draft.query(t.name).steps?.list.map((s) => s.name) ?? []];
            } catch {
              return [t.name, []];
            }
          }),
        );
        setModel({ info, steps, parameters: modelParameters(files) });
      },
      (e) => live && setError(errorText(e)),
    );
    return () => void (live = false);
  }, [dataset?.id, dataset?.access, version]); // eslint-disable-line react-hooks/exhaustive-deps
  return { model, error };
}

/**
 * Datasets: the semantic models the app knows (added here by link, or created from Excel or CSV files).
 * Tabs: the schema (tables, columns, measures, relationships, parameters), a DAX console, the linked
 * contexts. The AI changes the model's structure (saved when its answer is done); the data is refreshed
 * by the user only (about 8 refreshes a day on Pro).
 */
function Datasets() {
  const { dataset: datasetId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/datasets" });
  useLast("datasets:dataset", datasetId, (dataset) => navigate({ search: { dataset }, replace: true }));
  const dataset = useCollection("datasets").find((d) => d.id === datasetId);
  const [version, setVersion] = useState(0);
  const { model, error } = useModel(dataset, version);
  const [refresh, setRefresh] = useState<{ busy?: string; done?: string; error?: string }>({});
  const [creating, setCreating] = useState<"excel" | "csv">();
  const live = useRef(dataset);
  live.current = dataset;
  /** The last refresh of this screen, for the AI's page state (it checks the data after it). */
  const lastRefresh = useRef<{ id: string; text: string }>(undefined);
  const ref = () => live.current && { id: live.current.id, groupId: live.current.groupId };

  const agent = useMemo<Agent | undefined>(() => {
    if (!dataset) return undefined;
    const editable = dataset.access === "edit";
    return {
      instructions: `Screen: DATASETS. ${
        editable
          ? `You change the STRUCTURE of a Power BI semantic model: columns (settings, calculated DAX columns), measures, relationships, and Power Query steps (structured: you never write a whole query; ${STEPS_TEXT}). You never refresh the data: tell the user to click "Refresh the data" after changes to columns or steps.
You can see the data: preview_table (row count, blanks, first rows) and run_dax read the model as of its last refresh. When the user asks whether something worked, or after a refresh (see the page state), check it and say what you see.
Images from a URL column: step_image_base64 (never Web.Contents in step_custom_column: Power BI cannot refresh a source built from a column); images on SharePoint / OneDrive come resized; to redo an image step, remove_last_step then step_image_base64 with the same "into". After the user's refresh, check with preview_table.
Your changes go to a draft saved at the end. Measures and calculated columns are evaluated on the model first.`
          : `The user can only read this model (right: ${dataset.access ?? "unknown"}): you explain it, find fields and answer with DAX; changing it needs the edit right.`
      }
${DAX_RULES}`,
      tools: [...(editable ? modelEditorTools(ref) : [daxTool(() => ref())]), ...modelTools(() => live.current), daxEditorTool(() => navigate({ search: (x) => ({ ...x, tab: "dax" }), replace: true }))],
      // The structure and examples of the data come at once: no exploring call for a simple question.
      context: async () => withContexts([await datasetText(live.current ?? dataset)], [dataset.contextIds]),
      state: async () => {
        const editor = daxEditor.current ? `DAX editor (DAX tab) now:\n${daxEditor.current.query().slice(0, 1500)}` : "";
        const model = editable ? await modelState(ref(), lastRefresh.current?.id === live.current?.id ? lastRefresh.current?.text : undefined) : "";
        return [model, editor].filter(Boolean).join("\n\n");
      },
      finish: editable
        ? async () => {
            const r = ref();
            if (!r) return undefined;
            const text = await saveModelDraft(r);
            if (!text) return undefined;
            setVersion((v) => v + 1);
            void refreshDataset(live.current!).catch(() => undefined);
            return `✓ ${text}`;
          }
        : undefined,
    };
  }, [dataset?.id, dataset?.access]); // eslint-disable-line react-hooks/exhaustive-deps

  // Refreshes used of the day's quota (shown on the button, read again after each refresh).
  const [used, setUsed] = useState<number>();
  useEffect(() => {
    setUsed(undefined);
    if (dataset) refreshesToday(dataset).then(setUsed, () => undefined);
  }, [dataset?.id, refresh.busy]); // eslint-disable-line react-hooks/exhaustive-deps
  const runRefresh = () => {
    if (!dataset) return;
    setRefresh({ busy: "Refreshing…" });
    refreshModel({ data: { datasetId: dataset.id, groupId: dataset.groupId } }).then(
      (r) => {
        const at = new Date().toLocaleTimeString();
        lastRefresh.current = { id: dataset.id, text: r.status === "Completed" ? `completed at ${at}` : `${r.status} at ${at}${r.error ? `: ${r.error}` : ""}` };
        setRefresh(r.status === "Completed" ? { done: "Data refreshed." } : { error: `Refresh ${r.status}${r.error ? `: ${r.error}` : ""}` });
      },
      (e) => {
        lastRefresh.current = { id: dataset.id, text: `failed at ${new Date().toLocaleTimeString()}: ${errorText(e)}` };
        setRefresh({ error: errorText(e) });
      },
    );
  };
  const rename = () => {
    if (!dataset) return;
    const name = prompt("Name in this app (empty: the name in Power BI)", dataset.name);
    if (name === null) return;
    if (name.trim()) return void patch("datasets", dataset.id, { name: name.trim(), renamed: true });
    void patch("datasets", dataset.id, { renamed: false }).then(() => withToast("Reading…", refreshDataset(dataset), "Name read again."));
  };
  const open = (id: string) => navigate({ search: { dataset: id } });

  return (
    <>
      <Workspace
        title="Datasets"
        icon={DatabaseIcon}
        picker={<DatasetPicker value={datasetId} onChange={open} />}
        actions={
          <>
            <StatusText state={refresh} />
            {dataset && dataset.access !== "view" && (
              <Button variant="outline" size="sm" disabled={!!refresh.busy} title={`Reads the sources again in Power BI. On Pro, ${REFRESH_QUOTA} refreshes a day per model (this button or scheduled, over 24 hours).`} onClick={runRefresh}>
                <RefreshCwIcon className={refresh.busy ? "animate-spin" : ""} /> Refresh the data
                {used !== undefined && !refresh.busy && <span className="text-muted-foreground tabular-nums">{used}/{REFRESH_QUOTA}</span>}
              </Button>
            )}
            {dataset && (
              <Button variant="outline" size="sm" asChild>
                <a href={`https://app.powerbi.com/groups/${dataset.groupId ?? "me"}/datasets/${dataset.id}/details`} target="_blank" rel="noreferrer">
                  <ExternalLinkIcon /> Power BI
                </a>
              </Button>
            )}
            <EntityMenu
              kind="datasets"
              entry={dataset}
              onRename={rename}
              onRefresh={() => dataset && void withToast("Reading the model…", refreshDataset(dataset).then(() => setVersion((v) => v + 1)), "Read again.")}
              onRemove={() => dataset && confirm(`Remove "${dataset.name}" from the app? (the model stays in Power BI)`) && void remove("datasets", dataset.id).then(() => navigate({ search: {} }))}
            >
              <DropdownMenuItem onSelect={() => setCreating("excel")}>
                <FileSpreadsheetIcon /> New from Excel tables…
              </DropdownMenuItem>
              <DropdownMenuItem onSelect={() => setCreating("csv")}>
                <FileUpIcon /> New from a CSV file…
              </DropdownMenuItem>
            </EntityMenu>
          </>
        }
        ai={{ scope: `datasets:${dataset?.id}`, agent, placeholder: "Ask to add a measure, a relationship, clean a column…", empty: "Choose a dataset, or add one from its link." }}
        tab={tab}
        onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
        tabs={[
          {
            id: "schema",
            label: "Schema",
            icon: LayersIcon,
            content: !dataset ? <Empty icon={DatabaseIcon}>Choose a dataset, add one from its Power BI link, or create one from Excel or CSV files (menu at the top).</Empty> : error ? <p className="p-6 text-[13px] text-destructive">{error}</p> : !model ? <Empty>Reading the model…</Empty> : <SchemaTab model={model} />,
          },
          { id: "dax", label: "DAX", icon: TerminalSquareIcon, content: dataset ? <DaxConsole dataset={dataset} /> : <Empty>Choose a dataset.</Empty> },
          { id: "context", label: "Context", icon: BookTextIcon, count: dataset?.contextIds.length, content: dataset ? <ContextLinks what="dataset" ids={dataset.contextIds} onChange={(contextIds) => void patch("datasets", dataset.id, { contextIds })} /> : <Empty>Choose a dataset.</Empty> },
        ]}
      />
      <AddDialog
        open={creating === "excel"}
        onOpenChange={(v) => !v && setCreating(undefined)}
        title="New dataset from Excel tables"
        description="A semantic model in My workspace reading Excel tables of SharePoint / OneDrive (one model table per Excel table). Its data comes with its first refresh."
        fields={[
          { name: "name", label: "Name" },
          { name: "links", label: "Excel file links", placeholder: "One link per line or separated by spaces", hint: "Each file's first table is used." },
        ]}
        submit="Create"
        onSubmit={async ({ name, links }) => open((await createDatasetFromExcel(name, links.split(/\s+/).filter(Boolean).map((link) => ({ link })))).id)}
      />
      <CsvDialog open={creating === "csv"} onOpenChange={(v) => !v && setCreating(undefined)} onCreated={open} />
    </>
  );
}

const TYPE: Record<string, string> = { string: "Text", int64: "Whole number", double: "Decimal", decimal: "Fixed decimal", dateTime: "Date / time", boolean: "True / false" };

/** The model, read only: tables, columns, measures, relationships, parameters. */
function SchemaTab({ model }: { model: Model }) {
  const { info, steps, parameters } = model;
  const columns = info.tables.flatMap((t) => t.columns.map((c) => ({ ...c, table: t.name })));
  const measures = info.tables.flatMap((t) => t.measures.map((m) => ({ ...m, table: t.name })));
  const fieldParameters = info.tables.filter((t) => t.kind === "parameter");
  return (
    <SubTabs
      tabs={[
        {
          id: "tables",
          label: "Tables",
          count: info.tables.length,
          content: (
            <DataTable
              head={["Table", "Kind", "Columns", "Measures", "Power Query steps"]}
              rows={info.tables.map((t) => [
                <span key="n" className="flex items-center gap-2 font-medium">
                  <LayersIcon className="size-4 text-muted-foreground" aria-hidden /> {t.name}
                  {t.hidden && <Badge variant="outline">hidden</Badge>}
                </span>,
                t.kind ? <Badge key="k" variant="secondary">{t.kind}</Badge> : <span key="k" className="text-muted-foreground">table</span>,
                t.columns.length,
                t.measures.length,
                <span key="s" className="flex flex-wrap items-center gap-1 text-[12px]">
                  {(steps[t.name] ?? []).map((s, i) => (
                    <span key={i} className="flex items-center gap-1">
                      {i > 0 && <ArrowRightIcon className="size-3 text-muted-foreground" />}
                      <span className="rounded bg-muted px-1.5 py-0.5">{s}</span>
                    </span>
                  ))}
                </span>,
              ])}
              empty="No table."
            />
          ),
        },
        {
          id: "columns",
          label: "Columns",
          count: columns.length,
          content: (
            <DataTable
              head={["Column", "Table", "Type", "Format / category", "Description"]}
              rows={columns.map((c) => [
                <span key="n" className={`flex items-center gap-2 font-medium ${c.hidden ? "opacity-60" : ""}`}>
                  {c.expression ? <CalculatorIcon className="size-4 text-violet-600" aria-label="Calculated" /> : /^id$/i.test(c.name) ? <KeyRoundIcon className="size-4 text-muted-foreground" /> : <DatabaseIcon className="size-4 text-muted-foreground" />}
                  {c.name}
                </span>,
                c.table,
                <Badge key="t" variant="secondary">{TYPE[c.type ?? ""] ?? c.type ?? "?"}</Badge>,
                <span key="f" className="text-[12px] text-muted-foreground">{[c.format, c.category].filter(Boolean).join(" · ") || "—"}</span>,
                <span key="d" className="text-[12.5px] text-muted-foreground">{c.expression ? <code className="font-mono">{c.expression}</code> : (c.description ?? "")}</span>,
              ])}
              empty="No column."
            />
          ),
        },
        {
          id: "measures",
          label: "Measures",
          count: measures.length,
          content: (
            <DataTable
              head={["Measure", "Table", "DAX", "Format"]}
              rows={measures.map((m) => [
                <span key="n" className="flex items-center gap-2 font-medium">
                  <SigmaIcon className="size-4 text-success" aria-hidden /> {m.name}
                </span>,
                m.table,
                <code key="e" className="line-clamp-4 font-mono text-[12px] whitespace-pre-wrap">{m.expression}</code>,
                <span key="f" className="text-[12px] text-muted-foreground">{m.format ?? "—"}</span>,
              ])}
              empty="No measure."
            />
          ),
        },
        {
          id: "relations",
          label: "Relationships",
          count: info.relationships.length,
          content: (
            <DataTable
              head={["From (many)", "", "To (one)", "Cardinality", "Filter"]}
              rows={info.relationships.map((r) => [
                <code key="f" className="font-mono text-[12.5px]">{r.from}</code>,
                <ArrowRightIcon key="a" className="size-4 text-muted-foreground" />,
                <code key="t" className="font-mono text-[12.5px]">{r.to}</code>,
                <Badge key="c" variant="secondary">{r.cardinality}</Badge>,
                <span key="d" className="flex gap-1">
                  {r.both ? <Badge>both directions</Badge> : <span className="text-muted-foreground">single</span>}
                  {!r.active && <Badge variant="outline">inactive</Badge>}
                </span>,
              ])}
              empty="No relationship."
            />
          ),
        },
        {
          id: "parameters",
          label: "Parameters",
          count: parameters.length + fieldParameters.length,
          content: (
            <DataTable
              head={["Parameter", "Kind", "Value / definition"]}
              rows={[
                ...parameters.map((p) => [<span key="n" className="font-medium">{p.name}</span>, <Badge key="k" variant="secondary">{p.parameter ? "Power Query parameter" : "Shared M expression"}</Badge>, <code key="v" className="line-clamp-3 font-mono text-[12px] whitespace-pre-wrap">{p.value}</code>]),
                ...fieldParameters.map((t) => [<span key="n" className="font-medium">{t.name}</span>, <Badge key="k" variant="secondary">Field parameter</Badge>, <code key="v" className="line-clamp-3 font-mono text-[12px]">{t.extra}</code>]),
              ]}
              empty="No parameter in this model."
            />
          ),
        },
      ]}
    />
  );
}

/** Runs DAX queries on the model (Ctrl+Enter), results in a table; the last query is kept in this browser. */
/** Writes a query in the DAX tab's editor (and runs it): the user sees it and can change it. */
const daxEditorTool = (show: () => void) =>
  tool({
    name: "set_dax_editor",
    description: "Writes a DAX query in the user's DAX editor (the DAX tab, shown) and runs it unless run is false; returns the first rows. Use it when the user wants a query to see, keep or change.",
    args: z.object({ query: z.string().describe("DAX query starting with EVALUATE (or DEFINE)"), run: z.boolean().optional() }),
    readOnly: true,
    run: async ({ query, run = true }) => {
      show();
      // The DAX tab mounts its editor when shown.
      for (let i = 0; i < 20 && !daxEditor.current; i++) await new Promise((r) => setTimeout(r, 100));
      if (!daxEditor.current) throw new Error("The DAX editor is not open.");
      const rows = await daxEditor.current.set(query, run);
      return rows ? `Query written and run (${rows.length} rows):\n${rowsText(rows, 10)}` : "Query written in the editor (not run).";
    },
  });

/** The DAX editor on screen, for the AI's set_dax_editor tool (the console registers itself). */
const daxEditor: { current?: { query: () => string; set: (query: string, run: boolean) => Promise<Rows | undefined> } } = {};

function DaxConsole({ dataset }: { dataset: DatasetEntry }) {
  const key = `dax:${dataset.id}`;
  const [query, setQuery] = useState(() => last.get(key) ?? `EVALUATE\nTOPN(20, '${dataset.info?.tables[0]?.name ?? "Table"}')`);
  const [result, setResult] = useState<{ rows?: Rows; ms?: number; busy?: boolean; error?: string }>({});
  const { resolved } = useTheme();
  const execute = (text: string) => {
    last.set(key, text);
    setResult({ busy: true });
    const t0 = performance.now();
    return dax(dataset, text).then(
      (rows) => (setResult({ rows, ms: Math.round(performance.now() - t0) }), rows),
      (e) => {
        setResult({ error: errorText(e) });
        throw e;
      },
    );
  };
  const run = () => void execute(query).catch(() => undefined);
  const current = useRef(query);
  current.current = query;
  useEffect(() => {
    daxEditor.current = {
      query: () => current.current,
      set: async (text, andRun) => {
        setQuery(text);
        last.set(key, text);
        return andRun ? execute(text) : undefined;
      },
    };
    return () => void (daxEditor.current = undefined);
  }, [dataset.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const columns = result.rows?.[0] ? Object.keys(result.rows[0]) : [];
  return (
    <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-hidden p-6">
      <div className="flex flex-col overflow-hidden rounded-xl border bg-card shadow-xs" onKeyDown={(e) => (e.metaKey || e.ctrlKey) && e.key === "Enter" && (e.preventDefault(), run())}>
        <CodeMirror value={query} onChange={setQuery} height="180px" theme={resolved} basicSetup={{ lineNumbers: true, foldGutter: false }} className="text-[13px]" aria-label="DAX query" />
        <div className="flex items-center gap-3 border-t bg-muted/40 px-3 py-2">
          <Button size="sm" onClick={run} disabled={result.busy || !query.trim()}>
            <PlayIcon /> {result.busy ? "Running…" : "Run"}
          </Button>
          <span className="text-[12px] text-muted-foreground">Ctrl / ⌘ + Enter</span>
          {result.rows && (
            <span className="ml-auto text-[12px] text-muted-foreground tabular-nums">
              {result.rows.length} row{result.rows.length === 1 ? "" : "s"} · {result.ms} ms
            </span>
          )}
        </div>
      </div>
      {result.error && <p className="rounded-lg bg-destructive/10 px-4 py-3 font-mono text-[12.5px] whitespace-pre-wrap text-destructive">{result.error}</p>}
      {result.rows && (
        <div className="min-h-0 flex-1 overflow-auto">
          <DataTable head={columns} rows={result.rows.slice(0, 500).map((r) => columns.map((c) => <span key={c} className="tabular-nums">{r[c] === null || r[c] === undefined ? <span className="text-muted-foreground">blank</span> : String(r[c]).slice(0, 300)}</span>))} empty="No row." />
        </div>
      )}
    </div>
  );
}
