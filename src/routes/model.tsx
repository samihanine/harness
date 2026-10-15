import { useEffect, useMemo, useRef, useState } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { ArrowRightIcon, CalculatorIcon, DatabaseIcon, ExternalLinkIcon, KeyRoundIcon, RefreshCwIcon, SigmaIcon, SparklesIcon, TableIcon, WorkflowIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { Badge } from "@/components/field-input";
import { Picker } from "@/components/picker";
import { Empty, PageHeader, SidePanel } from "@/components/side-panel";
import { AI_PANEL } from "@/config";
import { withContexts } from "@/agent/context";
import type { Agent } from "@/agent/loop";
import { STEPS_TEXT, modelEditorTools, modelState, saveModelDraft } from "@/agent/model-editor-tools";
import { DAX_RULES } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { refreshDataset } from "@/lib/library";
import { fromTmdl, type ModelInfo } from "@/lib/model";
import { errorText } from "@/lib/ms";
import { useCollection, type DatasetEntry } from "@/lib/store";
import { readModelFiles } from "@/lib/tmdl";
import { ModelDraft } from "@/lib/tmdl-edit";
import { refreshModel } from "@/server/ms";
import { REFRESH_QUOTA, refreshesToday } from "@/lib/pbi";

export const Route = createFileRoute("/model")({
  validateSearch: (s: Record<string, unknown>) => ({ dataset: (s.dataset as string) || undefined }),
  component: ModelEditor,
});

type Model = { info: ModelInfo; steps: Record<string, string[]> };

/** The model as saved (its definition): structure and Power Query steps; read again after each save. */
function useModel(dataset?: DatasetEntry, version = 0) {
  const [model, setModel] = useState<Model>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    if (!dataset) return;
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
        setModel({ info, steps });
      },
      (e) => live && setError(errorText(e)),
    );
    return () => void (live = false);
  }, [dataset?.id, version]); // eslint-disable-line react-hooks/exhaustive-deps
  return { model, error };
}

/**
 * Dataset editor: an existing semantic model's structure — columns, measures, relationships, Power
 * Query steps — changed by the AI (saved when its answer is done) and shown read only. Its data is
 * refreshed by the user (a button: about 8 refreshes a day on Pro), never by the AI.
 */
function ModelEditor() {
  const { dataset: datasetId } = Route.useSearch();
  const navigate = useNavigate({ from: "/model" });
  useLast("model:dataset", datasetId, (dataset) => navigate({ search: { dataset }, replace: true }));
  const datasets = useCollection("datasets").filter((d) => d.access === "edit");
  const dataset = datasets.find((d) => d.id === datasetId);
  const [version, setVersion] = useState(0);
  const { model, error } = useModel(dataset, version);
  const [refresh, setRefresh] = useState<{ busy?: boolean; done?: string; error?: string }>({});
  const live = useRef(dataset);
  live.current = dataset;
  /** The last refresh of this screen, for the AI's page state (it checks the data after it). */
  const lastRefresh = useRef<{ id: string; text: string }>(undefined);
  const ref = () => live.current && { id: live.current.id, groupId: live.current.groupId };

  const agent = useMemo<Agent | undefined>(
    () =>
      dataset && {
        instructions: `You change the STRUCTURE of an existing Power BI semantic model: columns (settings, calculated DAX columns), measures, relationships, and Power Query steps (structured: you never write a whole query; ${STEPS_TEXT}). You do not add tables (that is done in the Library) and you never refresh the data: tell the user to click "Refresh the data" after changes to columns or steps.
You can see the data: preview_table (row count, blanks, first rows of some columns) and run_dax read the model as of its last refresh. When the user asks whether something worked, or after a refresh (see the page state), check it with them and say what you see (e.g. how many rows of a new column are filled) instead of saying you cannot check.
Images from a URL column (procedure):
1. step_image_base64 (column = the URL column, into = the new column): never Web.Contents in step_custom_column (Power BI cannot refresh a source built from a column). Images on SharePoint / OneDrive come resized by SharePoint (any file size, quick to refresh); images of other sites only up to about 24 KB: advise putting them on SharePoint.
2. To redo an image step (an old one, or images now in another site or library: the step knows the libraries of the URLs present when it was made), remove_last_step then step_image_base64 with the same "into": the column is kept.
3. Tell the user to click "Refresh the data"; then check with preview_table(table, [URL column, image column]): a blank image where a URL is given means a file that cannot be read (wrong URL, no access, or a site the step does not know: redo the step).
Your changes go to a draft saved when your answer is done. Measures and calculated columns are evaluated on the model first. Do the whole request in one go, then answer with what you changed.
${DAX_RULES}`,
        tools: modelEditorTools(ref),
        context: () => withContexts([`Semantic model "${dataset.name}" (id ${dataset.id}); the user can edit it.`], [dataset.contextIds]),
        state: () => modelState(ref(), lastRefresh.current?.id === live.current?.id ? lastRefresh.current?.text : undefined),
        finish: async () => {
          const r = ref();
          if (!r) return undefined;
          const text = await saveModelDraft(r);
          if (!text) return undefined;
          setVersion((v) => v + 1);
          void refreshDataset(live.current!).catch(() => undefined);
          return `✓ ${text}`;
        },
      },
    [dataset?.id], // eslint-disable-line react-hooks/exhaustive-deps
  );
  // Refreshes used of the day's quota (shown on the button, read again after each refresh).
  const [used, setUsed] = useState<number>();
  useEffect(() => {
    setUsed(undefined);
    if (dataset) refreshesToday(dataset).then(setUsed, () => undefined);
  }, [dataset?.id, refresh.busy]); // eslint-disable-line react-hooks/exhaustive-deps
  const runRefresh = () => {
    if (!dataset) return;
    setRefresh({ busy: true });
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

  return (
    <div className="flex h-full">
      {AI_PANEL && <SidePanel tabs={[{ id: "ai", label: "AI assistant", icon: SparklesIcon, content: agent ? <Chat scope={`model:${dataset!.id}`} agent={agent} placeholder="Ask to add a measure, a relationship, clean a column…" /> : <Empty>Choose a dataset you can edit.</Empty> }]} />}
      <div className="flex min-w-0 flex-1 flex-col">
        <PageHeader
          screen="Dataset editor"
          picker={
            <Picker
              items={datasets.map((d) => ({ id: d.id, label: d.name }))}
              value={datasetId}
              onChange={(dataset) => navigate({ search: { dataset } })}
              placeholder="Choose a dataset…"
              empty={<>No editable dataset: add one (or create one from Excel files) in the <Link to="/" search={{ tab: "datasets" }} className="underline">Library</Link>.</>}
            />
          }
          actions={
            dataset && (
              <>
                {refresh.done && <span className="text-[13px] text-green-700">{refresh.done}</span>}
                {refresh.error && <span className="max-w-md truncate text-[13px] text-destructive" title={refresh.error}>{refresh.error}</span>}
                <button type="button" className="btn" disabled={refresh.busy} title={`Reads the sources again in Power BI. On Pro, ${REFRESH_QUOTA} refreshes a day per model (by this button or scheduled, over 24 hours); "Refresh now" in Power BI is not counted.`} onClick={runRefresh}>
                  <RefreshCwIcon className={refresh.busy ? "animate-spin" : ""} /> {refresh.busy ? "Refreshing…" : "Refresh the data"}
                  {used !== undefined && !refresh.busy && <span className="text-muted-foreground">{used}/{REFRESH_QUOTA}</span>}
                </button>
                <a className="btn" href={`https://app.powerbi.com/groups/${dataset.groupId ?? "me"}/datasets/${dataset.id}/details`} target="_blank" rel="noreferrer">
                  <ExternalLinkIcon /> Power BI
                </a>
              </>
            )
          }
        />
        {!dataset && <Empty>Choose a dataset you can edit (its definition must be readable: see its access in the Library).</Empty>}
        {dataset && error && <p className="px-6 py-3 text-destructive">{error}</p>}
        {dataset && !model && !error && <Empty>Reading the model…</Empty>}
        {model && <ModelView model={model} />}
      </div>
    </div>
  );
}

const TYPE_COLOR: Record<string, "blue" | "green" | "amber" | "violet" | "gray" | "teal"> = { string: "gray", int64: "blue", double: "blue", decimal: "blue", dateTime: "amber", boolean: "violet" };

/** The model, read only: relationships, then one card per table (columns, measures, Power Query steps). */
function ModelView({ model }: { model: Model }) {
  const { info, steps } = model;
  return (
    <div className="min-h-0 flex-1 overflow-y-auto p-6">
      <section className="card mb-5 p-5">
        <h2 className="mb-3 flex items-center gap-2 font-semibold">
          <WorkflowIcon className="size-4 text-muted-foreground" /> Relationships
        </h2>
        {info.relationships.length ? (
          <ul className="flex flex-col gap-2">
            {info.relationships.map((r, i) => (
              <li key={i} className="flex flex-wrap items-center gap-2">
                <span className="font-mono text-[13px]">{r.from}</span>
                <ArrowRightIcon className="size-4 text-muted-foreground" />
                <span className="font-mono text-[13px]">{r.to}</span>
                <Badge color="gray">{r.cardinality}</Badge>
                {r.both && <Badge color="violet">both directions</Badge>}
                {!r.active && <Badge color="amber">inactive</Badge>}
              </li>
            ))}
          </ul>
        ) : (
          <p className="text-muted-foreground">No relationship.</p>
        )}
      </section>
      <div className="grid grid-cols-[repeat(auto-fill,minmax(340px,1fr))] gap-5">
        {info.tables.map((t) => (
          <section key={t.name} className="card flex flex-col">
            <h2 className="flex items-center gap-2.5 border-b px-5 py-4 font-semibold">
              <span className="flex size-8 items-center justify-center rounded-full bg-muted">
                <TableIcon className="size-4 text-muted-foreground" />
              </span>
              {t.name}
              {t.kind && <Badge color="teal">{t.kind}</Badge>}
              {t.hidden && <Badge color="gray">hidden</Badge>}
            </h2>
            <ul className="flex flex-col px-5 py-3">
              {t.columns.map((c) => (
                <li key={c.name} className={`flex h-9 items-center gap-2 ${c.hidden ? "opacity-50" : ""}`} title={c.description}>
                  {c.expression ? <CalculatorIcon className="size-4 shrink-0 text-violet-600" /> : /^id$/i.test(c.name) ? <KeyRoundIcon className="size-4 shrink-0 text-muted-foreground" /> : <DatabaseIcon className="size-4 shrink-0 text-muted-foreground" />}
                  <span className="min-w-0 flex-1 truncate">{c.name}</span>
                  <Badge color={TYPE_COLOR[c.type ?? ""] ?? "gray"}>{c.type ?? "?"}</Badge>
                </li>
              ))}
            </ul>
            {t.measures.length > 0 && (
              <ul className="flex flex-col border-t px-5 py-3">
                {t.measures.map((m) => (
                  <li key={m.name} className="flex flex-col py-1.5" title={m.expression}>
                    <span className="flex items-center gap-2">
                      <SigmaIcon className="size-4 shrink-0 text-emerald-600" />
                      <span className="min-w-0 flex-1 truncate font-medium">{m.name}</span>
                      {m.format && <span className="text-[12px] text-muted-foreground">{m.format}</span>}
                    </span>
                    <code className="ml-6 truncate text-[12px] text-muted-foreground">{m.expression?.replace(/\s+/g, " ")}</code>
                  </li>
                ))}
              </ul>
            )}
            {(steps[t.name]?.length ?? 0) > 0 && (
              <div className="mt-auto flex flex-wrap items-center gap-1.5 border-t px-5 py-3 text-[12.5px]">
                <span className="mr-1 text-muted-foreground">Power Query</span>
                {steps[t.name].map((s, i) => (
                  <span key={i} className="flex items-center gap-1.5">
                    {i > 0 && <ArrowRightIcon className="size-3 text-muted-foreground" />}
                    <span className="rounded-md bg-muted px-2 py-0.5">{s}</span>
                  </span>
                ))}
              </div>
            )}
          </section>
        ))}
      </div>
    </div>
  );
}
