import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { z } from "zod";
import { PlayIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { DAX_TABS, visible } from "@/config";
import { Picker } from "@/components/picker";
import { Empty, SidePanel } from "@/components/side-panel";
import { tool, type Agent } from "@/agent/loop";
import { daxTool } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { addDataset } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { dax, rowsText, type Rows } from "@/lib/pbi";
import { last, useLast } from "@/lib/last";
import { useCollection, type DatasetEntry } from "@/lib/store";

export const Route = createFileRoute("/dax")({
  validateSearch: (s: Record<string, unknown>) => ({ dataset: (s.dataset as string) || undefined }),
  component: DaxBuilder,
});

const EXAMPLE = "EVALUATE\nTOPN(10, …)";
/** Rows drawn in the table (the query result itself is kept whole). */
const SHOWN = 1000;

type Result = { rows?: Rows; error?: string; ms?: number };

function DaxBuilder() {
  const { dataset: datasetId } = Route.useSearch();
  const navigate = useNavigate({ from: "/dax" });
  useLast("dax:dataset", datasetId, (dataset) => navigate({ search: { dataset }, replace: true }));
  const datasets = useCollection("datasets");
  const dataset = datasets.find((d) => d.id === datasetId);
  // The query of each model is kept in this browser.
  const [query, setQueryState] = useState("");
  useEffect(() => setQueryState(last.get(`dax:query:${datasetId}`) ?? ""), [datasetId]);
  const setQuery = (q: string) => {
    setQueryState(q);
    last.set(`dax:query:${datasetId}`, q);
  };
  const [result, setResult] = useState<Result>({});
  const [running, setRunning] = useState(false);
  useEffect(() => setResult({}), [datasetId]);

  const run = async (q = query): Promise<Result> => {
    if (!dataset || !q.trim()) return {};
    setRunning(true);
    const t0 = performance.now();
    const r = await dax(dataset, q).then(
      (rows) => ({ rows, ms: Math.round(performance.now() - t0) }),
      (e) => ({ error: errorText(e), ms: Math.round(performance.now() - t0) }),
    );
    setResult(r);
    setRunning(false);
    return r;
  };

  // The agent reads the editor and the last result as they are when it is called.
  const live = useRef({ query, result, setQuery, run });
  live.current = { query, result, setQuery, run };
  const agent = useMemo(() => (dataset ? daxAgent(dataset, live) : undefined), [dataset]);

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible([
          { id: "ai", label: "AI", content: agent ? <Chat scope={`dax:${dataset!.id}`} agent={agent} placeholder="Describe the data you want…" /> : <Empty>Choose a semantic model.</Empty> },
        ], DAX_TABS)}
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
          <Empty>Choose a semantic model (add some in the Library).</Empty>
        ) : (
          <>
            <div className="flex shrink-0 flex-col gap-2 border-b bg-card p-3">
              <textarea
                className="input h-48 resize-y py-2 font-mono text-[12px] leading-relaxed"
                spellCheck={false}
                placeholder={EXAMPLE}
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) (e.preventDefault(), void run());
                }}
              />
              <div className="flex items-center gap-3">
                <button type="button" className="btn-primary" disabled={running || !query.trim()} onClick={() => void run()}>
                  <PlayIcon /> {running ? "Running…" : "Run"}
                </button>
                <span className="text-[11px] text-muted-foreground">⌘ / Ctrl + Enter</span>
                <span className="ml-auto text-[12px] text-muted-foreground tabular-nums">
                  {result.rows && `${result.rows.length} row${result.rows.length === 1 ? "" : "s"}${result.rows.length > SHOWN ? ` (first ${SHOWN} shown)` : ""}`}
                  {result.ms !== undefined && ` · ${result.ms} ms`}
                </span>
              </div>
            </div>
            {result.error && <p className="border-b bg-red-500/5 px-4 py-2 font-mono text-[12px] whitespace-pre-wrap text-destructive">{result.error}</p>}
            <ResultTable rows={result.rows} />
          </>
        )}
      </div>
    </div>
  );
}

function ResultTable({ rows }: { rows?: Rows }) {
  if (!rows) return <Empty>Write a DAX query (EVALUATE …) and run it, or ask the AI.</Empty>;
  if (rows.length === 0) return <Empty>The query returned no rows.</Empty>;
  const columns = Object.keys(rows[0]);
  return (
    <div className="min-h-0 flex-1 overflow-auto bg-card">
      <table className="w-max min-w-full border-separate border-spacing-0 text-[12px]">
        <thead className="sticky top-0 z-10 bg-card/95 backdrop-blur">
          <tr>
            {columns.map((c) => (
              <th key={c} className="border-r border-b px-2 py-1.5 text-left font-medium text-muted-foreground">
                {c}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.slice(0, SHOWN).map((r, i) => (
            <tr key={i} className="hover:bg-muted/40">
              {columns.map((c) => (
                <td key={c} className={`max-w-96 truncate border-r border-b px-2 py-1 ${typeof r[c] === "number" ? "text-right tabular-nums" : ""}`}>
                  {r[c] === null || r[c] === undefined ? <span className="text-muted-foreground/60">null</span> : String(r[c])}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ----------------------------------- AI ------------------------------------ */

function daxAgent(dataset: DatasetEntry, live: { current: { query: string; result: Result; setQuery: (q: string) => void; run: (q?: string) => Promise<Result> } }): Agent {
  return {
    instructions:
      "You write DAX queries for the user, on one semantic model. Read the model with its tools (describe_table, search_model, get_measure), test with run_dax if needed, then put the final query in the user's editor with set_query (it runs it there). Queries start with EVALUATE (or DEFINE); prefer SUMMARIZECOLUMNS, existing measures, TOPN for large results. Explain the query briefly.",
    context: async () => {
      const { query, result } = live.current;
      return [
        `Semantic model "${dataset.name}"${dataset.context ? `\nNotes: ${dataset.context}` : ""}\n${dataset.model ?? "(structure not read: use the model tools)"}`,
        `The user's editor:\n${query.trim() || "(empty)"}`,
        result.error ? `Last run failed: ${result.error}` : result.rows ? `Last run: ${result.rows.length} rows\n${rowsText(result.rows, 10)}` : "Not run yet.",
      ].join("\n\n");
    },
    tools: [
      tool({
        name: "set_query",
        description: "Writes a DAX query in the user's editor and runs it there; returns its first rows or its error.",
        args: z.object({ query: z.string().describe("Full DAX query (EVALUATE …)") }),
        run: async ({ query }) => {
          live.current.setQuery(query);
          const r = await live.current.run(query);
          if (r.error) throw new Error(r.error);
          return `Query set and run: ${r.rows?.length ?? 0} rows.\n${rowsText(r.rows ?? [], 10)}`;
        },
      }),
      daxTool(() => dataset),
      ...modelTools(() => dataset),
    ],
  };
}
