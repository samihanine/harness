import { useEffect, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { PlayIcon } from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures, useLatest } from "@/agent/server";
import { NeedsDb } from "@/components/guards";
import { ModelTree } from "@/components/model-tree";
import { ResultTable } from "@/components/result-table";
import { SearchSelect } from "@/components/search-select";
import { Button } from "@/components/ui/button";
import { useDb } from "@/db/db";
import { useSources } from "@/lib/use-source";
import type { QueryResult } from "@/pbi/api";
import { modelSummary } from "@/pbi/model";
import { runQuery, sampleQuery, toCsv } from "@/query/engine";

export const Route = createFileRoute("/dax")({
  component: () => (
    <NeedsDb>
      <DaxRunner />
    </NeedsDb>
  ),
});

type Run = { result?: QueryResult; error?: string; ms?: number; running?: boolean };

function DaxRunner() {
  const navigate = useNavigate();
  const data = useDb();
  const datasets = data.rows("datasets");
  const [datasetId, setDatasetId] = useState<string | null>(() => localStorage.getItem("pbi-agent:dax-dataset"));
  const { sources } = useSources(datasetId ? [datasetId] : []);
  const source = sources[0];
  const [query, setQuery] = useState("");
  const [run, setRun] = useState<Run>({});
  const editor = useRef<HTMLTextAreaElement>(null);
  const latest = useLatest({ source, query, run });

  useEffect(() => {
    if (!datasetId) return;
    localStorage.setItem("pbi-agent:dax-dataset", datasetId);
    setRun({});
  }, [datasetId]);
  useEffect(() => {
    if (source) setQuery(localStorage.getItem(`pbi-agent:dax:${source.id}`) ?? sampleQuery(source));
  }, [source]);
  useEffect(() => {
    if (source) localStorage.setItem(`pbi-agent:dax:${source.id}`, query);
  }, [source, query]);

  const execute = async (text = query) => {
    if (!source || !text.trim()) return;
    const start = performance.now();
    setRun({ running: true });
    try {
      const result = await runQuery(source, text);
      setRun({ result, ms: performance.now() - start });
      return result;
    } catch (error) {
      setRun({ error: error instanceof Error ? error.message : String(error), ms: performance.now() - start });
      throw error;
    }
  };

  /** Inserts a field reference at the cursor. */
  const insert = (text: string) => {
    const el = editor.current;
    if (!el) return;
    const { selectionStart: a, selectionEnd: b } = el;
    setQuery(query.slice(0, a) + text + query.slice(b));
    requestAnimationFrame(() => {
      el.focus();
      el.setSelectionRange(a + text.length, a + text.length);
    });
  };

  useAgentFeatures(
    source
      ? `Query runner on dataset "${source.title}" (${source.kind === "excel" ? "local Excel model: SQL with [brackets] (AlaSQL)" : "Power BI semantic model: DAX, EVALUATE …"}). dax://model describes the model, dax://editor holds the user's query and its last result. Use run_query to test queries, set_editor to put a query in the user's editor.`
      : "Query runner: no dataset selected yet.",
    (server) =>
      source
        ? [
            server.registerResource(
              "model",
              "dax://model",
              { title: `Model of ${source.title}`, mimeType: "text/plain", annotations: { priority: 1 } },
              async (uri) => ({ contents: [{ uri: uri.href, text: `${source.context ? `Context: ${source.context}\n` : ""}${modelSummary(source.model)}` }] }),
            ),
            server.registerResource(
              "editor",
              "dax://editor",
              { title: "Editor and last result", mimeType: "text/plain", annotations: { priority: 1 } },
              async (uri) => {
                const { query: q, run: r } = latest.current;
                const result = r.error ? `Error: ${r.error}` : r.result ? `${r.result.rows.length} rows:\n${toCsv(r.result, 20)}` : "(not run)";
                return { contents: [{ uri: uri.href, text: `Query (${source.language}):\n${q}\n\nLast result: ${result}` }] };
              },
            ),
            server.registerTool(
              "run_query",
              {
                description: `Runs a ${source.language} query on the dataset and returns the rows as CSV (does not change the user's editor).`,
                inputSchema: z.object({ query: z.string() }),
                annotations: { readOnlyHint: true },
              },
              async ({ query: q }) => {
                try {
                  const result = await runQuery(source, q);
                  return textResult(`${result.rows.length} rows\n${toCsv(result, 200)}`);
                } catch (error) {
                  return errorResult(error);
                }
              },
            ),
            server.registerTool(
              "set_editor",
              { description: "Puts a query in the user's editor (and runs it when run is true).", inputSchema: z.object({ query: z.string(), run: z.boolean().optional() }) },
              async ({ query: q, run: go }) => {
                const before = latest.current.query;
                setQuery(q);
                let summary = "Query set in the editor.";
                if (go) summary = await execute(q).then((r) => `Ran: ${r?.rows.length ?? 0} rows.`, (e: Error) => `Ran with an error: ${e.message}`);
                return textResult(summary, { name: "set_editor", arguments: { query: before } });
              },
            ),
          ]
        : [],
    [source],
  );

  return (
    <div className="flex h-full">
      <div className="flex w-64 shrink-0 flex-col border-r">
        <div className="p-2">
          <SearchSelect
            size="sm"
            value={datasetId}
            onChange={setDatasetId}
            placeholder="Choose a dataset"
            onAdd={() => void navigate({ to: "/database" })}
            addLabel="Add a dataset"
            options={datasets.map((d) => ({ value: d.id, label: String(d.title), recent: String(d.updated_at ?? "") }))}
          />
        </div>
        {source && <ModelTree sources={[source]} onPick={(_, f) => insert(f.ref)} />}
      </div>
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="relative shrink-0 border-b">
          <textarea
            ref={editor}
            value={query}
            spellCheck={false}
            disabled={!source}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
                e.preventDefault();
                void execute().catch(() => undefined);
              }
            }}
            placeholder={source ? `${source.language} query…` : "Choose a dataset first."}
            className="h-56 w-full resize-y bg-transparent p-3 font-mono text-[12px] leading-relaxed outline-none"
          />
          <div className="absolute right-2 bottom-2 flex items-center gap-2">
            {source && <span className="rounded bg-muted px-1.5 py-0.5 font-mono text-[10px] text-muted-foreground">{source.language}</span>}
            <span className="text-[11px] text-muted-foreground">⌘↵</span>
            <Button size="sm" disabled={!source || run.running} onClick={() => void execute().catch(() => undefined)}>
              <PlayIcon className="fill-current" /> Run
            </Button>
          </div>
        </div>
        {run.running && <p className="shimmer p-3 text-[12px]">Running…</p>}
        {run.error && <pre className="m-3 rounded-lg border border-destructive/20 bg-destructive/5 p-3 text-[12px] whitespace-pre-wrap text-destructive">{run.error}</pre>}
        {run.result && <ResultTable result={run.result} name={source?.title ?? "result"} />}
        {run.ms !== undefined && !run.running && <p className="border-t px-3 py-1 text-right text-[11px] text-muted-foreground">{Math.round(run.ms)} ms</p>}
      </div>
    </div>
  );
}
