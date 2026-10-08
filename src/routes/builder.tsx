import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Chat } from "@/components/chat";
import { Zoom } from "@/components/zoom";
import { Empty, SidePanel } from "@/components/side-panel";
import { authorTools, daxTool } from "@/agent/pbi-tools";
import { addReport } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { embedNewReport, pageText, resetEmbed } from "@/lib/pbi";
import { useCollection } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { useLast } from "@/lib/last";
import { libraryContext } from "./viewer";

export const Route = createFileRoute("/builder")({
  validateSearch: (s: Record<string, unknown>) => ({ report: (s.report as string) || undefined, newOn: (s.newOn as string) || undefined }),
  component: Builder,
});

/** "'Table': col:Type, [Measure] = …" lines of the library → "Table[Measure]" names. */
const measuresOf = (model = "") =>
  new Set([...model.matchAll(/^'([^']+)': (.*)$/gm)].flatMap(([, table, items]) => [...items.matchAll(/(?:^|, )\[([^\]]+)\]/g)].map((m) => `${table}[${m[1]}]`)));

function Builder() {
  const { report: reportId, newOn } = Route.useSearch();
  const navigate = useNavigate({ from: "/builder" });
  useLast("builder:report", reportId, (report) => navigate({ search: { report, newOn: undefined }, replace: true }), !!(reportId || newOn));
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const entry = reports.find((r) => r.id === reportId);
  const dataset = datasets.find((d) => d.id === (entry?.datasetId ?? newOn));
  const { element, report, error } = useEmbed(entry, "edit");
  const [inspiration, setInspiration] = useState<string[]>([]);
  const measures = useMemo(() => measuresOf(dataset?.model), [dataset?.model]);

  const agent = {
    instructions:
      "You build a Power BI report with the user, who edits it at the same time in the Power BI editor. Use the model fields exactly as listed. Lay visuals out on a 1280×720 page without overlap. Save the report when a requested change is done. For data questions, run DAX.",
    tools: [...authorTools(() => report, () => measures), daxTool(() => dataset && { id: dataset.id, groupId: dataset.groupId })],
    context: async () =>
      [
        libraryContext(entry, dataset),
        ...reports.filter((r) => inspiration.includes(r.id)).map((r) => `Report given as inspiration: "${r.name}"\n${r.context}\n${r.snapshot ?? ""}`),
        report ? `Report being edited, current state:\n${await pageText(report, false)}` : "No report open.",
      ].join("\n\n"),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={[
          { id: "chat", label: "AI", content: entry ? <Chat scope={`builder:${entry.id}`} agent={agent} placeholder="Describe the visuals you want…" /> : <Empty>Open or create a report.</Empty> },
          {
            id: "inspiration",
            label: `Inspiration${inspiration.length ? ` (${inspiration.length})` : ""}`,
            content: (
              <div className="flex flex-col gap-1 p-3">
                <p className="label mb-1">Reports of the library given to the AI as examples</p>
                {reports
                  .filter((r) => r.id !== reportId)
                  .map((r) => (
                    <label key={r.id} className="flex items-center gap-2 rounded-md px-1 py-1 hover:bg-muted">
                      <input type="checkbox" checked={inspiration.includes(r.id)} onChange={(e) => setInspiration(e.target.checked ? [...inspiration, r.id] : inspiration.filter((id) => id !== r.id))} />
                      {r.name}
                    </label>
                  ))}
              </div>
            ),
          },
        ]}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex items-center gap-2 px-2">
          <select className="input my-2 w-auto" value={reportId ?? ""} onChange={(e) => navigate({ search: { report: e.target.value, newOn: undefined } })}>
            <option value="" disabled>
              Open a report…
            </option>
            {reports.map((r) => (
              <option key={r.id} value={r.id}>
                {r.name}
                {r.editable === false ? " (view only)" : ""}
              </option>
            ))}
          </select>
          <span className="text-muted-foreground">or new report on</span>
          <select className="input my-2 w-auto" value={newOn ?? ""} onChange={(e) => navigate({ search: { report: undefined, newOn: e.target.value } })}>
            <option value="" disabled>
              a semantic model…
            </option>
            {datasets.map((d) => (
              <option key={d.id} value={d.id}>
                {d.name}
              </option>
            ))}
          </select>
          <span className="ml-auto" />
          <Zoom report={report} />
        </div>
        {error && <p className="px-3 text-destructive">{error}</p>}
        {newOn && !entry && dataset ? (
          <NewReport dataset={dataset} onSaved={(id) => navigate({ search: { report: id, newOn: undefined } })} />
        ) : (
          <>
            {!entry && <Empty>Open a report of the library, or start a new one on a semantic model.</Empty>}
            <div ref={element} className={entry ? "min-h-0 flex-1" : "hidden"} />
          </>
        )}
      </div>
    </div>
  );
}

/** New report on a semantic model: created and saved by code (blank report + saveAs), then opened in the editor. */
function NewReport({ dataset, onSaved }: { dataset: { id: string; groupId?: string; name: string }; onSaved: (id: string) => void }) {
  const element = useRef<HTMLDivElement>(null);
  const [name, setName] = useState(`${dataset.name} report`);
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  useEffect(() => {
    const el = element.current!;
    return () => void resetEmbed(el);
  }, []);
  const create = async () => {
    setState({ busy: true });
    try {
      const embed = await embedNewReport(element.current!, dataset);
      const saved = new Promise<string>((resolve) => embed.on("saved", (e: any) => resolve(e.detail.reportObjectId)));
      await (embed as any).saveAs({ name });
      const entry = await addReport(await saved);
      onSaved(entry.id);
    } catch (e) {
      setState({ error: errorText(e) });
    }
  };
  return (
    <div className="flex flex-col gap-2 p-3">
      <label className="label">Name of the new report on {dataset.name}</label>
      <div className="flex max-w-lg gap-2">
        <input className="input" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && name && void create()} />
        <button type="button" className="btn-primary" disabled={!name || state.busy} onClick={() => void create()}>
          {state.busy ? "Creating…" : "Create"}
        </button>
      </div>
      {state.error && <p className="text-[12px] text-destructive">{state.error}</p>}
      {/* The blank report is embedded off screen only to be saved. */}
      <div ref={element} style={{ position: "fixed", left: -10000, top: 0, width: 1280, height: 720 }} />
    </div>
  );
}
