import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { Chat } from "@/components/chat";
import { BUILDER_TABS, visible } from "@/config";
import { Picker } from "@/components/picker";
import { ReportActions } from "@/components/report-actions";
import { Zoom } from "@/components/zoom";
import { Empty, SidePanel } from "@/components/side-panel";
import { authorTools, daxTool } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { reportModelTools, rewrite } from "@/agent/report-model-tools";
import { fixExtensionRefs, readExtension } from "@/lib/report-def";
import { measureNames } from "@/lib/model";
import { addDataset, addReport } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { embedNewReport, pageText, resetEmbed } from "@/lib/pbi";
import { useCollection } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { useLast } from "@/lib/last";
import { ReportPicker, libraryContext } from "./viewer";

export const Route = createFileRoute("/builder")({
  validateSearch: (s: Record<string, unknown>) => ({ report: (s.report as string) || undefined, newOn: (s.newOn as string) || undefined }),
  component: Builder,
});


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
  // Report measures (stored in the report): known to the AI and usable as visual fields.
  const [extension, setExtension] = useState<{ names: string[]; list: { table: string; name: string }[]; text: string }>({ names: [], list: [], text: "" });
  // The agent keeps the closures of the turn's start: it reads the latest report measures from this ref.
  const extensionRef = useRef(extension);
  const loadExtension = async () => {
    if (!entry) return extensionRef.current;
    const next = await readExtension(entry).then(
      (ext) => {
        const list = (ext.entities ?? []).flatMap((e) => (e.measures ?? []).map((m) => ({ table: e.name, ...m })));
        return { names: list.map((m) => `${m.table}[${m.name}]`), list, text: list.map((m) => `[${m.name}] (in '${m.table}') = ${m.expression}`).join("\n") };
      },
      () => ({ names: [], list: [], text: "" }),
    );
    extensionRef.current = next;
    setExtension(next);
    return next;
  };
  useEffect(() => void loadExtension(), [entry?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const measures = { has: (name: string) => measureNames(dataset?.info).has(name) || extensionRef.current.names.includes(name) } as Set<string>;

  // Report measures bound by the authoring API need their references fixed in the definition.
  const fixReportMeasures = async (page: string, visual: string, fields: string[]) => {
    if (!report || !entry) return;
    // Names that are not model columns / measures may be report measures created during this turn.
    const known = new Set(dataset?.info?.tables.flatMap((t) => [...t.columns.map((c) => c.name), ...t.measures.map((m) => m.name)]) ?? []);
    const names = fields.map((f) => f.match(/\[([^\]]+)\]\)?$/)?.[1]).filter((n): n is string => !!n && !known.has(n));
    if (!names.length) return;
    const list = names.every((n) => extensionRef.current.list.some((m) => m.name === n)) ? extensionRef.current.list : (await loadExtension()).list;
    const used = list.filter((m) => names.includes(m.name));
    if (!used.length) return;
    return rewrite(report, () => fixExtensionRefs(entry, page, visual, used));
  };

  const agent = {
    instructions:
      `You build a Power BI report with the user, who edits it at the same time in the Power BI editor. Use the model fields exactly as listed (describe_table for details). Lay visuals out on the page (its size is in the context) without overlap. Save the report when a requested change is done. For data questions, run DAX.
Power BI facts to rely on:
- New calculations: add_report_measure (works without model rights). "Parameters" (field / numeric) need model write access; otherwise use a report measure reading a slicer, e.g. [Selected item] = SELECTEDVALUE('T'[title], "None").
- Selecting one item: a slicer on the item column (role "Values"), the detail visuals then show only the selection; a card / multiRowCard shows one row when one item is selected. Clicking a row of a table cross-filters the other visuals.
- There is no per-row button in a table visual: say so and propose the slicer pattern.
- "A click on a row only affects the detail": set_interactions (source = table, detail → filter, others → none). Slicer single / multiple selection: set_slicer_selection.
- "The detail shows one item, the most recent by default": set_visual_filter kind "top" on the detail visual (column = item id or title, order_by = the date, n = 1), plus set_interactions (table → detail: filter). Never fake it with a slicer value that happens to match one item: it also filters the other visuals and breaks when the data changes.
- A table cannot have a row selected by default. A slicer shown "like a select": set_slicer_selection style "dropdown".
- Only claim what the tools did; never present a coincidence in the data as a solution.
- Refer to visuals by their name (name=…), titles can change. Visuals marked hidden without fields are leftovers of failed attempts: delete them.
- Per-series colors: set_theme. Page / report background: set_page_background (set_theme background only colors visuals and panes).
- Titles and text: create_visual type "textbox", then set_textbox (text, size, bold, color, alignment).
- Layout: every visual's position is in the context (x y w h z) with the layout problems (overlaps, outside the page). Plan the whole page: visuals inside the page size given in the context, no overlap, ~20 px gaps. "Align vertically" (one above the other) = same x AND same width; "align horizontally" = same y (and usually the same height). Move with arrange_visuals (all visuals in one call), read the positions it returns, fix any remaining problem, and only then tell the user — never say "already aligned" without comparing the numbers.
- A colored rectangle / panel behind visuals: add_shape (around = those visuals, behind). Never fake it by coloring the visuals' backgrounds.`,
    tools: [...authorTools(() => report, () => measures, fixReportMeasures), daxTool(() => dataset && { id: dataset.id, groupId: dataset.groupId }), ...modelTools(() => dataset), ...reportModelTools({ report: () => report, entry: () => entry, dataset: () => dataset, changed: loadExtension })],
    context: async () =>
      [
        libraryContext(entry, dataset),
        ...reports.filter((r) => inspiration.includes(r.id)).map((r) => `Report given as inspiration: "${r.name}"\n${r.context}\n${r.snapshot ?? ""}`),
        extensionRef.current.text ? `Report measures (stored in this report):\n${extensionRef.current.text}` : "",
        `Model write access: ${dataset?.info?.source === "tmdl" ? "yes (parameters possible)" : "no (use report measures; parameters impossible)"}`,
        report ? `Report being edited, current state:\n${await pageText(report, false)}` : "No report open.",
      ].join("\n\n"),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible([
          { id: "ai", label: "AI", content: entry ? <Chat scope={`builder:${entry.id}`} agent={agent} placeholder="Describe the visuals you want…" /> : <Empty>Open or create a report.</Empty> },
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
        ], BUILDER_TABS)}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex h-12 shrink-0 items-center gap-2 border-b bg-card px-2">
          <ReportPicker value={reportId} onChange={(report) => navigate({ search: { report, newOn: undefined } })} />
          <span className="text-muted-foreground">or new report on</span>
          <Picker
            className="my-2 w-64"
            items={datasets.map((d) => ({ id: d.id, label: d.name }))}
            value={newOn}
            onChange={(id) => navigate({ search: { report: undefined, newOn: id } })}
            placeholder="a semantic model…"
            add={{ label: "Add a semantic model", placeholder: "Semantic model link or id", run: async (link) => (await addDataset(link)).id }}
          />
          <span className="ml-auto" />
          <ReportActions entry={entry} report={report} />
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
