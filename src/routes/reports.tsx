import { useEffect, useRef, useState, type ReactNode } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { z } from "zod";
import { BookTextIcon, RefreshCwIcon, ChartPieIcon, ExternalLinkIcon, EyeOffIcon, FileChartColumnIcon, LayersIcon, PencilRulerIcon } from "lucide-react";
import { ContextLinks } from "@/components/context-links";
import { EntityMenu, withToast } from "@/components/entity";
import { ReportPicker } from "@/components/pickers";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { DataTable } from "@/components/data-table";
import { Empty, SubTabs, Workspace } from "@/components/workspace";
import { reportContext } from "@/agent/context";
import { tool, type Agent } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, daxTool } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { refreshReport } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { describeReport, schemaText, type ReportSchema } from "@/lib/report-schema";
import { all, patch, remove, useCollection, type ReportEntry } from "@/lib/store";

export const Route = createFileRoute("/reports")({
  validateSearch: (s: Record<string, unknown>): { report?: string; tab?: string } => ({ report: (s.report as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Reports,
});

const ACCESS: Record<string, { label: string; variant: "default" | "secondary" | "outline" }> = { edit: { label: "Edit", variant: "default" }, build: { label: "Build", variant: "secondary" }, view: { label: "View only", variant: "outline" } };

/**
 * The report's descriptive schema: the one kept with it, made once when missing (from the definition,
 * else through an embed for read-only reports); `regenerate` makes it again.
 */
function useSchema(entry: ReportEntry | undefined) {
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const regenerate = async () => {
    if (!entry) return;
    setState({ busy: true });
    await describeReport(entry).then(
      () => setState({}),
      (e) => setState({ error: errorText(e) }),
    );
  };
  useEffect(() => {
    if (entry && !entry.schema && !state.busy) void regenerate();
  }, [entry?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  return { schema: entry?.schema, error: state.error, busy: !!state.busy, regenerate };
}

/**
 * Reports: the Power BI reports the app knows (added here by link): their structure read from the
 * definition (report, pages, visuals) and the contexts linked to them. The AI documents and explains
 * them, and links contexts.
 */
function Reports() {
  const { report: reportId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/reports" });
  useLast("reports:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const entry = useCollection("reports").find((r) => r.id === reportId);
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const { schema, error, busy, regenerate } = useSchema(entry);
  const live = useRef({ schema, entry });
  live.current = { schema, entry };
  const datasetRef = entry && { id: entry.datasetId, groupId: dataset?.groupId ?? entry.datasetGroupId };

  const agent: Agent | undefined = entry && {
    instructions: `Screen: REPORTS. The user looks at the structure of a Power BI report (pages, visuals and their fields, filters, report measures) and the contexts linked to it. You explain and document the report (what each page shows, where a field is used, what a measure computes), check figures with DAX, and link or unlink contexts. You do not change the report itself (that is the Builder screen: say so when asked).
${DAX_RULES}`,
    tools: [
      tool({
        name: "set_contexts",
        description: "Sets which contexts are linked to this report (the whole list of context names).",
        args: z.object({ names: z.array(z.string()) }),
        run: async ({ names }) => {
          const contexts = await all("contexts");
          const unknown = names.filter((n) => !contexts.some((c) => c.name.toLowerCase() === n.toLowerCase()));
          if (unknown.length) throw new Error(`Unknown context(s): ${unknown.join(", ")}. Contexts: ${contexts.map((c) => c.name).join(", ") || "none"}.`);
          await patch("reports", entry.id, { contextIds: contexts.filter((c) => names.some((n) => n.toLowerCase() === c.name.toLowerCase())).map((c) => c.id) });
          return `Linked contexts: ${names.join(", ") || "none"}.`;
        },
      }),
      daxTool(() => datasetRef),
      ...modelTools(() => dataset),
    ],
    context: () => reportContext(entry, dataset),
    state: async () => {
      const s = live.current.schema;
      const contexts = (await all("contexts")).map((c) => `${c.name}${live.current.entry?.contextIds.includes(c.id) ? " (linked)" : ""}`);
      return [s ? schemaText(s) : error ? `Structure not readable: ${error}` : "Reading the structure…", `Contexts: ${contexts.join(", ") || "none"}.`].join("\n\n");
    },
  };

  const rename = () => {
    if (!entry) return;
    const name = prompt("Name in this app (empty: the name in Power BI)", entry.name);
    if (name === null) return;
    if (name.trim()) return void patch("reports", entry.id, { name: name.trim(), renamed: true });
    void patch("reports", entry.id, { renamed: false }).then(() => withToast("Reading…", refreshReport(entry), "Name read again."));
  };

  return (
    <Workspace
      title="Reports"
      icon={FileChartColumnIcon}
      picker={<ReportPicker value={reportId} onChange={(report) => navigate({ search: { report } })} />}
      actions={
        entry && (
          <>
            <Button variant="outline" size="sm" disabled={busy} title={schema?.readAt ? `Read ${new Date(schema.readAt).toLocaleString()} (${schema.source === "embed" ? "through an embed" : "from the definition"})` : undefined} onClick={() => void regenerate()}>
              <RefreshCwIcon className={busy ? "animate-spin" : ""} /> {busy ? "Reading…" : "Regenerate the schema"}
            </Button>
            <Button variant="outline" size="sm" asChild>
              <Link to="/viewer" search={{ report: entry.id }}>
                <ChartPieIcon /> View
              </Link>
            </Button>
            {entry.access === "edit" && (
              <Button variant="outline" size="sm" asChild>
                <Link to="/editor" search={{ report: entry.id }}>
                  <PencilRulerIcon /> Edit
                </Link>
              </Button>
            )}
            <EntityMenu
              kind="reports"
              entry={entry}
              onRename={rename}
              onRefresh={() => void withToast("Reading the report…", refreshReport(entry).then(regenerate), "Read again.")}
              onRemove={() => confirm(`Remove "${entry.name}" from the app? (the report stays in Power BI)`) && void remove("reports", entry.id).then(() => navigate({ search: {} }))}
            />
          </>
        )
      }
      ai={{ scope: `reports:${entry?.id}`, agent, placeholder: "Ask what a page shows, where a field is used…", empty: "Choose a report, or add one from its link." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        { id: "schema", label: "Schema", icon: LayersIcon, content: entry ? <SchemaTab entry={entry} datasetName={dataset?.name} schema={schema} error={error} /> : <Empty icon={FileChartColumnIcon}>Choose a report, or add one from its Power BI link (menu at the top).</Empty> },
        { id: "context", label: "Context", icon: BookTextIcon, count: entry?.contextIds.length, content: entry ? <ContextLinks what="report" ids={entry.contextIds} onChange={(contextIds) => void patch("reports", entry.id, { contextIds })} /> : <Empty>Choose a report.</Empty> },
      ]}
    />
  );
}

function SchemaTab({ entry, datasetName, schema, error }: { entry: ReportEntry; datasetName?: string; schema?: ReportSchema; error?: string }) {
  const access = ACCESS[entry.access ?? ""] ?? { label: "…", variant: "outline" as const };
  const pageName = (name: string) => schema?.pages.find((p) => p.name === name)?.displayName ?? name;
  const info = (label: string, value: ReactNode) => (
    <div className="flex flex-col gap-1">
      <dt className="text-[12px] font-medium text-muted-foreground">{label}</dt>
      <dd className="text-[13.5px] break-words">{value}</dd>
    </div>
  );
  return (
    <SubTabs
      tabs={[
        {
          id: "report",
          label: "Report",
          content: (
            <div className="flex max-w-4xl flex-col gap-4">
              <dl className="grid grid-cols-2 gap-5 rounded-xl border bg-card p-5 shadow-xs md:grid-cols-3">
                {info("Name", <span className="font-semibold">{entry.name}</span>)}
                {info("Access", <Badge variant={access.variant}>{access.label}</Badge>)}
                {info("Semantic model", datasetName ?? entry.datasetId)}
                {info("Workspace", entry.groupId ?? (entry.appId ? `App ${entry.appId}` : "My workspace"))}
                {info("Theme", schema ? (schema.theme ?? schema.baseTheme ?? "Default") : "…")}
                {info("Report filters", schema ? schema.filters.join(", ") || "None" : "…")}
                {info(
                  "Open",
                  <a className="inline-flex items-center gap-1 text-primary hover:underline" href={entry.url} target="_blank" rel="noreferrer">
                    Power BI <ExternalLinkIcon className="size-3.5" />
                  </a>,
                )}
                {info("Pages", schema?.pages.length ?? "…")}
                {info("Visuals", schema?.visuals.length ?? "…")}
              </dl>
              {entry.accessNote && <p className="rounded-lg border border-warning/30 bg-warning/10 px-4 py-3 text-[13px]">{entry.accessNote}</p>}
              {error && <p className="rounded-lg bg-destructive/10 px-4 py-3 text-[13px] text-destructive">{error}</p>}
            </div>
          ),
        },
        {
          id: "pages",
          label: "Pages",
          count: schema?.pages.length,
          content: (
            <DataTable
              head={["Page", "Name (id)", "Size", "Visuals", "Filters"]}
              rows={(schema?.pages ?? []).map((p) => [
                <span key="n" className="flex items-center gap-2 font-medium">
                  {p.displayName}
                  {p.hidden && <EyeOffIcon className="size-3.5 text-muted-foreground" aria-label="Hidden" />}
                </span>,
                <code key="i" className="font-mono text-[12px] text-muted-foreground">{p.name}</code>,
                `${p.width} × ${p.height}`,
                p.visuals,
                p.filters,
              ])}
              empty={schema ? "No page." : error ? "Not readable." : "Reading…"}
            />
          ),
        },
        {
          id: "visuals",
          label: "Visuals",
          count: schema?.visuals.length,
          content: (
            <DataTable
              head={["Visual", "Type", "Page", "Fields", "Position"]}
              rows={(schema?.visuals ?? []).map((v) => [
                <span key="t" className="flex items-center gap-2 font-medium">
                  {v.title ?? <span className="text-muted-foreground">{v.name}</span>}
                  {v.hidden && <EyeOffIcon className="size-3.5 text-muted-foreground" aria-label="Hidden" />}
                </span>,
                <Badge key="ty" variant="secondary">{v.type}</Badge>,
                pageName(v.page),
                <span key="f" className="flex flex-col gap-0.5 text-[12px]">
                  {Object.entries(v.fields).map(([role, fields]) => (
                    <span key={role}>
                      <span className="text-muted-foreground">{role}:</span> {fields.join(", ")}
                    </span>
                  ))}
                </span>,
                <span key="p" className="text-[12px] text-muted-foreground tabular-nums">{`${v.x}, ${v.y} · ${v.width}×${v.height}`}</span>,
              ])}
              empty={schema ? "No visual." : error ? "Not readable." : "Reading…"}
            />
          ),
        },
      ]}
    />
  );
}
