import { useEffect, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { CopyIcon, PencilRulerIcon, SigmaIcon } from "lucide-react";
import { ReportPicker } from "@/components/pickers";
import { ReportActions } from "@/components/report-actions";
import { Button } from "@/components/ui/button";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Empty, Workspace } from "@/components/workspace";
import { Zoom } from "@/components/zoom";
import { builderAgent } from "@/agent/builder";
import type { Agent } from "@/agent/loop";
import { useLast } from "@/lib/last";
import { notEditable } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { definition, partText } from "@/lib/report-def";
import { useCollection } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";

export const Route = createFileRoute("/editor")({
  validateSearch: (s: Record<string, unknown>): { report?: string; tab?: string } => ({ report: (s.report as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Editor,
});

type ReportMeasure = { table: string; name: string; expression: string; formatString?: string; dataType?: string };

/** The report measures (stored in the report, reportExtensions.json), read again after each save. */
function useReportMeasures(ref: { id: string; groupId?: string } | undefined, version: number) {
  const [measures, setMeasures] = useState<ReportMeasure[]>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    if (!ref) return;
    let live = true;
    setError(undefined);
    definition(ref).then(
      (parts) => {
        const ext = parts.find((p) => p.path === "definition/reportExtensions.json");
        const json = ext ? JSON.parse(partText(ext)) : { entities: [] };
        const list = (json.entities ?? []).flatMap((e: { name: string; measures?: Omit<ReportMeasure, "table">[] }) => (e.measures ?? []).map((m) => ({ ...m, table: e.name })));
        if (live) setMeasures(list);
      },
      (e) => live && setError(errorText(e)),
    );
    return () => void (live = false);
  }, [ref?.id, version]); // eslint-disable-line react-hooks/exhaustive-deps
  return { measures, error };
}

/**
 * Builder: the report in Power BI's own editor (authoring embed) and the AI, which changes its
 * definition (PBIR) without limit — pages, visuals, formatting, theme, report measures — in a draft
 * saved when its answer is done; the editor is then reloaded. Tabs: the report, its measures.
 */
function Editor() {
  const { report: reportId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/editor" });
  useLast("editor:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const chosen = useCollection("reports").find((r) => r.id === reportId);
  const entry = chosen?.access === "edit" ? chosen : undefined;
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const { element, report, error } = useEmbed(entry, "edit");
  const [version, setVersion] = useState(0);
  const { measures, error: measuresError } = useReportMeasures(entry, version);
  // Read by the tools at call time (the agent object is made at each render).
  const live = useRef({ report, entry, dataset });
  live.current = { report, entry, dataset };
  const agent: Agent | undefined = entry && builderAgent({ entry: () => live.current.entry, dataset: () => live.current.dataset, report: () => live.current.report, onSaved: () => setVersion((v) => v + 1) });

  return (
    <Workspace
      title="Builder"
      icon={PencilRulerIcon}
      picker={<ReportPicker editable value={reportId} onChange={(report) => navigate({ search: { report } })} />}
      actions={
        <>
          <ReportActions entry={entry} report={report} />
          <Zoom report={report} />
        </>
      }
      ai={{ scope: `editor:${entry?.id}`, agent, placeholder: "Ask to add a page, build visuals, change the formatting…", empty: "Choose a report you can edit." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        {
          id: "report",
          label: "Report",
          icon: PencilRulerIcon,
          content: (
            <>
              {error && <p className="px-6 py-2 text-[13px] text-destructive">{error}</p>}
              {!entry && <Empty icon={PencilRulerIcon}>{chosen ? notEditable(chosen) : 'Choose a report you can edit, add one from its link, or create one from a semantic model (list at the top › "New report from a semantic model…").'}</Empty>}
              <div ref={element} className={entry ? "min-h-0 flex-1" : "hidden"} />
            </>
          ),
        },
        { id: "measures", label: "Measures", icon: SigmaIcon, count: measures?.length, content: <MeasuresTab measures={measures} error={measuresError} ready={!!entry} /> },
      ]}
    />
  );
}

/** The measures created on the report (not in the model): name, home table, DAX, format. */
function MeasuresTab({ measures, error, ready }: { measures?: ReportMeasure[]; error?: string; ready: boolean }) {
  if (!ready) return <Empty icon={SigmaIcon}>Choose a report.</Empty>;
  if (error) return <p className="p-6 text-[13px] text-destructive">{error}</p>;
  if (!measures) return <Empty>Reading the report…</Empty>;
  if (!measures.length) return <Empty icon={SigmaIcon}>No report measure yet. Ask the assistant to create one (e.g. "a measure for the margin rate").</Empty>;
  return (
    <div className="min-h-0 flex-1 overflow-auto p-6">
      <div className="overflow-hidden rounded-xl border bg-card shadow-xs">
        <Table>
          <TableHeader>
            <TableRow className="bg-muted/50">
              <TableHead className="w-56">Measure</TableHead>
              <TableHead className="w-40">Table</TableHead>
              <TableHead>DAX</TableHead>
              <TableHead className="w-32">Format</TableHead>
              <TableHead className="w-12" />
            </TableRow>
          </TableHeader>
          <TableBody>
            {measures.map((m) => (
              <TableRow key={`${m.table}/${m.name}`}>
                <TableCell className="font-medium">
                  <span className="flex items-center gap-2">
                    <SigmaIcon className="size-4 text-success" aria-hidden /> {m.name}
                  </span>
                </TableCell>
                <TableCell className="text-muted-foreground">{m.table}</TableCell>
                <TableCell>
                  <code className="line-clamp-3 font-mono text-[12px] whitespace-pre-wrap">{m.expression}</code>
                </TableCell>
                <TableCell className="text-[12.5px] text-muted-foreground">{m.formatString ?? "—"}</TableCell>
                <TableCell>
                  <Button variant="ghost" size="icon" className="size-7" aria-label={`Copy the DAX of ${m.name}`} onClick={() => void navigator.clipboard.writeText(`${m.name} = ${m.expression}`)}>
                    <CopyIcon />
                  </Button>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
