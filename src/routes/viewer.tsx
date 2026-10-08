import { useCallback, useEffect, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, ExternalLinkIcon, LinkIcon, PencilIcon, PlusIcon } from "lucide-react";
import { Markdown } from "@/components/markdown";
import { Chat } from "@/components/chat";
import { Zoom } from "@/components/zoom";
import { Empty, SidePanel } from "@/components/side-panel";
import { daxTool, viewTools } from "@/agent/pbi-tools";
import { addRows, readTable, updateRows, type Table } from "@/lib/excel";
import { errorText } from "@/lib/ms";
import { pageText } from "@/lib/pbi";
import { useCollection, type DatasetEntry, type ExcelEntry, type ReportEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { useLast } from "@/lib/last";

export const Route = createFileRoute("/viewer")({
  validateSearch: (s: Record<string, unknown>) => ({ report: (s.report as string) || undefined }),
  component: Viewer,
});

export function ReportPicker({ value, onChange }: { value?: string; onChange: (id: string) => void }) {
  const reports = useCollection("reports");
  return (
    <select className="input m-2 w-auto" value={value ?? ""} onChange={(e) => onChange(e.target.value)}>
      <option value="" disabled>
        Choose a report…
      </option>
      {reports.map((r) => (
        <option key={r.id} value={r.id}>
          {r.name}
        </option>
      ))}
    </select>
  );
}

/** Reads an Excel table of the library (null when none is chosen). */
export function useExcelTable(excel?: ExcelEntry) {
  const [table, setTable] = useState<Table | null>(null);
  const [error, setError] = useState<string>();
  const reload = useCallback(() => {
    if (!excel) return setTable(null);
    readTable(excel).then(setTable, (e) => setError(errorText(e)));
  }, [excel?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(reload, [reload]);
  return { table, error, reload };
}

/** AI context shared by the viewer and the builder: report, model and their notes. */
export const libraryContext = (report?: ReportEntry, dataset?: DatasetEntry) =>
  [
    report && `Report "${report.name}"${report.context ? `\nNotes: ${report.context}` : ""}`,
    dataset && `Semantic model "${dataset.name}"${dataset.context ? `\nNotes: ${dataset.context}` : ""}\n${dataset.model ?? ""}`,
  ]
    .filter(Boolean)
    .join("\n\n");

function Viewer() {
  const { report: reportId } = Route.useSearch();
  const navigate = useNavigate({ from: "/viewer" });
  useLast("viewer:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const entry = useCollection("reports").find((r) => r.id === reportId);
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const excels = useCollection("excels");
  const { element, report, error, selection } = useEmbed(entry, "view");

  const agent = {
    instructions:
      "You help the user read a Power BI report. You see the page they are looking at (visuals with their data). You can change page, filters and slicers, and run DAX on the model to answer precisely. Explain numbers simply.",
    tools: [...viewTools(() => report), daxTool(() => entry && { id: entry.datasetId, groupId: entry.datasetGroupId })],
    context: async () => [libraryContext(entry, dataset), report ? `What the user sees now:\n${await pageText(report)}` : "The report is not loaded."].join("\n\n"),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={[
          { id: "info", label: "Info", content: <InfoTab excel={excels.find((e) => e.id === entry?.infoExcelId)} /> },
          { id: "guide", label: "Guides", content: <GuideTab excel={excels.find((e) => e.id === entry?.guidesExcelId)} selection={selection} /> },
          { id: "chat", label: "AI", content: entry ? <Chat scope={`viewer:${entry.id}`} agent={agent} placeholder="Ask about this page…" /> : <Empty>Choose a report.</Empty> },
        ]}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex items-center gap-2 pr-2">
          <ReportPicker value={reportId} onChange={(report) => navigate({ search: { report } })} />
          <span className="ml-auto" />
          <Zoom report={report} />
        </div>
        {error && <p className="px-3 text-destructive">{error}</p>}
        {!entry && <Empty>Choose a report (add reports in the Library).</Empty>}
        <div ref={element} className={entry ? "min-h-0 flex-1" : "hidden"} />
      </div>
    </div>
  );
}

const IconOf = ({ name }: { name?: unknown }) => {
  const key = String(name ?? "").replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons;
  const Icon = icons[key] ?? LinkIcon;
  return <Icon className="size-4 shrink-0" />;
};

/** Links listed in an Excel table: label, url, icon (lucide name), color, description. */
function InfoTab({ excel }: { excel?: ExcelEntry }) {
  const { table, error } = useExcelTable(excel);
  if (!excel) return <Empty>No info links: choose an Excel file for this report in the Library (columns label, url, icon, color, description).</Empty>;
  if (error) return <Empty>{error}</Empty>;
  return (
    <div className="flex flex-col gap-1.5 p-3">
      {table?.rows.map((row, i) => (
        <a key={i} href={String(row.url ?? "#")} target="_blank" rel="noreferrer" className="flex items-center gap-3 rounded-lg border p-2.5 hover:bg-muted" style={{ borderLeft: `4px solid ${row.color ?? "#a1a1aa"}` }}>
          <span style={{ color: String(row.color ?? "") }}>
            <IconOf name={row.icon} />
          </span>
          <span className="min-w-0 flex-1">
            <span className="block truncate font-medium">{String(row.label ?? row.title ?? row.url)}</span>
            {row.description != null && <span className="block truncate text-[11px] text-muted-foreground">{String(row.description)}</span>}
          </span>
          <ExternalLinkIcon className="size-3.5 text-muted-foreground" />
        </a>
      ))}
      <a href={excel.link} target="_blank" rel="noreferrer" className="mt-2 text-[11px] text-muted-foreground underline">
        Edit these links in Excel
      </a>
    </div>
  );
}

/** Guides in an Excel table: title, content (markdown), links ("pageName" or "pageName/visualName", separated by ;). */
function GuideTab({ excel, selection }: { excel?: ExcelEntry; selection: { page?: string; visual?: string } }) {
  const { table, error, reload } = useExcelTable(excel);
  const [editing, setEditing] = useState<{ index?: number; title: string; content: string; links: string }>();
  if (!excel) return <Empty>No guides: choose an Excel file for this report in the Library (columns title, content, links).</Empty>;
  if (error) return <Empty>{error}</Empty>;
  if (!table) return <Empty>Loading…</Empty>;

  const linksOf = (row: Record<string, unknown>) => String(row.links ?? "").split(/\s*;\s*/).filter(Boolean);
  const here = selection.visual ? `${selection.page}/${selection.visual}` : undefined;
  const guides = table.rows
    .map((row, index) => ({ row, index, links: linksOf(row) }))
    .filter((g) => g.links.some((l) => l === here || l === selection.page || l.startsWith(`${selection.page}/`)))
    .sort((a, b) => Number(b.links.includes(here ?? "")) - Number(a.links.includes(here ?? "")));
  const save = async () => {
    if (!editing) return;
    const values = { title: editing.title, content: editing.content, links: editing.links };
    if (editing.index === undefined) await addRows(excel, table, [values]);
    else await updateRows(excel, table, [{ index: editing.index, values }]);
    setEditing(undefined);
    reload();
  };

  if (editing)
    return (
      <div className="flex flex-col gap-2 p-3">
        <input className="input font-medium" value={editing.title} onChange={(e) => setEditing({ ...editing, title: e.target.value })} />
        <textarea className="input field-sizing-content min-h-48 py-1 font-mono text-[12px]" placeholder="Markdown: what it shows, how to read it, caveats…" value={editing.content} onChange={(e) => setEditing({ ...editing, content: e.target.value })} />
        <label className="label">Linked to (pageName/visualName; …)</label>
        <input className="input" value={editing.links} onChange={(e) => setEditing({ ...editing, links: e.target.value })} />
        <div className="flex gap-2">
          <button type="button" className="btn-primary" onClick={() => void save()}>
            Save
          </button>
          <button type="button" className="btn" onClick={() => setEditing(undefined)}>
            Cancel
          </button>
        </div>
      </div>
    );

  return (
    <div className="flex flex-col gap-3 p-3">
      {guides.length === 0 && <Empty>No guide for this {here ? "visual" : "page"}.</Empty>}
      {guides.map(({ row, index, links }) => (
        <article key={index} className="flex flex-col gap-1.5 rounded-lg border p-3">
          <div className="flex items-center gap-2">
            <h3 className="flex-1 font-medium">{String(row.title ?? "")}</h3>
            <button type="button" className="icon-btn" title="Edit" onClick={() => setEditing({ index, title: String(row.title ?? ""), content: String(row.content ?? ""), links: links.join("; ") })}>
              <PencilIcon />
            </button>
          </div>
          <Markdown>{String(row.content ?? "")}</Markdown>
          {here && !links.includes(here) && (
            <button type="button" className="btn h-7 w-fit" onClick={() => void updateRows(excel, table, [{ index, values: { links: [...links, here].join("; ") } }]).then(reload)}>
              <LinkIcon /> Link the selected visual
            </button>
          )}
        </article>
      ))}
      <button type="button" className="btn w-fit" onClick={() => setEditing({ title: "New guide", content: "", links: here ?? selection.page ?? "" })}>
        <PlusIcon /> New guide for this {here ? "visual" : "page"}
      </button>
    </div>
  );
}
